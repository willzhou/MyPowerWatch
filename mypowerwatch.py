import os
import sys
import time
import psutil
import GPUtil
import asyncio
import subprocess
import wmi
from typing import Dict, Tuple
from textual.app import App, ComposeResult
from textual.containers import Container, Vertical
from textual.reactive import reactive
from textual.widgets import Header, Footer, Static, ProgressBar, DataTable
from textual import events, log
from textual.widgets import Button
from textual.screen import Screen

def is_frozen() -> bool:
    """检查是否在打包环境中"""
    return getattr(sys, 'frozen', False)

class MyPowerWatch(App):
    """使用Textual框架的Windows终端UI功耗监控程序"""
    
    CSS = """
    Screen {
        layout: vertical;
    }

    #stats-container {
        height: 1fr;
        padding: 1;
        border: solid $accent;
    }

    .power-value {
        text-style: bold;
        color: $success;
    }

    .component-name {
        width: 12;
    }

    ProgressBar {
        width: 16;
    }

    DataTable {
        height: 10;
    }

    #power-chart {
        height: 10;
        border: solid $accent;
        margin: 1 0;
    }

    .developer-info {
        padding: 1;
        border: solid $accent;
        margin: 1;
    }

    #close {
        width: 10;
        margin: 1 2;
        background: $primary;
        color: $text;
    }

    """
    
    BINDINGS = [
        ("q", "quit", "退出程序"),
        ("ctrl+c", "quit", "退出程序"),
        ("h", "show_developer_info", "开发者信息"),  # 新增h键绑定
        # ("d", "toggle_dark", "切换暗黑模式")
    ]
    
    # 反应式数据
    total_power = reactive(0.0)
    components = reactive(dict)
    run_time = reactive(0.0)
    energy_consumption = reactive(0.0)
    power_history = reactive(list)
    
    def __init__(self):
        super().__init__()
        if is_frozen():
            sys.stdout = sys.__stdout__
            sys.stderr = sys.__stderr__
        
        self.wmi_conn = wmi.WMI()
        self.ohm_conn = self._connect_to_ohm()
        self.hardware_specs = self._detect_hardware()
        self.start_time = time.time()
        self.last_update = time.time()
        self.max_history_points = 60

        # Windows 特定初始化
        if sys.platform == "win32":
            self._init_windows_specific()
    
    def _init_windows_specific(self):
        """Windows 特定的初始化"""
        try:
            # 检查是否安装了 NVIDIA SMI
            self.has_nvidia_smi = bool(subprocess.run(['nvidia-smi', '-L'], 
                                      capture_output=True, shell=True).returncode == 0)
        except:
            self.has_nvidia_smi = False
    
    def _connect_to_ohm(self):
        """连接到 Open Hardware Monitor"""
        try:
            return wmi.WMI(namespace="root\OpenHardwareMonitor")
        except:
            return None
    
    def _detect_hardware(self) -> Dict:
        """检测硬件信息"""
        specs = {
            'cpu': self._detect_cpu_info(),
            'gpu': self._detect_gpu_info(),
            'ram': psutil.virtual_memory().total / (1024**3),
            'disks': self._detect_disks_info(),
            'temperatures': self._get_temperatures(),
            'fan_speeds': self._get_fan_speeds(),
            'display': 20,
            'motherboard': 30,
            'fans': 10,
            'peripherals': 15,
            'battery': self._get_battery_info()
        }
        return specs
    
    def _detect_cpu_info(self) -> Dict:
        """检测CPU信息"""
        try:
            cpu = self.wmi_conn.Win32_Processor()[0]
            name = cpu.Name
            
            # 更精确的 TDP 估算
            if 'Intel' in name:
                if 'i9' in name: base, turbo = 65, 250
                elif 'i7' in name: base, turbo = 65, 200
                elif 'i5' in name: base, turbo = 65, 150
                elif 'i3' in name: base, turbo = 35, 110
                else: base, turbo = 35, 95
            elif 'AMD' in name:
                if 'Ryzen 9' in name: base, turbo = 105, 230
                elif 'Ryzen 7' in name: base, turbo = 65, 180
                elif 'Ryzen 5' in name: base, turbo = 65, 150
                else: base, turbo = 35, 95
            else:
                base, turbo = 35, 95
            
            return {
                'name': name,
                'cores': cpu.NumberOfCores,
                'threads': cpu.NumberOfLogicalProcessors,
                'base_tdp': base,
                'max_tdp': turbo,
                'current_clock': cpu.CurrentClockSpeed,
                'max_clock': cpu.MaxClockSpeed
            }
        except:
            return {
                'name': 'Unknown CPU',
                'cores': psutil.cpu_count(logical=False),
                'threads': psutil.cpu_count(logical=True),
                'base_tdp': 65,
                'max_tdp': 150
            }
    
    def _detect_gpu_info(self) -> Dict:
        """检测GPU信息 - 改进版"""
        try:
            # 优先使用 NVIDIA SMI
            if self.has_nvidia_smi:
                result = subprocess.run(
                    ['nvidia-smi', '--query-gpu=name,utilization.gpu,memory.used,memory.total', 
                    '--format=csv,noheader,nounits'],
                    capture_output=True, text=True, shell=True
                )
                if result.returncode == 0:
                    gpu_data = result.stdout.strip().split(',')
                    load = float(gpu_data[1])/100
                    return {
                        'name': gpu_data[0].strip(),
                        'load': load if load > 0 else self._get_gpu_load_win(),  # 双重检查
                        'memory_used': float(gpu_data[2]),
                        'memory_total': float(gpu_data[3]),
                        'tdp': self._estimate_gpu_tdp(gpu_data[0])
                    }
            
            # 回退到 WMI 检测
            gpus = self.wmi_conn.Win32_VideoController()
            if gpus:
                load = self._get_gpu_load_win()  # 使用改进的负载检测
                return {
                    'name': gpus[0].Name,
                    'load': load,
                    'memory_used': int(gpus[0].AdapterRAM)/(1024**3) if gpus[0].AdapterRAM else 0,
                    'memory_total': 0,
                    'tdp': self._estimate_gpu_tdp(gpus[0].Name)
                }
        except Exception as e:
            log.error(f"检测GPU信息失败: {e}")
        
        return {
            'name': 'Integrated GPU',
            'load': 0,
            'memory_used': 0,
            'memory_total': 0,
            'tdp': 15
        }
    
    def _get_gpu_load_win(self) -> float:
        """Windows 下获取 GPU 负载 - 改进版"""
        try:
            # 方法1: 使用 NVIDIA SMI (更可靠)
            if self.has_nvidia_smi:
                result = subprocess.run(
                    ['nvidia-smi', '--query-gpu=utilization.gpu', '--format=csv,noheader,nounits'],
                    capture_output=True, text=True, shell=True
                )
                if result.returncode == 0:
                    return float(result.stdout.strip()) / 100
            
            # 方法2: 使用 WMI 性能计数器 (适用于大多数GPU)
            perf_data = self.wmi_conn.Win32_PerfFormattedData_GPUPerformanceCounters_GPUEngine()
            if perf_data:
                total_utilization = sum(float(p.UtilizationPercentage) for p in perf_data)
                return total_utilization / len(perf_data) / 100
            
            # 方法3: 使用 Open Hardware Monitor
            if self.ohm_conn:
                for sensor in self.ohm_conn.Sensor():
                    if sensor.SensorType == 'Load' and 'GPU' in sensor.Name:
                        return float(sensor.Value) / 100
        except Exception as e:
            log.error(f"获取GPU负载失败: {e}")
        
        return 0.0  # 默认返回0

    def _get_gpu_load_with_retry(self, retries=3) -> float:
        """带重试的GPU负载获取"""
        for i in range(retries):
            try:
                load = self._get_gpu_load_win()
                if load > 0:  # 只有大于0才认为是有效值
                    return load
                time.sleep(0.5)  # 短暂延迟后重试
            except:
                time.sleep(0.5)
        return 0.0
    
    def _estimate_gpu_tdp(self, gpu_name: str) -> int:
        """估算GPU TDP"""
        gpu_name = gpu_name.lower()
        if 'rtx 4090' in gpu_name: return 450
        if 'rtx 4080' in gpu_name: return 320
        if 'rtx 3090' in gpu_name: return 350
        if 'rtx 3080' in gpu_name: return 320
        if 'rtx 3070' in gpu_name: return 220
        if 'rtx 3060' in gpu_name: return 170
        if 'rx 7900' in gpu_name: return 300
        if 'rx 6800' in gpu_name: return 250
        return 150
    
    def _detect_disks_info(self) -> list:
        """检测磁盘信息"""
        disks = []
        for part in psutil.disk_partitions():
            try:
                usage = psutil.disk_usage(part.mountpoint)
                disk = {
                    'device': part.device,
                    'type': 'SSD' if 'ssd' in part.opts.lower() else 'HDD',
                    'size_gb': usage.total / (1024**3),
                    'used_gb': usage.used / (1024**3)
                }
                disks.append(disk)
            except:
                continue
        return disks
    
    def _get_temperatures(self) -> Dict[str, float]:
        """获取温度信息"""
        temps = {}
        try:
            # 使用 Open Hardware Monitor
            if self.ohm_conn:
                for sensor in self.ohm_conn.Sensor():
                    if sensor.SensorType == 'Temperature':
                        temps[sensor.Name] = float(sensor.Value)
            
            # 使用 WMI 基础温度
            if not temps and hasattr(self.wmi_conn, 'Win32_TemperatureProbe'):
                for probe in self.wmi_conn.Win32_TemperatureProbe():
                    temps[probe.Name] = float(probe.CurrentReading)
        except:
            pass
        return temps
    
    def _get_fan_speeds(self) -> Dict[str, float]:
        """获取风扇转速"""
        fans = {}
        try:
            # 使用 Open Hardware Monitor
            if self.ohm_conn:
                for sensor in self.ohm_conn.Sensor():
                    if sensor.SensorType == 'Fan':
                        fans[sensor.Name] = float(sensor.Value)
        except:
            pass
        return fans
    
    def _get_battery_info(self) -> Dict:
        """获取电池信息"""
        battery = psutil.sensors_battery()
        if battery:
            return {
                'percent': battery.percent,
                'power_plugged': battery.power_plugged,
                'power_consumption': self._estimate_battery_power(battery)
            }
        return None
    
    def _estimate_battery_power(self, battery) -> float:
        """估算电池供电时的功耗"""
        if not battery.power_plugged:
            return (battery.percent / 100) * 50  # 假设最大功耗约50W
        return 0
    
    def _get_real_power_data(self) -> Dict[str, float]:
        """从硬件传感器获取实际功耗数据"""
        power_data = {}
        try:
            # 使用 Open Hardware Monitor
            if self.ohm_conn:
                for sensor in self.ohm_conn.Sensor():
                    if sensor.SensorType == 'Power':
                        if 'CPU' in sensor.Name:
                            power_data['cpu'] = float(sensor.Value)
                        elif 'GPU' in sensor.Name:
                            power_data['gpu'] = float(sensor.Value)
            
            # 使用 NVIDIA SMI 获取 GPU 功耗
            if 'gpu' not in power_data and self.has_nvidia_smi:
                result = subprocess.run(
                    ['nvidia-smi', '--query-gpu=power.draw', '--format=csv,noheader,nounits'],
                    capture_output=True, text=True
                )
                if result.returncode == 0:
                    power_data['gpu'] = float(result.stdout.strip())
        except:
            pass
        return power_data
    
    def _calculate_cpu_power(self) -> Tuple[float, float]:
        """计算CPU功耗"""
        # 尝试获取实际功耗
        real_power = self._get_real_power_data()
        if 'cpu' in real_power:
            return real_power['cpu'], psutil.cpu_percent()
        
        # 估算功耗
        cpu_usage = psutil.cpu_percent(interval=1) / 100
        cpu_info = self.hardware_specs['cpu']
        power = cpu_info['base_tdp'] + (cpu_info['max_tdp'] - cpu_info['base_tdp']) * cpu_usage
        power *= min(1.0, cpu_info['cores'] / 8)
        return power, cpu_usage * 100
    
    def _calculate_gpu_power(self) -> Tuple[float, float]:
        """计算GPU功耗"""
        # 尝试获取实际功耗
        real_power = self._get_real_power_data()
        if 'gpu' in real_power:
            return real_power['gpu'], self.hardware_specs['gpu']['load'] * 100
        
        # 估算功耗
        gpu_info = self.hardware_specs['gpu']
        if gpu_info['name'] != 'Integrated GPU':
            return gpu_info['tdp'] * gpu_info['load'], gpu_info['load'] * 100
        return 10 + 5 * gpu_info['load'], gpu_info['load'] * 100
    
    def _calculate_disk_power(self) -> Tuple[float, float]:
        """计算磁盘功耗"""
        power = 0
        usage = 0
        for disk in self.hardware_specs['disks']:
            if disk['type'] == 'SSD':
                power += 2 + 1 * (disk['used_gb'] / disk['size_gb'])
            else:
                power += 5 + 3 * (disk['used_gb'] / disk['size_gb'])
            usage += disk['used_gb'] / disk['size_gb']
        avg_usage = usage / len(self.hardware_specs['disks']) if self.hardware_specs['disks'] else 0
        return power, avg_usage * 100
    
    def _calculate_ram_power(self) -> Tuple[float, float]:
        """计算内存功耗"""
        ram_usage = psutil.virtual_memory().percent / 100
        return 3 + 0.5 * self.hardware_specs['ram'] * ram_usage, ram_usage * 100
    
    def update_power_consumption(self) -> Tuple[float, Dict]:
        """更新功耗计算"""
        current_time = time.time()
        time_elapsed = current_time - self.last_update
        self.last_update = current_time
        
        # 计算各组件功耗
        cpu_power, cpu_usage = self._calculate_cpu_power()
        gpu_power, gpu_usage = self._calculate_gpu_power()
        disk_power, disk_usage = self._calculate_disk_power()
        ram_power, ram_usage = self._calculate_ram_power()
        
        components = {
            'CPU': {'power': cpu_power, 'usage': cpu_usage},
            'GPU': {'power': gpu_power, 'usage': gpu_usage},
            'RAM': {'power': ram_power, 'usage': ram_usage},
            'Disks': {'power': disk_power, 'usage': disk_usage},
            'Motherboard': {'power': self.hardware_specs['motherboard'], 'usage': 0},
            'Fans': {'power': self.hardware_specs['fans'], 'usage': 0},
            'Display': {'power': self.hardware_specs['display'], 'usage': 0},
            'Peripherals': {'power': self.hardware_specs['peripherals'], 'usage': 0}
        }
        
        # 如果是笔记本电脑且使用电池，调整总功耗
        if self.hardware_specs['battery'] and not self.hardware_specs['battery']['power_plugged']:
            total_power = self.hardware_specs['battery']['power_consumption']
        else:
            total_power = sum(comp['power'] for comp in components.values())
        
        # 更新能耗和运行时间
        self.energy_consumption += total_power * (time_elapsed / 3600)
        self.run_time = current_time - self.start_time
        
        # 更新历史数据
        self.power_history.append(total_power)
        if len(self.power_history) > self.max_history_points:
            self.power_history.pop(0)
        
        # 检查功耗阈值
        self._check_power_threshold(total_power)
        
        return total_power, components
    
    def _check_power_threshold(self, current_power: float):
        """检查功耗是否超过阈值"""
        if current_power > 200:  # 200W 阈值
            self._show_windows_notification(
                "高功耗警告", 
                f"当前系统功耗已达到 {current_power:.1f}W"
            )
    
    def _show_windows_notification(self, title: str, message: str):
        """显示 Windows 通知"""
        try:
            from win10toast import ToastNotifier
            toaster = ToastNotifier()
            toaster.show_toast(title, message, duration=5)
        except:
            pass
    
    def compose(self) -> ComposeResult:
        """创建UI布局"""
        yield Header()
        yield Container(
            Vertical(
                Static("MyPowerWatch 电脑功耗实时监控", id="title"),
                Static(id="summary"),
                Static("功耗趋势:", classes="section-title"),
                Static(id="power-chart"),
                Static("组件功耗详情:", classes="section-title"),
                DataTable(id="components-table"),
            ),
            id="stats-container"
        )
        yield Footer()
    
    async def on_mount(self) -> None:
        """初始化UI元素"""
        table = self.query_one("#components-table", DataTable)
        table.add_columns("组件", "功耗(W)", "占比(%)")
        
        # 设置定时更新
        self.set_interval(1, self.update_display)
        
        # 强制刷新界面
        self.refresh(layout=True)
    
    def update_display(self) -> None:
        """更新UI显示 - 确保GPU数据同步"""
        self.total_power, self.components = self.update_power_consumption()
        
        # 强制刷新GPU数据
        if self.components['GPU']['usage'] == 0:
            self.hardware_specs['gpu']['load'] = self._get_gpu_load_win()
            _, gpu_usage = self._calculate_gpu_power()
            self.components['GPU']['usage'] = gpu_usage
        
        # 更新运行时间和总功耗显示
        hours, remainder = divmod(self.run_time, 3600)
        minutes, seconds = divmod(remainder, 60)
        time_str = f"{int(hours):02d}:{int(minutes):02d}:{int(seconds):02d}"
        
        summary = self.query_one("#summary", Static)
        summary.update(
            f"运行时间: {time_str} | "
            f"总功耗: [bold green]{self.total_power:.1f}W[/] | "
            f"能耗: [bold green]{self.energy_consumption/1000:.3f}kWh[/]"
        )
        
        # 更新组件表格
        table = self.query_one("#components-table", DataTable)
        table.clear()
        for name, data in self.components.items():
            power = data['power']
            percentage = (power / self.total_power) * 100 if self.total_power > 0 else 0
            table.add_row(name, f"{power:.1f}", f"{percentage:.1f}%")
        
        # 更新图表
        self.update_charts()
    
    def update_charts(self):
        """更新图表显示"""
        # 确保有足够的数据
        if not hasattr(self, 'power_history') or not self.power_history:
            return
        
        try:
            # 只保留功耗趋势图
            max_power = max(self.power_history)
            chart_lines = []
            for i in range(10, 0, -1):
                threshold = max_power * (i / 10)
                line = " ".join("█" if p >= threshold else " " for p in self.power_history)
                chart_lines.append(line)
            chart = "\n".join(chart_lines)
            self.query_one("#power-chart", Static).update(chart)
        except Exception as e:
            log.error(f"更新图表时出错: {e}")
    
    async def action_quit(self):
        """自定义退出动作"""
        self.exit()
    
    async def action_toggle_dark(self):
        """切换暗黑模式"""
        self.dark = not self.dark
    
    async def on_key(self, event: events.Key) -> None:
        """处理键盘事件"""
        if event.key == "q":
            await self.action_quit()
        elif event.key == "h":  # 新增h键处理
            await self.show_developer_info()

    async def show_developer_info(self):
        """显示开发者信息"""   
        # 创建一个简单的弹窗显示开发者信息
        class DeveloperScreen(Screen):
            def compose(self) -> ComposeResult:
                developer_info = """开发者：Will Zhou\n公司：汇视创影科技\n首发时间：2025-08-16"""
                yield Vertical(
                    Static(developer_info, classes="developer-info"),
                    Button("关闭", variant="primary", id="close"),
                )
            
            async def on_button_pressed(self, event: Button.Pressed) -> None:
                if event.button.id == "close":
                    self.dismiss()
        
        ds_id = "developer-info-screen"
        await self.push_screen(DeveloperScreen(id=ds_id))
            
async def run_app():
    """修复后的运行函数"""
    try:
        app = MyPowerWatch()
        
        # 在Windows打包环境中特殊处理
        if is_frozen() and sys.platform == "win32":
            # 确保标准流存在
            if sys.stdin is None:
                sys.stdin = open(os.devnull)
            if sys.stdout is None:
                sys.stdout = open(os.devnull, 'w')
            if sys.stderr is None:
                sys.stderr = open(os.devnull, 'w')
            
            # 设置Windows特定的事件循环策略
            asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
        
        await app.run_async()
    except Exception as e:
        log.error(f"应用程序错误: {e}")
        sys.exit(1)

if __name__ == "__main__":
    asyncio.run(run_app())