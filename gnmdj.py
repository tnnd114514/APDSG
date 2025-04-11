import os
import pythoncom
import psutil
import subprocess
import time
import win32api
import win32security
import wmi
from pycaw.pycaw import (
    AudioUtilities,
    IAudioEndpointVolume,
    CLSID_MMDeviceEnumerator
)
from pycaw.constants import DEVICE_STATE_ACTIVE  # 关键修改点

# 配置参数
THRESHOLD_HIGH = 60    # 启动音量阈值
THRESHOLD_LOW = 40     # 关闭音量阈值
VOLUME_BUFFER = 1.5    # 缓冲区间防止抖动
TARGET_EXE = r"D:\Program Files\AtHomeVideoStreamer\AtHomeVideoStreamer.exe"
EXE_NAME = os.path.basename(TARGET_EXE)
CHECK_INTERVAL = 0.2   # 检测间隔(秒)
HEADPHONE_NAMES = ["AT2", "Headphones", "Headset"]  # 多语言设备名称

class AudioCommander:
    def __init__(self):
        self.volume_control = None
        self.last_volume = None
        self._escalate_privileges()
        self._init_audio_device()
        self._validate_permissions()

    def _escalate_privileges(self):
        """提升至系统内核权限"""
        try:
            token = win32security.OpenProcessToken(
                win32api.GetCurrentProcess(),
                win32security.TOKEN_ALL_ACCESS
            )
            new_priv = (
                (win32security.LookupPrivilegeValue(None, "SeDebugPrivilege"), 
                 win32security.SE_PRIVILEGE_ENABLED),
                (win32security.LookupPrivilegeValue(None, "SeTcbPrivilege"),
                 win32security.SE_PRIVILEGE_ENABLED)
            )
            win32security.AdjustTokenPrivileges(token, False, new_priv)
        except Exception as e:
            print(f"⛔ 权限提升失败: {str(e)}")
            raise

    def _init_audio_device(self):
        """智能初始化耳机设备"""
        pythoncom.CoInitialize()
        enumerator = AudioUtilities.GetDeviceEnumerator()
        
        # 优先获取默认通信设备
        default_dev = enumerator.GetDefaultAudioEndpoint(0, 1)
        if self._is_headphone(default_dev):
            self._bind_volume_control(default_dev)
            return
        
        # 遍历所有活动设备
        devices = AudioUtilities.GetAllDevices()
        for dev in devices:
            if dev.state == DEVICE_STATE_ACTIVE and self._is_headphone(dev):
                self._bind_volume_control(dev)
                return
                
        raise RuntimeError("未检测到可用耳机设备")

    def _is_headphone(self, device):
        """多语言设备检测"""
        return any(name in device.FriendlyName for name in HEADPHONE_NAMES)

    def _bind_volume_control(self, device):
        """绑定音量控制接口"""
        interface = device.Activate(
            IAudioEndpointVolume._iid_,
            pythoncom.CLSCTX_ALL,
            None
        )
        self.volume_control = interface.QueryInterface(IAudioEndpointVolume)
        print(f"🎧 已连接音频设备: {device.FriendlyName}")

    def _validate_permissions(self):
        """验证进程控制权限"""
        try:
            psutil.Process().kill()
        except psutil.AccessDenied:
            print("⚠️ 需要以管理员身份运行!")
            os._exit(1)

    def get_volume(self):
        """获取精确耳机音量"""
        return round(self.volume_control.GetMasterVolumeLevelScalar() * 100, 1)

    def _should_trigger(self, current_vol):
        """触发条件判断（补充完整方法）"""
        if self.last_volume is None:
            self.last_volume = current_vol
            return False

        # 音量升高超过阈值
        if current_vol > THRESHOLD_HIGH + VOLUME_BUFFER:
            return "start"
        # 音量降低超过阈值
        elif current_vol < THRESHOLD_LOW - VOLUME_BUFFER:
            return "stop"
        else:
            return False

    def monitor_loop(self):
        """主监控循环"""
        last_action = None
        while True:
            try:
                current_vol = self.get_volume()
                action = self._should_trigger(current_vol)

                if action == "start" and last_action != "start":
                    if not self._is_process_running():
                        self._start_process()
                        last_action = "start"
                        print(f"🚀 已启动 {EXE_NAME}")

                elif action == "stop" and last_action != "stop":
                    if self._is_process_running():
                        self._kill_process()
                        last_action = "stop"
                        print(f"🛑 已终止 {EXE_NAME}")

                time.sleep(CHECK_INTERVAL)

            except Exception as e:
                print(f"🔴 监控异常: {str(e)}")
                time.sleep(1)

    def _is_process_running(self):
        """检测目标进程状态"""
        for proc in psutil.process_iter(['pid', 'name', 'exe']):
            if proc.info['exe'] == TARGET_EXE:
                return True
        return False

    def _start_process(self):
        """启动目标程序"""
        subprocess.Popen(
            TARGET_EXE,
            creationflags=subprocess.CREATE_NO_WINDOW,
            shell=True
        )

    def _kill_process(self):
        """终止目标进程"""
        for proc in psutil.process_iter(['pid', 'name', 'exe']):
            if proc.info['exe'] == TARGET_EXE:
                proc.terminate()
                proc.wait()

if __name__ == "__main__":
    commander = AudioCommander()
    commander.monitor_loop()
