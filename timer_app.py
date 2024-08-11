import customtkinter as ctk
import threading
import time
import psutil
import win32process
import win32gui
import win32api
import win32con
import elevate
import keyboard
from CTkMessagebox import CTkMessagebox

from keyboard_code import VK_CODE

ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

class TimerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Timer App")
        self.root.geometry("400x650")

        self.process_id = None
        self.hwnd = None
        self.remaining_time = None
        self.timer_duration = 300  # default timer duration is 5 minutes
        self.is_run = False
        self.end_time = None

        self.create_widgets()
        self.setup_hotkeys()

    def create_widgets(self):
        self.status_label = ctk.CTkLabel(self.root, text="Статус: Не захвачено")
        self.status_label.pack(pady=5)

        self.timer_label = ctk.CTkLabel(self.root, text="Таймер: 00:00")
        self.timer_label.pack(pady=5)

        self.timer_duration_label = ctk.CTkLabel(self.root, text=f"Текущая длительность таймера: {self.timer_duration} секунд")
        self.timer_duration_label.pack(pady=5)

        self.start_button = ctk.CTkButton(self.root, text="Старт", command=self.start)
        self.start_button.pack(pady=5)

        self.reset_button = ctk.CTkButton(self.root, text="Сброс", command=self.reset)
        self.reset_button.pack(pady=5)

        self.capture_window_button = ctk.CTkButton(self.root, text="Захватить окно",
                                                   command=self.find_and_capture_window)
        self.capture_window_button.pack(pady=5)

        self.set_timer_button = ctk.CTkButton(self.root, text="Установить таймер", command=self.set_timer)
        self.set_timer_button.pack(pady=5)

        self.description_label = ctk.CTkLabel(self.root, text="Описание функционала программы:",
                                              font=("Arial", 12, "bold"))
        self.description_label.pack(pady=5)

        self.description_text = ctk.CTkTextbox(self.root, width=380, height=125)
        self.description_text.pack(pady=5)
        self.description_text.insert(ctk.END,
                                     "1. При нажатии 'ctrl+q' - захватывает HWID (PID) и HWND текущего активного окна.\n")
        self.description_text.insert(ctk.END,
                                     "2. При нажатии 'ctrl+w' - отправляет клавишу 'z' в текущий активный процесс и запускает таймер.\n")
        self.description_text.insert(ctk.END, "3. При нажатии 'ctrl+e' - сбрасывает таймер.\n")
        self.description_text.configure(state="disabled")

        self.description_label_console = ctk.CTkLabel(self.root, text="Консоль:",
                                              font=("Arial", 12, "bold"))
        self.description_label_console.pack(pady=5)

        self.console = ctk.CTkTextbox(self.root, width=380, height=150)
        self.console.pack(pady=5)

    def setup_hotkeys(self):
        keyboard.add_hotkey('ctrl+q', self.capture_window)
        keyboard.add_hotkey('ctrl+w', self.start)
        keyboard.add_hotkey('ctrl+e', self.reset)

    def get_active_window_pid(self):
        hwnd = win32gui.GetForegroundWindow()
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        process = psutil.Process(pid)
        if process.name() == "EndlessWar.exe":
            return pid, hwnd
        else:
            return None, None

    def press_key_z(self):
        try:
            time.sleep(1)
            win32api.PostMessage(self.hwnd, win32con.WM_KEYDOWN, int(VK_CODE['z']), 0)
            win32api.PostMessage(self.hwnd, win32con.WM_KEYUP, int(VK_CODE['z']), 0)
            self.log('Нажата клавиша z')
        except Exception as e:
            self.log('Ошибка нажатия клавиши')
            self.log(f"Ошибка: {e}")

    def timer_function(self):
        self.log('Таймер завершен')
        self.is_run = False
        self.press_key_z()

    def capture_window(self):
        self.process_id, self.hwnd = self.get_active_window_pid()
        if self.hwnd:
            self.status_label.configure(text="Статус: Захвачено")
            self.log('Окно захвачено')
        else:
            self.status_label.configure(text="Статус: Не захвачено")
            self.log('Окно не является EndlessWar.exe')

    def find_and_capture_window(self):
        processes = [proc for proc in psutil.process_iter(['pid', 'name']) if proc.info['name'] == "EndlessWar.exe"]
        if len(processes) > 1:
            CTkMessagebox(title="Ошибка", message="Найдено более одного процесса EndlessWar.exe. Используйте горячую клавишу 'ctrl+q' для захвата окна.")
            return
        elif len(processes) == 1:
            self.process_id = processes[0].info['pid']
            self.hwnd = self.get_hwnd_from_pid(self.process_id)
            if self.hwnd:
                self.status_label.configure(text="Статус: Захвачено")
                self.log('Окно EndlessWar.exe захвачено')
                return
        self.status_label.configure(text="Статус: Не захвачено")
        self.log('Окно EndlessWar.exe не найдено')

    def get_hwnd_from_pid(self, pid):
        def callback(hwnd, hwnds):
            _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
            if found_pid == pid:
                hwnds.append(hwnd)
            return True

        hwnds = []
        win32gui.EnumWindows(callback, hwnds)
        return hwnds[0] if hwnds else None

    def set_timer(self):
        try:
            self.timer_duration = int(ctk.CTkInputDialog(title="Установить таймер", text="Введите длительность таймера в секундах:").get_input())
            if self.timer_duration is not None:
                self.timer_duration_label.configure(text=f"Текущая длительность таймера: {self.timer_duration} секунд")
                self.log(f"Таймер установлен на {self.timer_duration} секунд")
            else:
                CTkMessagebox(title="Ошибка", message="Ошибка: введено не целое число.")
        except ValueError:
            CTkMessagebox(title="Ошибка", message="Ошибка: введено не целое число.")

    def start(self):
        if self.hwnd:
            self.log('Старт таймера')
            self.timer = threading.Timer(self.timer_duration, self.timer_function)
            self.press_key_z()
            self.start_time = time.time()
            self.end_time = self.start_time + self.timer_duration
            self.is_run = True
            self.timer.start()
            self.update_timer()
        else:
            self.log('Требуется захват окна - ctrl+q')

    def reset(self):
        if self.is_run:
            self.timer.cancel()
            self.is_run = False
            self.remaining_time = None
            self.timer_label.configure(text="Таймер: 00:00")
            self.log('Таймер сброшен')
        else:
            self.log('Таймер не запущен')

    def update_timer(self):
        if self.is_run:
            remaining_time = self.end_time - time.time()
            if remaining_time > 0:
                minutes, seconds = divmod(int(remaining_time), 60)
                self.timer_label.configure(text=f"Таймер: {minutes:02d}:{seconds:02d}")
                self.root.after(1000, self.update_timer)
            else:
                self.timer_label.configure(text="Таймер: 00:00")

    def log(self, message):
        self.console.insert(ctk.END, message + "\n")
        self.console.see(ctk.END)

elevate.elevate()

root = ctk.CTk()
app = TimerApp(root)

root.mainloop()