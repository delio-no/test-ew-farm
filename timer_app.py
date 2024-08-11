import keyboard
import threading
import time

import psutil
import win32process
import win32gui
import win32api
import win32con
import elevate

from keyboard_code import VK_CODE


def require_hwnd(func):
    def wrapper(self, *args, **kwargs):
        if self.hwnd:
            return func(self, *args, **kwargs)
        else:
            print('-----')
            print('Требуется захват окна - ctrl+q')

    return wrapper


class TimerApp:
    def __init__(self):
        self.process_id = None
        self.hwnd = None
        self.remaining_time = None
        self.timer_duration = 300  # default timer duration is 5 minutes
        self.is_run = False
        self.end_time = None

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
            print('Нажата клавиша z')
        except Exception as e:
            print('Ошибка нажатия клавиши')
            print(f"Ошибка: {e}")

    def timer_function(self):
        print('Таймер завершен')
        self.is_run = False
        self.press_key_z()

    def capture_window(self):
        self.process_id, self.hwnd = self.get_active_window_pid()
        if self.hwnd:
            print('-----')
            print('Окно захвачено')
        else:
            print('-----')
            print('Окно не является EndlessWar.exe')

    def set_timer(self):
        while True:
            try:
                self.timer_duration = int(input("Введите длительность таймера в секундах: "))
                print(f"Таймер установлен на {self.timer_duration} секунд")
                break
            except ValueError:
                print("Ошибка: введено не целое число. Попробуйте еще раз.")

    @require_hwnd
    def start(self):
        if self.is_run:
            print('-----')
            print('Таймер уже запущен')
            return
        if self.remaining_time:
            print('-----')
            print('Старт таймера')
            self.timer = threading.Timer(self.remaining_time, self.timer_function)
            self.start_time = time.time()
            self.end_time = self.start_time + self.remaining_time
            self.press_key_z()
        else:
            print('-----')
            print('Старт таймера')
            self.timer = threading.Timer(self.timer_duration, self.timer_function)
            self.start_time = time.time()
            self.end_time = self.start_time + self.timer_duration
            self.press_key_z()
        self.is_run = True
        self.timer.start()

    @require_hwnd
    def stop(self):
        if self.is_run:
            self.remaining_time = self.get_remaining_time()
            self.timer.cancel()
            self.is_run = False
            print('-----')
            print('Таймер остановлен')
            print(f'Время таймера: {self.get_remaining_time()}')
            self.press_key_z()
        else:
            print('-----')
            print('Таймер не запущен')

    @require_hwnd
    def reset(self):
        if self.is_run:
            self.timer.cancel()
            self.is_run = False
            self.remaining_time = None
            print('-----')
            print('Таймер сброшен')
            self.press_key_z()
        else:
            print('Таймер не запущен')

    def get_info_timer(self):
        print('-----')
        print(f'Интервал: {self.timer.interval} секунд')
        if self.is_run:
            print('Таймер запущен')
        else:
            print('Таймер остановлен')
        print(f"Время таймера: {self.get_remaining_time()}")

    def get_remaining_time(self):
        if self.timer.is_alive():
            return self.end_time - time.time()
        else:
            return self.remaining_time


elevate.elevate()

print('Как пользоваться:')
print('1. Ставим персонажа в точку для запуска фарма')
print('2. В активном окне EW нажимаем клаишу ctrl+Q для захвата окна')
print('3. Нажимаем клавишу ctrl+W для запуска фарма и таймера по дефолту 5 минут')

print('\n')
print('-------------------------')
print('\n')

# Описание функционала программы
print("Описание функционала программы:")
print("1. При нажатии 'ctrl+q' - захватывает HWID (PID) и HWND текущего активного окна.")
print("2. При нажатии 'ctrl+w' - отправляет клавишу 'z' в текущий активный процесс и запускает таймер.")
print("3. При нажатии 'ctrl+e' - останавливает таймер, выводит оставшееся время на таймере и отправляет клавишу 'z'.")
print("4. При нажатии 'ctrl+r' - сбрасывает таймер и отправляет клавишу 'z'.")
print("5. При нажатии 'ctrl+t' - устанавливает длительность таймера.")
print("5. При нажатии 'ctrl+y' - показать информацию о таймере.")

print('\n')
print('-------------------------')
print('\n')

app = TimerApp()

keyboard.add_hotkey('ctrl+q', app.capture_window)
keyboard.add_hotkey('ctrl+w', app.start)
keyboard.add_hotkey('ctrl+e', app.stop)
keyboard.add_hotkey('ctrl+r', app.reset)
keyboard.add_hotkey('ctrl+t', app.set_timer)
keyboard.add_hotkey('ctrl+y', app.get_info_timer)

keyboard.wait()
