import math
import subprocess
import os
import time
from tkinter import *
import tkinter.ttk as ttk
from threading import Thread
import requests as requests
import sys
from zipfile import ZipFile
import winshell
from win32com.client import Dispatch


class Updater:
    def __init__(self):
        self.cur_path = "."
        self.internal_share = "\\\\local_ip\\slt"
        try:
            responds = requests.get("https://site.ru/slt_ver.ini")
            self.last_ver = (responds.text + ".zip")
        except requests.exceptions.ConnectionError:
            self.last_ver = self.if_no_inet()
        self.update_file = f"{self.internal_share}\\{self.last_ver}"

    def if_no_inet(self):
        names = []
        for root, dirs, files in os.walk(self.internal_share):
            for filename in files:
                names.append(filename)
        return max(names)

    def clear_catalog(self):
        reserve = ["update", "slt.ini"]
        for root, dirs, files in os.walk(f"../{self.cur_path}", topdown=False):
            for f in files:
                for r in reserve:
                    if f == r:
                        break
                else:
                    try:
                        print("Удален файл: " + f)
                        os.remove(os.path.join(root, f))
                    except PermissionError:
                        pass
            for d in dirs:
                for r in reserve:
                    if d == r:
                        break
                else:
                    print("Удалена папка: " + d)
                    os.rmdir(os.path.join(root, d))

    def download_new_version(self):
        def progress():
            def unzip():
                zf = ZipFile(self.last_ver)
                uncompress_size = sum((file.file_size for file in zf.infolist()))
                extracted_size = 0

                for file in zf.infolist():
                    print("Извлечено " + str(file))
                    extracted_size += file.file_size
                    percentage = extracted_size * 100 / uncompress_size
                    pb["value"] = percentage
                    zf.extract(file, "../")

            try:
                fsize = int(os.path.getsize(self.update_file))
                with open(self.update_file, 'rb') as f:
                    with open(self.last_ver, 'ab') as n:
                        buffer = bytearray()
                        b = fsize / 8192
                        step = 100 / b
                        percent = 0
                        while True:
                            buf = f.read(8192)
                            pb["value"] = math.floor(percent)
                            percent += step
                            n.write(buf)
                            if len(buf) == 0:
                                break
                            buffer += buf
            except FileNotFoundError:
                s.configure('text.Horizontal.TProgressbar', text='Файл обновления не найден', font="14")
                time.sleep(5)
                root.destroy()
            pb.stop()
            s.configure('text.Horizontal.TProgressbar', text='Установка', font="Bahnschrift 14", background="#4d4d4d",
                        troughcolor='#333', troughrelief='flat', foreground='#DCDCDC')
            unzip()
            os.remove(self.last_ver)
            root.destroy()

        root = Tk()
        root.geometry("532x110")
        root["bg"] = "#333"

        label1 = ttk.Label(text="Обновление программы СЛТ, пожалуйста подождите...",
                           font="Bahnschrift 14", compound='center', foreground='#DCDCDC', background='#333')
        label1.pack(padx=(20, 20), pady=(10, 10))

        s = ttk.Style(root)
        s.theme_use('alt')
        s.layout('text.Horizontal.TProgressbar',
                 [('text.Horizontal.TProgressbar.trough',
                   {'children': [('text.Horizontal.TProgressbar.pbar', {'side': 'left', 'sticky': 'ns'})],
                    'sticky': 'nswe'}), ('Horizontal.TProgressbar.label', {'sticky': ''})])

        s.configure('text.Horizontal.TProgressbar', text=f'Загрузка', font="Bahnschrift 14", background="#A34239",
                    troughcolor='#333', troughrelief='flat', foreground='#DCDCDC')

        pb = ttk.Progressbar(root, style='text.Horizontal.TProgressbar', length=500, mode="determinate")
        pb.pack(ipady=10, pady=(0, 20))

        x = (root.winfo_screenwidth() - root.winfo_reqwidth() - 266) / 2
        y = (root.winfo_screenheight() - root.winfo_reqheight() - 55) / 2
        root.wm_geometry("+%d+%d" % (x, y))
        root.overrideredirect(1)
        root.attributes("-topmost", True)
        root.lift()
        Thread(target=progress, daemon=True).start()
        root.mainloop()

    def start_new_process(self):
        subprocess.Popen([f'{self.cur_dir}\slt.exe'], cwd=f'{self.cur_dir}\\')
        sys.exit(0)

    def create_shortcut(self):
        os.chdir('..')
        self.cur_dir = os.getcwd()
        desktop = winshell.desktop()
        lnk = os.path.exists(f"{desktop}\СЛТ.lnk")
        if lnk:
            pass
        else:
            path = os.path.join(desktop, "СЛТ.lnk")
            target = fr"{self.cur_dir}\slt.exe"
            icon = fr"{self.cur_dir}\img\icon.ico"
            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(path)
            shortcut.Targetpath = target
            shortcut.WorkingDirectory = self.cur_dir
            shortcut.IconLocation = icon
            shortcut.save()


if __name__ == '__main__':
    updater = Updater()
    updater.clear_catalog()
    updater.download_new_version()
    updater.create_shortcut()
    updater.start_new_process()
