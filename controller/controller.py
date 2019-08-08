import os
import threading
import time
import win32api
import win32gui
from os.path import exists, join
from subprocess import Popen, PIPE, STDOUT
from xml.dom.minidom import parse

import win32con
import xlrd


class Converter:
    def __init__(self, input, exe, format='xml', timeout=60, output=None):
        self.input = input
        self.exe = exe
        self.format = format
        self.timeout = timeout
        self.output = output
        if self.output is None:
            self.output = join(os.getcwd(), 'output')

    def check(self, file):
        check_time = time.time()
        if self.format == 'xml':
            while time.time() - check_time < self.timeout:
                self.blocking()
                try:
                    _ = parse(file)
                    return True
                except:
                    continue
            return False

        if self.format == 'xlsx':
            while time.time() - check_time < self.timeout:
                self.blocking()
                try:
                    _ = xlrd.open_workbook(file)
                    return True
                except:
                    continue
            return False


    def pdf2html(self,cmd):
        Popen(cmd, shell=True)

    def kill_tasks(self):
        task_list = ["pdf2html.exe", "Acrobat.exe", "AcroCEF.exe"]
        for task in task_list:
            cmd = "tasklist | findstr \"%s\"" % task
            kill = Popen(cmd, stdout=PIPE, stderr=STDOUT, shell=True)
            info = kill.stdout.readlines()
            for each in info:
                each = each.decode('utf-8')
                Popen("taskkill /pid %s -t -f" % each.split()
                [1], stdout=PIPE, stderr=STDOUT, shell=True)

    def get_child_windows(self,parent):
        '''
         获得parent的所有子窗口句柄
         返回子窗口句柄列表
         '''
        if not parent:
            return
        hwndChildList = []
        try:
            win32gui.EnumChildWindows(
                parent, lambda hwnd, param: param.append(hwnd), hwndChildList)
        except:
            return []
        return hwndChildList

    def blocking(self):
        whd = win32gui.FindWindow(0, 'Adobe Acrobat')
        if whd > 0:
            hwndChildList = self.get_child_windows(whd)
            for hwnd in hwndChildList:
                try:
                    title = win32gui.GetWindowText(hwnd)
                    class_name = win32gui.GetClassName(hwnd)
                    if (title == "确定(&O)" or title == "确定") and class_name == "Button":
                        win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, 0, 0)
                        win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, 0)
                        break
                except:
                    continue

    def convert(self):
        self.kill_tasks()
        if not exists(self.input):
            raise FileNotFoundError

        if not exists(self.output):
            os.makedirs(self.output)

        print("############ Starting Converting ############")
        file = self.input
        name, _ = os.path.splitext(os.path.basename(file))
        save_path = join(self.output, name + '.' + self.format)
        if os.path.exists(save_path):
            os.remove(save_path)
        cmd = "%s -i %s -o %s -f %s" % (self.exe,
                                        file, self.output, self.format)
        thread = threading.Thread(target=self.pdf2html, args=(cmd,))
        thread.start()

        start_time = time.time()
        while True:
            if os.path.exists(save_path):
                if self.check(save_path):
                    print('%s ok!' % file)
                    print("Time cost: %.2fs" % (time.time() - start_time))
                    self.kill_tasks()
                    return save_path
                else:
                    print('%s fail!' % file)
                    self.kill_tasks()
                    return None
            if time.time() - start_time > self.timeout:
                return None
            self.blocking()
