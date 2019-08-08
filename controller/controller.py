import argparse
import os
import shutil
import threading
import time
import win32api
import win32gui
from os.path import exists, join
from subprocess import Popen, PIPE, STDOUT
from xml.dom.minidom import parse

import win32con
import xlrd
from docx import Document


def check(file, args):
    check_time = time.time()
    if args.format == 'xml':
        while time.time() - check_time < args.timeout:
            blocking()
            try:
                _ = parse(file)
                return True
            except:
                continue
        return False
    if args.format == 'docx':
        while time.time() - check_time < args.timeout:
            blocking()
            try:
                _ = Document(file)
                return True
            except:
                continue
        return False
    if args.format == 'xlsx':
        while time.time() - check_time < args.timeout:
            blocking()
            try:
                _ = xlrd.open_workbook(file)
                return True
            except:
                continue
        return False


def iter_files(path):
    """Walk through all files located under a root path."""
    if os.path.isfile(path):
        yield path
    elif os.path.isdir(path):
        for dir_path, _, file_names in os.walk(path):
            for f in file_names:
                yield os.path.join(dir_path, f)
    else:
        raise RuntimeError('Path %s is invalid' % path)


def pdf2html(cmd):
    Popen(cmd, shell=True)


def movefile(file, path):
    try:
        shutil.copy(file, path)
        os.remove(file)
        # shutil.rmtree(each[1])
    except Exception as e:
        print(e)


def kill_tasks():
    task_list = ["pdf2html.exe", "Acrobat.exe", "AcroCEF.exe"]
    for task in task_list:
        cmd = "tasklist | findstr \"%s\"" % task
        kill = Popen(cmd, stdout=PIPE, stderr=STDOUT, shell=True)
        info = kill.stdout.readlines()
        for each in info:
            each = each.decode('utf-8')
            Popen("taskkill /pid %s -t -f" % each.split()
            [1], stdout=PIPE, stderr=STDOUT, shell=True)


def get_child_windows(parent):
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


def blocking():
    whd = win32gui.FindWindow(0, 'Adobe Acrobat')
    if whd > 0:
        hwndChildList = get_child_windows(whd)
        for hwnd in hwndChildList:
            try:
                title = win32gui.GetWindowText(hwnd)
                class_name = win32gui.GetClassName(hwnd)
                # print(title,class_name)
                if (title == "确定(&O)" or title == "确定") and class_name == "Button":
                    win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, 0, 0)
                    win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, 0)
                    break
            except:
                continue


def main(args):
    kill_tasks()
    if not exists(args.input):
        raise FileNotFoundError

    if not exists(args.output):
        os.makedirs(args.output)

    done = open(args.done_path, 'a+')
    fail = open(args.fail_path, 'a+')
    converted = open(args.done_path).read() + open(args.fail_path).read()

    converted = converted.split('\n')
    if len(converted[-1]) < 2:
        converted = converted[:-1]
    converted = set(converted)
    HAVE_DONE = len(converted)
    # print("Files have been converted will be move to:\n%s" % args.done_path)
    print("Having converted %d" % HAVE_DONE)
    print("############ Starting Converting ############")
    files = list(iter_files(args.input))
    files = list(set(files) - converted)
    files.sort()
    for id, file in enumerate(files):
        print(file)
        name, _ = os.path.splitext(os.path.basename(file))
        save_path = join(args.output, name + '.' + args.format)
        if os.path.exists(save_path):
            print('exists!')
            if check(save_path, args) and not args.replace:
                print('%d/%d done!' % (id + 1, len(files)))
                done.write(file + '\n')
                done.flush()
                continue
            else:
                os.remove(save_path)
                print('error %s, converted again!' % args.format)

        cmd = "%s -i %s -o %s -f %s" % (args.exe,
                                        file, args.output, args.format)
        thread = threading.Thread(target=pdf2html, args=(cmd,))
        thread.start()

        start_time = time.time()
        while True:
            if os.path.exists(save_path):
                if check(save_path, args):
                    print('%d/%d ok!' % (id + 1, len(files)))
                    done.write(file + '\n')
                    done.flush()
                    print("Time cost: %.2fs" % (time.time() - start_time))
                    kill_tasks()
                    break
                else:
                    print('Timeout! %d/%d fail!' % (id + 1, len(files)))
                    kill_tasks()
                    fail.write(file + '\n')
                    fail.flush()
                    break
            if time.time() - start_time > args.timeout:
                break
            blocking()

    done.close()
    fail.close()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='bat run of the pdf2html')

    parser.add_argument('--input', '-i', required=True, action='store',
                        help='Path of files you want to convert')
    parser.add_argument('--output', '-o', default=join(os.getcwd(), 'output'), action='store',
                        help='Path to store converted files')
    parser.add_argument('--done_path', default=join(os.getcwd(), 'done.txt'),
                        help='Path to store done pdfs')
    parser.add_argument('--fail_path', default=join(os.getcwd(), 'fail.txt'),
                        help='Path to store fail pdfs')
    parser.add_argument('--exe', '-e', required=True, action='store',
                        help='Path of pdf2html.exe')
    parser.add_argument('--replace', '-r', default=True, action='store',
                        help='Whether replace exist files')
    parser.add_argument('--timeout', '-t', type=float, default=180, action='store',
                        help='Timeout for a single convert')
    parser.add_argument('--format', '-f', default='xml', action='store',
                        help='Format you want to convert: '
                             'xml, txt, doc, docx, '
                             'ps, jpeg, jpe, jpg, jpf, '
                             'jpx, j2k, j2c, jpc, rtf, xlsx'
                             'accesstext, tif, tiff')
    args = parser.parse_args()
    main(args)
