# -*- coding=utf-8 -*-
import csv
import datetime
import os
import re
import shutil
import sys
import uuid
import zipfile
from ftplib import FTP, all_errors
import pythoncom
import subprocess
from PyQt5.QtCore import QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox, QDesktopWidget
from win32com.shell import shell
from main_window_2 import Ui_MainWindow


# 主窗口
class MyMainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MyMainWindow, self).__init__(parent)
        # 选中的软件名字
        self.select_module_soft_name = None

        # ftp 连接标志，True: 已连接； False: 未连接。
        self.ftp_link_status = False
        # ftp 内所有的软件
        self.all_module_soft_list = []
        # 搜索结果
        self.search_result_dict = {}
        # 创建 FTP
        self.ftp = FTP()
        # 本机 MAC 地址
        self.mac = None
        # 已下载文件大小
        self.download_file_size = None
        # 下载进度
        self.download_process_num = None
        # 定时 1 分钟
        self.time = 60
        # 创建下载线程，并绑定下载成功提示函数
        self.thread = DownloadThread()
        self.thread.sinOut.connect(self.download_module_soft)

        # 创建日志线程
        self.log_thread = LogThread()

        # 设置程序信息
        self.setupUi(self)
        self.setWindowTitle('FTP 下载工具')
        self.setWindowIcon(QIcon(r'./images/cartoon5.ico'))

        # 获取本机 MAC 地址
        self.get_mac()
        self.statusbar.showMessage(r'本机 MAC 地址： ' + self.mac)

        # ip 写死
        self.ip_1.setEnabled(False)

        # 点 “连接” 连接到 FTP 地址，
        self.link_ftp_btn.clicked.connect(self.link_ftp)

        # 点 “搜索” 搜索模组
        self.search_module_btn.clicked.connect(self.search_module)

        # 点 “清除” 清除搜索框
        self.clear_btn.clicked.connect(self.clear_search_soft)

        # 点 “下载” 下载模组
        self.download_btn.clicked.connect(self.start_download_thread)

        # 搜索的内容全部转换成大写
        self.module_name.textChanged.connect(self.upper_module_name)

        # 隐藏下载进度条
        self.progressBar.hide()

        # 创建一个下载进度定时器
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.get_download_process)

        # 创建一个连接后，1 分钟不操作就断开连接的定时器
        self.timer2 = QTimer(self)
        self.timer2.timeout.connect(self.disconnect_ftp)

        # 让主窗口居中显示
        self.center()

    # 主窗口居中显示函数
    def center(self):
        screen = QDesktopWidget().screenGeometry()
        size = self.geometry()
        self.move((screen.width() - size.width()) / 2, (screen.height() - size.height()) / 2)

    # 下载定时器，用来获取下载进度
    def get_download_process(self):
        l_size = os.path.getsize(r'D:\专用软件\\' + self.select_module_soft_name)
        self.download_process_num = l_size / self.download_file_size / 2 * 100
        self.progressBar.setValue(int(self.download_process_num))

    # 断开 FTP 定时器
    def disconnect_ftp(self):
        self.time -= 1
        if self.time == 0:
            self.timer2.stop()
            self.ftp.quit()

    # 获取本机 mac 地址
    def get_mac(self):
        mac = uuid.uuid1().hex[-12:].upper()
        mac = re.findall(r'.{2}', mac)
        self.mac = ':'.join(mac)

    # 搜索的字符变成大写
    def upper_module_name(self, text):
        self.module_name.setText(text.upper())

    # 连接到 FTP 地址函数
    def link_ftp(self):
        ftp_addr = self.ip_1.text()
        try:
            '''
            self.ftp.connect(ftp_addr, port=26, timeout=2.5)
            self.ftp.login(user="rjk", passwd="rjk123@")
            '''
            self.ftp.connect('192.168.64.105', port=22, timeout=2.5)
            self.ftp.login(user="rjk", passwd="123.")
        except all_errors:
            # 连接失败
            reply = QMessageBox.information(self, '连接失败', '连接失败，请检查网络')
        else:
            # 连接成功
            if not self.ftp_link_status:
                self.ftp_link_status = True
                reply = QMessageBox.information(self, '连接成功', '连接成功')
                self.timer2.start(1000)
                self.link_ftp_btn.setText('已连接')
                self.link_ftp_btn.setEnabled(False)
            self.ftp.encoding = 'GB18030'
            self.ftp.cwd(r'\.\Software')

    # 搜索模组
    def search_module(self):
        if self.ftp_link_status:
            self.timer2.stop()
            self.link_ftp()
            # 获取当前目录下所有文件
            self.all_module_soft_list = self.ftp.nlst()
            # 初始化搜索结果字典
            self.search_result_dict = {}
            module_name = self.module_name.text()
            if len(module_name.strip()) == 0:
                reply = QMessageBox.information(self, '搜索失败', '请输入模组名称')
            else:
                for item in self.all_module_soft_list:
                    if re.search(module_name, item.rpartition('.')[0], flags=re.I):
                        self.search_result_dict[item.rpartition('.')[0]] = item
                if self.search_result_dict:
                    # 清楚显示框的内容
                    self.module_soft_list.clear()
                    self.oqc_soft_list.clear()
                    self.restart_soft_list.clear()
                    # 把搜索到的软件放到显示框里面
                    for key in self.search_result_dict.keys():
                        if re.search('OQC', key, flags=re.I):
                            self.oqc_soft_list.addItem(key)
                        elif re.search('重工', key, flags=re.I) or re.search('返修', key, flags=re.I):
                            self.restart_soft_list.addItem(key)
                        else:
                            self.module_soft_list.addItem(key)
                    # 获取选中的软件的名字
                    self.module_soft_list.itemClicked.connect(self.select_module_soft)
                    self.oqc_soft_list.itemClicked.connect(self.select_module_soft)
                    self.restart_soft_list.itemClicked.connect(self.select_module_soft)
                else:
                    reply = QMessageBox.information(self, '搜索失败', '没有找到该模组，请确认模组名称')
                self.time = 60
                self.timer2.start(1000)
        else:
            reply = QMessageBox.information(self, '搜索失败', '请先连接 IP 地址')

    # 清空模组搜索框的内容
    def clear_search_soft(self, event):
        self.module_name.setText('')
        self.module_name.setFocus()

    # 获取选中的软件的名字
    def select_module_soft(self, item):
        self.select_module_soft_name = self.search_result_dict[item.text()]

    # 点下载按钮，启动下载线程
    def start_download_thread(self, msg):
        if self.ftp_link_status:
            self.link_ftp()

            if self.select_module_soft_name:
                self.download_file_size = self.ftp.size(self.select_module_soft_name)
                # 启动下载线程
                self.thread.select_module_soft_name = self.select_module_soft_name
                self.thread.ftp = self.ftp
                self.thread.start()
                self.download_btn.setEnabled(False)
                self.progressBar.show()
                self.download_file.setText(self.select_module_soft_name)
            else:
                reply = QMessageBox.information(self, '下载失败', '请先选中下载内容')
        else:
            reply = QMessageBox.information(self, '提示', '请先连接服务器')

    # 下载线程绑定槽
    def download_module_soft(self, msg):
        if msg == '开始下载':
            self.progressBar.setValue(0)
            self.timer.start(500)
            self.download_process.setText('正在下载中，请稍后')
        elif msg == '开始解压':
            self.timer.stop()
            self.download_process.setText('正在解压中，请稍后')
            # 创建并启动日志线程
            self.log_thread.select_module_soft_name = self.select_module_soft_name
            self.log_thread.mac = self.mac
            self.log_thread.start()
        elif msg == '下载成功':
            self.download_process.setText(msg)

            # 删除 zip
            os.remove(r'D:\专用软件\\' + self.select_module_soft_name)
            reply = QMessageBox.information(self, "提示", "下载成功！")
            self.select_module_soft_name = None
            try:
                subprocess.Popen('explorer.exe /n, D:\专用软件')
            except Exception as e:
                print(e)
            sys.exit()
        else:
            self.progressBar.setValue(int(msg))


# 下载线程
class DownloadThread(QThread):
    sinOut = pyqtSignal(str)

    def __init__(self, select_module_soft_name=None, ftp=None):
        super(DownloadThread, self).__init__()
        self.select_module_soft_name = select_module_soft_name
        self.ftp = ftp

    def run(self):
        try:
            self.sinOut.emit('开始下载')
            # 清空放模组软件的文件夹
            if os.path.isdir(r'D:\专用软件'):
                shutil.rmtree(r'D:\专用软件')
                os.mkdir(r'D:\专用软件')
            else:
                os.mkdir(r'D:\专用软件')

            # 清空桌面
            ls = os.listdir(r'C:\Users')
            for i in ls:
                try:
                    if os.listdir(r'C:\Users\\' + i + '\Desktop') and i != 'Public':
                        addr = r'C:\Users\\' + i + '\Desktop\\'
                except:
                    pass

            desktop_ls = os.listdir(addr)
            for i in desktop_ls:
                if os.path.isfile(addr + i):
                    try:
                        os.remove(addr + i)
                    except Exception as e:
                        print(e)
                elif os.path.isdir(addr + i):
                    try:
                        shutil.rmtree(addr + i)
                    except Exception as e:
                        print(e)
        except Exception as e:
            print(e)
        try:
            pythoncom.CoInitialize()
            # 创建放模组软件的文件夹的快捷方式
            filename = r"D:\专用软件"  # 要创建快捷方式的文件的完整路径
            lnkname = addr + r"专用软件.lnk"  # 将要在此路径创建快捷方式
            shortcut = pythoncom.CoCreateInstance(
                shell.CLSID_ShellLink, None,
                pythoncom.CLSCTX_INPROC_SERVER, shell.IID_IShellLink)
            shortcut.SetPath(filename)
            shortcut.SetWorkingDirectory(r"D:\\")  # 设置快捷方式的起始位置, 不然会出现找不到辅助文件的情况
            shortcut.QueryInterface(pythoncom.IID_IPersistFile).Save(lnkname, 0)
        except Exception as e:
            print(e)
        try:
            # 下载 zip
            with open(r'D:\专用软件\\' + self.select_module_soft_name, 'wb') as fp:
                self.ftp.retrbinary('retr ' + self.select_module_soft_name, fp.write)
            self.sinOut.emit('开始解压')

            # 解压
            f = zipfile.ZipFile(r'D:\专用软件\\' + self.select_module_soft_name, 'r',)
            f_name_list = f.namelist()
            f_name_list_len = len(f_name_list)
            i = 0
            for file in f_name_list:
                f.extract(file, r'D:\专用软件\\')
                i += 1
                process = int(i / f_name_list_len / 2 * 100 + 50)
                if process in [55, 60, 65, 70, 75, 80, 85, 90, 95, 100]:
                    self.sinOut.emit(str(process))
            f.close()
            self.sinOut.emit('下载成功')
            self.quit()
        except Exception as e:
            print(e)


# 日志线程
class LogThread(QThread):
    sinOut = pyqtSignal(str)

    def __init__(self, select_module_soft_name=None, mac=None):
        super(LogThread, self).__init__()
        self.select_module_soft_name = select_module_soft_name
        self.mac = mac

    def run(self):
        try:
            if not os.path.isdir(r'D:\\' + str(self.mac).replace(':', '') + '_Log'):
                os.mkdir(r'D:\\' + str(self.mac).replace(':', '') + '_Log')
            # 添加日志
            with open(r'D:\\' + str(self.mac).replace(':', '') + '_Log\\' + str(datetime.date.today()) + '.csv', 'a', newline='') as fp:
                fieldnames = ['软件名称', '下载日期']
                writer = csv.DictWriter(fp, fieldnames)
                writer.writerow({'软件名称': str(self.select_module_soft_name), '下载日期': str(datetime.datetime.now())})
        except Exception as e:
            print(e)
        self.sinOut.emit('日志上传成功')
        self.quit()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_win = MyMainWindow()
    style = '''
        #MainWindow {
            background-color: #cae1ff;
        }
        '''
    main_win.setStyleSheet(style)
    main_win.show()
    sys.exit(app.exec_())