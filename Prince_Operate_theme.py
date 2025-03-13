import sys
import os
import re
import time
import math
import pandas as pd
import csv
import numpy as np
# import win32com.client
import datetime
import chicon  # 引用图标
# from PyQt5 import QtCore, QtGui, QtWidgets
# from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtWidgets import QApplication, QFileDialog, QMainWindow, QMessageBox, QVBoxLayout, QPushButton, QAction
# from PyQt5.QtCore import *
from PyQt5.QtGui import QIcon
from Get_Data import *
from Prince_Operate_Ui import Ui_MainWindow
from Table_Operate import *
from Logger import *
from theme_manager_theme import ThemeManager
from File_Operate import *
from Browser_operation import *


class MyMainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MyMainWindow, self).__init__(parent)
        self.setupUi(self)

        self.theme_manager = ThemeManager(QApplication.instance())
        self.init_theme_action()

        self.setGeometry(100, 100, 300, 200)

        self.theme_manager = ThemeManager(QApplication.instance())

        layout = QVBoxLayout()

        toggle_button = QPushButton("Toggle Theme")
        toggle_button.clicked.connect(self.theme_manager.toggle_theme)
        layout.addWidget(toggle_button)

        self.actionExport.triggered.connect(self.exportConfig)
        self.actionImport.triggered.connect(self.importConfig)
        self.actionView.triggered.connect(lambda: self.viewData('%s\\config_prince.csv' % configFileUrl))
        self.actionExit.triggered.connect(MyMainWindow.close)
        self.actionHelp.triggered.connect(self.showVersion)
        self.actionAuthor.triggered.connect(self.showAuthorMessage)
        self.theme_manager.set_theme("blue")  # 设置默认主题
        self.pushButton.clicked.connect(self.prince_op)
        self.pushButton_4.clicked.connect(self.textBrowser.clear)
        self.pushButton_2.clicked.connect(self.getPrinceFile)
        self.pushButton_5.clicked.connect(lambda: self.viewData(self.lineEdit.text()))

    def init_theme_action(self):
        theme_action = QAction(QIcon('theme_icon.png'), 'Toggle Theme', self)
        theme_action.setStatusTip('Toggle Theme')
        theme_action.triggered.connect(self.toggle_theme)

        # 将 action 添加到菜单（如果有的话）
        if hasattr(self, 'menuBar'):
            view_menu = self.menuBar().addMenu('Theme')
            view_menu.addAction(theme_action)

        # # 将 action 添加到工具栏
        # toolbar = self.addToolBar('主题')
        # toolbar.addAction(theme_action)

    def toggle_theme(self):
        self.theme_manager.set_random_theme()
        # 可以在这里添加其他需要在主题切换后更新的UI元素

    def getConfig(self):
        # 初始化，获取或生成配置文件
        global configFileUrl
        global desktopUrl
        global now
        global last_time
        global today
        global fileUrl

        now = int(time.strftime('%Y'))
        last_time = now - 1
        today = time.strftime('%Y.%m.%d')
        desktopUrl = os.path.join(os.path.expanduser("~"), 'Desktop')
        configFileUrl = '%s\\config' % desktopUrl
        configFile = os.path.exists('%s/config_prince.csv' % configFileUrl)
        # print(desktopUrl,configFileUrl,configFile)
        if not configFile:  # 判断是否存在文件夹如果不存在则创建为文件夹
            reply = QMessageBox.question(self, '信息', '确认是否要创建配置文件', QMessageBox.Yes | QMessageBox.No,
                                         QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                if not os.path.exists(configFileUrl):
                    os.makedirs(configFileUrl)
                MyMainWindow.createConfigContent(self)
                MyMainWindow.getConfigContent(self)
                self.textBrowser.append("创建并导入配置成功")
            else:
                exit()
        else:
            MyMainWindow.getConfigContent(self)

    # 获取配置文件内容
    def getConfigContent(self):
        # 配置文件
        csvFile = pd.read_csv('%s/config_prince.csv' % configFileUrl, names=['A', 'B', 'C'])
        global configContent
        global username
        global role

        configContent = {}
        username = list(csvFile['A'])
        number = list(csvFile['B'])
        role = list(csvFile['C'])
        for i in range(len(username)):
            configContent['%s' % username[i]] = number[i]

        # 新增校验逻辑
        required_keys = ('Account', 'Password', 'Files_Import_URL', 'Files_Name', 'Files_Export_URL', 'Browser_URL')
        missing_keys = [k for k in required_keys if k not in configContent]
        if missing_keys:
            reply = QMessageBox.question(self, '信息', f"缺少必要配置项: {', '.join(missing_keys)}",
                                         QMessageBox.Yes | QMessageBox.No,
                                         QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                MyMainWindow.createConfigContent(self)
                self.textBrowser.append("创建并导入配置成功")
                self.textBrowser.append('----------------------------------')
                app.processEvents()
        MyMainWindow.getDefaultInformation(self)
        try:
            self.textBrowser.append("配置获取成功")
            self.textBrowser.append('----------------------------------')
        except AttributeError:
            QMessageBox.information(self, "提示信息", "已获取配置文件内容", QMessageBox.Yes)
        else:
            pass

    # 创建配置文件
    def createConfigContent(self):
        global monthAbbrev
        months = "JanFebMarAprMayJunJulAugSepOctNovDec"
        n = time.strftime('%m')
        pos = (int(n) - 1) * 3
        monthAbbrev = months[pos:pos + 3]

        configContent = [
            ['信息', '内容', '备注'],
            ['Account', 'chen-fr@cn001.itgr.net', '账号'],
            ['Password', 'As123123', '密码'],
            ['Files_Import_URL',
             'N:\\XM Softlines\\6. Personel\\5. Personal\Supporting Team\\2.财务\\7.外部分包\\4.软件数据',
             '文件数据输入路径'],
            ['Files_Name', '采购录入.csv', '文件名称'],
            ['Files_Export_URL',
             'N:\\XM Softlines\\6. Personel\\5. Personal\Supporting Team\\2.财务\\7.外部分包\\4.软件数据',
             '文件导出路径'],
            ['Browser_URL',
             'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
             '浏览器路径'],
        ]
        config = np.array(configContent)
        df = pd.DataFrame(config)
        df.to_csv('%s/config_prince.csv' % configFileUrl, index=0, header=0, encoding='utf_8_sig')
        self.textBrowser.append("配置文件创建成功")
        QMessageBox.information(self, "提示信息",
                                "默认配置文件已经创建好，\n如需修改请在用户桌面查找config文件夹中config_prince.csv，\n将相应的文件内容替换成用户需求即可，修改后记得重新导入配置文件。",
                                QMessageBox.Yes)

    # 导出配置文件
    def exportConfig(self):
        # 重新导出默认配置文件
        reply = QMessageBox.question(self, '信息', '确认是否要创建默认配置文件', QMessageBox.Yes | QMessageBox.No,
                                     QMessageBox.Yes)
        if reply == QMessageBox.Yes:
            MyMainWindow.createConfigContent(self)
        else:
            QMessageBox.information(self, "提示信息", "没有创建默认配置文件，保留原有的配置文件", QMessageBox.Yes)

    # 导入配置文件
    def importConfig(self):
        # 重新导入配置文件
        reply = QMessageBox.question(self, '信息', '确认是否要导入配置文件', QMessageBox.Yes | QMessageBox.No,
                                     QMessageBox.Yes)
        if reply == QMessageBox.Yes:
            MyMainWindow.getConfigContent(self)
        else:
            QMessageBox.information(self, "提示信息", "没有重新导入配置文件，将按照原有的配置文件操作", QMessageBox.Yes)

    # 界面设置默认配置文件信息
    def getDefaultInformation(self):
        # 默认登录界面信息
        try:
            # data处理
            self.lineEdit.setText('%s\\%s' % (configContent['Files_Import_URL'], configContent['Files_Name']))
        except Exception as msg:
            self.textBrowser.append("错误信息：%s" % msg)
            self.textBrowser.append('----------------------------------')
            app.processEvents()
            reply = QMessageBox.question(self, '信息', '错误信息：%s。\n是否要重新创建配置文件' % msg,
                                         QMessageBox.Yes | QMessageBox.No,
                                         QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                MyMainWindow.createConfigContent(self)
                self.textBrowser.append("创建并导入配置成功")
                self.textBrowser.append('----------------------------------')
                app.processEvents()

    def showAuthorMessage(self):
        # 关于作者
        QMessageBox.about(self, "关于",
                          "人生苦短，码上行乐。\n\n\n        ----Frank Chen")

    def showVersion(self):
        # 关于作者
        QMessageBox.about(self, "版本",
                          "V 25.01.01\n\n\n 2025-02-27")

    # 获取文件
    def getFile(self):
        selectBatchFile = QFileDialog.getOpenFileName(self, '选择采购录入文件',
                                                      '%s' % configContent['Files_Import_URL'],
                                                      'files(*.docx;*.xls*;*.csv)')
        fileUrl = selectBatchFile[0]
        return fileUrl

    def getPrinceFile(self):
        file_path = myWin.getFile()
        self.lineEdit.setText(file_path)

    # 获取需要Combine文件的路径
    def getCombineFileUrl(self):
        fileUrl = MyMainWindow.getFile(self)
        if fileUrl:
            self.lineEdit_8.setText(fileUrl)
            app.processEvents()
        else:
            self.textBrowser_2.append("请重新选择ODM文件")
            QMessageBox.information(self, "提示信息", "请重新选择ODM文件", QMessageBox.Yes)

    # 查看SAP操作数据详情
    def viewData(self, url):
        try:
            fileUrl = url
            data_obj = Get_Data()
            df = data_obj.getFileData(fileUrl)
            myTable.createTable(df)
            myTable.showMaximized()
        except Exception as msg:
            self.textBrowser.append("错误信息：%s" % msg)

    def prince_op(self):
        prince_msg = {}
        prince_msg['flag'] = True
        data_file = self.lineEdit.text()
        web_url = self.lineEdit_2.text()
        now = time.strftime('%Y%m%d %H%M%S')
        log_file_path = '%s\\log %s.xlsx' % (configContent['Files_Export_URL'], now)
        log_file = Logger(log_file_path,
                          ['Update', 'request_id', 'Table row count', 'Prince Order Number', 'Prince 金额',
                           'OPdEX Order Number',
                           'OPdEX 未税金额(/1+税点)', 'remark', 'Order check', 'Revenue check'])
        # 判断文件是否存在
        if data_file != '' and web_url != '':
            log_list = {}
            data_obj = Get_Data()
            df = data_obj.getFileData(data_file)
            # 判断数据是否包含所有必要字段
            required_columns = ['Order Number', '未税金额(/1+税点)', '检测内容描述']
            if set(required_columns).issubset(df.columns):
                self.textBrowser.append("包含所有必要字段")
                try:
                    self.textBrowser.append("开始执行操作")
                    app.processEvents()
                    browser_obj = Browser(browser_path=configContent['Browser_URL'])
                    msg_login = browser_obj.login(web_url, configContent)
                    if msg_login['flag']:
                        self.textBrowser.append("登录成功")
                        app.processEvents()
                        for index, row in df.iterrows():
                            time.sleep(5)
                            log_list['OPdEX Order Number'] = row['Order Number']
                            log_list['OPdEX 未税金额(/1+税点)'] = round(float(row['未税金额(/1+税点)']), 2)
                            self.textBrowser.append(f"第 {index + 1}行开始处理")
                            app.processEvents()
                            process_msg = browser_obj.process_data_flow(row.to_dict())
                            if not process_msg['flag']:
                                self.textBrowser.append("<font color='red'>第%s行开始处理 %s 处理失败：%s</font>" % (
                                index + 1, row['Order Number'], process_msg['info']))
                                log_list['remark'] = process_msg.get['error_step', ''] + ';' + process_msg.get['info', ''] + ';' + process_msg.get['error', '']
                                app.processEvents()
                                browser_obj.close_iframe()
                            else:
                                log_list['remark'] = '完整流程跑完'
                            # 情况 1: 原始值为空或不存在时
                            # process_msg['data'] 中没有 'Prince Order Number' 字段
                            # 会使用默认值 '0' -> 最终得到 0
                            prince_order_str = str(process_msg['data'].get('Prince Order Number', '0')).strip()
                            # 移除逗号等非数字字符
                            prince_order_clean = prince_order_str.replace(',', '').replace(' ', '')
                            log_list['Prince Order Number'] = int(
                                prince_order_clean) if prince_order_clean.isdigit() else 0
                            prince_amount_str = str(process_msg['data'].get('Prince 金额', '0')).replace(',', '')
                            log_list['Prince 金额'] = float(prince_amount_str) if prince_amount_str else 0.0
                            log_list['Table row count'] = process_msg['Table row count']
                            log_list['request_id'] = process_msg['data'].get('request_id')
                            # 'Order check', 'Revenue check'
                            log_list['Order check'] = log_list['Prince Order Number'] == log_list['OPdEX Order Number']
                            log_list['Revenue check'] = log_list['Prince 金额'] == log_list['OPdEX 未税金额(/1+税点)']
                            # 记录成功日志
                            log_file.log(log_list)
                            self.textBrowser.append(f"第 {index + 1}行,{row['Order Number']}处理成功")
                            self.textBrowser.append("Request ID:%s Order No.:%s Prince 金额:%s" % (log_list['request_id'], str(log_list['Prince Order Number']), str(log_list['Prince 金额'])))
                            app.processEvents()

                        self.textBrowser.append("所有数据处理完成")
                        app.processEvents()
                        log_file.save_log_to_excel()
                        self.textBrowser.append("日志记录完成，保存路径%s" % log_file_path)
                        browser_obj.close_browser()
                        os.startfile(log_file_path)
                    else:
                        self.textBrowser.append("<font color='red'>登录失败</font>")
                        browser_obj.close_browser()
                except Exception as msg:
                    log_file.save_log_to_excel()
                    self.textBrowser.append("日志记录完成，保存路径%s" % log_file_path)
                    os.startfile(log_file_path)
                    browser_obj.close_browser()
                    self.textBrowser.append("<font color='red'>错误信息：%s</font>" % msg)
                    self.textBrowser.append("<font color='red'>退出采购录入</font>")
                    app.processEvents()

            else:
                missing = set(required_columns) - set(df.columns)
                self.textBrowser.append(f"<font color='red'>缺少必要字段: {missing}</font>")
                self.textBrowser.append(f"-----------------------------")
                app.processEvents()
        elif web_url == '':
            self.textBrowser.append(f"<font color='red'>没有采购网址</font>")
            self.textBrowser.append(f"-----------------------------")
            app.processEvents()
        else:
            self.textBrowser.append(f"<font color='red'>没有采购文件</font>")
            self.textBrowser.append(f"-----------------------------")
            app.processEvents()


if __name__ == "__main__":
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
    app = QApplication(sys.argv)
    myWin = MyMainWindow()
    myTable = MyTableWindow()
    myWin.show()
    myWin.resize(989, 732)
    myWin.getConfig()
    sys.exit(app.exec_())
