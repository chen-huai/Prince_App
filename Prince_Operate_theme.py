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
        self.pushButton_4.clicked.connect(self.textBrowser.clear)
        self.pushButton_2.clicked.connect(self.getPrinceFile)

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
        global oneWeekday
        global fileUrl

        date = datetime.datetime.now() + datetime.timedelta(days=1)
        now = int(time.strftime('%Y'))
        last_time = now - 1
        today = time.strftime('%Y.%m.%d')
        oneWeekday = (datetime.datetime.now() + datetime.timedelta(days=7)).strftime('%Y.%m.%d')
        desktopUrl = os.path.join(os.path.expanduser("~"), 'Desktop')
        configFileUrl = '%s\\config' % desktopUrl
        configFile = os.path.exists('%s/config_prince.csv' % configFileUrl)
        # print(desktopUrl,configFileUrl,configFile)
        if not configFile:  # 判断是否存在文件夹如果不存在则创建为文件夹
            reply = QMessageBox.question(self, '信息', '确认是否要创建配置文件', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
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
        required_keys = ('Date_URL', 'Invoice_File_URL', 'Billing_List_URL')
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
        else:
            MyMainWindow.getDefaultInformation(self)
            try:
                self.textBrowser.append("配置获取成功")
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
            ['Account', 'chen-fr@cn001.itgr.net','文件数据路径'],
            ['Password', 'As123123','文件数据路径'],
            ['Files_Import_URL', 'N:\\XM Softlines\\6. Personel\\5. Personal\\Supporting Team\\收样\\3.Sap\\ODM Data - XM',
             '文件数据路径'],
            ['Files_Name','采购录入.csv','特殊开票文件路径'],
            ['Files_Export_URL',
             'N:\\XM Softlines\\6. Personel\\5. Personal\\Supporting Team\\收样\\3.Sap\\ODM Data - XM\\2.特殊开票',
             '特殊开票文件路径'],
            ['Browser_URL',
             'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
             '特殊开票文件路径'],
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
        reply = QMessageBox.question(self, '信息', '确认是否要创建默认配置文件', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if reply == QMessageBox.Yes:
            MyMainWindow.createConfigContent(self)
        else:
            QMessageBox.information(self, "提示信息", "没有创建默认配置文件，保留原有的配置文件", QMessageBox.Yes)

    # 导入配置文件
    def importConfig(self):
        # 重新导入配置文件
        reply = QMessageBox.question(self, '信息', '确认是否要导入配置文件', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
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
            self.textBrowser.append("配置获取成功")
            self.textBrowser.append('----------------------------------')
        except Exception as msg:
            self.textBrowser.append("错误信息：%s" % msg)
            self.textBrowser.append('----------------------------------')
            app.processEvents()
            reply = QMessageBox.question(self, '信息', '错误信息：%s。\n是否要重新创建配置文件' % msg, QMessageBox.Yes | QMessageBox.No,
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
        fileUrl = url
        data_obj = Get_Data()
        df = data_obj.getFileData(fileUrl)
        myTable.createTable(df)
        myTable.showMaximized()

    def prince_op(self):
        data_file = self.lineEdit.text()
        web_url = self.lineEdit_2.text()
        if data_file:
            self.textBrowser.append("开始执行操作")
            app.processEvents()
            data_obj = Get_Data()
            df = data_obj.getFileData(data_file)
            if df.empty:
                self.textBrowser.append("文件为空")
                app.processEvents()
            else:
                pass




if __name__ == "__main__":
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
    app = QApplication(sys.argv)
    myWin = MyMainWindow()
    myTable = MyTableWindow()
    myWin.show()
    myWin.resize(989, 732)
    myWin.getConfig()
    sys.exit(app.exec_())
