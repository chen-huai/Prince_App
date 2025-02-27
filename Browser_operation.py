import time

from playwright.sync_api import sync_playwright
import logging
import pandas as pd
from datetime import datetime
import os
import re

class Browser():
    def __init__(self, browser_path, output_path):
        self.browser_path = browser_path
        self.output_path = output_path
        self.page_title = None
        self.playwright = None
        self.browser = None
        self.msg = {}
        self.msg['flag'] = True
        self.msg['error'] = []
        self.msg['info'] = []
        # # 登录网址
        # self.playwright = sync_playwright().start()
        # self.msg = {}
        # self.msg['error'] = []
        # self.msg['info'] = []
        # try:
        #     # 启动浏览器
        #     self.browser = self.playwright.chromium.launch(headless=False, executable_path=self.browser_path)
        #     # 创建新的浏览器上下文
        #     self.context = self.browser.new_context(user_agent='chen-fr@cn001.itgr.net')
        #     # 创建新的页面
        #     self.page = self.context.new_page()
        #     # 导航到指定URL
        #     self.page.goto(url)
        #     # 填写用户名和密码
        #     self.page.locator('#i0116').fill('chen-fr@cn001.itgr.net')
        #     self.page.locator('#idSIButton9').click()
        #     self.page.locator('#passwordInput').fill('As123123')
        #     self.page.locator('#submitButton').click()
        #     self.page.locator('#idSIButton9').click()
        #
        #     # 等待导航完成
        #     self.page.wait_for_url(url, timeout=600000)
        #     self.msg['info'] = '登录成功'
        # except Exception:
        #     self.msg['info'] = '网页打开出现错误，请稍后重试'
    def login(self, url, userinfo):
        # 登录网址
        self.playwright = sync_playwright().start()
        try:
            # 启动浏览器
            self.browser = self.playwright.chromium.launch(headless=False, executable_path=self.browser_path)
            # 创建新的浏览器上下文
            self.context = self.browser.new_context(user_agent='chen-fr@cn001.itgr.net')
            # 创建新的页面
            self.page = self.context.new_page()
            # 导航到指定URL
            self.page.goto(url)
            # 填写用户名和密码
            self.page.locator('#i0116').fill(userinfo['account'])
            self.page.locator('#idSIButton9').click()
            self.page.locator('#passwordInput').fill(userinfo['password'])
            self.page.locator('#submitButton').click()
            self.page.locator('#idSIButton9').click()
            # 等待导航完成
            self.page.wait_for_url(url, timeout=600000)
            self.msg['info'] = '登录成功'
        except Exception as e:
            self.msg['info'] = '网页打开出现错误，请稍后重试'
            self.msg['error'] = e
            self.msg['flag'] = False

    def add_line(self, data):
        flag = Browser.Verify_main_web()
        if self.msg['flag']:

            self.page.locator('#body_x_tabc_identity_prxidentity_x_proxyItemControl_x_grdItems_grd').click()



    def Verify_main_web(self):
        max_retries = 5
        self.page_tbody = None
        for attempt in range(max_retries):
            try:
                # 定位到表格元素
                self.page_table = self.page.locator('#body_x_tabc_identity_prxidentity_x_proxyItemControl_x_grdItems_grd')
                # 等待表格元素加载完成
                self.page_table.wait_for(timeout=60000)
                # 定位到表格的 tbody 元素
                self.page_tbody = self.page_table.locator('tbody')
                break
            except Exception as e:
                if attempt == max_retries - 1:
                    self.msg['info'] = '当前网络严重拥堵，请稍后重试'
                    self.msg['error'] = e
                    self.msg['flag'] = False

    def test_1(self):
        self.page.page_table = self.page.locator(
            '#body_x_tabc_identity_prxidentity_x_proxyItemControl_x_grdItems_grd')

    def test_2(self):
        self.page.page_table = self.page.locator(
            '#body_x_tabc_identity_prxidentity_x_proxyItemControl_x_grdItems_grd')


if __name__ == "__main__":
    browser_path = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"  # 实际浏览器路径
    output_path = "C:\\Users\\chen-fr\\Desktop\\config\\data"
    file_data = pd.read_excel(r"C:\Users\chen-fr\Downloads\采购.xlsx")
    url = 'https://prince.ivalua.app/page.aspx/zh/ord/basket_manage/124025'
    browser_obj = Browser(browser_path=browser_path, output_path=output_path)
    userinfo ={
        'account': 'chen-fr@cn001.itgr.net',
        'password': 'As123123',
    }
    browser_obj.login(url, userinfo)
    browser_obj.test_1()
    print('aaa')
    browser_obj.test_2()
    print('dddd')
    input("按回车键退出并关闭浏览器...")
    # browser_obj.test2()
    # print(data)