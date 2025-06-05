import time

from playwright.sync_api import sync_playwright
import logging
import pandas as pd
from datetime import datetime
import os
import re

class Browser():
    def __init__(self, browser_path):
        self.browser_path = browser_path
        self.page_title = None
        self.playwright = None
        self.browser = None
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

    # 初始化消息
    def initialize_msg(self):
        msg = {}
        msg['flag'] = True
        msg['error'] = ''
        msg['error_step'] = ''
        msg['info'] = ''
        msg['Table row count'] = ''
        msg['data'] = {
            'request_id': '',
            'Prince Order Number': '',
            'Prince 金额': '',
        }
        return msg

    def login(self, web_url, userinfo):
        # 登录网址
        # 1. 登录成功
        # 2. 网页打开出现错误，请稍后重试
        msg = self.initialize_msg()
        self.playwright = sync_playwright().start()
        try:
            # 启动浏览器
            self.browser = self.playwright.chromium.launch(headless=False, executable_path=self.browser_path)
            # 创建新的浏览器上下文
            self.context = self.browser.new_context(user_agent='chen-fr@cn001.itgr.net')
            # 创建新的页面
            self.page = self.context.new_page()
            # 导航到指定web_url
            self.page.goto(web_url)
            # 填写用户名和密码
            self.page.locator('#i0116').fill(userinfo['Account'])
            self.page.locator('#idSIButton9').click()
            self.page.locator('#passwordInput').fill(userinfo['Password'])
            self.page.locator('#submitButton').click()
            self.page.locator('#idSIButton9').click()
            # 等待导航完成
            self.page.wait_for_url(web_url, timeout=600000)
            msg['info'] = '登录成功'
        except Exception as e:
            msg['info'] = '网页打开出现错误，请稍后重试'
            msg['error'] = e
            msg['flag'] = False
            self.browser.close()

        return msg

    def add_line(self):
        # 验证主页面是否成功加载
        # 1. '复制成功' - 当表格复制操作成功完成
        # 2. '勾选第一行出现错误，当前网络严重拥堵，请稍后重试' - 重试5次后仍无法勾选复选框
        # 3. '表格加载失败，有可能表格未复制成功，也有可能是网络繁忙，请稍后重试' - 30次内表格行数未变化
        # 4. '当前网络严重拥堵，请稍后重试'
        msg = self.verify_main_web()
        try:
            if msg['flag']:
                # 计算老表格的行数
                old_row_count = self.page_tbody.locator('tr').count()
                count = 0
                # 等待加载主页面，并点击第一条line
                while True:
                    try:
                        # 定位到表格元素
                        self.page_table = self.page.locator(
                            '#body_x_tabc_identity_prxidentity_x_proxyItemControl_x_grdItems_grd')
                        # 等待表格元素加载完成
                        self.page_table.wait_for(timeout=60000)
                        # 定位到表格的 tbody 元素
                        self.page_tbody = self.page_table.locator('tbody')
                        # 定位到第一行
                        self.page_tbody_first_row = self.page_tbody.locator('tr:nth-child(1)')
                        # 勾选第一行的复选框
                        self.page_tbody_first_row.locator('td:nth-child(1)').locator(
                            'input[type="checkbox"]').set_checked(True)
                        # 勾选成功，退出循环
                        if self.page_tbody_first_row.locator('td:nth-child(1)').locator(
                                'input[type="checkbox"]').is_checked():
                            break
                    except Exception as e:
                        # 若勾选第一行出现错误，提示重试并记录日志
                        count += 1
                        if count == 5:  # 重试 5 次，超时关闭
                            msg['info'] = '勾选第一行出现错误，当前网络严重拥堵，请稍后重试'
                            msg['error'] = e
                            msg['flag'] = False
                            return msg
                # 定位到复制按钮并点击
                self.btn_copy = self.page.locator(
                    '//*[@id="body_x_tabc_identity_prxidentity_x_proxyItemControl_x_grdItems_btnCopyItem"]')
                self.btn_copy.click()
                # 复制后，等待主页面加载
                msg = self.verify_main_web()
                if msg['flag']:
                    msg = self.verify_table_add_line(old_row_count)
                    if msg['flag']:
                        msg['info'] = '复制成功'
        except Exception as e:
            msg['info'] = '报错了'
            msg['error'] = e
            msg['flag'] = False
        return msg

    def edit_line(self, data):
        # 验证主页面是否成功加载
        # 1. 填写成功
        # 2. 当前网络严重拥堵，编辑表格弹窗未成功加载
        # 3. iframe 弹窗未关闭成功，可能是网络繁忙，请稍后重试
        # 4. iframe弹窗关闭成功
        # 5. 该Order No无法在采购系统中使用,请解锁后重新填写

        row = data
        msg = self.verify_iframe_web()
        try:
            if msg['flag']:
                self.page.wait_for_timeout(5000)
                # 在 iframe 中找到名为 "Account Assignment Category Sales Order" 的下拉框，并填充订单号
                self.page_frame.get_by_role("combobox", name="Account Assignment Category Sales Order").fill(
                    str(row['Order Number']))
                # 等待 5 秒，确保下拉框操作完成
                self.page.wait_for_timeout(5000)
                # 点击该下拉框
                self.page_frame.get_by_role("combobox",
                                       name="Account Assignment Category Sales Order").click()
                # 等待 5 秒，确保下拉框操作完成
                self.page.wait_for_timeout(5000)
                # 使用更精确的下拉菜单定位器
                # self.page_frame.locator("ul.scrolling.menu.visible li.item:first-child")
                self.page.keyboard.press('Enter')
                # # 点击 iframe 中的下拉框第一条数据
                # self.page_frame.locator("div[id^='dropdown-']").first.click()
                # 清空 iframe 中 id 为 'body_x_txtItemLabel' 的输入框
                self.page_frame.locator('#body_x_txtItemLabel').clear()
                # 填充该输入框为检测内容描述
                self.page_frame.locator('#body_x_txtItemLabel').fill(row['检测内容描述'])
                # 清空 iframe 中 id 为 'body_x_txtPrice' 的输入框
                self.page_frame.locator('#body_x_txtPrice').clear()
                # 填充该输入框为未税金额
                rev = round(float(row['未税金额(/1+税点)']), 2)
                self.page_frame.locator('#body_x_txtPrice').fill(str(rev))
                # 点击该输入框
                self.page_frame.locator('#body_x_txtPrice').click()
                # 等待 2 秒，确保输入操作完成
                self.page.wait_for_timeout(2000)
                # 清空 iframe 中标签为 "%" 的输入框
                self.page_frame.locator(
                    'div.data_to_update.percent-allocation-line input[type="text"]'
                ).clear()
                # 填充该输入框为 100
                self.page_frame.locator(
                    'div.data_to_update.percent-allocation-line input[type="text"]'
                ).fill('100')
                # 等待 5 秒，确保输入操作完成
                self.page.wait_for_timeout(5000)
                # 获取下拉框关联的 div 元素的文本内容
                sales_order = self.page_frame.get_by_role("combobox",
                                                     name="Account Assignment Category Sales Order").locator(
                    '+div').inner_text()
                revenue = self.page_frame.locator('#body_x_txtTotalPriceDiscounted').input_value()
                if row['Order Number'] == int(sales_order):
                    # 等待 5 秒，确保操作完成
                    self.page.wait_for_timeout(1000)
                    if msg['flag']:
                        msg['info'] = '填写成功'
                        msg['data'] = {
                            'request_id': self.request_id,
                            'Prince Order Number': sales_order,
                            # 'OPdEX Order Number': row['Order Number'],
                            'Prince 金额': revenue,
                            # 'OPdEX 未税金额(/1+税点)': row['未税金额(/1+税点)'],
                        }
                else:
                    msg['info'] = '该Order No无法在采购系统中使用,请解锁后重新填写'
                    msg['flag'] = False
                    raise msg
        except Exception as e:
            msg['info'] = '报错了'
            msg['error'] = e
            msg['flag'] = False
        return msg

    # 验证主页面是否成功加载
    def refresh(self, old_row_count):
        # 1. 复制成功
        # 2. 当前网络严重拥堵，请稍后重试
        # 3. 表格加载失败，有可能表格未复制成功，也有可能是网络繁忙，请稍后重试
        self.page.reload()
        msg = self.verify_main_web()
        try:
            if msg['flag']:
                msg = self.verify_table_add_line(old_row_count)
        except Exception as e:
            msg['info'] = '报错了'
            msg['error'] = e
            msg['flag'] = False
        return msg

    def close_iframe(self):
        # 1. iframe弹窗关闭成功
        # 2. iframe 弹窗未关闭成功，可能是网络繁忙，请稍后重试

        msg = self.initialize_msg()
        try:
            flag = True
            num = 0
            while flag:
                num += 1
                if num <= 10:
                    # 点击 iframe 中的保存并关闭
                    self.page_frame.locator('//*[@id="proxyActionBar_x__cmdEnd"]').click()
                    msg = self.verify_close_iframe()
                    if msg['flag']:
                        flag = False
                else:
                    msg['info'] = 'iframe 弹窗未关闭成功，可能是网络繁忙，请稍后重试'
                    msg['flag'] = False
                    flag = False
        except Exception as e:
            msg['info'] = '报错了'
            msg['error'] = e
            msg['flag'] = False
        return msg

    def close_browser(self):
        # 关闭浏览器
        msg = self.initialize_msg()
        try:
            self.browser.close()
            self.playwright.stop()
            msg['info'] = '浏览器已关闭'
        except Exception as e:
            msg['info'] = '报错了'
            msg['error'] = e
            msg['flag'] = False
        return msg

    def process_data_flow(self, row_data):
        """
        完整数据流程处理
        :param row_data: 单行数据字典
        :return: 包含全流程状态的 msg 字典
        """
        msg = self.initialize_msg()

        # 步骤1: 添加行
        add_msg = self.add_line()
        if not add_msg['flag']:
            msg.update({
                'error_step': 'add_line',
                'error': add_msg['error'],
                'flag': False,
                'info': f"添加行失败: {add_msg['info']}"
            })
            return msg
        msg['Table row count'] = add_msg['Table row count']

        # 步骤2: 编辑行
        edit_msg = self.edit_line(row_data)
        if not edit_msg['flag']:
            msg.update({
                'error_step': 'edit_line',
                'error': edit_msg['error'],
                'flag': False,
                'info': f"编辑行失败: {edit_msg['info']}"
            })
            return msg

        # 步骤3: 关闭iframe
        close_msg = self.close_iframe()
        if not close_msg['flag']:
            msg.update({
                'error_step': 'close_iframe',
                'error': close_msg['error'],
                'flag': False,
                'info': f"关闭弹窗失败: {close_msg['info']}"
            })
            return msg

        # 全部成功
        msg.update({
            'info': '完整数据流程执行成功',
            'data': edit_msg.get('data', {})
        })
        return msg

    def select_final_delivery(self):
        """批量选中所有最终收货复选框"""
        msg = self.initialize_msg()
        try:
            final_web = self.verify_final_delivery()
            if not final_web['flag']:
                msg.update({
                    'flag': False,
                    'error': final_web['error'],
                    'info': '最终收货页面加载失败',
                    'error_step':'verify_final_delivery'
                })
                return msg
            # 定位所有目标checkbox
            checkboxes = self.page.locator('td[data-iv-role="cell"] input[type="checkbox"][aria-label="最终收货"]')
            
            # 获取总数
            count = checkboxes.count()
            
            # 批量操作
            for i in range(count):
                checkbox = checkboxes.nth(i)
                # 添加延迟逻辑
                if i % 10 == 0:  # 每10个暂停2秒
                    self.page.wait_for_timeout(2000)
                else:  # 其他操作暂停500毫秒
                    self.page.wait_for_timeout(500)
                # 添加状态检查
                if checkbox.is_checked():
                    continue  # 跳过已选中的复选框
                # 通过DOM关系定位关联元素
                checkbox.element_handle().evaluate('''
                    checkbox => {
                        // 查找相邻的隐藏域
                        const hiddenInput = checkbox.nextElementSibling.nextElementSibling;
                        if (hiddenInput && hiddenInput.type === 'hidden') {
                            hiddenInput.value = 'True';
                            hiddenInput.dispatchEvent(new Event('change'));
                        }
                        // 触发点击事件
                        checkbox.click();
                    }
                ''')

            # 验证选中数量
            checked_count = checkboxes.evaluate_all('nodes => nodes.filter(n => n.checked).length')
            if checked_count != count:
                raise Exception(f"成功选中 {checked_count}/{count} 个复选框")
            msg['info'] = f'已成功选中 {checked_count}/{count} 个最终收货复选框'
        except Exception as e:
            msg.update({
                'flag': False,
                'error': str(e),
                'info': '批量选中复选框时发生错误',
                'error_step': 'select_all_final_deliveries'
            })
        return msg

    # 验证主页面是否成功加载
    def verify_main_web(self):
        # 1. 主页加载成功
        # 2. 当前网络严重拥堵，请稍后重试
        msg = self.initialize_msg()
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
                msg['info'] = '主页加载成功'
                break
            except Exception as e:
                if attempt == max_retries - 1:
                    msg['info'] = '当前网络严重拥堵，请稍后重试'
                    msg['error'] = e
                    msg['flag'] = False
        return msg

    def verify_iframe_web(self):
        # 1. iframe页面加载成功
        # 2. 当前网络严重拥堵，编辑表格弹窗未成功加载

        msg = self.initialize_msg()
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
                # 定位到最后一行
                self.page_tbody_last_row = self.page_tbody.locator('tr:nth-last-child(1)')
                # 获取编号
                self.request_id = self.page_tbody_last_row.locator('td:nth-child(5)').inner_text()
                # 点击最后一行的修改链接
                self.page_tbody_last_row.locator('td:nth-child(2)').locator('a').click()
                # 等待 iframe 加载完成
                self.page.wait_for_selector("iframe", timeout=10000)
                # 定位到 iframe
                self.page_frame = self.page.frame_locator("iframe")
                msg['info'] = 'iframe页面加载成功'
                break
            except Exception as e:
                if attempt == max_retries - 1:
                    msg['info'] = '当前网络严重拥堵，编辑表格弹窗未成功加载'
                    msg['error'] = e
                    msg['flag'] = False
        return msg

    def verify_table_add_line(self, old_row_count):
        # 验证表格是否增加成功
        # 1. 表格加载失败，有可能表格未复制成功，也有可能是网络繁忙，请稍后重试
        # 2. 复制成功

        msg = self.initialize_msg()
        flag = True
        num = 0
        while flag:
            num += 1
            new_row_count = self.page_tbody.locator('tr').count()
            # 若新表格行数与老表格行数相同，提示重试并记录日志
            if (new_row_count == old_row_count or new_row_count == 0) and num <= 30:
                time.sleep(1)
                flag = True
            elif num > 30:
                msg['info'] = '表格加载失败，有可能表格未复制成功，也有可能是网络繁忙，请稍后重试'
                msg['Table row count'] = new_row_count
                msg['flag'] = False
                flag = False
            else:
                msg['info'] = '复制成功'
                msg['Table row count'] = new_row_count
                flag = False
        return msg


    def verify_close_iframe(self):
        msg = self.initialize_msg()
        max_retries = 5
        self.page_tbody = None
        for attempt in range(max_retries):
            try:
                self.page.locator('//button[@id="proxyActionBar_x__cmdEnd"]/span[text()="保存并关闭"]')
                msg['info'] = 'iframe弹窗关闭成功'
                break
            except Exception as e:
                if attempt == max_retries - 1:
                    msg['info'] = 'iframe 弹窗未关闭成功，可能是网络繁忙，请稍后重试'
                    msg['error'] = e
                    msg['flag'] = False
        return msg

    def verify_final_delivery(self):
        # 验证最终收货复选框是否选中
        # 1. 最终收货复选框未选中
        # 2. 最终收货复选框已选中
        msg = self.initialize_msg()
        try:
            # 更精确的定位策略
            content_div = self.page.locator(
                'div[id^="body_x_tabc_prxDelivery_prxprxDelivery_x_prxItem_x_gridDeliveryItems_phcgridDeliveryItems_content"]'
            )
            
            # 添加更严格的验证
            if not content_div.get_attribute("id").endswith("_content"):
                raise Exception("定位到非内容容器")
                
            # 等待内容加载完成
            content_div.wait_for(state="visible", timeout=30000)
            
            msg['data']['container'] = content_div
            msg['info'] = '成功定位到动态表格容器'
        
        except Exception as e:
            msg.update({
                'flag': False,
                'error': str(e),
                'info': '定位动态表格容器失败',
                'error_step': 'verify_final_delivery'
            })
        return msg

# if __name__ == "__main__":
#     browser_path = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"  # 实际浏览器路径
#     output_path = "C:\\Users\\chen-fr\\Desktop\\config\\data"
#     file_data = pd.read_excel(r"C:\Users\chen-fr\Downloads\采购.xlsx")
#     web_url = 'https://prince.ivalua.app/page.aspx/zh/ord/basket_manage/124025'
#     browser_obj = Browser(browser_path=browser_path)
#     userinfo ={
#         'Account': 'chen-fr@cn001.itgr.net',
#         'Password': 'As123123',
#     }
#     msg_login = browser_obj.login(web_url, userinfo)
#     print(msg_login)
#     msg_line = browser_obj.add_line()
#     print('msg_line')
#     msg_edit = browser_obj.edit_line(file_data.iloc[1].to_dict())
#     print(msg_edit)
#     print('dddd')
#     msg_close = browser_obj.close_iframe()
#     print(msg_close)
#     browser_obj.close_browser()
#     # input("按回车键退出并关闭浏览器...")
#     # browser_obj.test2()
#     # print(data)
#     # 21-3