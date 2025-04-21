import sys
import os
import platform
import json
import keyboard
import signal
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QTextEdit, QSystemTrayIcon, QMenu, QAction, QHBoxLayout, QLabel, QGroupBox, QMessageBox
from PyQt5.QtGui import QIcon, QFont, QPalette, QColor
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QPoint

# 用于检查和修改自启动项的函数
def add_to_autostart():
    # 根据操作系统设置自启动项
    if platform.system() == "Windows":
        # 在Windows上将程序加入到注册表中的启动项
        import winreg as reg
        path = sys.executable  # 当前Python程序的路径
        key = r"Software\Microsoft\Windows\CurrentVersion\Run"
        try:
            reg_key = reg.OpenKey(reg.HKEY_CURRENT_USER, key, 0, reg.KEY_WRITE)
            reg.SetValueEx(reg_key, "MapPasswordApp", 0, reg.REG_SZ, path)
            reg.CloseKey(reg_key)
            print("程序已添加到Windows自启动项")
        except Exception as e:
            print(f"无法添加到自启动项: {e}")
    elif platform.system() == "Darwin":
        # 在macOS上创建plist文件来设置自启动
        pass  # 具体实现需要根据macOS的要求
    else:
        print("不支持的操作系统")

# 爬虫类
class Crawler:
    def __init__(self, url):
        self.url = url

    def fetch_data(self):
        try:
            options = webdriver.ChromeOptions()
            options.add_argument('--headless')  # 无头模式
            options.add_argument('--disable-gpu')
            
            # 使用 WebDriver Manager 安装驱动
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            driver.get(self.url)

            # 等待页面加载并找到地图密码部分
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'overview-bonus-door-container')))
            
            # 获取页面内容
            page_source = driver.page_source
            soup = BeautifulSoup(page_source, 'html.parser')

            # 提取地图密码
            bonus_container = soup.find(id="overview-bonus-door-container")
            if not bonus_container:
                return "未找到地图密码部分"
            
            # 获取所有地图信息卡片
            cards = bonus_container.find_all('div', class_='layui-col-xs6')
            result_data = []
            for card in cards:
                title = card.find('p', class_='overview-bd-t').text.strip()
                password = card.find('p', class_='overview-bd-p').text.strip()
                date = card.find('p', class_='overview-bd-ud').text.strip()
                result_data.append({
                    "title": title,
                    "password": password,
                    "date": date
                })

            driver.quit()
            return result_data
        
        except Exception as e:
            driver.quit()
            return f"发生错误: {e}"

# 爬虫线程类
class FetchDataThread(QThread):
    update_signal = pyqtSignal(list)  # 用于更新UI的信号，传递数据列表
    status_signal = pyqtSignal(str)  # 用于更新状态的信号

    def __init__(self, crawler):
        super().__init__()
        self.crawler = crawler

    def run(self):
        # 更新为“正在抓取中”
        self.status_signal.emit("正在抓取数据中...")
        result_data = self.crawler.fetch_data()
        self.status_signal.emit("")  # 清空状态消息
        self.update_signal.emit(result_data)

# 创建主窗口
class MapPasswordApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("每日地图密码爬取")
        self.setGeometry(100, 100, 500, 100)  # 稍微调整窗口大小

        # 设置窗口为无边框、透明且始终置顶
        self.setWindowFlags(self.windowFlags() | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)

        # 创建布局
        layout = QVBoxLayout()

        # 创建爬虫实例和线程
        self.crawler = Crawler("http://www.kkrb.net")
        self.fetch_thread = FetchDataThread(self.crawler)
        self.fetch_thread.update_signal.connect(self.update_display)
        self.fetch_thread.status_signal.connect(self.update_status)

        # 设置系统托盘图标
        self.tray_icon = QSystemTrayIcon(QIcon("icon.png"), self)
        self.tray_icon.setVisible(True)

        # 创建托盘菜单
        tray_menu = QMenu(self)

        # 创建退出菜单项
        exit_action = QAction("退出", self)
        exit_action.triggered.connect(self.close)
        tray_menu.addAction(exit_action)

        # 创建托盘菜单并显示
        self.tray_icon.setContextMenu(tray_menu)

        # 启用全局按键监听
        self.listen_global_keys()

        # 添加状态标签（显示抓取状态）
        self.status_label = QLabel("正在初始化...", self)
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label)

        # 横向模块布局容器
        self.modules_layout = QHBoxLayout()
        layout.addLayout(self.modules_layout)

        # 设置主窗口布局
        self.setLayout(layout)

        # 判断是否第一次启动
        self.check_first_run()

        # 自动抓取数据
        self.show_window_and_update()

        # 用于拖动窗口的初始化
        self.setMouseTracking(True)
        self.drag_pos = None

        # 设置全局字体
        self.set_global_font()

        # 设置窗口位置为右上角
        screen_geometry = QApplication.desktop().availableGeometry()
        self.move(screen_geometry.right() - self.width(), screen_geometry.top())

    def check_first_run(self):
        # 读取配置文件判断是否第一次启动
        config_file = "config.json"
        if not os.path.exists(config_file):
            # 第一次启动，显示询问框
            self.ask_for_autostart()
        else:
            with open(config_file, 'r') as f:
                config = json.load(f)
                if not config.get('autostart_set', False):
                    self.ask_for_autostart()

    def ask_for_autostart(self):
        # 弹出对话框询问用户是否启用自启动
        reply = QMessageBox.question(self, '设置自启动',
                                     "是否希望程序自启动？",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            add_to_autostart()
            self.save_first_run_config(True)
        else:
            self.save_first_run_config(False)

    def save_first_run_config(self, autostart_set):
        # 保存配置，记录是否启用了自启动
        config = {
            "autostart_set": autostart_set
        }
        with open("config.json", 'w') as f:
            json.dump(config, f)

    def listen_global_keys(self):
        # 使用keyboard库监听Home和End键
        keyboard.add_hotkey('end', self.hide_window)
        keyboard.add_hotkey('home', self.show_window_and_update)
        keyboard.add_hotkey('pagedown', self.close_and_cleanup)  # 监听 PageDown 键

    def hide_window(self):
        self.setVisible(False)  # 隐藏窗口

    def show_window_and_update(self):
        self.setVisible(True)   # 显示窗口
        self.raise_()           # 确保窗口显示在最前面
        print("开始抓取数据...")  # 调试信息
        self.fetch_thread.start()  # 程序打开时自动抓取数据

    def update_status(self, status_text):
        # 更新状态标签
        self.status_label.setText(status_text)

    def update_display(self, result_data):
        # 检查爬取的数据
        print("接收到数据:", result_data)
        if not result_data:
            print("没有抓取到任何数据.")
            self.status_label.setText("未抓取到数据.")
        
        # 清空之前的模块
        for i in reversed(range(self.modules_layout.count())):
            widget = self.modules_layout.itemAt(i).widget()
            if widget:
                widget.deleteLater()

        # 根据抓取的数据动态创建模块
        for data in result_data:
            module = self.create_module(data)
            self.modules_layout.addWidget(module)

    def create_module(self, data):
        # 创建一个模块显示每个数据项
        module = QGroupBox(f"{data['title']}")
        module_layout = QVBoxLayout()

        # 地图密码
        password_label = QLabel(f"{data['password']}")
        password_label.setAlignment(Qt.AlignCenter)  # 设置居中对齐
        module_layout.addWidget(password_label)

        module.setLayout(module_layout)
        module.setAlignment(Qt.AlignCenter)  # 设置模块标题居中
        return module

    def close_and_cleanup(self):
        # 关闭程序并进行清理
        print("关闭程序并清理...")
        self.close()  # 关闭窗口
        QApplication.quit()  # 退出应用
        sys.exit(0)  # 强制退出

    def set_global_font(self):
        # 设置字体为黑体，字号为14
        font = QFont("黑体", 14)  # 设置黑体，字号14
        self.setFont(font)
        # 应用到所有控件
        self.status_label.setFont(font)

        # 给文本框、模块等控件设置字体
        for i in range(self.modules_layout.count()):
            widget = self.modules_layout.itemAt(i).widget()
            if widget:
                widget.setFont(font)

        # 设置字体颜色为白色
        palette = self.palette()
        palette.setColor(QPalette.WindowText, QColor(255, 255, 255))  # 设置文字颜色为白色
        self.setPalette(palette)

# 启动应用
if __name__ == "__main__":
    app = QApplication([])  # 使用默认应用
    window = MapPasswordApp()
    window.show()
    window.show_window_and_update()  # 程序打开时自动抓取数据
    app.exec_()
