from flask import Flask, render_template, jsonify
import win32gui
import win32con
import os
import pyautogui
import time
import re
import logging
import ctypes, sys

current_directory = os.path.dirname(os.path.realpath(sys.argv[0]))
template_path = os.path.join(current_directory, 'templates')
static_path = os.path.join(current_directory, 'static')
app = Flask(__name__,template_folder=template_path,static_folder=static_path)

# 配置日志
def configure_logging():
    # 创建日志记录器
    logger = logging.getLogger('werkzeug')
    logger.setLevel(logging.INFO)  # 设置日志级别

    # 创建文件处理器，并设置级别和格式
    file_handler = logging.FileHandler('app.log')  # 日志文件路径
    file_handler.setLevel(logging.INFO)  # 设置文件处理器的日志级别
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)

    # 添加文件处理器到 Werkzeug 记录器
    logger.addHandler(file_handler)

# 调用配置函数
configure_logging()

def get_m2server_window():
    window_caption_file = "hwnd.txt"
    if os.path.exists(window_caption_file):
        try:
            with open(window_caption_file, "r") as f:
                hwnd = int(f.read().strip())  # 确保转换为整数
                window_title = win32gui.GetWindowText(hwnd)
                # 让用户确认窗口标题是否正确
                print(f"读取到的窗口标题是: {window_title}")
                confirm = input("是否正确？(y/n): ")
                if confirm.lower() == "y":
                    return hwnd
                else:
                    print("用户确认窗口标题不正确，将尝试通过鼠标点击获取窗口句柄。")
                    throw(Exception("用户确认窗口标题不正确"))
        except Exception as e:
            print("读取文件时出现错误:", e)
    else:
        print("窗口标题文件不存在，将尝试通过鼠标点击获取窗口句柄。")

    print("请在3秒内点击M2Server窗口的标题栏...")
    time.sleep(3)
    position = pyautogui.position()
    hwnd = win32gui.WindowFromPoint((position.x, position.y))
    window_title = win32gui.GetWindowText(hwnd)
    print("获取到的窗口标题是:", window_title)

    with open(window_caption_file, "w") as f:  # 移除 encoding="utf-8"，因为写入整数不需要编码
        f.write(str(hwnd))  # 写入编码后的字符串
    return hwnd


def get_m2server_log(callback2):
    def callback(child_hwnd, lParam):
        try:
            # 获取子窗口的类名
            class_name_buf = win32gui.GetClassName(child_hwnd)
            if class_name_buf == "TRichView":
                # 获取窗口文本内容
                length = win32gui.SendMessage(child_hwnd, win32con.WM_GETTEXTLENGTH) *2
                buffer = win32gui.PyMakeBuffer(length)
                win32gui.SendMessage(child_hwnd, win32con.WM_GETTEXT, length, buffer)
                callback2(buffer[:length].tobytes().decode("utf-16", errors='ignore'))
        except Exception as e:
            # 打印异常信息
            print(f"处理窗口 {child_hwnd} 时发生异常: {e}")
            # 可以选择返回None或者继续执行
            return None
    win32gui.EnumChildWindows(hwnd, callback, None)


def reload_menu_option(hwnd, submenu_index, item_index):
    menu = win32gui.GetMenu(hwnd)
    if menu:
        control = win32gui.GetSubMenu(menu, 0)
        reloads = win32gui.GetSubMenu(control, submenu_index)
        item_id = win32gui.GetMenuItemID(reloads, item_index)
        win32gui.PostMessage(hwnd, win32con.WM_COMMAND, item_id, 0)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/getTRichViewContent', methods=['GET'])
def get_trichview_content():
    content = {'data': ''}
    def content_callback(text):
        # 在日期时间前添加换行标签
        text = re.sub(r'(\d{4}/\d{1,2}/\d{1,2} \d{2}:\d{2}:\d{2})', r'<br>\1', text)
        content['data'] = text
    get_m2server_log(content_callback)
    return jsonify({"content": content['data']})


@app.route('/reloadGoos', methods=['POST'])
def reloadGoods():
    reload_menu_option(hwnd, 1, 0)  # For reload_Goods
    return "重载物品操作已触发!"


@app.route('/reloadSkiis', methods=['POST'])
def reloadSkiis():
    reload_menu_option(hwnd, 1, 1)  # For reload_Skiis
    return "重载技能操作已触发!"


@app.route('/reloadMonster', methods=['POST'])
def reloadMonster():
    reload_menu_option(hwnd, 1, 2)
    return "重载怪物操作已触发!"


@app.route('/reloadQManage', methods=['POST'])
def reloadQManage():
    reload_menu_option(hwnd, 1, 11)
    return "重载QManage操作已触发!"


@app.route('/reloadQFunction', methods=['POST'])
def reloadQFunction():
    reload_menu_option(hwnd, 1, 12)
    return "重载QFunction操作已触发!"


@app.route('/reloadNPC', methods=['POST'])
def reloadNPC():
    reload_menu_option(hwnd, 1, 16)
    return "重载NPC操作已触发!"


if __name__ == '__main__':
    print("请使用管理员权限运行此程序!")
    hwnd = get_m2server_window()
    try:
        app.run(host='0.0.0.0', debug=False)
        print("服务器已启动,地址: http://127.0.0.1:5000/")
    except Exception as e:
        print(f"Error: {e}")