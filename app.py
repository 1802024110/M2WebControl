from flask import Flask, render_template, jsonify
import win32gui
import win32con
import os
import pyautogui
import time
import re
import logging
import ctypes
import sys
import threading

current_directory = os.path.dirname(os.path.realpath(sys.argv[0]))
template_path = os.path.join(current_directory, 'templates')
static_path = os.path.join(current_directory, 'static')
app = Flask(__name__, template_folder=template_path, static_folder=static_path)

# 检查是否有管理员权限
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

# 如果没有管理员权限，则重新以管理员权限运行
if not is_admin():
    print("没有管理员权限，尝试以管理员权限重新运行脚本...")
    script = sys.argv[0]
    params = ' '.join([script] + sys.argv[1:])
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)
    sys.exit(0)

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

class HwndManager:
    def __init__(self):
        self._hwnd = None
        self._lock = threading.Lock()
    
    def set_hwnd(self, hwnd):
        with self._lock:
            self._hwnd = hwnd
    
    def get_hwnd(self):
        with self._lock:
            return self._hwnd


def get_m2server_window():
    window_caption_file = "./hwnd.txt"
    if os.path.exists(window_caption_file):
        # 检查文件是否存在，不存在则创建
        if not os.path.isfile(window_caption_file):
            open(window_caption_file, 'w').close()
        with open(window_caption_file, "r") as f:
            windonws_title = f.read().strip()
            if windonws_title == "" or windonws_title is None:
                print("请在5秒内点击M2Server窗口的标题栏,不要使用嵌入到控制台中显示，会导致找不到窗口：")
                time.sleep(5)
                position = pyautogui.position()
                hwnd = win32gui.WindowFromPoint((position.x, position.y))
                window_title = win32gui.GetWindowText(hwnd)
                print("获取到的窗口标题是:", window_title)
                with open(window_caption_file, "w") as f:
                    f.write(window_title)
                return hwnd
            hwnds = win32gui.FindWindow(None, windonws_title)
            if hwnds:
                return hwnds

def get_m2server_log(callback2):
    def callback(child_hwnd, lParam):
        try:
            class_name_buf = win32gui.GetClassName(child_hwnd)
            if class_name_buf == "TRichView":
                length = win32gui.SendMessage(child_hwnd, win32con.WM_GETTEXTLENGTH) * 2
                buffer = win32gui.PyMakeBuffer(length)
                win32gui.SendMessage(child_hwnd, win32con.WM_GETTEXT, length, buffer)
                callback2(buffer[:length].tobytes().decode("utf-16", errors='ignore'))
        except Exception as e:
            print(f"处理窗口 {child_hwnd} 时发生异常: {e}")
            return None
    win32gui.EnumChildWindows(hwnd_manager.get_hwnd(), callback, None)


def reload_menu_option(hwnd_manager, submenu_index, item_index):
    hwnd = hwnd_manager.get_hwnd()
    if hwnd:
        menu = win32gui.GetMenu(hwnd)
        if menu:
            control = win32gui.GetSubMenu(menu, 0)
            reloads = win32gui.GetSubMenu(control, submenu_index)
            item_id = win32gui.GetMenuItemID(reloads, item_index)
            win32gui.PostMessage(hwnd, win32con.WM_COMMAND, item_id, 0)

def check_hwnd_validity(hwnd_manager):
    while True:
        hwnd = hwnd_manager.get_hwnd()
        if not win32gui.IsWindow(hwnd):
            print("检测到窗口无效，重新获取窗口句柄...")
            hwnd = get_m2server_window()
            if hwnd:
                hwnd_manager.set_hwnd(hwnd)
                print(f"重新获取到的窗口标题是: <{win32gui.GetWindowText(hwnd)}>，若窗口标题不对请删除hwnd.txt文件后重新运行程序")
        time.sleep(5)  # 每5秒检查一次窗口有效性


@app.route('/')
def index():
    return render_template('index.html')

@app.route('/getTRichViewContent', methods=['GET'])
def get_trichview_content():
    content = {'data': ''}
    def content_callback(text):
        text = re.sub(r'(\d{4}/\d{1,2}/\d{1,2} \d{2}:\d{2}:\d{2})', r'<br>\1', text)
        content['data'] = text
    get_m2server_log(content_callback)
    return jsonify({"content": content['data']})

@app.route('/reloadGoos', methods=['POST'])
def reloadGoods():
    reload_menu_option(hwnd_manager, 1, 0)
    return "重载物品操作已触发!"

@app.route('/reloadSkiis', methods=['POST'])
def reloadSkiis():
    reload_menu_option(hwnd_manager, 1, 1)
    return "重载技能操作已触发!"

@app.route('/reloadMonster', methods=['POST'])
def reloadMonster():
    reload_menu_option(hwnd_manager, 1, 2)
    return "重载怪物操作已触发!"

@app.route('/reloadQManage', methods=['POST'])
def reloadQManage():
    reload_menu_option(hwnd_manager, 1, 11)
    return "重载QManage操作已触发!"

@app.route('/reloadQFunction', methods=['POST'])
def reloadQFunction():
    reload_menu_option(hwnd_manager, 1, 12)
    return "重载QFunction操作已触发!"

@app.route('/reloadNPC', methods=['POST'])
def reloadNPC():
    reload_menu_option(hwnd_manager, 1, 16)
    return "重载NPC操作已触发!"

def run_server():
    app.run(host='0.0.0.0', debug=False)


if __name__ == '__main__':
    hwnd_manager = HwndManager()
    hwnd = get_m2server_window()
    hwnd_manager.set_hwnd(hwnd)
    print(f"获取到的窗口标题是: <{win32gui.GetWindowText(hwnd)}>, 若窗口标题不对请删除hwnd.txt文件后重新运行程序")
    print("服务器已启动, 地址: http://127.0.0.1:5000/")
  
    server_thread = threading.Thread(target=run_server)
    server_thread.start()

    check_hwnd_thread = threading.Thread(target=check_hwnd_validity, args=(hwnd_manager,))
    check_hwnd_thread.start()