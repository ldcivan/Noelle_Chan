import cv2
import pytesseract
import pyautogui
import keyboard
import time
import numpy as np
import win32api
import win32con
import win32gui
import functools
import re
import math
import json
import traceback


def move_window(window_titles, x, y):
    hwnd = None

    # 遍历窗口标题列表，按顺序查找窗口句柄
    for title in window_titles:
        hwnd = win32gui.FindWindow(None, title)
        if hwnd:
            break

    # 如果找到窗口句柄，则移动窗口到指定位置
    if hwnd:
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x, y, 0, 0, win32con.SWP_NOSIZE)
        win32gui.SetForegroundWindow(hwnd)
        return title
    else:
        print("未找到窗口", title)


def move_admin_cmd_window(window_width, window_height):
    def enum_windows_callback(hwnd, _, window_width, window_height):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetClassName(hwnd) == 'ConsoleWindowClass':
            # 检查窗口标题是否包含"管理员"字样
            window_title = win32gui.GetWindowText(hwnd)
            if '管理员' in window_title:
                # 将窗口移动
                print(window_title)
                move_window([window_title], window_width, window_height)
                # return False  # 停止遍历窗口

    # 遍历所有窗口
    callback = functools.partial(enum_windows_callback, window_width=window_width, window_height=window_height)
    win32gui.EnumWindows(callback, None)


def get_window_size(window_title):
    hwnd = win32gui.FindWindow(None, window_title)
    rect = win32gui.GetWindowRect(hwnd)
    width = rect[2] - rect[0]
    height = rect[3] - rect[1]
    return width, height

# 定义窗口标题列表
window_titles = ["原神", "Genshin", "genshin", "Genshin Impact", "genshin impact", "Yuanshen", "yuanshen"]

# 调用函数移动窗口并获取窗口大小
window_title = move_window(window_titles, -10, -10)
window_width, window_height = get_window_size(window_title)
print(window_title, "窗口大小：", window_width, "x", window_height)

# 移动cmd窗口
move_admin_cmd_window(window_width+20, 0)

# 设置Tesseract的路径（如果你的Tesseract安装在非默认路径下，请修改此处）
pytesseract.pytesseract.tesseract_cmd = 'D:/Program Files/Tesseract-OCR/tesseract.exe'

# 设置聊天区域的坐标（根据你的窗口大小和位置进行调整）
chatbox_coordinates = (0.225*window_width, 0, 0.48*window_width, window_height*0.9)  # (左上角x坐标, 左上角y坐标, 宽度, 高度)


def get_latest_chat():
    # 截取聊天区域的图像
    screenshot = pyautogui.screenshot()
    chatbox_image = screenshot.crop(chatbox_coordinates)
    chatbox_image.save('screenshot.png')

    # 将图像转换为灰度图像
    chatbox_image_gray = cv2.cvtColor(np.array(chatbox_image), cv2.COLOR_RGB2GRAY)

    # 使用Tesseract进行文本识别
    text = pytesseract.image_to_string(chatbox_image_gray, lang='chi_sim')

    # 按行分割文本，获取最新的一行
    lines = text.split('\n')
    index = -1
    latest_chat = lines[index].strip()
    while latest_chat == '':
        index = index-1
        try:
            latest_chat = lines[index].strip()
        except:
            print("未检测到聊天区，尝试定位到聊天区")
            init()


    return latest_chat


def send_response(response):
    # 模拟按下Enter键开始输入
    keyboard.press_and_release('enter')

    # 输入内容
    time.sleep(1)
    keyboard.write(response)

    # 模拟按下Enter键发送内容
    time.sleep(1)
    keyboard.press_and_release('enter')

    print('已发送：' + response)


def click_pic(target_pic):
    time.sleep(2)
    # 在屏幕上查找给定图像
    position = pyautogui.locateOnScreen(target_pic, confidence=0.8)

    # 如果找到了图像
    if position is not None:
        # 获取图像的中心坐标
        x, y, width, height = position
        target_x = x + width / 2
        target_y = y + height / 2

        # 点击图像的中心位置
        pyautogui.click(target_x, target_y)
    else:
        print("未找到图像")


# 加载玩家编号图标和踢出按钮的图像
player_icon_imgs = [
    cv2.imread('./src/1p_icon.png'),
    cv2.imread('./src/2p_icon.png'),
    cv2.imread('./src/3p_icon.png'),
    cv2.imread('./src/4p_icon.png')
]
kick_button_img = cv2.imread('./src/kick_icon.png')


def locate_player_icon(player_icon_img):
    time.sleep(2)
    # 截取屏幕图像
    screenshot = pyautogui.screenshot()
    screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

    # 在屏幕图像中搜索玩家编号图标
    result = cv2.matchTemplate(screenshot, player_icon_img, cv2.TM_CCOEFF_NORMED)
    _, max_val, _, max_loc = cv2.minMaxLoc(result)

    if max_val >= 0.5:
        # 返回玩家编号图标的位置
        return max_loc
    else:
        # 未找到玩家编号图标
        return None

def locate_kick_button(player_icon_loc):
    time.sleep(1)
    # 截取屏幕图像
    screenshot = pyautogui.screenshot()
    screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

    # 在屏幕图像中搜索踢出按钮
    result = cv2.matchTemplate(screenshot, kick_button_img, cv2.TM_CCOEFF_NORMED)
    _, max_val, _, max_loc = cv2.minMaxLoc(result)

    if max_val >= 0.8:
        # 将生成器对象转换为列表
        kick_button_locs = list(pyautogui.locateAllOnScreen('./src/kick_icon.png', confidence=0.8))

        if len(kick_button_locs) == 0:
            # 未找到踢出按钮
            return None

        # 计算踢出按钮与玩家编号图标之间的距离
        distances = []
        for loc in kick_button_locs:
            distance = math.sqrt((loc[1] - player_icon_loc[1]) ** 2)
            distances.append(distance)

        # 找到距离最近的踢出按钮位置
        closest_index = distances.index(min(distances))
        closest_loc = kick_button_locs[closest_index]

        # 计算踢出按钮的中心位置
        kick_button_loc = (closest_loc[0] + kick_button_img.shape[1] // 2, closest_loc[1] + kick_button_img.shape[0] // 2)

        # 返回踢出按钮的位置
        return kick_button_loc
    else:
        # 未找到踢出按钮
        return None

def kick_out_player(player_number):
    if player_number < 2 or player_number > len(player_icon_imgs):
        print("无效的玩家编号")
        return

    keyboard.press_and_release('esc')
    time.sleep(2)
    keyboard.press_and_release('esc')
    time.sleep(2)
    keyboard.press_and_release('esc')
    click_pic('./src/multi.png')

    # 定位玩家编号图标的位置
    player_icon_img = player_icon_imgs[player_number - 1]
    player_icon_loc = locate_player_icon(player_icon_img)
    if player_icon_loc is None:
        print("未找到玩家编号图标")
        return

    # 计算玩家编号图标所在行的踢出按钮位置
    kick_button_loc = locate_kick_button(player_icon_loc)

    # 点击踢出按钮
    pyautogui.click(kick_button_loc)
    time.sleep(1)
    click_pic('./src/confirm.png')

    time.sleep(3)
    keyboard.press_and_release('esc')
    click_pic('./src/team.png')
    time.sleep(5)
    keyboard.press_and_release('enter')


def add_friend():
    keyboard.press_and_release('esc')
    time.sleep(2)
    keyboard.press_and_release('esc')
    time.sleep(2)
    keyboard.press_and_release('esc')
    click_pic('./src/friend.png')
    click_pic('./src/add_friend.png')
    click = 0
    while click <= 5:
        try:
            click_pic('./src/tick.png')
        except:
            pass
        click = click + 1
    time.sleep(3)
    keyboard.press_and_release('esc')
    click_pic('./src/team.png')
    time.sleep(5)
    keyboard.press_and_release('enter')


# 添加项目到JSON文件
def add_item(time_now, item_name, location):
    data = {
        'time': time_now,
        'item_name': item_name,
        'location': location
    }

    # 读取现有的JSON数据集
    with open('./data/items.json', 'r') as file:
        json_data = json.load(file)

    # 将新数据添加到JSON数据集中
    json_data.append(data)

    # 将更新后的JSON数据集写入文件
    with open('./data/items.json', 'w') as file:
        json.dump(json_data, file)

# 读取最近两天的项目名称
def read_item():
    current_time = time.time()

    with open('./data/items.json', 'r') as file:
        data = json.load(file)
        items = []
        for line in data:
            item_time = line['time']
            item_name = line['item_name']
            location = line['location']

            # 检查项目是否在最近两天内
            if current_time - item_time <= 2 * 24 * 60 * 60:
                items.append([item_name, location])

        return items


def main():
    # 设置上次检测到的最新聊天内容
    last_chat = ""

    while True:
        win32gui.SetForegroundWindow(win32gui.FindWindow(None, window_title))
        time.sleep(0.5)
        # 获取最新的聊天内容
        latest_chat = get_latest_chat()

        # 如果有新消息且不同于上次的内容
        if latest_chat and latest_chat != last_chat:
            print("最新的聊天内容：", latest_chat)

            if '进入了大世界' in latest_chat:
                time.sleep(7)
                send_response('【诺艾尔】欢迎来到本世界，本世界由见习骑士诺艾尔，也就是在下看管。')
                send_response('【诺艾尔】想知道诺艾尔能为您做什么，请发送“#帮助”！')
                send_response('【诺艾尔】采集材料时请告诉在下“#在何处采集了xx”，如“#在枫丹水下采集了海露花”。')
                send_response('【诺艾尔】在告诉在下要采集了什么，您就可随意获取大世界资源了！祝您游戏愉快！')
            elif '#加好友' in latest_chat:
                send_response('【诺艾尔】请您稍后！')
                add_friend()
                time.sleep(2)
                send_response('【诺艾尔】已经操作完毕了，请检查是否添加；若没能成功，请等待主人手动添加。')
            elif re.search(r"#请离玩家(\d+)", latest_chat):
                player_number = int(re.search(r"#请离玩家(\d+)", latest_chat).group(1))
                send_response('【诺艾尔】正在踢除玩家' + str(player_number) + '，但若发现恶意使用该功能，使用者将被纳入黑名单。')
                kick_out_player(player_number)
                time.sleep(2)
                send_response('【诺艾尔】已经操作完毕了。')
            elif re.findall(r"#在(.*)采集了(.*)", latest_chat):
                location = str(re.findall(r"#在(.*)采集了(.*)", latest_chat)[0][0])
                item_name = str(re.findall(r"#在(.*)采集了(.*)", latest_chat)[0][1])
                add_item(time.time(), item_name, location)
                time.sleep(1)
                send_response('【诺艾尔】谢谢，已经记录好了！')
            elif '#已采集名单' in latest_chat:
                results = read_item()
                for result in results:
                    send_response(f'【诺艾尔】{result[1]}的{result[0]}已被采集。')
                send_response('【诺艾尔】报告完毕！')
            elif '#帮助' in latest_chat:
                send_response('【诺艾尔】需要知道哪些材料被采集过，请输入“#已采集名单”。')
                send_response('【诺艾尔】需要添加主人好友的话，请在发出请求后输入“#加好友”，在下会帮您添加。')
                send_response('【诺艾尔】需要踢掉某位恶意玩家的话，请输入“#请离'
                              '玩家2/3/4”，在下会帮您踢出。')
                send_response('【诺艾尔】不要滥用功能，不然诺艾尔会生气的！')

            # 更新上次检测到的最新聊天内容
            last_chat = latest_chat

        # 等待一段时间后再进行下一次检测
        time.sleep(3)

def init():
    while not pyautogui.locateOnScreen('./src/paimon.png', confidence=0.8):
        keyboard.press_and_release('esc')
        time.sleep(1)
    time.sleep(3)
    keyboard.press_and_release('esc')
    click_pic('./src/team.png')
    time.sleep(5)
    keyboard.press_and_release('enter')
    time.sleep(3)
    print('已重载')
    try:
        main()
    except:
        traceback.print_exc()
        init()


try:
    main()
except:
    traceback.print_exc()
    init()
