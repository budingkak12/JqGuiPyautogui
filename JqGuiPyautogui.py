import pyautogui
import pyperclip
import subprocess
from tkinter import filedialog, Tk, Label, Button
import pandas as pd
import time
import os
from pyautogui import ImageNotFoundException

# 全局变量，用于存储文件路径和数据
file_path = None
df = None


# 选择 Excel 文件的函数
def select_file():
    global file_path, df
    file_path = filedialog.askopenfilename(
        title="请选择 Excel 文件",
        filetypes=[("Excel files", "*.xlsx")],
        initialdir="/"
    )
    if file_path:
        df = pd.read_excel(file_path)
        label.config(text=f"已选择文件: {os.path.basename(file_path)}")


# 执行动作的函数，增加 image_file 参数
def perform_action(image, image_file, action, txt=None):
    max_attempts = 150  # 最大尝试次数
    for attempt in range(max_attempts):
        try:
            # 使用完整路径来查找图片
            location = pyautogui.locateCenterOnScreen(image, confidence=0.9)
            if location:
                if action == '点击':
                    print(f"{action} {image_file}")
                    pyautogui.click(location)
                elif action == '右击':
                    print(f"{action} {image_file}")
                    pyautogui.rightClick(location)
                elif action == '双击':
                    print(f"{action} {image_file}")
                    pyautogui.doubleClick(location, interval=0.25)
                elif action == '移动':
                    pyautogui.moveTo(location)
                elif action == '点击并输入':
                    print(f"点击{image_file},输入{txt}")
                    pyperclip.copy(txt)
                    pyautogui.click(location)
                    # pyautogui.write(txt)
                    pyautogui.hotkey('ctrl', 'v')


                    # txt = str(row['要输入的内容'])
                    # print(f"{action} {txt}")
                    # pyperclip.copy(txt)
                    # pyautogui.hotkey('ctrl', 'v')
                else:
                    print(f"不支持的操作: {action}")
                return  # 找到图像后退出函数
            else:
                print(f"第 {attempt + 1} 次尝试：未找到图片: {image}")
        except ImageNotFoundException:
            print(f"第 {attempt + 1} 次尝试：未找到图片（{image_file}）")
        time.sleep(0.1)  # 等待0.1秒再重试
    print('未能找到图像，结束尝试')


# 开始执行的函数
def start_execution():
    # time.sleep(2)
    if df is None:
        label.config(text="请先选择文件")
        return

    # 倒计时逻辑
    for i in range(3, 0, -1):
        label.config(text=f"即将开始，倒计时 {i} 秒")
        root.update()  # 更新界面以显示倒计时
        time.sleep(1)

    label.config(text="开始执行...")  # 更新倒计时结束后的提示

    image_directory = os.path.join(os.path.dirname(file_path), "png").replace("/", "\\")

    for index, row in df.iterrows():
        action = row['操作']  # 获取操作类型
        excldel = row['是否逻辑删除']  # 获取是否逻辑删除
        image_file = str(row['图片名']) + '.png'  # 获取图片名并添加后缀
        image_path = os.path.join(image_directory, image_file)  # 获取图片的完整路径

        # 判断是否逻辑删除
        if excldel != 1:
            time.sleep(0.1)
            if action in ['点击', '右击', '双击']:
                perform_action(image_path, image_file, action)
            elif action == '点击并输入':
                try:
                    txt = str(int(row['要输入的内容']))  # 尝试将内容转换为整数再转为字符串
                except ValueError:
                    txt = str(row['要输入的内容'])  # 如果无法转换为整数，则直接作为字符串处理
                perform_action(image_path, image_file, action, txt)
            elif action == '输入':
                txt = str(row['要输入的内容'])
                print(f"{action} {txt}")
                pyperclip.copy(txt)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(0.4)
            elif action == '按键':
                txt = row['要输入的内容']  # 获取要输入的内容
                print(f"{action} {txt}")
                keys = txt.split('+')  # 按 '+' 分割为单独的按键列表
                pyautogui.hotkey(*keys)
                time.sleep(0.5)
            elif action == '等待':
                txt = float(row['要输入的内容'])  # 获取等待时间
                printtxt = str(int(row['要输入的内容']))  # 尝试将内容转换为整数再转为字符串
                print(f"{action} {txt}秒")
                time.sleep(txt)
            else:
                pass  # 如果action不是上述任何一种，不执行任何操作
        else:
            pass  # 如果excldel为0，不执行任何操作
    print('执行完成')

# 创建 GUI 窗口
root = Tk()
root.title("自动化脚本")

# 设置窗口大小
window_width = 300
window_height = 150

# 获取屏幕的宽度和高度
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# 计算窗口位置：右上角
# x = screen_width - window_width
x = 0
y = 0  # 顶部对齐

# 设置窗口的位置和大小
root.geometry(f"{window_width}x{window_height}+{x}+{y}")

# 创建 GUI 元素
label = Label(root, text="请选择一个 Excel 文件")
label.pack(pady=10)

select_button = Button(root, text="选择文件", command=select_file)
select_button.pack(pady=5)

start_button = Button(root, text="开始", command=start_execution)
start_button.pack(pady=5)

# 运行主循环
root.mainloop()
