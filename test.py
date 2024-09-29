import pyautogui
import pyperclip
import json
from tkinter import filedialog, Tk, Label, Button, Frame
import pandas as pd
import time
import os
from threading import Thread
import subprocess
import tempfile
import shutil


# 读取和写入任务状态的配置文件
config_file = 'config.json'


def load_config():
    if os.path.exists(config_file):
        with open(config_file, 'r') as f:
            return json.load(f)
    return {'tasks': [{'file_path': '', 'status': 'idle'} for _ in range(5)]}


def save_config(config):
    with open(config_file, 'w') as f:
        json.dump(config, f, indent=4)


# 全局变量
tasks = []  # 记录所有任务的状态


class Task:
    def __init__(self, row_frame, task_index):
        self.file_path = None
        self.df = None
        self.running = False
        self.thread = None
        self.row_frame = row_frame
        self.task_index = task_index
        self.config = load_config()
        self.temp_dir = None  # 将临时目录初始化为 None

    def select_file(self):
        self.file_path = filedialog.askopenfilename(
            title="请选择 Excel 文件",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if self.file_path:
            self.df = pd.read_excel(self.file_path)
            self.row_frame.file_name_label.config(text=f"文件名: {os.path.basename(self.file_path)}")  # 更新文件名
            self.config['tasks'][self.task_index]['file_path'] = self.file_path
            save_config(self.config)

    def start_execution(self):
        # 确保每次任务开始时创建一个新的临时目录
        self.temp_dir = tempfile.mkdtemp()  # 创建新的临时目录

        # 其他任务启动逻辑保持不变
        if self.file_path is None:
            self.row_frame.status_label.config(text="请先选择文件", fg="black")
            return

        self.df = pd.read_excel(self.file_path)

        if self.df.empty:
            self.row_frame.status_label.config(text="文件为空或无有效数据", fg="black")
            return

        if not self.running:
            self.row_frame.start_button.config(text="终止", state="normal")
            self.row_frame.status_label.config(text="即将开始，倒计时 2 秒...", fg="black")
            self.running = True
            self.config['tasks'][self.task_index]['status'] = 'running'
            save_config(self.config)

            # 启动倒计时线程
            Thread(target=self.countdown_start).start()
        else:
            self.terminate_execution()

    def countdown_start(self):
        for i in range(1, 0, -1):  # 倒计时2秒
            self.row_frame.status_label.config(text=f"即将开始，倒计时 {i} 秒...")
            time.sleep(1)  # 暂停1秒
        self.row_frame.status_label.config(text="开始执行...")
        self.disable_other_tasks()
        # 启动任务线程
        self.thread = Thread(target=self.run_task)
        self.thread.start()

    def terminate_execution(self):
        # 设置标志位为 False，线程将优雅退出
        self.running = False
        self.row_frame.status_label.config(text="任务已终止")
        self.enable_all_tasks()

    def disable_other_tasks(self):
        for task in tasks:
            if task != self:
                task.row_frame.start_button.config(state="disabled")

    def enable_all_tasks(self):
        for task in tasks:
            task.row_frame.start_button.config(state="normal", text="开始")
            task.running = False
            self.config['tasks'][task.task_index]['status'] = 'idle'
            save_config(self.config)

    def run_task(self):
        image_directory = os.path.join(os.path.dirname(self.file_path), "png")

        # 将图片文件复制到临时目录并重命名为英文名
        temp_image_map = self.copy_images_to_temp(image_directory)

        for _, row in self.df.iterrows():
            if not self.running:
                print("任务被终止")
                break

            action = row['操作']
            excldel = row['是否逻辑删除']
            time.sleep(0.3)

            if excldel != 1:
                image_file = f"{row['图片名']}.png"
                image_path = temp_image_map.get(image_file)  # 获取临时路径

                if action in ['点击', '右击', '双击', '点击并输入']:
                    txt = str(row.get('要输入的内容', ''))
                    print(f"{action}, '{image_file}'")
                    perform_action(image_path, image_file, action, txt, self)

                elif action == '输入':
                    txt = str(row['要输入的内容'])
                    pyperclip.copy(txt)
                    print(f" {action},  '{txt}'")
                    pyautogui.hotkey('ctrl', 'v')

                elif action == '按键':
                    key_combo = row['要输入的内容'].split('+')
                    print(f" {action}, '{' + '.join(key_combo)}'")
                    pyautogui.hotkey(*key_combo)

                elif action == '等待':
                    wait_time = float(row['要输入的内容'])
                    print(f" {action},  {wait_time} 秒")
                    time.sleep(wait_time)

                elif action == '代码':
                    program_path = row['要输入的内容']
                    print(f": {action},  '{program_path}'")

                    # 关闭程序
                    subprocess.run(["taskkill", "/F", "/IM", "360ChromeX.exe"], stdout=subprocess.PIPE,
                                   stderr=subprocess.PIPE)

                    # 等待程序完全关闭
                    time.sleep(0.5)

                    # 重新打开程序
                    subprocess.Popen(program_path)
                    print(f"重新打开程序: '{program_path}'")

        # 更新状态标签为“执行完毕”，并设置为绿色
        self.row_frame.status_label.config(text="执行完毕", fg="green")


        # 恢复按钮的初始颜色（默认黑色）
        self.row_frame.start_button.config(fg="black")
        self.enable_all_tasks()
        self.row_frame.status_label.config(text="执行完毕", fg="green")
        self.clean_up_temp()  # 清理临时文件

    def copy_images_to_temp(self, image_directory):
        temp_image_map = {}
        for image_file in os.listdir(image_directory):
            if image_file.endswith(".png"):
                temp_image_name = f"temp_{len(temp_image_map)}.png"
                temp_image_path = os.path.join(self.temp_dir, temp_image_name)
                shutil.copyfile(os.path.join(image_directory, image_file), temp_image_path)
                temp_image_map[image_file] = temp_image_path
        return temp_image_map

    def clean_up_temp(self):
        if self.temp_dir and os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)  # 删除临时目录及其内容
        self.temp_dir = None  # 清理后将临时目录设置为 None

class RowFrame(Frame):
    def __init__(self, master, task):
        super().__init__(master)

        # 文件名标签
        self.file_name_label = Label(self, text="文件名: ")
        self.file_name_label.pack(side="left", padx=5)

        # 状态标签
        self.status_label = Label(self, text="")
        self.status_label.pack(side="left", padx=5)

        self.select_button = Button(self, text="选择文件", command=task.select_file)
        self.select_button.pack(side="left", padx=5)

        self.start_button = Button(self, text="开始", command=task.start_execution)
        self.start_button.pack(side="left", padx=5)

        self.pack(pady=5, padx=5)


# 通用的点击、输入、热键处理函数
def perform_action(image, image_file, action, txt=None, task=None):
    for attempt in range(150):
        if not task.running:  # 检查是否需要终止任务
            print("任务被终止，停止操作")
            return  # 立即退出操作

        try:
            location = pyautogui.locateCenterOnScreen(image, confidence=0.9)
            if location:
                if action == '点击':
                    pyautogui.click(location)
                elif action == '右击':
                    pyautogui.rightClick(location)
                elif action == '双击':
                    pyautogui.doubleClick(location)
                elif action == '点击并输入':
                    pyperclip.copy(txt)
                    pyautogui.click(location)
                    pyautogui.hotkey('ctrl', 'v')
                return
            print(f"第 {attempt + 1} 次尝试：未找到图片: {image}")
        except pyautogui.ImageNotFoundException:
            print(f"第 {attempt + 1} 次尝试：未找到图片（{image_file}）")
        time.sleep(0.1)
    print('未能找到图像，结束尝试')


# 创建 GUI 窗口
root = Tk()
root.title("自动化脚本")
root.geometry("600x400")

# 添加多行任务按钮
config = load_config()
for i in range(len(config['tasks'])):
    task = Task(None, i)
    row_frame = RowFrame(root, task)
    task.row_frame = row_frame
    tasks.append(task)

root.mainloop()
#可以