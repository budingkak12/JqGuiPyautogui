import pyautogui
import pyperclip
import json
from tkinter import filedialog, Tk, Label, Button, Frame
import pandas as pd
import time
import os
from threading import Thread

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
        self.thread = None  # 新增：用于保存线程对象
        self.row_frame = row_frame
        self.task_index = task_index
        self.config = load_config()

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
        if self.file_path is None:  # 检查是否选择了文件
            self.row_frame.status_label.config(text="请先选择文件")
            return

        # 每次点击“开始”时重新加载Excel文件中的DataFrame
        self.df = pd.read_excel(self.file_path)

        if self.df.empty:  # 检查 DataFrame 是否为空
            self.row_frame.status_label.config(text="文件为空或无有效数据")
            return

        if not self.running:
            self.row_frame.start_button.config(text="终止", state="normal")
            self.row_frame.status_label.config(text="即将开始，倒计时 2 秒...")
            self.running = True
            self.config['tasks'][self.task_index]['status'] = 'running'
            save_config(self.config)

            # 启动倒计时线程
            Thread(target=self.countdown_start).start()
        else:
            self.terminate_execution()

    def countdown_start(self):
        for i in range(2, 0, -1):  # 倒计时2秒
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

    def run_taska(self):
        image_directory = os.path.join(os.path.dirname(self.file_path), "png")

        for _, row in self.df.iterrows():
            if not self.running:  # 检查是否被终止
                print("任务被终止")
                break  # 退出循环，终止任务

            action = row['操作']
            excldel = row['是否逻辑删除']
            if excldel != 1:
                image_file = f"{row['图片名']}.png"
                image_path = os.path.join(image_directory, image_file)

                if action in ['点击', '右击', '双击', '点击并输入']:
                    txt = str(row.get('要输入的内容', ''))
                    print(f"{action} {txt}")
                    perform_action(image_path, image_file, action, txt, self)
                elif action == '输入':
                    pyperclip.copy(str(row['要输入的内容']))
                    print(f"{action} {txt}")

                    pyautogui.hotkey('ctrl', 'v')
                elif action == '按键':
                    print(f"{action} {txt}")

                    pyautogui.hotkey(*row['要输入的内容'].split('+'))
                elif action == '等待':
                    print(f"{action} {txt}")

                    time.sleep(float(row['要输入的内容']))

        # 更新状态标签为“执行完毕”，并设置为绿色
        self.row_frame.status_label.config(text="执行完毕", fg="green")
        self.enable_all_tasks()


    def run_task(self):
        image_directory = os.path.join(os.path.dirname(self.file_path), "png")

        for _, row in self.df.iterrows():
            if not self.running:  # 检查是否被终止
                print("任务被终止")
                break  # 退出循环，终止任务

            action = row['操作']
            excldel = row['是否逻辑删除']

            if excldel != 1:
                image_file = f"{row['图片名']}.png"
                image_path = os.path.join(image_directory, image_file)

                # 执行对应的操作，并打印日志
                if action in ['点击', '右击', '双击', '点击并输入']:
                    txt = str(row.get('要输入的内容', ''))
                    print(f" {action},  '{image_file}'")
                    perform_action(image_path, image_file, action, txt, self)

                elif action == '输入':
                    txt = str(row['要输入的内容'])
                    pyperclip.copy(txt)
                    print(f" {action},  '{txt}'")
                    pyautogui.hotkey('ctrl', 'v')

                elif action == '按键':
                    key_combo = row['要输入的内容'].split('+')
                    print(f" {action},  '{' + '.join(key_combo)}'")
                    pyautogui.hotkey(*key_combo)

                elif action == '等待':
                    wait_time = float(row['要输入的内容'])
                    print(f" {action},  {wait_time} 秒")
                    time.sleep(wait_time)

        # 更新状态标签为“执行完毕”，并设置为绿色
        self.row_frame.status_label.config(text="执行完毕", fg="green")
        self.enable_all_tasks()


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