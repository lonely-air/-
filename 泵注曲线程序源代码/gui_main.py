import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
from PIL import Image, ImageTk # 需要安装 Pillow 库：pip install Pillow
import os
import threading
import glob
import time
import sys # 导入 sys 模块

# 导入 main.py 和 mymodel_v1_8.py 中的功能
# 确保 mymodel_v1_8.py 和 main.py 在同一目录下
from main import process_pump_injection_curve, process_pressure_flow_ratio, process_shutdown_curve

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("压裂数据处理系统")
        self.geometry("1200x700") # 调整窗口大小以容纳左右两边

        # 配置网格布局，使其在窗口大小改变时能够扩展
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1) # 左侧控制区
        self.grid_columnconfigure(1, weight=2) # 右侧图片显示区

        # --- 左侧控制区 ---
        self.control_frame = ttk.Frame(self, padding="10")
        self.control_frame.grid(row=0, column=0, sticky="nsew")
        self.control_frame.grid_rowconfigure(5, weight=1) # 让日志文本框可以扩展

        # 按钮
        self.btn_pump_injection = ttk.Button(self.control_frame, text="处理泵注曲线", command=lambda: self.run_function_in_thread(self.process_pump_injection_curve_gui))
        self.btn_pump_injection.grid(row=0, column=0, pady=5, sticky="ew")

        self.btn_pressure_flow = ttk.Button(self.control_frame, text="处理压排比", command=lambda: self.run_function_in_thread(self.process_pressure_flow_ratio_gui))
        self.btn_pressure_flow.grid(row=1, column=0, pady=5, sticky="ew")

        self.btn_shutdown_curve = ttk.Button(self.control_frame, text="处理停泵曲线", command=lambda: self.run_function_in_thread(self.process_shutdown_curve_gui))
        self.btn_shutdown_curve.grid(row=2, column=0, pady=5, sticky="ew")

        self.btn_full_process = ttk.Button(self.control_frame, text="全流程处理", command=lambda: self.run_function_in_thread(self.process_full_gui))
        self.btn_full_process.grid(row=3, column=0, pady=5, sticky="ew")

        self.btn_exit = ttk.Button(self.control_frame, text="退出程序", command=self.quit_app)
        self.btn_exit.grid(row=4, column=0, pady=5, sticky="ew")

        # 日志输出区域
        self.log_text = scrolledtext.ScrolledText(self.control_frame, wrap=tk.WORD, height=15)
        self.log_text.grid(row=5, column=0, pady=10, sticky="nsew")
        self.log_text.config(state=tk.DISABLED) # 初始设置为只读

        # 重定向 print 输出到日志文本框
        self.console = ConsoleRedirector(self.log_text)
        sys.stdout = self.console
        sys.stderr = self.console

        # --- 右侧图片显示区 ---
        self.image_frame = ttk.Frame(self, padding="10")
        self.image_frame.grid(row=0, column=1, sticky="nsew")
        self.image_frame.grid_rowconfigure(0, weight=1) # 图片画布可以扩展
        self.image_frame.grid_columnconfigure(0, weight=1)

        self.image_canvas = tk.Canvas(self.image_frame, bg="lightgray")
        self.image_canvas.grid(row=0, column=0, sticky="nsew")

        # 图片选择下拉框
        self.image_files = []
        self.image_combobox = ttk.Combobox(self.image_frame, values=self.image_files, state="readonly")
        self.image_combobox.grid(row=1, column=0, pady=10, sticky="ew")
        self.image_combobox.bind("<<ComboboxSelected>>", self.display_selected_image)

        self.current_image_index = -1
        self.update_image_list() # 初始化图片列表

    def run_function_in_thread(self, func):
        """在单独线程中运行函数，避免GUI卡顿"""
        self.log_text.config(state=tk.NORMAL) # 允许写入
        self.log_text.delete(1.0, tk.END) # 清空日志
        self.log_text.config(state=tk.DISABLED) # 再次设置为只读
        self.set_buttons_state(tk.DISABLED)
        self.log_message(f"开始执行：{func.__name__.replace('_gui', '')}...")
        thread = threading.Thread(target=func)
        thread.start()

    def set_buttons_state(self, state):
        """设置所有功能按钮的状态"""
        self.btn_pump_injection.config(state=state)
        self.btn_pressure_flow.config(state=state)
        self.btn_shutdown_curve.config(state=state)
        self.btn_full_process.config(state=state)
        # 退出按钮始终可用
        # self.btn_exit.config(state=state)

    def log_message(self, message):
        """将消息输出到日志文本框"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END) # 滚动到最新消息
        self.log_text.config(state=tk.DISABLED)

    def update_image_list(self):
        """更新图片下拉框列表并显示最新图片"""
        figs_dir = "figs"
        if not os.path.exists(figs_dir):
            os.makedirs(figs_dir) # 确保 figs 目录存在
            self.image_files = []
        else:
            # 获取 figs 目录下所有 png 图片
            self.image_files = sorted(glob.glob(os.path.join(figs_dir, "*.png")))
            # 只保留文件名部分
            self.image_files = [os.path.basename(f) for f in self.image_files]

        self.image_combobox['values'] = self.image_files
        if self.image_files:
            # 默认选中并显示最新生成的图片（按文件名排序的最后一张）
            self.image_combobox.set(self.image_files[-1])
            self.display_image(os.path.join(figs_dir, self.image_files[-1]))
        else:
            self.image_combobox.set("") # 清空选择
            self.image_canvas.delete("all") # 清空画布

    def display_selected_image(self, event=None):
        """显示下拉框中选中的图片"""
        selected_image_name = self.image_combobox.get()
        if selected_image_name:
            image_path = os.path.join("figs", selected_image_name)
            self.display_image(image_path)

    def display_image(self, image_path):
        """在画布上显示指定路径的图片"""
        try:
            img = Image.open(image_path)
            # 调整图片大小以适应画布
            canvas_width = self.image_canvas.winfo_width()
            canvas_height = self.image_canvas.winfo_height()

            if canvas_width == 1 or canvas_height == 1: # 初始winfo_width/height可能为1，需要等待更新
                self.after(100, lambda: self.display_image(image_path))
                return

            img_width, img_height = img.size
            
            # 计算缩放比例
            ratio_w = canvas_width / img_width
            ratio_h = canvas_height / img_height
            ratio = min(ratio_w, ratio_h)

            new_width = int(img_width * ratio)
            new_height = int(img_height * ratio)

            img = img.resize((new_width, new_height), Image.LANCZOS)
            self.photo = ImageTk.PhotoImage(img)
            
            self.image_canvas.delete("all") # 清除旧图片
            # 将图片放置在画布中心
            x = (canvas_width - new_width) / 2
            y = (canvas_height - new_height) / 2
            self.image_canvas.create_image(x, y, anchor=tk.NW, image=self.photo)
        except FileNotFoundError:
            self.log_message(f"错误：图片文件未找到 - {image_path}")
            self.image_canvas.delete("all")
        except Exception as e:
            self.log_message(f"显示图片时发生错误：{e}")
            self.image_canvas.delete("all")

    # 包装 main.py 中的功能，以便在GUI中调用并处理完成后的逻辑
    def process_pump_injection_curve_gui(self):
        try:
            process_pump_injection_curve()
            self.log_message("泵注曲线处理完成！")
        except Exception as e:
            self.log_message(f"泵注曲线处理失败：{e}")
            messagebox.showerror("错误", f"泵注曲线处理失败：{e}\n请检查输入文件是否存在且格式正确！")
        finally:
            self.update_image_list()
            self.set_buttons_state(tk.NORMAL)

    def process_pressure_flow_ratio_gui(self):
        try:
            process_pressure_flow_ratio()
            self.log_message("压排比处理完成！")
        except Exception as e:
            self.log_message(f"压排比处理失败：{e}")
            messagebox.showerror("错误", f"压排比处理失败：{e}\n请检查输入文件是否存在且格式正确！")
        finally:
            self.update_image_list()
            self.set_buttons_state(tk.NORMAL)

    def process_shutdown_curve_gui(self):
        try:
            process_shutdown_curve()
            self.log_message("停泵曲线处理完成！")
        except Exception as e:
            self.log_message(f"停泵曲线处理失败：{e}")
            messagebox.showerror("错误", f"停泵曲线处理失败：{e}\n请检查输入文件是否存在且格式正确！")
        finally:
            self.update_image_list()
            self.set_buttons_state(tk.NORMAL)

    def process_full_gui(self):
        try:
            self.log_message("开始全流程处理...")
            process_pump_injection_curve()
            process_pressure_flow_ratio()
            process_shutdown_curve()
            self.log_message("全流程处理完成！")
        except Exception as e:
            self.log_message(f"全流程处理失败：{e}")
            messagebox.showerror("错误", f"全流程处理失败：{e}\n请检查输入文件是否存在且格式正确！")
        finally:
            self.update_image_list()
            self.set_buttons_state(tk.NORMAL)

    def quit_app(self):
        """退出应用程序"""
        self.log_message("程序已退出。")
        self.destroy()
        sys.exit() # 确保完全退出

class ConsoleRedirector:
    """将控制台输出重定向到Tkinter文本框"""
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.stdout = sys.__stdout__ # 保存原始stdout
        self.stderr = sys.__stderr__ # 保存原始stderr

    def write(self, message):
        self.text_widget.config(state=tk.NORMAL)
        self.text_widget.insert(tk.END, message)
        self.text_widget.see(tk.END)
        self.text_widget.config(state=tk.DISABLED)
        self.stdout.write(message) # 同时输出到原始控制台
        self.stdout.flush()

    def flush(self):
        self.stdout.flush()

if __name__ == "__main__":
    app = Application()
    app.mainloop()
