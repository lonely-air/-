import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from PIL import Image, ImageTk
import os
import threading
import glob
import time
import sys
from pickle import TRUE
from openpyxl import load_workbook
from openpyxl import Workbook
from concurrent.futures import ThreadPoolExecutor
import datetime
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import cm, rcParams
from mpl_toolkits.mplot3d import Axes3D
from numpy.core.function_base import linspace
import pandas as pd
import numpy as np
from scipy import linalg
import math
from docx import Document
from docx.shared import Cm, Pt, Inches
from docx.oxml.ns import qn
from docx.shared import RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_PARAGRAPH_ALIGNMENT
from docx.oxml.shared import OxmlElement
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from scipy.interpolate import interp1d

# ++++++++++++++++++++ UTILITY FUNCTIONS (for PyInstaller) ++++++++++++++++++++

def resource_path(relative_path):
    """ 获取资源的绝对路径，适用于开发环境和PyInstaller打包环境 """
    try:
        # PyInstaller 创建一个临时文件夹，并将路径存储在 _MEIPASS 中
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# ++++++++++++++++++++ MODEL FUNCTIONS (from mymodel_v1_8.py) ++++++++++++++++++++

def Initialize(plt):
    plt.rcParams['font.sans-serif'] = ['SimSun']
    plt.rcParams['axes.unicode_minus'] = False
    plt.rcParams['figure.figsize']=(8, 6)
    plt.rcParams['font.size'] = 14

def copy_data(source_sheet, target_sheet, source_col, target_col, start_row, end_row,target_start_row=2):
    for row in range(start_row, end_row):
        value = source_sheet[f'{source_col}{row}'].value
        target_sheet[f'{target_col}{target_start_row}'] = value
        target_start_row += 1

def find_repeated_value(ws):
    value_count = {}
    for idx, row in enumerate(ws.iter_rows(min_row=2, min_col=7, max_col=7, values_only=True), start=2):
        value = row[0]
        if value is not None and value > 10:
            if value in value_count:
                value_count[value][0] += 1
            else:
                value_count[value] = [1, idx]
            if value_count[value][0] == 600:
                return value, value_count[value][1]
    return None, None

def find_value_in_column(ws, col_letter, start_row, value_to_find):
    for row_idx in range(start_row, ws.max_row + 1):
        cell_value = ws[f'{col_letter}{row_idx}'].value
        if cell_value == value_to_find:
            return row_idx
    return None

def find_last_numeric_value_in_column(ws, col_letter):
    for row_idx in range(ws.max_row, 1, -1):
        cell_value = ws[f'{col_letter}{row_idx}'].value
        if isinstance(cell_value, (int, float)):
            return cell_value
    return None

def filter_sheet_by_column(source_sheet, target_wb, sheet_name, col_letter, start_row):
    try:
        target_sheet = target_wb[sheet_name]
    except KeyError:
        target_sheet = target_wb.create_sheet(sheet_name)
    for col_idx in range(1, source_sheet.max_column + 1):
        target_sheet.cell(row=1, column=col_idx).value = source_sheet.cell(row=start_row, column=col_idx).value
    target_row_idx = 2
    for row_idx in range(start_row, source_sheet.max_row + 1):
        cell_value = source_sheet[f'{col_letter}{row_idx}'].value
        if cell_value is not None and isinstance(cell_value, (int, float)) and cell_value > 0:
            for col_idx in range(1, source_sheet.max_column + 1):
                cell_to_copy = source_sheet.cell(row=row_idx, column=col_idx)
                target_sheet.cell(row=target_row_idx, column=col_idx).value = cell_to_copy.value
            target_row_idx += 1

def count(ws):
    cnt=0
    for cell in ws['C']:
        if cnt>=1 and cell.value is None:
            break
        cnt+=1
    return cnt

def IsStr(x):
    return isinstance(x,str) and (x[0:2] == "压力" or x[0:4] == "施工压力")
def IsPai(x):
    return isinstance(x,str) and x[0:2] == "排量"
def IsYe(x):
    return isinstance(x,str) and (x[0:4] == "累计液量" or x[0:3] == "总液量")
def IsSB(x):
    return isinstance(x,str) and x[0:2] == "砂比"
def IsSN(x):
    return isinstance(x,str) and x[0:3] == "砂浓度"

def get_sha(ws):
    result =1
    for col in ws['1']:
        if IsSB(col.value):
            result=0
    return result

def setfont(run):
    font=run.font
    font.name = 'Times New Roman'
    font.size = Pt(10.5)
    font.color.rgb = RGBColor(0, 0, 0)
    run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

def get_deep_sha(ws):
    result =1
    lst = list(str(i) for i in range (1,101))
    for j in lst:
        for col in ws[j]:
            if IsSB(col.value):
                result=0
                return result
    return result

def find_data(ws,lb):
    for i in ws[lb]:
        if isinstance(i.value,int) or isinstance(i.value,float):
            result = int(i.coordinate[1:])
            return result
    return -1

def get_data(ws):
    choose = get_sha(ws)
    压力, 排量, 累计液量, 砂 = [None, None], [None, None], [None, None], [None, None]
    for col in ws['1']:
        result1=col.coordinate[0]
        result2=find_data(ws,result1)
        if IsStr(col.value):
            压力=[result1,result2]
        elif IsPai(col.value):
            排量=[result1,result2]
        elif IsYe(col.value):
            累计液量=[result1,result2]
        if choose == 0:
            if IsSB(col.value):
                砂=[result1,result2]
        elif choose == 1:
            if IsSN(col.value):
                砂=[result1,result2]
    return [压力,排量,累计液量,砂]

def deep_scan (ws):
    result = 1
    choose = get_sha(ws)
    p=[0,0,0,0]
    for col in ws['1']:
        if IsStr(col.value): p[0]=1
        elif IsPai(col.value): p[1]=1
        elif IsYe(col.value): p[2]=1
        else:
            if choose == 0:
                if IsSB(col.value): p[3]=1
            elif choose == 1:
                if IsSN(col.value): p[3]=1
    if p[0]==1 and p[1]==1 and p[2]==1 and p[3]==1 :
        result = 0
    return result

def get_deep_data(ws):
    choose = get_deep_sha(ws)
    deep=20
    lst = list(str(i) for i in range (1,deep))
    p=[0,0,0,0]
    压力, 排量, 累计液量, 砂 = [None, None], [None, None], [None, None], [None, None]
    for j in lst :
        for col in ws[j]:
            if p[0]==1 and p[1]==1 and p[2]==1 and p[3]==1 : break
            result1=col.coordinate[0]
            result2=find_data(ws,result1)
            if IsStr(col.value):
                压力=[result1,result2]; p[0]=1
            elif IsPai(col.value):
                排量=[result1,result2]; p[1]=1
            elif IsYe(col.value):
                累计液量=[result1,result2]; p[2]=1
            else:
                if choose == 0:
                    if IsSB(col.value): 砂=[result1,result2]; p[3]=1
                elif choose == 1:
                    if IsSN(col.value): 砂=[result1,result2]; p[3]=1
    return [压力,排量,累计液量,砂]

def find_end(ws):
    for i in range(ws.max_row,0,-1):
        if type(ws[f'C{i}'].value) in [int, float, str]:
            return i
    return 0

def data_clean(wb,target):
    sheet = wb.sheetnames
    for id in range(len(sheet)):
        print(f"正在处理第{id+1}页")
        ws = wb[sheet[id]]
        end=find_end(ws)
        for k in range(end, 0, -1):
            if ws[f'C{k}'].value is None:
                print(f"正在处理第{k}行")
                ws.delete_rows(idx=k)
    print("清洗完成")
    wb.save(target)

def data_solve(source,target,mode = 0):
    if mode == 0 :
        wb = load_workbook(source)
        sheets = wb.sheetnames
        st=0
        ed=len(sheets)
    elif mode == 1:
        st=0
        ed=100
    print("数据整理开始")
    try:
        target_wb = load_workbook(target)
    except FileNotFoundError:
        target_wb = Workbook()
        target_wb.save(target)
        target_wb = load_workbook(target)
    
    data_clean(wb,source)
    if 'Sheet' in target_wb.sheetnames:
        del target_wb['Sheet']
    for i in range (st,ed):
        if mode == 0:
            print(f"第{i+1}段数据处理")
            ws=wb[sheets[i]]
        elif mode == 1:
            try:
                print(f"第{i+1}段数据处理")
                wb_m = load_workbook(f'm{i+1}.xlsx')
                ws=wb_m[wb_m.sheetnames[0]]
            except FileNotFoundError:
                print(f"不存在第{i+1}段,结束执行")
                break
        
        dp = deep_scan(ws)
        if dp==0 :
            砂浓度转化=get_sha(ws)
            压力,排量,累计液量标签,砂=get_data(ws)
        elif dp==1 :
            砂浓度转化=get_deep_sha(ws)
            压力,排量,累计液量标签,砂=get_deep_data(ws)
        
        try:
            target_sheet = target_wb[str(i+1)]
        except KeyError:
            target_sheet = target_wb.create_sheet(str(i+1))

        砂列='E'
        if 砂浓度转化==True: 砂列='F'

        with ThreadPoolExecutor(max_workers=30) as executor:
            executor.submit(copy_data, ws, target_sheet, 压力[0], 'C', 压力[1], 200001,2)
            executor.submit(copy_data, ws, target_sheet, 排量[0], 'D', 排量[1], 200001,2)
            executor.submit(copy_data, ws, target_sheet, 砂[0], 砂列, 砂[1], 200001,2)
            executor.submit(copy_data, ws, target_sheet, 累计液量标签[0], 'G', 累计液量标签[1], 200001,2)
        
        target_row = 2
        if 砂浓度转化==True:
            for row in range(2, 200001):
                value = target_sheet[f'F{row}'].value
                if value is not None:
                    target_sheet[f'E{target_row}'] = value / 15
                    target_row += 1
        
        累计液量=find_last_numeric_value_in_column(target_sheet, 'G')
        行号=find_value_in_column(target_sheet, 'G', 2, 累计液量)
        if 累计液量 is not None and 行号 is not None and 行号 > 0:
            泵注时间=行号
            平均排量=累计液量/泵注时间
            停泵压力=target_sheet[f'C{行号+3}'].value
            p_i=target_sheet[f'C{行号+900}'].value
            t_i=行号+900
            p_j=target_sheet[f'C{行号+1200}'].value
            t_j=行号+1200
            G函数压力序列=[平均排量,泵注时间,停泵压力,p_i,t_i,p_j,t_j]
            start_col = 'H'
            start_row = 1
            for idx, value in enumerate(G函数压力序列):
                target_sheet[f'{chr(ord(start_col) + idx)}{start_row}'] = value
    
    target_wb.save(target)
    target_wb.close()
    print('数据整理处理完成！')

def fig_solve(target):
    Initialize(plt)
    print("图片生成开始运行")
    figs_path = "figs"
    try:
        os.makedirs(figs_path, exist_ok=True)
    except FileExistsError:
        print("图片文件夹已存在，继续运行")

    wb = load_workbook(target)
    sheets = wb.sheetnames
    for i in range (0,len(sheets)):
        print(f"正在进行第{i+1}张图片处理")
        ws = wb[sheets[i]]
        wsl=count(ws)
        压力=np.array([cell.value for cell in ws['C'][1:wsl]], dtype=np.float64)
        排量=np.array([cell.value for cell in ws['D'][1:wsl]], dtype=np.float64)
        砂比=np.array([cell.value for cell in ws['E'][1:wsl]], dtype=np.float64)
        sec=np.array(range(1,wsl), dtype=np.int64)
        min_val=sec/60

        fig=plt.figure()
        ax1=fig.subplots()
        ax2=ax1.twinx()
        ax1.plot(min_val, 压力,'r',label="施工压力(MPa)")
        ax1.plot(min_val, 排量,'g',label="排量(m$^3$/min)")
        ax1.plot(min_val, 砂比,'black',alpha=0.5,label="砂比(%)")
        ax1.set_xlabel("泵注时间(min)")
        ax1.set_ylabel("施工压力(MPa)\n排量(m$^3$/min)")
        ax2.plot(min_val, 压力,'black',alpha=0,label="砂比(%)")
        ax2.set_ylabel("砂比(%)")
        ax1.set_yticks(np.arange(0, 81, 10))
        ax2.set_yticks(np.arange(0, 81, 10))
        ax1.legend()
        ax1.grid(True)
        plt.savefig(os.path.join(figs_path, f'施工曲线第{i+1}段.png'), dpi=600)
        plt.close(fig)
    print("图片处理完成")

def YP_solve(source,target):
    wb = load_workbook(source)
    sheets = wb.sheetnames
    try:
        target_wb = load_workbook(target)
    except FileNotFoundError:
        target_wb = Workbook()
    if 'Sheet' in target_wb.sheetnames:
        del target_wb['Sheet']
        
    for i in range(len(sheets)):
        ws = wb[sheets[i]]
        print(f'正在处理第{i+1}段')
        filter_sheet_by_column(ws, target_wb, str(i+1), 'E', 2)
    target_wb.save(target)
    target_wb.close()
    print('处理完成！')
    
def YP_fig(target):
    Initialize(plt)
    print("图片生成开始运行")
    figs_path = "figs"
    os.makedirs(figs_path, exist_ok=True)

    wb = load_workbook(target)
    sheets = wb.sheetnames
    for i in range (len(sheets)):
        print(f"正在进行第{i+1}张图片处理")
        ws = wb[sheets[i]]
        wsl=count(ws)
        压力=np.array([cell.value for cell in ws['C'][1:wsl]], dtype=np.float64)
        排量=np.array([cell.value for cell in ws['D'][1:wsl]], dtype=np.float64)
        砂比=np.array([cell.value for cell in ws['E'][1:wsl]], dtype=np.float64)
        min_val=np.array(range(1,wsl), dtype=np.int64)/60
        压排比=np.divide(压力, 排量, out=np.zeros_like(压力), where=排量!=0)
        
        fig=plt.figure()
        ax1=fig.subplots()
        ax2=ax1.twinx()
        ax1.plot(min_val, 压排比,'r',label="压排比(MPa/(m$^3$/min))")
        ax2.plot(min_val, 砂比,'black',alpha=0.5,label="砂比(%)")
        ax1.set_xlabel("泵注时间(min)")
        ax1.set_ylabel("压力(MPa)/排量(m$^3$/min)")
        ax2.set_ylabel("砂比(%)")
        ax1.set_ylim(0,5)
        ax2.set_ylim(0,100)
        ax1.set_yticks(np.arange(0, 5.1, 0.5))
        ax2.set_yticks(np.arange(0, 101, 10))
        line1, = ax1.plot([],[],'r',label="压排比(MPa/(m$^3$/min))")
        line2, = ax2.plot([],[],'black',alpha=0.5,label="砂比(%)")
        ax1.legend(handles = [line1, line2])
        ax1.grid(True)
        plt.savefig(os.path.join(figs_path, f'压排比曲线第{i+1}段.png'), dpi=600)
        plt.close(fig)
    print("图片处理完成")

def fig_doc(target):
    wb = load_workbook(target)
    sheets = wb.sheetnames
    figs_path = "figs"

    doc = Document()
    heading = doc.add_heading(level=2)
    run = heading.add_run("3.5 施工曲线图")
    font=run.font
    font.name = 'Times New Roman'
    font.size = Pt(14)
    font.color.rgb = RGBColor(0, 0, 0)
    run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    for i in range (len(sheets)):
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run()
        run.add_picture(os.path.join(figs_path, f'施工曲线第{i+1}段.png'), width=Cm(12))

        para = doc.add_paragraph()
        para.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = para.add_run(f"图3-5-{i+1}    第{i+1}级泵注曲线")
        setfont(run)
        run.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
    doc.save(os.path.join(figs_path, '施工曲线图.docx'))
    print("文件已生成")

def TBsolve(source):
    target = "停泵曲线.xlsx"
    try:
        wb = load_workbook(source)
    except FileNotFoundError:
        print(f"文件 {source} 不存在")
        return
    
    # Always create a new workbook to avoid issues with existing files or empty workbooks
    twb = Workbook()
    ws_t = twb.active
    ws_t.title = "停泵曲线"
    ws_t['A1'].value='时间'
    ws_t['B1'].value='压力'

    k=2
    for j, sheet_name in enumerate(wb.sheetnames):
        print(f"处理第 {j+1} 段")
        ws_s=wb[sheet_name]
        start_row_val = ws_s['I1'].value
        if not isinstance(start_row_val, int):
            print(f"第 {j+1} 段的 I1 单元格不是有效行号，跳过。")
            continue
        
        cnt = ws_s.max_row - start_row_val-10
        for i in range(cnt):
            ws_t[f'B{k}'].value=ws_s[f'C{i+start_row_val}'].value
            k+=1
        k+=50

    for i in range(1, ws_t.max_row):
        ws_t[f'A{i+1}'].value=i
        if ws_t[f'B{i+1}'].value == 0:
            ws_t[f'B{i+1}'].value = None

    wb.close()
    twb.save(target)
    twb.close()
    print('完成')

def TBfig(source):
    wb = load_workbook(source)
    sheets = wb.sheetnames
    ws = wb[sheets[0]]
    Initialize(plt)
    l1 = np.array(list( i.value for i in ws['A'])[1:],dtype=np.float64)
    l2 = np.array(list( i.value for i in ws['B'])[1:],dtype=np.float64)
    fig=plt.figure()
    ax1=fig.subplots()
    ax1.plot(l1, l2,'r',label="压力(MPa)")
    ax1.set_xlabel("时间(s)")
    ax1.set_ylabel("压力(MPa)")
    #ax1.set_ylim(0,15)
    ax1.set_yticks(np.arange(0, 50, 5))
    ax1.title.set_text("FK11-X2井3-4段压降数据（MPa）")
    #设置标签用
    ax1.legend()
    ax1.grid(True)
    plt.savefig(f'figs/停泵曲线',dpi=600)
    wb.close()
    print("图片处理完成")
    
def process_pump_injection_curve():
    source = '泵注曲线待处理.xlsx'
    target = '泵注曲线.xlsx'
    print("\n开始处理泵注曲线数据...")
    data_solve(source, target)
    print("\n开始生成泵注曲线图片...")
    fig_solve(target)
    print("\n开始生成曲线文档...")
    fig_doc(target)
    print("\n泵注曲线处理完成！")

def process_pressure_flow_ratio():
    source = '泵注曲线.xlsx'
    target = '压排比.xlsx'
    print("\n开始处理压排比数据...")
    YP_solve(source, target)
    print("\n开始生成压排比图片...")
    YP_fig(target)
    print("\n压排比处理完成！")

def process_shutdown_curve():
    source = '泵注曲线.xlsx'
    target = '停泵曲线.xlsx'
    print("\n开始处理停泵曲线数据...")
    TBsolve(source)
    print("\n开始生成停泵曲线图片...")
    TBfig(target)
    print("\n停泵曲线处理完成！")

# ++++++++++++++++++++ GUI APPLICATION CLASS ++++++++++++++++++++

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("压裂数据处理系统")
        self.geometry("1200x700")
        self.source_file_path = None

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=2)

        self.control_frame = ttk.Frame(self, padding="10")
        self.control_frame.grid(row=0, column=0, sticky="nsew")
        self.control_frame.grid_rowconfigure(6, weight=1) # Adjust row for log text

        # --- File Selection ---
        self.btn_select_file = ttk.Button(self.control_frame, text="选择源文件", command=self.select_file)
        self.btn_select_file.grid(row=0, column=0, pady=5, sticky="ew")

        self.file_path_label = ttk.Label(self.control_frame, text="未选择文件", anchor="w", relief="sunken")
        self.file_path_label.grid(row=1, column=0, pady=2, sticky="ew")

        # --- Processing Buttons ---
        self.btn_pump_injection = ttk.Button(self.control_frame, text="处理泵注曲线", command=lambda: self.run_function_in_thread(self.process_pump_injection_curve_gui))
        self.btn_pump_injection.grid(row=2, column=0, pady=5, sticky="ew")

        self.btn_pressure_flow = ttk.Button(self.control_frame, text="处理压排比", command=lambda: self.run_function_in_thread(self.process_pressure_flow_ratio_gui))
        self.btn_pressure_flow.grid(row=3, column=0, pady=5, sticky="ew")

        self.btn_shutdown_curve = ttk.Button(self.control_frame, text="处理停泵曲线", command=lambda: self.run_function_in_thread(self.process_shutdown_curve_gui))
        self.btn_shutdown_curve.grid(row=4, column=0, pady=5, sticky="ew")

        self.btn_full_process = ttk.Button(self.control_frame, text="全流程处理", command=lambda: self.run_function_in_thread(self.process_full_gui))
        self.btn_full_process.grid(row=5, column=0, pady=5, sticky="ew")

        self.btn_exit = ttk.Button(self.control_frame, text="退出程序", command=self.quit_app)
        self.btn_exit.grid(row=6, column=0, pady=5, sticky="ew")

        self.log_text = scrolledtext.ScrolledText(self.control_frame, wrap=tk.WORD, height=15)
        self.log_text.grid(row=7, column=0, pady=10, sticky="nsew")
        self.log_text.config(state=tk.DISABLED)

        sys.stdout = ConsoleRedirector(self.log_text)
        sys.stderr = ConsoleRedirector(self.log_text)

        self.image_frame = ttk.Frame(self, padding="10")
        self.image_frame.grid(row=0, column=1, sticky="nsew")
        self.image_frame.grid_rowconfigure(0, weight=1)
        self.image_frame.grid_columnconfigure(0, weight=1)

        self.image_canvas = tk.Canvas(self.image_frame, bg="lightgray")
        self.image_canvas.grid(row=0, column=0, sticky="nsew")

        self.image_files = []
        self.image_combobox = ttk.Combobox(self.image_frame, values=self.image_files, state="readonly")
        self.image_combobox.grid(row=1, column=0, pady=10, sticky="ew")
        self.image_combobox.bind("<<ComboboxSelected>>", self.display_selected_image)

        self.set_buttons_state(tk.DISABLED) # Initially disable buttons
        self.update_image_list()

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="选择 '泵注曲线待处理.xlsx' 文件",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            if os.path.basename(file_path) != '泵注曲线待处理.xlsx':
                messagebox.showwarning("文件名错误", "请选择名为 '泵注曲线待处理.xlsx' 的文件。")
                return
            self.source_file_path = file_path
            self.file_path_label.config(text=file_path)
            self.set_buttons_state(tk.NORMAL)
            self.log_message(f"已选择文件: {file_path}")
            self.update_image_list() # Update images from new directory

    def run_function_in_thread(self, func):
        if not self.source_file_path:
            messagebox.showerror("错误", "请先选择一个源文件！")
            return

        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.set_buttons_state(tk.DISABLED)
        self.log_message(f"开始执行...")
        
        # Run in thread with dynamic working directory
        thread = threading.Thread(target=self.run_with_changed_dir, args=(func,))
        thread.start()

    def run_with_changed_dir(self, func):
        original_dir = os.getcwd()
        target_dir = os.path.dirname(self.source_file_path)
        try:
            os.chdir(target_dir)
            func()
        finally:
            os.chdir(original_dir)
            self.after(0, self.set_buttons_state, tk.NORMAL)

    def set_buttons_state(self, state):
        self.btn_pump_injection.config(state=state)
        self.btn_pressure_flow.config(state=state)
        self.btn_shutdown_curve.config(state=state)
        self.btn_full_process.config(state=state)

    def log_message(self, message):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def update_image_list(self):
        if not self.source_file_path:
            self.image_files = []
        else:
            figs_dir = os.path.join(os.path.dirname(self.source_file_path), "figs")
            if not os.path.exists(figs_dir):
                os.makedirs(figs_dir)
            image_paths = sorted(glob.glob(os.path.join(figs_dir, "*.png")))
            self.image_files = [os.path.basename(f) for f in image_paths]

        self.image_combobox['values'] = self.image_files
        if self.image_files:
            self.image_combobox.set(self.image_files[-1])
            self.display_selected_image()
        else:
            self.image_combobox.set("")
            self.image_canvas.delete("all")

    def display_selected_image(self, event=None):
        selected_image_name = self.image_combobox.get()
        if selected_image_name and self.source_file_path:
            image_path = os.path.join(os.path.dirname(self.source_file_path), "figs", selected_image_name)
            self.display_image(image_path)

    def display_image(self, image_path):
        try:
            self.image_canvas.update_idletasks()
            canvas_width = self.image_canvas.winfo_width()
            canvas_height = self.image_canvas.winfo_height()

            if canvas_width <= 1 or canvas_height <= 1:
                self.after(100, lambda: self.display_image(image_path))
                return

            img = Image.open(image_path)
            img_width, img_height = img.size
            
            ratio = min(canvas_width / img_width, canvas_height / img_height)
            new_width = int(img_width * ratio * 0.95) # Add some padding
            new_height = int(img_height * ratio * 0.95)

            img = img.resize((new_width, new_height), Image.LANCZOS)
            self.photo = ImageTk.PhotoImage(img)
            
            self.image_canvas.delete("all")
            x = (canvas_width - new_width) / 2
            y = (canvas_height - new_height) / 2
            self.image_canvas.create_image(x, y, anchor=tk.NW, image=self.photo)
        except Exception as e:
            self.log_message(f"显示图片时发生错误：{e}")
            self.image_canvas.delete("all")

    def process_pump_injection_curve_gui(self):
        try:
            process_pump_injection_curve()
            self.log_message("泵注曲线处理完成！")
        except Exception as e:
            self.log_message(f"泵注曲线处理失败：{e}")
            messagebox.showerror("错误", f"泵注曲线处理失败：{e}")
        finally:
            self.after(0, self.update_image_list)

    def process_pressure_flow_ratio_gui(self):
        try:
            process_pressure_flow_ratio()
            self.log_message("压排比处理完成！")
        except Exception as e:
            self.log_message(f"压排比处理失败：{e}")
            messagebox.showerror("错误", f"压排比处理失败：{e}")
        finally:
            self.after(0, self.update_image_list)

    def process_shutdown_curve_gui(self):
        try:
            process_shutdown_curve()
            self.log_message("停泵曲线处理完成！")
        except Exception as e:
            self.log_message(f"停泵曲线处理失败：{e}")
            messagebox.showerror("错误", f"停泵曲线处理失败：{e}")
        finally:
            self.after(0, self.update_image_list)

    def process_full_gui(self):
        try:
            self.log_message("开始全流程处理...")
            process_pump_injection_curve()
            process_pressure_flow_ratio()
            process_shutdown_curve()
            self.log_message("全流程处理完成！")
        except Exception as e:
            self.log_message(f"全流程处理失败：{e}")
            messagebox.showerror("错误", f"全流程处理失败：{e}")
        finally:
            self.after(0, self.update_image_list)

    def quit_app(self):
        self.log_message("程序已退出。")
        self.destroy()
        sys.exit(0)

class ConsoleRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.stdout = sys.__stdout__
        self.stderr = sys.__stderr__

    def write(self, message):
        # Use 'after' to ensure thread-safety with Tkinter
        self.text_widget.after(0, self.insert_text, message)
        if self.stdout:
            self.stdout.write(message)
            self.stdout.flush()

    def insert_text(self, message):
        self.text_widget.config(state=tk.NORMAL)
        self.text_widget.insert(tk.END, message)
        self.text_widget.see(tk.END)
        self.text_widget.config(state=tk.DISABLED)

    def flush(self):
        if self.stdout:
            self.stdout.flush()

if __name__ == "__main__":
    app = Application()
    app.mainloop()
