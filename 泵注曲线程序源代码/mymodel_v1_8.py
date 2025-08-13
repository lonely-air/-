from pickle import TRUE
from openpyxl import load_workbook
from openpyxl import Workbook
from concurrent.futures import ThreadPoolExecutor# 导入线程池模块
import datetime
import os
import matplotlib.pyplot as plt
from matplotlib import cm, rcParams
from mpl_toolkits.mplot3d import Axes3D
from numpy.core.function_base import linspace
import pandas as pd
from rich import console#终端美化输出(rich)
from rich.table import Column, Table
import numpy as np
from scipy import linalg
import math
import sys
from rich.console import Console
# from itertools import combinations
# from scipy.special import comb
from docx import Document  # 用来建立一个word对象
from docx.shared import Cm, Pt  # 用来设置字体的大小
from docx.shared import Inches
from docx.oxml.ns import qn  # 设置字体
from docx.shared import RGBColor  # 设置字体的颜色
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_PARAGRAPH_ALIGNMENT  # 设置对其方式
from docx.oxml.shared import OxmlElement, qn
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from scipy.interpolate import interp1d

#字体初始化
def Initialize(plt): #导入plt确保导入全局图片变量
    plt.rcParams['font.sans-serif'] = ['SimSun']  # 用来正常显示中文标签
    plt.rcParams['axes.unicode_minus'] = False  # 用来正常显示负号
    plt.rcParams['figure.figsize']=(8, 6)  #全局图片大小
    plt.rcParams['font.size'] = 14         #全局字体大小

def copy_data(source_sheet, target_sheet, source_col, target_col, start_row, end_row,target_start_row=2):
    for row in range(start_row, end_row):
        value = source_sheet[f'{source_col}{row}'].value
        target_sheet[f'{target_col}{target_start_row}'] = value
        target_start_row += 1  # 每次写入后递增目标行
        
def find_repeated_value(ws):
    # 创建一个字典来存储值出现的次数和第一次出现的行号
    value_count = {}
    
    # 遍历G列的所有单元格
    for idx, row in enumerate(ws.iter_rows(min_row=2, min_col=7, max_col=7, values_only=True), start=2):
        value = row[0]  # 获取单元格的值
        # print(value)
        if value is not None and value > 10:  # 只处理大于10的值
            if value in value_count:  # 如果值已经存在，则增加计数器
                value_count[value][0] += 1
            else:
                value_count[value] = [1, idx]  # 记录第一次出现的行号
            
            # 检查是否有200个相同的值
            if value_count[value][0] == 600:
                # print(f"找到200个相同的值: {value_count}")
                return value, value_count[value][1]
    
    return None, None

###计算停泵时间2
def find_value_in_column(ws, col_letter, start_row, value_to_find):
    """
    在指定工作表的指定列中查找特定数值，并返回第一次出现的行号。
    :param ws: openpyxl工作表对象。
    :param col_letter: 要查找的列的字母（例如 'A', 'B', 'C'）。
    :param start_row: 开始查找的行号。
    :param value_to_find: 要查找的数值。
    :return: 数值第一次出现的行号，如果未找到则返回None。
    """
    for row_idx in range(start_row, ws.max_row + 1):
        cell_value = ws[f'{col_letter}{row_idx}'].value
        if cell_value == value_to_find:
            return row_idx
    return None
###计算累计液量
def find_last_numeric_value_in_column(ws, col_letter):
    """
    在指定工作表的指定列中查找最后一个数字值。
    :param ws: openpyxl工作表对象。
    :param col_letter: 要查找的列的字母（例如 'A', 'B', 'C'）。
    :return: 列中最后一个数字值，如果未找到则返回None。
    """
    for row_idx in range(ws.max_row, 1, -1): # 从最大行开始向上遍历，总是反映实际有数据存在的最大行号
        cell_value = ws[f'{col_letter}{row_idx}'].value
        if isinstance(cell_value, (int, float)):
            return cell_value
    return None
#计算有多少个有效数据

def filter_sheet_by_column(source_sheet, target_wb, sheet_name, col_letter, start_row):
    """
    根据指定列的条件过滤工作表数据，并将结果写入目标工作簿的指定工作表。

    参数:
    source_sheet (openpyxl.worksheet.worksheet.Worksheet): 源工作表对象。
    target_wb (openpyxl.workbook.workbook.Workbook): 目标工作簿对象。
    sheet_name (str): 目标工作表中要写入数据的工作表名称。
    col_letter (str): 要检查的列的字母（例如 'A', 'B'）。
    start_row (int): 开始检查的行号。
    """
    # 获取或创建目标工作表
    try:
        target_sheet = target_wb[sheet_name]
    except KeyError:
        target_sheet = target_wb.create_sheet(sheet_name)

    # 清空目标工作表原有内容（可选，取决于您的需求）
    # for row in target_sheet.iter_rows():
    #     for cell in row:
    #         cell.value = None

    # 复制标题行 (如果需要)
    for col_idx in range(1, source_sheet.max_column + 1):
        target_sheet.cell(row=1, column=col_idx).value = source_sheet.cell(row=start_row, column=col_idx).value

    target_row_idx = 2 # 新工作表的起始行（假设从第二行开始写数据）

    for row_idx in range(start_row, source_sheet.max_row + 1):
        cell_value = source_sheet[f'{col_letter}{row_idx}'].value
        
        # 检查值是否大于0且不是空白格
        if cell_value is not None and isinstance(cell_value, (int, float)) and cell_value > 0:
            # 如果符合条件，复制整行数据到目标工作表
            for col_idx in range(1, source_sheet.max_column + 1):
                cell_to_copy = source_sheet.cell(row=row_idx, column=col_idx)
                target_sheet.cell(row=target_row_idx, column=col_idx).value = cell_to_copy.value
            target_row_idx += 1
            
    # 函数不再返回工作表，直接修改了 target_wb
    # return target_sheet 

def count(ws):        
    cnt=0
    for cell in ws['C']:
        if cnt>=1 and cell.value is None:     #首行标题为空，要跳过 不能用==要用is
            break
        cnt+=1
    return cnt
#判断压力格
def IsStr(x):
    return isinstance(x,str) and (x[0:2] == "压力" or x[0:4] == "施工压力")
#判断排量格
def IsPai(x):
    return isinstance(x,str) and x[0:2] == "排量"
#判断总液量格
def IsYe(x):
    return isinstance(x,str) and (x[0:4] == "累计液量" or x[0:3] == "总液量")
#判断砂比格
def IsSB(x):
    return isinstance(x,str) and x[0:2] == "砂比"
#判断砂浓度格
def IsSN(x):
    return isinstance(x,str) and x[0:3] == "砂浓度"
#确认是否有砂比,有为0，无为1

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
    
#输入列标签向下查找第一个数字的标签
def find_data(ws,lb):
    for i in ws[lb]:
        if isinstance(i.value,int) or isinstance(i.value,float):
            result = int(i.coordinate[1:])
            return result
    return -1
#自动捕获压力，排量,累计液量,砂比或砂浓度
def get_data(ws):
    choose = get_sha(ws)
    for col in ws['1']:       #保证数据标签在第一行即可
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

#扫描是否使用深度模式，是为1，不是为0
def deep_scan (ws):
    result = 1
    choose = get_sha(ws)
    p=[0,0,0,0]
    for col in ws['1']:
        if IsStr(col.value):
            p[0]=1
        elif IsPai(col.value):
            p[1]=1
        elif IsYe(col.value):
            p[2]=1
        else:
            if choose == 0:
                if IsSB(col.value): 
                    p[3]=1
            elif choose == 1:
                if IsSN(col.value):
                    p[3]=1
    if p[0]==1 and p[1]==1 and p[2]==1 and p[3]==1 :
        result = 0
    return result
        
#加上一个范围，无视数据标签在第一行的限制，但运行会变慢
def get_deep_data(ws):
    choose = get_deep_sha(ws)
    deep=20
    lst = list(str(i) for i in range (1,deep))
    p=[0,0,0,0]
    for j in lst :
        for col in ws[j]:
            if p[0]==1 and p[1]==1 and p[2]==1 and p[3]==1 :
                break 
            result1=col.coordinate[0]
            result2=find_data(ws,result1)
            if IsStr(col.value):
                压力=[result1,result2]
                p[0]=1
            elif IsPai(col.value):
                排量=[result1,result2]
                p[1]=1
            elif IsYe(col.value):
                累计液量=[result1,result2]
                p[2]=1
            else:
                if choose == 0:
                    if IsSB(col.value):
                        砂=[result1,result2]  
                        p[3]=1
                elif choose == 1:
                    if IsSN(col.value):
                        砂=[result1,result2]
                        p[3]=1  
    return [压力,排量,累计液量,砂]

#找到一个excel中一页里面最大的位数行
def find_end(ws):
    for i in range(ws.max_row,0,-1):
        if type(ws[f'C{i}'].value) == int or type(ws[f'C{i}'].value) == float or type(ws[f'C{i}'].value) == str:
            result = i
            break
    return result
#清洗空格行
def data_clean(wb,target):
    sheet = wb.sheetnames
    for id in range(len(sheet)):
        print(f"正在处理第{id+1}页")
        ws = wb[sheet[id]]
        end=find_end(ws)
        # 倒序循环以避免索引问题
        for k in range(end, 0, -1):
            if ws[f'C{k}'].value is None:
                print(f"正在处理第{k}行")
                ws.delete_rows(idx=k)
    print("清洗完成")
    wb.save(target)

#泵注曲线数据整理,mode =1 时为多m文件处理，mode = 0时为单泵注曲线待处理
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
        # 如果文件不存在则创建新的工作簿
        target_wb = Workbook()
        target_wb.save(target)
        target_wb = load_workbook(target)# 重新加载工作簿

    data_clean(wb,source)
    # 删除默认创建的工作表
    if 'Sheet' in target_wb.sheetnames:
        del target_wb['Sheet']       
    for i in range (st,ed):

        if mode == 0:
            print(f"第{i+1}段数据处理")
            ws=wb[sheets[i]]
        elif mode == 1:
            try:
                print(f"第{i+1}段数据处理")
                wb = load_workbook(f'm{i+1}.xlsx')
                sheets = wb.sheetnames
                ws=wb[sheets[0]]
            except FileNotFoundError:
                print(f"不存在第{i+1}段,结束执行")
                break
        dp = deep_scan(ws)
        #检测是否开启深度搜索标签模式
        if dp==0 :
            砂浓度转化=get_sha(ws)
            压力,排量,累计液量标签,砂=get_data(ws)
        elif dp==1 :
            砂浓度转化=get_deep_sha(ws)          
            压力,排量,累计液量标签,砂=get_deep_data(ws)
        
        

        try:
            target_sheet = target_wb[str(i+1)]
        except KeyError:
            # 如果工作表不存在则创建新的工作表
            target_sheet = target_wb.create_sheet(str(i+1))

        砂列='E'
        if 砂浓度转化==True:
            砂列='F'

        with ThreadPoolExecutor(max_workers=30) as executor:
            executor.submit(copy_data, ws, target_sheet, 压力[0], 'C', 压力[1], 200001,2)#压力
            executor.submit(copy_data, ws, target_sheet, 排量[0], 'D', 排量[1], 200001,2)#排量
            executor.submit(copy_data, ws, target_sheet, 砂[0], 砂列, 砂[1], 200001,2)#砂比e,0/砂浓度f,1
            executor.submit(copy_data, ws, target_sheet, 累计液量标签[0], 'G', 累计液量标签[1], 200001,2)#累计液量
        #处理砂浓度为砂比
        
        # 初始化目标行号
        target_row = 2
        # 处理砂浓度为砂比.处理为1，不处理为0

        if 砂浓度转化==True:
            for row in range(2, 200001):
                value = target_sheet[f'F{row}'].value
                if value is not None:
                    target_sheet[f'E{target_row}'] = value / 15
                    target_row += 1  # 每次写入后递增目标行
        
        #计算累计液量
        累计液量=find_last_numeric_value_in_column(target_sheet, 'G')
        行号=find_value_in_column(target_sheet, 'G', 2, 累计液量)
        泵注时间=行号
        平均排量=累计液量/泵注时间
        停泵压力=target_sheet[f'C{行号+3}'].value#停泵压力误差加5秒
        p_i=target_sheet[f'C{行号+900}'].value
        t_i=行号+900
        p_j=target_sheet[f'C{行号+1200}'].value#需要有至少1000秒停泵时间
        t_j=行号+1200
        G函数压力序列=[平均排量,泵注时间,停泵压力,p_i,t_i,p_j,t_j]

        start_col = 'H'
        start_row = 1
        for i, value in enumerate(G函数压力序列):
            target_sheet[f'{chr(ord(start_col) + i)}{start_row}'] = value
    # 保存目标文件
    target_wb.save(target)
    target_wb.close()
    print('数据整理处理完成！')

#泵注曲线图片生成
def fig_solve(target):
    Initialize(plt)
    print("图片生成开始运行")
    try:
        path = "figs"
        os.makedirs(path)
    except FileExistsError:
        print("图片文件夹已存在，继续运行")

    path=target
    wb = load_workbook(path)

    sheets = wb.sheetnames

    for i in range (0,len(sheets)):
        print(f"正在进行第{i+1}张图片处理")
        sheets = wb.sheetnames
        ws = wb[sheets[i]] #wb[i] i为sheet名称
        wsl=count(ws)
        dt1=np.dtype(np.float64)
        dt2=np.dtype(np.int64)
        压力_=ws['C'][1:wsl]
        排量_=ws['D'][1:wsl]
        砂比_=ws['E'][1:wsl]
        压力=np.array(list(cell.value for cell in 压力_),dt1)    #将结构体转换为np.array，为后面matplotlib画图做准备
        排量=np.array(list(cell.value for cell in 排量_),dt1)   
        砂比=np.array(list(cell.value for cell in 砂比_),dt1)   
        sec=np.array(list(i for i in range(1,wsl)),dt2)
        min=sec/60 #计算横轴长度(分钟)

        fig=plt.figure()
        ax1=fig.subplots()
        ax2=ax1.twinx()
        ax1.plot(min, 压力,'r',label="施工压力(MPa)")
        ax1.plot(min, 排量,'g',label="排量(m$^3$/min)")
        ax1.plot(min, 砂比,'black',alpha=0.5,label="砂比(%)")
        ax1.set_xlabel("泵注时间(min)")
        ax1.set_ylabel("施工压力(MPa)\n排量(m$^3$/min)")
        ax2.plot(min, 压力,'black',alpha=0,label="砂比(%)") #对其左右
        ax2.set_ylabel("砂比(%)")
        ax1.set_yticks(np.arange(0, 81, 10))
        ax2.set_yticks(np.arange(0, 81, 10))
        ax1.legend()
        ax1.grid(True)
        plt.savefig(f'figs/BZfig{i+1}',dpi=600)
    print("图片处理完成")
#输入泵注曲线，输出压排比文件
def getsheets(target):
    wb = load_workbook(target)
    sheets = wb.sheetnames
    return sheets

def YP_solve(source,target):
    wb = load_workbook(source)
    sheets = wb.sheetnames
    st=0
    ed=len(sheets)
    砂比标题格=['E',2]
    # 读取Excel文件

    try:
        target_wb = load_workbook(target)
    except FileNotFoundError:
        # 如果文件不存在则创建新的工作簿
        target_wb = Workbook()
        # target_wb.save(目标文件) # 第一次创建时不需要立即保存，最后统一保存
        # target_wb = load_workbook(目标文件)# 重新加载工作簿 # 也不需要重新加载

    # 删除默认创建的工作表 (如果存在且是新创建的工作簿)
    # if 'Sheet' in target_wb.sheetnames and len(target_wb.sheetnames) == 1:
    #     del target_wb['Sheet']
    # 如果文件是新创建的，并且有默认的'Sheet'，可以删除它
    if 'Sheet' in target_wb.sheetnames and target_wb.sheetnames.index('Sheet') == 0 and len(target_wb.sheetnames) > 1:
        # 如果'Sheet'是第一个工作表且不止一个工作表，说明可能是加载旧文件时保留的，不删除
        pass
    elif 'Sheet' in target_wb.sheetnames:
        # 如果是新创建的工作簿，或者'Sheet'不是第一个工作表，则删除
        del target_wb['Sheet']
        
    for i in range(st, ed):
        ws = wb[sheets[i]]
        print('正在处理：', i+1)
        filter_sheet_by_column(ws, target_wb, str(i+1), 砂比标题格[0], 砂比标题格[1])#标题行：F10
    # 在循环结束后统一保存目标文件
    target_wb.save(target)
    target_wb.close()
    print('处理完成！')
    
def YP_fig(target):
    Initialize(plt)
    print("图片生成开始运行")
    try:
        path = "figs"
        os.makedirs(path)
    except FileExistsError:
        print("图片文件夹已存在，继续运行")

    path=target
    wb = load_workbook(path)

    sheets = wb.sheetnames

    for i in range (0,len(sheets)):
        print(f"正在进行第{i+1}张图片处理")
        sheets = wb.sheetnames
        ws = wb[sheets[i]] #wb[i] i为sheet名称
        wsl=count(ws)
        dt1=np.dtype(np.float64)
        dt2=np.dtype(np.int64)
        压力_=ws['C'][1:wsl]
        排量_=ws['D'][1:wsl]
        砂比_=ws['E'][1:wsl]
        压力=np.array(list(cell.value for cell in 压力_),dt1)    #将结构体转换为np.array，为后面matplotlib画图做准备
        排量=np.array(list(cell.value for cell in 排量_),dt1)   
        砂比=np.array(list(cell.value for cell in 砂比_),dt1)   
        sec=np.array(list(i for i in range(1,wsl)),dt2)
        min=sec/60 #计算横轴长度(分钟)
        压排比=压力/排量
        
        fig=plt.figure()
        ax1=fig.subplots()
        ax2=ax1.twinx()
        ax1.plot(min, 压排比,'r',label="压排比(MPa/(m$^3$/min))")
        ax2.plot(min, 砂比,'black',alpha=0.5,label="砂比(%)")
        ax1.set_xlabel("泵注时间(min)")
        ax1.set_ylabel("压力(MPa)/排量(m$^3$/min)")
        ax2.set_ylabel("砂比(%)")
        
        ax1.set_ylim(0,5)
        ax2.set_ylim(0,100)
        ax1.set_yticks(np.arange(0, 5.1, 0.5))
        ax2.set_yticks(np.arange(0, 101, 10))
        #设置标签用
        line1, = plt.plot([1],'r',label="压排比(MPa/(m$^3$/min))")
        line2, = plt.plot([1],'black',alpha=0.5,label="砂比(%)")
        ax1.legend(handles = [line1, line2])
        ax1.grid(True)
        plt.savefig(f'figs/YPfig{i+1}',dpi=600)
    print("图片处理完成")

def fig_doc(target):
    path=target
    wb = load_workbook(path)
    sheets = wb.sheetnames

    doc = Document()
    heading = doc.add_heading(level=2)
    run = heading.add_run("3.5 施工曲线图")
    font=run.font
    font.name = 'Times New Roman' 
    font.size = Pt(14) 
    font.color.rgb = RGBColor(0, 0, 0) 
    run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    for i in range (0,len(sheets)):
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run()
        run.add_picture(f'figs/BZfig{i+1}.png',width=Cm(12))

        para = doc.add_paragraph()
        paragraph_format = para.paragraph_format
        paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = para.add_run(f"图3-5-{i+1}    第{i+1}级泵注曲线")
        setfont(run)
        font = run.font
        run.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
    doc.save('figs/施工曲线图.docx')
    print("文件已生成")

def TBsolve(source):
    target = "停泵曲线.xlsx"
    #工作簿='CMG2-12L井泵注曲线3.xlsx'

    # wb=load_workbook(工作簿)
    try:
        wb = load_workbook(source)
    except FileNotFoundError:
        # 如果文件不存在则创建新的工作簿
        print(f"文件 {source} 不存在")
    try:
        twb = load_workbook(target)
    except FileNotFoundError:
        # 如果文件不存在则创建新的工作簿
        twb = Workbook()
        twb.save(target)
        twb = load_workbook(target)# 重新加载工作簿

    sheets = wb.sheetnames
    tsheets = twb.sheetnames
    段数=len(sheets)

    ws_t=twb[tsheets[0]]
    ws_t['A1'].value='时间'
    ws_t['B1'].value='压力'

    k=2
    for j in range(段数):
        print(j+1)
        ws_s=wb[sheets[j]]
        cnt = ws_s.max_row - ws_s[f'I{1}'].value
        for i in range(cnt):  
            ws_t[f'B{k}'].value=ws_s[f'C{i+ws_s[f'I{1}'].value}'].value
            k+=1
        k+=50

    for i in range(ws_t.max_row):
        ws_t[f'A{i+2}'].value=i+1
        if ws_t[f'B{i+1}'].value == 0  :
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
    plt.savefig(f'figs/TBfig',dpi=600)
    wb.close()
    print("图片处理完成")



    







    
    
