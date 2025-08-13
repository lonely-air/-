from mymodel_v1_8 import *

def process_pump_injection_curve():
    """泵注曲线处理功能"""
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
    """压排比处理功能"""
    source = '泵注曲线.xlsx'
    target = '压排比.xlsx'
    print("\n开始处理压排比数据...")
    YP_solve(source, target)
    print("\n开始生成压排比图片...")
    YP_fig(target)
    print("\n压排比处理完成！")

def process_shutdown_curve():
    """停泵曲线处理功能"""
    source = '泵注曲线.xlsx'
    print("\n开始处理停泵曲线数据...")
    TBsolve(source)
    print("\n开始生成停泵曲线图片...")
    TBfig('停泵曲线.xlsx')
    print("\n停泵曲线处理完成！")

def main():
    while True:
        print("\n=== 压裂数据处理系统 ===")
        print("1. 处理泵注曲线")
        print("2. 处理压排比")
        print("3. 处理停泵曲线")
        print("4. 全流程处理")
        print("0. 退出程序")
        
        choice = input("\n请选择功能（输入数字）：")
        
        try:
            if choice == "1":
                process_pump_injection_curve()
            elif choice == "2":
                process_pressure_flow_ratio()
            elif choice == "3":
                process_shutdown_curve()
            elif choice == "4":
                print("\n开始全流程处理...")
                process_pump_injection_curve()
                process_pressure_flow_ratio()
                process_shutdown_curve()
                print("\n全流程处理完成！")
            elif choice == "0":
                print("\n程序已退出")
                break
            else:
                print("\n无效的选择，请重新输入！")
        except Exception as e:
            print(f"\n处理过程中出现错误：{str(e)}")
            print("请检查输入文件是否存在且格式正确！")

if __name__ == "__main__":
    main()
