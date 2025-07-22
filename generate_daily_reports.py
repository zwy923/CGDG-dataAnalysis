import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.backends.backend_pdf import PdfPages
import os
import re
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
plt.rcParams['axes.unicode_minus'] = False

def extract_date_from_folder(folder_name):
    """从文件夹名称提取日期"""
    match = re.search(r'5月(\d+)日', folder_name)
    if match:
        return int(match.group(1))
    return None

def find_data_files(base_path, date_str):
    """查找指定日期的数据文件"""
    data_folder = os.path.join(base_path, "数据")
    
    # 查找日前日内文件
    day_ahead_file = None
    prediction_file = None
    
    if os.path.exists(data_folder):
        for file in os.listdir(data_folder):
            if f"日前日内{date_str}" in file and file.endswith('.xls'):
                day_ahead_file = os.path.join(data_folder, file)
            elif "网厂平台导出预测数据" in file and file.endswith('.xls'):
                prediction_file = os.path.join(data_folder, file)
    
    return day_ahead_file, prediction_file

def process_day_ahead_data(file_path):
    """处理日前日内数据"""
    if not file_path or not os.path.exists(file_path):
        return None
    
    try:
        df = pd.read_excel(file_path, sheet_name='Sheet1')
        
        # 清理数据
        df = df.dropna(subset=['时段'])
        
        # 计算关键指标
        summary_data = {
            '日前总电量': df['日前结果(MW)'].sum()/4,
            '实时总电量': df['实时结果(MW)'].sum(),
            '日前平均价格': df['日前价格(元/MWh)'].mean(),
            '实时平均价格': df['实时价格(元/MWh)'].mean(),
            '价格差异': df['实时价格(元/MWh)'].mean() - df['日前价格(元/MWh)'].mean()
        }
        
        return df, summary_data
        
    except Exception as e:
        print(f"处理日前日内数据时出错: {e}")
        return None, None

def process_prediction_data(file_path, date_str):
    """处理预测数据"""
    if not file_path or not os.path.exists(file_path):
        return None
    
    try:
        # 构造工作表名称
        sheet_name = f"202505{date_str.zfill(2)}"
        
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # 重命名列
        df.columns = ['时段', '预测出力', '实际出力', '可用功率', '限电时段']
        
        # 计算限电量
        df['限电量'] = np.maximum(0, df['预测出力'] - df['实际出力'])
        
        # 计算关键指标
        summary_data = {
            '总预测出力': df['预测出力'].sum(),
            '总实际出力': df['实际出力'].sum(),
            '总限电量': df['限电量'].sum(),
            '限电电量_kWh': df['限电量'].sum() * 0.25,  # 转换为kWh
            '限电时段数': len(df[df['限电量'] > 0])
        }
        
        return df, summary_data
        
    except Exception as e:
        print(f"处理预测数据时出错: {e}")
        return None, None

def create_summary_table(day_ahead_summary, prediction_summary, date_str):
    """创建汇总表格"""
    table_data = []
    
    if day_ahead_summary and prediction_summary:
        table_data = [
            ['日期', f'2025年5月{date_str}日'],
            ['日前总电量(MW)', f"{day_ahead_summary['日前总电量']:.2f}"],
            ['实时总电量(MW)', f"{day_ahead_summary['实时总电量']:.2f}"],
            ['预测总出力(MW)', f"{prediction_summary['总预测出力']:.2f}"],
            ['实际总出力(MW)', f"{prediction_summary['总实际出力']:.2f}"],
            ['总限电量(MW)', f"{prediction_summary['总限电量']:.2f}"],
            ['限电电量(kWh)', f"{prediction_summary['限电电量_kWh']:.2f}"],
            ['日前平均价格(元/MWh)', f"{day_ahead_summary['日前平均价格']:.2f}"],
            ['实时平均价格(元/MWh)', f"{day_ahead_summary['实时平均价格']:.2f}"],
            ['价格差异(元/MWh)', f"{day_ahead_summary['价格差异']:.2f}"]
        ]
    
    return table_data

def create_power_chart(day_ahead_df, prediction_df, date_str):
    """创建电量统计图表"""
    fig, ax = plt.subplots(figsize=(12, 6))
    
    if day_ahead_df is not None and prediction_df is not None:
        # 处理时间数据
        try:
            # 创建时间序列
            time_points = range(len(day_ahead_df))
            
            # 绘制电量曲线
            ax.plot(time_points, day_ahead_df['日前结果(MW)'], 
                   label='日前电量', color='blue', linewidth=2)
            ax.plot(time_points, day_ahead_df['实时结果(MW)'], 
                   label='实时电量', color='green', linewidth=2)
            
            # 如果数据长度匹配，绘制预测和实际出力
            if len(prediction_df) == len(day_ahead_df):
                ax.plot(time_points, prediction_df['预测出力'], 
                       label='预测出力', color='orange', linewidth=2)
                ax.plot(time_points, prediction_df['实际出力'], 
                       label='实际出力', color='red', linewidth=2)
            
            # 设置图表属性
            ax.set_title(f'2025年5月{date_str}日 电量统计图', fontsize=14, fontweight='bold')
            ax.set_xlabel('时段', fontsize=12)
            ax.set_ylabel('电量(MW)', fontsize=12)
            ax.legend(fontsize=10)
            ax.grid(True, alpha=0.3)
            
            # 设置x轴标签
            if len(time_points) > 0:
                step = max(1, len(time_points) // 10)
                ax.set_xticks(time_points[::step])
                ax.set_xticklabels([f"{i//4:02d}:{(i%4)*15:02d}" for i in time_points[::step]], 
                                 rotation=45)
            
        except Exception as e:
            ax.text(0.5, 0.5, f'数据处理错误: {str(e)}', 
                   transform=ax.transAxes, ha='center', va='center')
    else:
        ax.text(0.5, 0.5, '数据不可用', 
               transform=ax.transAxes, ha='center', va='center')
    
    plt.tight_layout()
    return fig

def create_price_chart(day_ahead_df, date_str):
    """创建电价统计图表"""
    fig, ax = plt.subplots(figsize=(12, 6))
    
    if day_ahead_df is not None:
        try:
            time_points = range(len(day_ahead_df))
            
            # 绘制电价曲线
            ax.plot(time_points, day_ahead_df['日前价格(元/MWh)'], 
                   label='日前电价', color='blue', linewidth=2)
            ax.plot(time_points, day_ahead_df['实时价格(元/MWh)'], 
                   label='实时电价', color='red', linewidth=2)
            
            # 设置图表属性
            ax.set_title(f'2025年5月{date_str}日 电价统计图', fontsize=14, fontweight='bold')
            ax.set_xlabel('时段', fontsize=12)
            ax.set_ylabel('电价(元/MWh)', fontsize=12)
            ax.legend(fontsize=10)
            ax.grid(True, alpha=0.3)
            
            # 设置x轴标签
            if len(time_points) > 0:
                step = max(1, len(time_points) // 10)
                ax.set_xticks(time_points[::step])
                ax.set_xticklabels([f"{i//4:02d}:{(i%4)*15:02d}" for i in time_points[::step]], 
                                 rotation=45)
            
        except Exception as e:
            ax.text(0.5, 0.5, f'数据处理错误: {str(e)}', 
                   transform=ax.transAxes, ha='center', va='center')
    else:
        ax.text(0.5, 0.5, '数据不可用', 
               transform=ax.transAxes, ha='center', va='center')
    
    plt.tight_layout()
    return fig

def create_revenue_chart(day_ahead_df, prediction_df, date_str):
    """创建收益统计图表"""
    fig, ax = plt.subplots(figsize=(12, 6))
    
    if day_ahead_df is not None:
        try:
            time_points = range(len(day_ahead_df))
            
            # 计算收益（简化计算）
            day_ahead_revenue = day_ahead_df['日前结果(MW)'] * day_ahead_df['日前价格(元/MWh)']
            realtime_revenue = day_ahead_df['实时结果(MW)'] * day_ahead_df['实时价格(元/MWh)']
            
            # 绘制收益柱状图
            width = 0.35
            ax.bar([x - width/2 for x in time_points[::4]], day_ahead_revenue[::4], 
                  width, label='日前收益', color='lightblue', alpha=0.7)
            ax.bar([x + width/2 for x in time_points[::4]], realtime_revenue[::4], 
                  width, label='实时收益', color='lightcoral', alpha=0.7)
            
            # 设置图表属性
            ax.set_title(f'2025年5月{date_str}日 收益统计图', fontsize=14, fontweight='bold')
            ax.set_xlabel('时段', fontsize=12)
            ax.set_ylabel('收益(元)', fontsize=12)
            ax.legend(fontsize=10)
            ax.grid(True, alpha=0.3)
            
            # 设置x轴标签
            if len(time_points) > 0:
                step = max(1, len(time_points) // 10)
                ax.set_xticks(time_points[::step])
                ax.set_xticklabels([f"{i//4:02d}:{(i%4)*15:02d}" for i in time_points[::step]], 
                                 rotation=45)
            
        except Exception as e:
            ax.text(0.5, 0.5, f'数据处理错误: {str(e)}', 
                   transform=ax.transAxes, ha='center', va='center')
    else:
        ax.text(0.5, 0.5, '数据不可用', 
               transform=ax.transAxes, ha='center', va='center')
    
    plt.tight_layout()
    return fig

def generate_daily_report():
    """生成每日交易报告"""
    base_dir = "2025年5月交易日报"
    
    # 获取所有日期文件夹
    date_folders = []
    for folder in os.listdir(base_dir):
        if folder.startswith("5月") and "日" in folder:
            date_num = extract_date_from_folder(folder)
            if date_num:
                date_folders.append((date_num, folder))
    
    # 按日期排序
    date_folders.sort(key=lambda x: x[0])
    
    # 创建PDF文件
    pdf_filename = "风电场每日交易报告_2025年5月.pdf"
    
    with PdfPages(pdf_filename) as pdf:
        for date_num, folder_name in date_folders:
            print(f"正在处理 {folder_name}...")
            
            # 构造文件路径
            folder_path = os.path.join(base_dir, folder_name)
            date_str = str(date_num)
            
            # 查找数据文件
            day_ahead_file, prediction_file = find_data_files(folder_path, date_str)
            
            # 处理数据
            day_ahead_df, day_ahead_summary = process_day_ahead_data(day_ahead_file)
            prediction_df, prediction_summary = process_prediction_data(prediction_file, date_str)
            
            # 创建报告页面
            fig = plt.figure(figsize=(16, 20))
            
            # 添加标题
            fig.suptitle(f'中电装备北镇市风电场交易日报 - 2025年5月{date_str}日', 
                        fontsize=18, fontweight='bold', y=0.95)
            
            # 创建汇总表格
            table_data = create_summary_table(day_ahead_summary, prediction_summary, date_str)
            
            if table_data:
                # 表格区域
                ax_table = fig.add_subplot(4, 1, 1)
                ax_table.axis('off')
                
                # 创建表格
                table = ax_table.table(cellText=table_data,
                                     colLabels=['指标', '数值'],
                                     cellLoc='center',
                                     loc='center',
                                     colWidths=[0.4, 0.6])
                table.auto_set_font_size(False)
                table.set_fontsize(10)
                table.scale(1, 2)
                
                # 设置表格样式
                for i in range(len(table_data) + 1):
                    for j in range(2):
                        cell = table[(i, j)]
                        if i == 0:  # 标题行
                            cell.set_facecolor('#4CAF50')
                            cell.set_text_props(weight='bold', color='white')
                        else:
                            cell.set_facecolor('#f0f0f0' if i % 2 == 0 else 'white')
            
            # 创建图表
            try:
                # 电量统计图
                power_fig = create_power_chart(day_ahead_df, prediction_df, date_str)
                ax_power = fig.add_subplot(4, 1, 2)
                power_fig.savefig(ax_power.figure, bbox_inches='tight')
                plt.close(power_fig)
                
                # 电价统计图
                price_fig = create_price_chart(day_ahead_df, date_str)
                ax_price = fig.add_subplot(4, 1, 3)
                price_fig.savefig(ax_price.figure, bbox_inches='tight')
                plt.close(price_fig)
                
                # 收益统计图
                revenue_fig = create_revenue_chart(day_ahead_df, prediction_df, date_str)
                ax_revenue = fig.add_subplot(4, 1, 4)
                revenue_fig.savefig(ax_revenue.figure, bbox_inches='tight')
                plt.close(revenue_fig)
                
            except Exception as e:
                print(f"创建图表时出错: {e}")
            
            # 调整布局
            plt.tight_layout()
            
            # 保存到PDF
            pdf.savefig(fig, bbox_inches='tight', dpi=300)
            plt.close(fig)
            
            print(f"已完成 {folder_name} 的处理")
    
    print(f"\n报告已生成: {pdf_filename}")

if __name__ == "__main__":
    generate_daily_report() 