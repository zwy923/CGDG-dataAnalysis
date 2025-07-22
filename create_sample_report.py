import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import os
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
plt.rcParams['axes.unicode_minus'] = False

def create_sample_report():
    """创建示例报告，基于5月1日数据"""
    
    # 读取5月1日的数据
    base_path = "2025年5月交易日报/5月1日（改预测曲线，多云，14：00-15：30平均发电指标18.45MW，最大限电电力19MW，限电量1.97万kWh。）/数据"
    
    # 读取日前日内数据
    day_ahead_file = os.path.join(base_path, "1、日前日内5.1.xls")
    prediction_file = os.path.join(base_path, "3、网厂平台导出预测数据2025.5.2（5.1-5.8）.xls")
    
    # 创建PDF
    pdf_filename = "风电场交易报告示例.pdf"
    
    with PdfPages(pdf_filename) as pdf:
        # 处理数据
        try:
            # 读取日前日内数据
            df_day_ahead = pd.read_excel(day_ahead_file, sheet_name='Sheet1')
            df_day_ahead = df_day_ahead.dropna(subset=['时段'])
            
            # 读取预测数据
            df_prediction = pd.read_excel(prediction_file, sheet_name='20250501')
            df_prediction.columns = ['时段', '预测出力', '实际出力', '可用功率', '限电时段']
            df_prediction['限电量'] = np.maximum(0, df_prediction['预测出力'] - df_prediction['实际出力'])
            
            # 创建图表
            fig = plt.figure(figsize=(16, 24))
            fig.suptitle('中电装备北镇市风电场交易日报 - 2025年5月1日', fontsize=20, fontweight='bold', y=0.96)
            
            # 1. 创建顶部汇总表格
            create_top_summary_table(fig, df_day_ahead, df_prediction)
            
            # 2. 创建电量统计图
            create_power_statistics_chart(fig, df_day_ahead, df_prediction)
            
            # 3. 创建电价统计图
            create_price_statistics_chart(fig, df_day_ahead)
            
            # 4. 创建收益统计图
            create_revenue_statistics_chart(fig, df_day_ahead, df_prediction)
            
            # 调整布局
            plt.tight_layout()
            
            # 保存到PDF
            pdf.savefig(fig, bbox_inches='tight', dpi=300)
            plt.close(fig)
            
            print(f"📊 示例报告已生成: {pdf_filename}")
            
        except Exception as e:
            print(f"创建示例报告时出错: {e}")

def create_top_summary_table(fig, df_day_ahead, df_prediction):
    """创建顶部汇总表格"""
    ax = fig.add_subplot(6, 1, 1)
    ax.axis('off')
    
    # 计算滚动撮合数据
    if len(df_day_ahead) > 0:
        # 模拟滚动撮合交易数据
        table_data = [
            ['中长期持仓', '中长期价格', '滚动撮合电量1', '滚动撮合电价1', '滚动撮合电量2', '滚动撮合电价2', '滚动撮合电量3', '滚动撮合电价3', '滚动撮合电量4', '滚动撮合电价4', '总合约电量', '总合约电价'],
            ['兆瓦时', '元/兆瓦时', '兆瓦时', '元/兆瓦时', '兆瓦时', '元/兆瓦时', '兆瓦时', '元/兆瓦时', '兆瓦时', '元/兆瓦时', '兆瓦时', '元/兆瓦时'],
            ['451.53', '374.66', '31.44', '27.24', '', '', '', '', '', '', '482.97', '352.04']
        ]
        
        # 创建表格
        table = ax.table(cellText=table_data,
                        cellLoc='center',
                        loc='center',
                        colWidths=[0.08] * 12)
        table.auto_set_font_size(False)
        table.set_fontsize(9)
        table.scale(1, 2)
        
        # 设置表格样式
        for i in range(len(table_data)):
            for j in range(12):
                cell = table[(i, j)]
                if i == 0:  # 标题行
                    cell.set_facecolor('#E3F2FD')
                    cell.set_text_props(weight='bold')
                elif i == 1:  # 单位行
                    cell.set_facecolor('#F5F5F5')
                else:
                    cell.set_facecolor('white')
                    
        ax.set_title('滚动撮合交易汇总', fontsize=14, fontweight='bold', pad=20)

def create_power_statistics_chart(fig, df_day_ahead, df_prediction):
    """创建电量统计图 - 图1"""
    ax = fig.add_subplot(6, 1, 2)
    
    # 创建时间轴
    time_points = range(len(df_day_ahead))
    
    # 绘制各种电量曲线
    ax.plot(time_points, df_day_ahead['日前结果(MW)'], 
           label='日前电量', color='blue', linewidth=2, linestyle='-')
    ax.plot(time_points, df_day_ahead['实时结果(MW)'], 
           label='实时电量', color='green', linewidth=2, linestyle='-')
    
    # 如果预测数据长度匹配，添加预测曲线
    if len(df_prediction) == len(df_day_ahead):
        ax.plot(time_points, df_prediction['预测出力'], 
               label='预测出力', color='orange', linewidth=2, linestyle='--')
        ax.plot(time_points, df_prediction['实际出力'], 
               label='实际出力', color='red', linewidth=2, linestyle='-')
    
    # 添加限电区域填充
    if len(df_prediction) == len(df_day_ahead):
        # 找到14:00-15:30的限电区域
        start_idx = 56  # 14:00
        end_idx = 62    # 15:30
        
        if start_idx < len(time_points) and end_idx < len(time_points):
            ax.fill_between(time_points[start_idx:end_idx], 
                           0, max(df_day_ahead['日前结果(MW)']) * 1.1,
                           alpha=0.2, color='orange', label='限电时段')
    
    ax.set_title('图1：合约、撮合、日前、实际电量统计', fontsize=14, fontweight='bold')
    ax.set_xlabel('时段', fontsize=12)
    ax.set_ylabel('电量(MW)', fontsize=12)
    ax.legend(loc='upper right', fontsize=10)
    ax.grid(True, alpha=0.3)
    
    # 设置x轴标签
    step = max(1, len(time_points) // 10)
    ax.set_xticks(time_points[::step])
    ax.set_xticklabels([f"{i//4:02d}:{(i%4)*15:02d}" for i in time_points[::step]], 
                       rotation=45)

def create_price_statistics_chart(fig, df_day_ahead):
    """创建电价统计图 - 图2"""
    ax = fig.add_subplot(6, 1, 3)
    
    time_points = range(len(df_day_ahead))
    
    # 绘制电价曲线
    ax.plot(time_points, df_day_ahead['日前价格(元/MWh)'], 
           label='日前电价', color='blue', linewidth=2)
    ax.plot(time_points, df_day_ahead['实时价格(元/MWh)'], 
           label='实时电价', color='red', linewidth=2)
    
    # 添加电价区间填充
    ax.fill_between(time_points, 
                   df_day_ahead['日前价格(元/MWh)'], 
                   df_day_ahead['实时价格(元/MWh)'], 
                   alpha=0.2, color='lightblue', label='价差区间')
    
    ax.set_title('图2：电价统计图', fontsize=14, fontweight='bold')
    ax.set_xlabel('时段', fontsize=12)
    ax.set_ylabel('电价(元/MWh)', fontsize=12)
    ax.legend(loc='upper right', fontsize=10)
    ax.grid(True, alpha=0.3)
    
    # 设置x轴标签
    step = max(1, len(time_points) // 10)
    ax.set_xticks(time_points[::step])
    ax.set_xticklabels([f"{i//4:02d}:{(i%4)*15:02d}" for i in time_points[::step]], 
                       rotation=45)

def create_revenue_statistics_chart(fig, df_day_ahead, df_prediction):
    """创建收益统计图 - 图3"""
    ax = fig.add_subplot(6, 1, 4)
    
    time_points = range(len(df_day_ahead))
    
    # 计算各类收益
    day_ahead_revenue = df_day_ahead['日前结果(MW)'] * df_day_ahead['日前价格(元/MWh)']
    realtime_revenue = df_day_ahead['实时结果(MW)'] * df_day_ahead['实时价格(元/MWh)']
    
    # 计算合约收益（示例）
    contract_revenue = day_ahead_revenue * 0.8  # 假设合约收益为日前收益的80%
    
    # 显示每4个时段的数据
    x_pos = list(range(0, len(time_points), 4))
    width = 0.25
    
    # 绘制柱状图
    bars1 = ax.bar([x - width for x in x_pos], contract_revenue[::4], 
                   width, label='合约收益', color='lightblue', alpha=0.7)
    bars2 = ax.bar([x for x in x_pos], day_ahead_revenue[::4], 
                   width, label='日前收益', color='lightgreen', alpha=0.7)
    bars3 = ax.bar([x + width for x in x_pos], realtime_revenue[::4], 
                   width, label='实时收益', color='lightcoral', alpha=0.7)
    
    # 添加总收入曲线
    total_revenue = contract_revenue + day_ahead_revenue + realtime_revenue
    ax2 = ax.twinx()
    ax2.plot(time_points, total_revenue, color='red', linewidth=3, label='总收入')
    ax2.set_ylabel('总收入(元)', fontsize=12)
    
    ax.set_title('图3：收益统计图', fontsize=14, fontweight='bold')
    ax.set_xlabel('时段', fontsize=12)
    ax.set_ylabel('收益(元)', fontsize=12)
    
    # 合并图例
    lines1, labels1 = ax.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax.legend(lines1 + lines2, labels1 + labels2, loc='upper right', fontsize=10)
    
    ax.grid(True, alpha=0.3)
    
    # 设置x轴标签
    step = max(1, len(x_pos) // 8)
    ax.set_xticks(x_pos[::step])
    ax.set_xticklabels([f"{i//4:02d}:{(i%4)*15:02d}" for i in x_pos[::step]], 
                       rotation=45)

if __name__ == "__main__":
    create_sample_report() 