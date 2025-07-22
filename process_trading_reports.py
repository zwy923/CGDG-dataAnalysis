#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
2025年5月交易日报处理脚本
从每日Excel文件的"统计图"表中提取"日前结算收益"列的负值并求和
"""

import pandas as pd
import glob
import os
from datetime import datetime
import warnings


def extract_negative_revenue(file_path):
    """
    从单个Excel文件中提取日前结算收益列的负值并求和
    
    Args:
        file_path: Excel文件路径
        
    Returns:
        tuple: (日期, 负值总和, 成功标志)
    """
    try:
        filename = os.path.basename(file_path)
        # 文件名格式: "2025年大市风电场现货交易日报(5月X日).xlsx"
        date_part = filename.split('(')[1].split(')')[0]  # 提取"5月X日"部分
        
        print(f"处理文件: {filename}")
        print(f"  提取日期: {date_part}")
        
        # 读取Excel文件的"统计图"工作表
        df = pd.read_excel(file_path, sheet_name="统计图", header=0)
        
        # 第一行包含真正的列标题，需要跳过第一行来获取数据
        # 重新读取，使用第一行作为header
        df = pd.read_excel(file_path, sheet_name="统计图", header=1)
        
        # 检查是否有"日前结算收益"列
        revenue_columns = [col for col in df.columns if "日前结算收益" in str(col)]
        
        if not revenue_columns:
            # 如果列名不匹配，尝试按位置获取（第15列，索引为14）
            if len(df.columns) > 14:
                revenue_col = df.iloc[:, 14]  # 第15列
                print(f"  按位置使用第15列作为日前结算收益列")
            else:
                print(f"  错误: 未找到日前结算收益列，总列数: {len(df.columns)}")
                return date_part, 0, False
        else:
            revenue_col = df[revenue_columns[0]]
            print(f"  找到日前结算收益列: {revenue_columns[0]}")
        
        # 将列转换为数值类型，非数值的转为NaN
        revenue_numeric = pd.to_numeric(revenue_col, errors='coerce')
        
        # 筛选出负值
        negative_values = revenue_numeric[revenue_numeric < 0]
        
        # 计算负值总和
        negative_sum = negative_values.sum()
        
        print(f"  负值数量: {len(negative_values)}")
        print(f"  负值总和: {negative_sum:.2f}")
        
        if len(negative_values) > 0:
            print(f"  负值范围: {negative_values.min():.2f} 到 {negative_values.max():.2f}")
        
        return date_part, negative_sum, True
        
    except Exception as e:
        print(f"  错误: {e}")
        return date_part if 'date_part' in locals() else "未知日期", 0, False

def process_all_files():
    """处理所有交易日报文件"""
    print("=" * 60)
    print("2025年5月交易日报 - 日前结算收益负值统计")
    print("=" * 60)
    
    # 查找所有Excel文件
    pattern = "2025年5月交易日报/*/2025年大市风电场现货交易日报*.xlsx"
    files = glob.glob(pattern)
    
    if not files:
        print("错误: 未找到匹配的Excel文件")
        print(f"搜索模式: {pattern}")
        return
    
    print(f"找到 {len(files)} 个Excel文件")
    print()
    
    # 存储结果
    daily_results = []
    total_negative_sum = 0
    successful_files = 0
    
    # 处理每个文件
    for file_path in sorted(files):
        date, negative_sum, success = extract_negative_revenue(file_path)
        
        if success:
            daily_results.append({
                'date': date,
                'negative_sum': negative_sum,
                'file_path': file_path
            })
            total_negative_sum += negative_sum
            successful_files += 1
        
        print()  # 空行分隔
    
    # 输出汇总结果
    print("=" * 60)
    print("汇总结果")
    print("=" * 60)
    print(f"成功处理文件数: {successful_files}/{len(files)}")
    print(f"2025年5月总的日前结算收益负值总和: {total_negative_sum:.2f}")
    print()
    
    # 输出每日详细结果
    print("每日详细结果:")
    print("-" * 40)
    for result in sorted(daily_results, key=lambda x: x['date']):
        print(f"{result['date']:10s}: {result['negative_sum']:12.2f}")
    
    print("-" * 40)
    print(f"{'月度总计':10s}: {total_negative_sum:12.2f}")
    
    # 保存结果到CSV文件
    if daily_results:
        df_results = pd.DataFrame(daily_results)
        df_results['negative_sum_formatted'] = df_results['negative_sum'].round(2)
        
        output_file = "2025年5月日前结算收益负值统计.csv"
        df_results[['date', 'negative_sum_formatted']].to_csv(
            output_file, 
            index=False, 
            encoding='utf-8-sig',
            columns=['date', 'negative_sum_formatted'],
            header=['日期', '日前结算收益负值总和']
        )
        print(f"\n结果已保存到: {output_file}")

if __name__ == "__main__":
    process_all_files() 