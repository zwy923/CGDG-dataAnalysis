#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
电力现货交易日前结算收益分析器
功能：分析Excel文件中的日前结算收益负值，支持按日、月、年查询
新增功能：原合约优化分析
"""

import pandas as pd
import os
import sys
import re
import numpy as np
from datetime import datetime
from collections import defaultdict
import argparse

class PowerSettlementAnalyzer:
    def __init__(self, data_dir='.'):
        self.data_dir = data_dir
        self.data = {}  # 存储所有数据：{date: negative_sum}
        self.detailed_data = {}  # 存储详细数据：{date: dataframe}
        
    def extract_date_from_filename(self, filename):
        """从文件名中提取日期"""
        # 匹配模式：2025年大市风电场现货交易日报(5月10日)
        pattern = r'(\d{4})年.*?(\d{1,2})月(\d{1,2})日'
        match = re.search(pattern, filename)
        if match:
            year, month, day = match.groups()
            try:
                return datetime(int(year), int(month), int(day))
            except ValueError:
                return None
        return None
    
    def process_excel_file(self, filepath):
        """处理单个Excel文件，提取日前结算收益负值总和"""
        try:
            xl_file = pd.ExcelFile(filepath)
            if '统计图' not in xl_file.sheet_names:
                print(f"警告：{filepath} 中未找到'统计图'工作表")
                return None
            
            # 读取数据，第一行作为真正的列名
            df = pd.read_excel(filepath, sheet_name='统计图', header=0)
            
            # 获取真正的列名（第一行的值）
            real_column_names = df.iloc[0].tolist()
            df.columns = [str(col) if pd.notna(col) else f'Unnamed_{i}' 
                         for i, col in enumerate(real_column_names)]
            
            # 删除第一行（现在已经是列名）
            df = df.iloc[1:].reset_index(drop=True)
            
            # 查找日前结算收益列
            settlement_col = None
            for col in df.columns:
                if '日前结算收益' in str(col):
                    settlement_col = col
                    break
            
            if settlement_col is None:
                print(f"警告：{filepath} 中未找到'日前结算收益'列")
                return None
            
            # 转换为数值并计算负值总和
            numeric_values = pd.to_numeric(df[settlement_col], errors='coerce')
            negative_values = numeric_values[numeric_values < 0]
            negative_sum = negative_values.sum() if len(negative_values) > 0 else 0
            
            return negative_sum
            
        except Exception as e:
            print(f"处理文件 {filepath} 时出错: {e}")
            return None
    
    def load_detailed_data(self, filepath):
        """加载详细数据用于原合约优化分析"""
        try:
            xl_file = pd.ExcelFile(filepath)
            if '统计图' not in xl_file.sheet_names:
                return None
            
            # 读取数据，第一行作为真正的列名
            df = pd.read_excel(filepath, sheet_name='统计图', header=0)
            
            # 获取真正的列名（第一行的值）
            real_column_names = df.iloc[0].tolist()
            df.columns = [str(col) if pd.notna(col) else f'Unnamed_{i}' 
                         for i, col in enumerate(real_column_names)]
            
            # 删除第一行（现在已经是列名）
            df = df.iloc[1:].reset_index(drop=True)
            
            return df
            
        except Exception as e:
            print(f"加载详细数据时出错: {e}")
            return None
    
    def calculate_total_revenue(self, df, original_contract):
        """计算总收入"""
        try:
            # 查找所需的列
            contract_price_col = None
            forward_price_col = None
            
            for col in df.columns:
                col_str = str(col).lower()
                if '合约电价' in col_str:
                    contract_price_col = col
                elif '日前电价' in col_str:
                    forward_price_col = col
            
            if contract_price_col is None or forward_price_col is None:
                print("未找到必要的电价列")
                return None
            
            # 转换为数值
            contract_prices = pd.to_numeric(df[contract_price_col], errors='coerce')
            forward_prices = pd.to_numeric(df[forward_price_col], errors='coerce')
            
            # 计算总收入：原合约*合约电价 - 原合约*日前电价
            total_revenue = original_contract * contract_prices - original_contract * forward_prices
            
            return total_revenue
            
        except Exception as e:
            print(f"计算总收入时出错: {e}")
            return None
    
    def optimize_original_contract(self, df, contract_range=(0, 12)):
        """优化原合约使总收入最大"""
        try:
            # 查找合约电价和日前电价列
            contract_price_col = None
            forward_price_col = None
            
            for col in df.columns:
                col_str = str(col).lower()
                if '合约电价' in col_str:
                    contract_price_col = col
                elif '日前电价' in col_str:
                    forward_price_col = col
            
            if contract_price_col is None or forward_price_col is None:
                print("未找到必要的电价列")
                return None, None
            
            # 转换为数值
            contract_prices = pd.to_numeric(df[contract_price_col], errors='coerce')
            forward_prices = pd.to_numeric(df[forward_price_col], errors='coerce')
            
            # 计算价格差
            price_diff = contract_prices - forward_prices
            
            # 优化原合约
            # 如果价格差为正，选择最大合约值；如果为负，选择最小合约值
            optimal_contract = np.where(price_diff > 0, contract_range[1], contract_range[0])
            
            # 计算最优总收入
            optimal_revenue = optimal_contract * price_diff
            
            return optimal_contract, optimal_revenue
            
        except Exception as e:
            print(f"优化原合约时出错: {e}")
            return None, None
    
    def get_monthly_contract_value(self, year, month):
        """获取指定月份的原合约固定值"""
        # 这里可以根据实际业务逻辑设置每月的固定值
        # 示例：5月的每天定值为451.527
        monthly_values = {
            5: 451.527,
            # 可以添加其他月份的值
        }
        return monthly_values.get(month, 0)
    
    def analyze_contract_optimization(self, target_date):
        """分析指定日期的原合约优化"""
        if isinstance(target_date, str):
            try:
                target_date = datetime.strptime(target_date, '%Y-%m-%d')
            except ValueError:
                print("日期格式错误，请使用 YYYY-MM-DD 格式")
                return None
        
        # 查找对应的Excel文件
        excel_files = [f for f in os.listdir(self.data_dir) if f.endswith('.xlsx')]
        target_file = None
        
        for filename in excel_files:
            date = self.extract_date_from_filename(filename)
            if date and date == target_date:
                target_file = filename
                break
        
        if target_file is None:
            print(f"未找到 {target_date.strftime('%Y-%m-%d')} 的数据文件")
            return None
        
        filepath = os.path.join(self.data_dir, target_file)
        df = self.load_detailed_data(filepath)
        
        if df is None:
            print("无法加载详细数据")
            return None
        
        # 优化原合约
        optimal_contract, optimal_revenue = self.optimize_original_contract(df)
        
        if optimal_contract is None:
            return None
        
        # 获取月度固定值
        monthly_value = self.get_monthly_contract_value(target_date.year, target_date.month)
        
        result = {
            'date': target_date,
            'data': df,
            'optimal_contract': optimal_contract,
            'optimal_revenue': optimal_revenue,
            'monthly_fixed_value': monthly_value,
            'total_optimal_revenue': optimal_revenue.sum(),
            'avg_optimal_contract': optimal_contract.mean()
        }
        
        return result
    
    def get_monthly_average_contract(self, year, month):
        """获取指定月份每天每个15分钟时间段的原合约数值均值"""
        monthly_data = []
        
        # 获取该月所有日期的数据
        for date, _ in self.data.items():
            if date.year == year and date.month == month:
                result = self.analyze_contract_optimization(date)
                if result and result['optimal_contract'] is not None:
                    monthly_data.append(result['optimal_contract'])
        
        if not monthly_data:
            return None
        
        # 计算均值
        monthly_avg = np.mean(monthly_data, axis=0)
        return monthly_avg
    
    def load_all_data(self):
        """加载所有Excel文件数据"""
        print("正在加载数据...")
        excel_files = [f for f in os.listdir(self.data_dir) if f.endswith('.xlsx')]
        
        for filename in excel_files:
            date = self.extract_date_from_filename(filename)
            if date is None:
                print(f"无法从文件名 {filename} 中提取日期")
                continue
            
            filepath = os.path.join(self.data_dir, filename)
            negative_sum = self.process_excel_file(filepath)
            
            if negative_sum is not None:
                self.data[date] = negative_sum
                print(f"已处理：{date.strftime('%Y-%m-%d')} -> {negative_sum:.2f}")
        
        print(f"\n总共加载了 {len(self.data)} 个文件的数据")
    
    def query_by_date(self, target_date):
        """查询指定日期的数据"""
        if isinstance(target_date, str):
            try:
                target_date = datetime.strptime(target_date, '%Y-%m-%d')
            except ValueError:
                print("日期格式错误，请使用 YYYY-MM-DD 格式")
                return None
        
        if target_date in self.data:
            return self.data[target_date]
        else:
            print(f"未找到 {target_date.strftime('%Y-%m-%d')} 的数据")
            return None
    
    def query_by_month(self, year, month):
        """查询指定月份的数据总和"""
        monthly_sum = 0
        count = 0
        
        for date, value in self.data.items():
            if date.year == year and date.month == month:
                monthly_sum += value
                count += 1
        
        return monthly_sum, count
    
    def query_by_year(self, year):
        """查询指定年份的数据总和"""
        yearly_sum = 0
        count = 0
        
        for date, value in self.data.items():
            if date.year == year:
                yearly_sum += value
                count += 1
        
        return yearly_sum, count
    
    def get_all_dates(self):
        """获取所有可用的日期"""
        return sorted(self.data.keys())
    
    def get_monthly_summary(self):
        """获取月度汇总"""
        monthly_data = defaultdict(lambda: {'sum': 0, 'count': 0, 'dates': []})
        
        for date, value in self.data.items():
            key = (date.year, date.month)
            monthly_data[key]['sum'] += value
            monthly_data[key]['count'] += 1
            monthly_data[key]['dates'].append(date.strftime('%Y-%m-%d'))
        
        return dict(monthly_data)
    
    def get_yearly_summary(self):
        """获取年度汇总"""
        yearly_data = defaultdict(lambda: {'sum': 0, 'count': 0})
        
        for date, value in self.data.items():
            yearly_data[date.year]['sum'] += value
            yearly_data[date.year]['count'] += 1
        
        return dict(yearly_data)

def main():
    parser = argparse.ArgumentParser(description='电力现货交易日前结算收益分析器')
    parser.add_argument('--date', '-d', help='查询指定日期 (YYYY-MM-DD)')
    parser.add_argument('--month', '-m', help='查询指定月份 (YYYY-MM)')
    parser.add_argument('--year', '-y', type=int, help='查询指定年份 (YYYY)')
    parser.add_argument('--all', '-a', action='store_true', help='显示所有数据汇总')
    parser.add_argument('--list', '-l', action='store_true', help='列出所有可用日期')
    
    args = parser.parse_args()
    
    # 创建分析器实例
    analyzer = PowerSettlementAnalyzer()
    analyzer.load_all_data()
    
    if not analyzer.data:
        print("没有找到任何有效数据")
        return
    
    # 根据参数执行查询
    if args.date:
        result = analyzer.query_by_date(args.date)
        if result is not None:
            print(f"\n{args.date} 的日前结算收益负值总和: {result:.2f}")
    
    elif args.month:
        try:
            year, month = map(int, args.month.split('-'))
            result, count = analyzer.query_by_month(year, month)
            print(f"\n{year}年{month}月的日前结算收益负值总和: {result:.2f}")
            print(f"包含 {count} 天的数据")
        except ValueError:
            print("月份格式错误，请使用 YYYY-MM 格式")
    
    elif args.year:
        result, count = analyzer.query_by_year(args.year)
        print(f"\n{args.year}年的日前结算收益负值总和: {result:.2f}")
        print(f"包含 {count} 天的数据")
    
    elif args.list:
        dates = analyzer.get_all_dates()
        print(f"\n可用日期列表 (共 {len(dates)} 天):")
        for date in dates:
            value = analyzer.data[date]
            print(f"{date.strftime('%Y-%m-%d')}: {value:.2f}")
    
    elif args.all:
        print("\n=== 数据汇总 ===")
        
        # 年度汇总
        yearly_summary = analyzer.get_yearly_summary()
        print("\n年度汇总:")
        for year in sorted(yearly_summary.keys()):
            data = yearly_summary[year]
            print(f"{year}年: {data['sum']:.2f} (共{data['count']}天)")
        
        # 月度汇总
        monthly_summary = analyzer.get_monthly_summary()
        print("\n月度汇总:")
        for (year, month) in sorted(monthly_summary.keys()):
            data = monthly_summary[(year, month)]
            print(f"{year}年{month}月: {data['sum']:.2f} (共{data['count']}天)")
    
    else:
        # 默认显示简单汇总
        print("\n=== 简单汇总 ===")
        total_sum = sum(analyzer.data.values())
        total_days = len(analyzer.data)
        print(f"总计: {total_sum:.2f} (共{total_days}天)")
        
        print(f"\n使用 --help 查看更多选项")
        print("示例用法:")
        print("  python power_analyzer.py --date 2025-05-10")
        print("  python power_analyzer.py --month 2025-05")
        print("  python power_analyzer.py --year 2025")
        print("  python power_analyzer.py --all")
        print("  python power_analyzer.py --list")

if __name__ == "__main__":
    main() 