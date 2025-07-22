#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
原合约优化分析器
功能：分析Excel文件中的统计图表数据，优化原合约使总收入最大
约束条件：每日原合约总量不超过用户设定的限制值
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
import sys
import re
from datetime import datetime
from collections import defaultdict
from scipy.optimize import linprog

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

class ContractOptimizer:
    def __init__(self, data_dir='.'):
        self.data_dir = data_dir
        self.contract_range = (0, 12)  # 原合约取值范围
        self.daily_total_limit = None  # 每日原合约总量限制
        
    def set_daily_total_limit(self, limit):
        """设置每日原合约总量限制"""
        self.daily_total_limit = limit
        
    def extract_date_from_filename(self, filename):
        """从文件名中提取日期"""
        pattern = r'(\d{4})年.*?(\d{1,2})月(\d{1,2})日'
        match = re.search(pattern, filename)
        if match:
            year, month, day = match.groups()
            try:
                return datetime(int(year), int(month), int(day))
            except ValueError:
                return None
        return None
    
    def load_data(self, filepath):
        """加载Excel文件中的统计图表数据"""
        try:
            # 检查文件是否存在
            if not os.path.exists(filepath):
                print(f"文件不存在: {filepath}")
                return None
            
            # 检查文件大小
            if os.path.getsize(filepath) == 0:
                print(f"文件为空: {filepath}")
                return None
            
            xl_file = pd.ExcelFile(filepath)
            
            # 查找合适的工作表
            available_sheets = xl_file.sheet_names
            if '统计图' not in available_sheets:
                # 尝试找到包含"统计"或"图"的工作表
                similar_sheets = [sheet for sheet in available_sheets if '统计' in sheet or '图' in sheet]
                if similar_sheets:
                    sheet_name = similar_sheets[0]
                else:
                    sheet_name = available_sheets[0]
            else:
                sheet_name = '统计图'
            
            # 读取数据
            df = pd.read_excel(filepath, sheet_name=sheet_name, header=0)
            
            if df.empty:
                print(f"  警告：工作表为空")
                return None
            
            # 获取真正的列名（第一行的值）
            if len(df) > 0:
                real_column_names = df.iloc[0].tolist()
                df.columns = [str(col) if pd.notna(col) else f'Unnamed_{i}' 
                             for i, col in enumerate(real_column_names)]
                
                # 删除第一行（现在已经是列名）
                df = df.iloc[1:].reset_index(drop=True)
            
            # 移除完全空的行和列
            df = df.dropna(axis=0, how='all')
            df = df.dropna(axis=1, how='all')
            
            # 确保只取96个时间点（0:15-24:00）
            if len(df) > 96:
                # 如果数据超过96行，只取前96行（对应0:15-24:00）
                df = df.iloc[:96].reset_index(drop=True)
            elif len(df) < 96:
                print(f"  警告：数据不足96个时间点，实际{len(df)}个")
            
            if df.empty:
                print(f"  警告：处理后数据为空")
                return None
            
            return df
            
        except PermissionError:
            print(f"权限错误：无法访问文件 {filepath}")
            return None
        except Exception as e:
            print(f"加载数据时出错: {e}")
            return None
    
    def find_price_columns(self, df):
        """查找电价相关列"""
        columns = {}
        
        for col in df.columns:
            col_str = str(col).lower()
            if '合约电价' in col_str:
                columns['contract_price'] = col
            elif '日前电价' in col_str:
                columns['forward_price'] = col
            elif '撮合电价' in col_str:
                columns['matching_price'] = col
            elif '实时电价' in col_str:
                columns['realtime_price'] = col
            elif '省间现货电价' in col_str:
                columns['interprovincial_price'] = col
        
        return columns
    
    def find_volume_columns(self, df):
        """查找电量相关列"""
        columns = {}
        
        for col in df.columns:
            col_str = str(col).lower()
            if '滚动撮合电量' in col_str:
                columns['matching_volume'] = col
            elif '日前出清' in col_str:
                columns['forward_clearing'] = col
            elif '日内实际' in col_str:
                columns['realtime_actual'] = col
            elif '省间现货电量' in col_str:
                columns['interprovincial_volume'] = col
        
        return columns
    
    def calculate_total_revenue(self, df, original_contract):
        """计算总收入"""
        try:
            price_cols = self.find_price_columns(df)
            volume_cols = self.find_volume_columns(df)
            
            if 'contract_price' not in price_cols or 'forward_price' not in price_cols:
                print("未找到必要的电价列")
                return None
            
            # 转换为数值
            contract_prices = pd.to_numeric(df[price_cols['contract_price']], errors='coerce')
            forward_prices = pd.to_numeric(df[price_cols['forward_price']], errors='coerce')
            
            # 计算总收入：原合约*合约电价 - 原合约*日前电价
            total_revenue = original_contract * contract_prices - original_contract * forward_prices
            
            return total_revenue
            
        except Exception as e:
            print(f"计算总收入时出错: {e}")
            return None
    
    def print_optimal_values(self, optimal_contract, date_str):
        """打印96个时间点的最优原合约值"""
        print(f"\n=== {date_str} 96个时间点的最优原合约值 ===")
        print("时间点\t原合约值")
        print("-" * 20)
        
        for i, value in enumerate(optimal_contract):
            # 计算时间（每15分钟一个点）
            hour = i // 4
            minute = (i % 4) * 15
            time_str = f"{hour:02d}:{minute:02d}"
            print(f"{i+1:2d} ({time_str})\t{value:.3f}")
        
        print("-" * 20)
        print(f"总计:\t{np.sum(optimal_contract):.3f}")
        print(f"平均值:\t{np.mean(optimal_contract):.3f}")
        print(f"最大值:\t{np.max(optimal_contract):.3f}")
        print(f"最小值:\t{np.min(optimal_contract):.3f}")
    
    def optimize_contract_with_constraint(self, df):
        """在总量约束下优化原合约"""
        try:
            price_cols = self.find_price_columns(df)
            
            if 'contract_price' not in price_cols or 'forward_price' not in price_cols:
                print("未找到必要的电价列")
                return None, None
            
            # 转换为数值
            contract_prices = pd.to_numeric(df[price_cols['contract_price']], errors='coerce')
            forward_prices = pd.to_numeric(df[price_cols['forward_price']], errors='coerce')
            
            # 检查数据有效性
            if contract_prices.isnull().all() or forward_prices.isnull().all():
                print("警告：电价数据全部无效")
                return None, None
            
            # 计算价格差 (收益系数)
            price_diff = contract_prices - forward_prices
            
            # 处理无效值：将NaN、inf替换为0
            price_diff = price_diff.fillna(0)
            price_diff = price_diff.replace([np.inf, -np.inf], 0)
            
            # 检查处理后的数据
            if len(price_diff) == 0:
                print("警告：没有有效的价格差数据")
                return None, None
            
            # 如果没有设置总量限制，使用原来的简单方法
            if self.daily_total_limit is None:
                optimal_contract = np.where(price_diff > 0, self.contract_range[1], self.contract_range[0])
                optimal_revenue = optimal_contract * price_diff
                return optimal_contract, optimal_revenue
            
            # 有总量约束的优化
            n_points = len(price_diff)
            
            # 目标函数：最大化总收益 (linprog默认最小化，所以取负数)
            c = -price_diff.values
            
            # 再次检查目标函数系数
            if np.any(~np.isfinite(c)):
                return self._greedy_optimization(price_diff)
            
            # 约束条件
            # 1. 总量约束：sum(x) <= daily_total_limit
            A_ub = np.ones((1, n_points))
            b_ub = np.array([self.daily_total_limit])
            
            # 2. 变量边界：0 <= x <= 12
            bounds = [(self.contract_range[0], self.contract_range[1]) for _ in range(n_points)]
            
            # 求解线性规划问题
            try:
                result = linprog(c, A_ub=A_ub, b_ub=b_ub, bounds=bounds, method='highs')
                
                if result.success:
                    optimal_contract = result.x
                    optimal_revenue = optimal_contract * price_diff
                    return optimal_contract, optimal_revenue
                else:
                    return self._greedy_optimization(price_diff)
                    
            except Exception as lp_error:
                return self._greedy_optimization(price_diff)
                
        except Exception as e:
            print(f"优化原合约时出错: {e}")
            try:
                # 尝试计算简单的价格差
                price_cols = self.find_price_columns(df)
                if 'contract_price' in price_cols and 'forward_price' in price_cols:
                    contract_prices = pd.to_numeric(df[price_cols['contract_price']], errors='coerce')
                    forward_prices = pd.to_numeric(df[price_cols['forward_price']], errors='coerce')
                    price_diff = (contract_prices - forward_prices).fillna(0)
                    price_diff = price_diff.replace([np.inf, -np.inf], 0)
                    return self._greedy_optimization(price_diff)
                else:
                    return None, None
            except Exception as fallback_error:
                print(f"贪心算法也失败了: {fallback_error}")
                return None, None
    
    def _greedy_optimization(self, price_diff):
        """贪心算法优化（当线性规划失败时使用）"""
        try:
            n_points = len(price_diff)
            optimal_contract = np.zeros(n_points)
            remaining_total = self.daily_total_limit if self.daily_total_limit else float('inf')
            
            # 确保price_diff是有效的数值
            if hasattr(price_diff, 'values'):
                price_values = price_diff.values
            else:
                price_values = np.array(price_diff)
            
            # 处理无效值
            price_values = np.nan_to_num(price_values, nan=0.0, posinf=0.0, neginf=0.0)
            
            # 按价格差排序，优先分配给收益最高的时间点
            sorted_indices = np.argsort(-price_values)  # 降序排列
            
            for idx in sorted_indices:
                if remaining_total <= 0:
                    break
                    
                if price_values[idx] > 0:  # 只对正收益的时间点分配
                    allocate = min(self.contract_range[1], remaining_total)
                    optimal_contract[idx] = allocate
                    remaining_total -= allocate
            
            # 重新创建price_diff Series用于计算收益
            if hasattr(price_diff, 'index'):
                optimal_revenue = optimal_contract * price_diff.fillna(0)
            else:
                optimal_revenue = optimal_contract * price_values
            
            return optimal_contract, optimal_revenue
            
        except Exception as e:
            print(f"贪心算法失败: {e}")
            return None, None
    
    def optimize_contract(self, df):
        """优化原合约使总收入最大（兼容旧接口）"""
        return self.optimize_contract_with_constraint(df)
    
    def analyze_daily_optimization(self, target_date):
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
        df = self.load_data(filepath)
        
        if df is None:
            return None
        
        # 询问用户输入每日总量限制
        if self.daily_total_limit is None:
            try:
                limit_input = input(f"\n请输入 {target_date.strftime('%Y年%m月%d日')} 的每日原合约总量限制: ").strip()
                if limit_input:
                    self.daily_total_limit = float(limit_input)
                    print(f"已设置每日总量限制为: {self.daily_total_limit}")
                else:
                    print("未设置总量限制，将使用无约束优化")
            except ValueError:
                print("输入无效，将使用无约束优化")
                self.daily_total_limit = None
        
        # 优化原合约
        optimal_contract, optimal_revenue = self.optimize_contract_with_constraint(df)
        
        if optimal_contract is None:
            return None
        
        # 打印96个时间点的最优原合约值
        date_str = target_date.strftime('%Y-%m-%d')
        self.print_optimal_values(optimal_contract, date_str)
        
        # 获取价格列信息
        price_cols = self.find_price_columns(df)
        
        result = {
            'date': target_date,
            'data': df,
            'optimal_contract': optimal_contract,
            'optimal_revenue': optimal_revenue,
            'total_optimal_revenue': optimal_revenue.sum(),
            'avg_optimal_contract': optimal_contract.mean(),
            'total_contract_amount': optimal_contract.sum(),
            'daily_total_limit': self.daily_total_limit,
            'price_columns': price_cols,
            'contract_range': self.contract_range
        }
        
        return result
    
    def plot_daily_optimization(self, result, save_path=None):
        """绘制每日优化结果"""
        if result is None:
            return
        
        fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(16, 12))
        
        date_str = result['date'].strftime('%Y-%m-%d')
        optimal_contract = result['optimal_contract']
        optimal_revenue = result['optimal_revenue']
        
        # 时间点
        time_points = range(1, len(optimal_contract) + 1)
        
        # 1. 最优原合约曲线
        ax1.plot(time_points, optimal_contract, 'b-', linewidth=2, marker='o', markersize=3)
        ax1.set_title(f'{date_str} 最优原合约曲线', fontsize=12)
        ax1.set_xlabel('时间点 (15分钟间隔)')
        ax1.set_ylabel('原合约值')
        ax1.grid(True, alpha=0.3)
        ax1.set_xticks(range(0, 97, 8))
        
        # 添加统计信息
        total_amount = np.sum(optimal_contract)
        avg_contract = np.mean(optimal_contract)
        limit_info = f"总量限制: {result['daily_total_limit']}" if result['daily_total_limit'] else "无总量限制"
        ax1.text(0.02, 0.98, f'总合约量: {total_amount:.3f}\n平均值: {avg_contract:.3f}\n{limit_info}', 
                transform=ax1.transAxes, verticalalignment='top',
                bbox=dict(boxstyle='round', facecolor='lightblue', alpha=0.8))
        
        # 2. 最优收益曲线
        ax2.plot(time_points, optimal_revenue, 'g-', linewidth=2, marker='s', markersize=3)
        ax2.set_title(f'{date_str} 最优收益曲线', fontsize=12)
        ax2.set_xlabel('时间点 (15分钟间隔)')
        ax2.set_ylabel('收益值')
        ax2.grid(True, alpha=0.3)
        ax2.set_xticks(range(0, 97, 8))
        
        # 添加统计信息
        total_revenue = np.sum(optimal_revenue)
        ax2.text(0.02, 0.98, f'总收益: {total_revenue:.2f}', 
                transform=ax2.transAxes, verticalalignment='top',
                bbox=dict(boxstyle='round', facecolor='lightgreen', alpha=0.8))
        
        # 3. 价格差分布
        price_cols = result['price_columns']
        if 'contract_price' in price_cols and 'forward_price' in price_cols:
            contract_prices = pd.to_numeric(result['data'][price_cols['contract_price']], errors='coerce')
            forward_prices = pd.to_numeric(result['data'][price_cols['forward_price']], errors='coerce')
            price_diff = contract_prices - forward_prices
            
            ax3.plot(time_points, price_diff, 'r-', linewidth=2, marker='^', markersize=3)
            ax3.set_title(f'{date_str} 价格差 (合约电价 - 日前电价)', fontsize=12)
            ax3.set_xlabel('时间点 (15分钟间隔)')
            ax3.set_ylabel('价格差')
            ax3.grid(True, alpha=0.3)
            ax3.set_xticks(range(0, 97, 8))
            ax3.axhline(y=0, color='black', linestyle='--', alpha=0.5)
        
        # 4. 原合约分布直方图
        ax4.hist(optimal_contract, bins=20, alpha=0.7, color='orange', edgecolor='black')
        ax4.set_title(f'{date_str} 最优原合约分布', fontsize=12)
        ax4.set_xlabel('原合约值')
        ax4.set_ylabel('频次')
        ax4.grid(True, alpha=0.3)
        
        plt.tight_layout()
        
        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
            print(f"图表已保存到: {save_path}")
        
        plt.show()
    
    def analyze_monthly_optimization(self, year, month):
        """分析指定月份的原合约优化"""
        monthly_data = []
        failed_dates = []
        
        # 询问用户输入月度总量限制
        if self.daily_total_limit is None:
            try:
                limit_input = input(f"\n请输入 {year}年{month}月 每日原合约总量限制: ").strip()
                if limit_input:
                    self.daily_total_limit = float(limit_input)
                    print(f"已设置每日总量限制为: {self.daily_total_limit}")
                else:
                    print("未设置总量限制，将使用无约束优化")
            except ValueError:
                print("输入无效，将使用无约束优化")
                self.daily_total_limit = None
        
        # 获取该月所有日期的数据
        excel_files = [f for f in os.listdir(self.data_dir) if f.endswith('.xlsx')]
        
        if not excel_files:
            print(f"数据目录中没有找到Excel文件")
            return None
        
        target_files = []
        for filename in excel_files:
            date = self.extract_date_from_filename(filename)
            if date and date.year == year and date.month == month:
                target_files.append((filename, date))
        
        if not target_files:
            print(f"没有找到 {year}年{month}月 的数据文件")
            return None
        
        # 按日期排序
        target_files.sort(key=lambda x: x[1])
        
        print(f"找到 {len(target_files)} 个数据文件，开始处理...")
        
        # 处理每个文件
        for i, (filename, date) in enumerate(target_files, 1):
            try:
                result = self.analyze_daily_optimization_internal(date)
                if result and result['optimal_contract'] is not None:
                    monthly_data.append(result['optimal_contract'])
                    # 简化进度显示：每10天显示一次，或显示失败的日期
                    if i % 10 == 0 or i == len(target_files):
                        print(f"  已处理: {i}/{len(target_files)} 天")
                else:
                    failed_dates.append(date.strftime('%Y-%m-%d'))
                    print(f"  ✗ {date.strftime('%m-%d')} 失败")
            except Exception as e:
                failed_dates.append(date.strftime('%Y-%m-%d'))
                print(f"  ✗ {date.strftime('%m-%d')} 错误: {e}")
        
        if not monthly_data:
            print("没有成功处理任何数据")
            if failed_dates:
                print(f"失败的日期: {', '.join(failed_dates)}")
            return None
        
        print(f"\n处理完成：成功 {len(monthly_data)} 天，失败 {len(failed_dates)} 天")
        if failed_dates:
            print(f"失败的日期: {', '.join(failed_dates)}")
        
        # 计算月度平均
        try:
            # 现在所有数据都应该是96个时间点，直接计算
            monthly_avg = np.mean(monthly_data, axis=0)
            monthly_std = np.std(monthly_data, axis=0)
            
            print(f"月度统计计算完成，使用{len(monthly_data)}天的数据")
            
            return {
                'year': year,
                'month': month,
                'daily_data': monthly_data,
                'monthly_average': monthly_avg,
                'monthly_std': monthly_std,
                'days_count': len(monthly_data),
                'failed_dates': failed_dates,
                'daily_total_limit': self.daily_total_limit
            }
        except Exception as e:
            print(f"计算月度统计时出错: {e}")
            # 如果仍有长度不一致问题，执行备用方案
            try:
                data_lengths = [len(data) for data in monthly_data]
                if len(set(data_lengths)) > 1:
                    print(f"  检测到数据长度不一致，执行标准化...")
                    # 统一到96个时间点
                    normalized_data = []
                    for data in monthly_data:
                        if len(data) >= 96:
                            normalized_data.append(data[:96])
                        else:
                            # 填充到96个时间点
                            padded_data = np.pad(data, (0, 96 - len(data)), 
                                               mode='constant', constant_values=0)
                            normalized_data.append(padded_data)
                    
                    monthly_avg = np.mean(normalized_data, axis=0)
                    monthly_std = np.std(normalized_data, axis=0)
                    
                    return {
                        'year': year,
                        'month': month,
                        'daily_data': normalized_data,
                        'monthly_average': monthly_avg,
                        'monthly_std': monthly_std,
                        'days_count': len(monthly_data),
                        'failed_dates': failed_dates,
                        'daily_total_limit': self.daily_total_limit
                    }
            except Exception as fallback_error:
                print(f"  备用计算也失败: {fallback_error}")
            return None
    
    def analyze_daily_optimization_internal(self, target_date):
        """内部使用的日分析函数，不询问用户输入"""
        try:
            if isinstance(target_date, str):
                try:
                    target_date = datetime.strptime(target_date, '%Y-%m-%d')
                except ValueError:
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
                return None
            
            filepath = os.path.join(self.data_dir, target_file)
            
            # 检查文件是否存在
            if not os.path.exists(filepath):
                return None
            
            df = self.load_data(filepath)
            
            if df is None or df.empty:
                return None
            
            # 优化原合约（使用已设置的限制）
            optimal_contract, optimal_revenue = self.optimize_contract_with_constraint(df)
            
            if optimal_contract is None or optimal_revenue is None:
                return None
            
            # 验证结果
            if len(optimal_contract) == 0 or len(optimal_revenue) == 0:
                return None
            
            result = {
                'date': target_date,
                'optimal_contract': optimal_contract,
                'optimal_revenue': optimal_revenue,
                'total_optimal_revenue': float(optimal_revenue.sum()) if hasattr(optimal_revenue, 'sum') else sum(optimal_revenue),
                'total_contract_amount': float(optimal_contract.sum()) if hasattr(optimal_contract, 'sum') else sum(optimal_contract)
            }
            
            return result
            
        except Exception as e:
            return None
    
    def plot_monthly_optimization(self, result, save_path=None):
        """绘制月度优化结果"""
        if result is None:
            return
        
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 8))
        
        year = result['year']
        month = result['month']
        monthly_avg = result['monthly_average']
        monthly_std = result['monthly_std']
        time_points = range(1, len(monthly_avg) + 1)
        
        # 1. 月度平均原合约曲线
        ax1.plot(time_points, monthly_avg, 'r-', linewidth=2, marker='o', markersize=4)
        ax1.fill_between(time_points, monthly_avg - monthly_std, monthly_avg + monthly_std, 
                        alpha=0.3, color='red', label='±1标准差')
        ax1.set_title(f'{year}年{month}月 平均原合约曲线', fontsize=14)
        ax1.set_xlabel('时间点 (15分钟间隔)')
        ax1.set_ylabel('平均原合约值')
        ax1.grid(True, alpha=0.3)
        ax1.set_xticks(range(0, 97, 8))
        ax1.legend()
        
        # 添加统计信息
        avg_total = np.sum(monthly_avg)
        avg_value = np.mean(monthly_avg)
        limit_info = f"每日总量限制: {result['daily_total_limit']}" if result['daily_total_limit'] else "无总量限制"
        
        ax1.text(0.02, 0.98, f'平均总合约量: {avg_total:.3f}\n月平均值: {avg_value:.3f}\n处理天数: {result["days_count"]}天\n{limit_info}', 
                transform=ax1.transAxes, verticalalignment='top',
                bbox=dict(boxstyle='round', facecolor='lightcoral', alpha=0.8))
        
        # 2. 月度原合约分布
        all_contracts = np.concatenate(result['daily_data'])
        ax2.hist(all_contracts, bins=30, alpha=0.7, color='blue', edgecolor='black')
        ax2.set_title(f'{year}年{month}月 原合约分布', fontsize=14)
        ax2.set_xlabel('原合约值')
        ax2.set_ylabel('频次')
        ax2.grid(True, alpha=0.3)
        
        # 添加统计信息
        ax2.axvline(avg_value, color='red', linestyle='--', linewidth=2, label=f'平均值: {avg_value:.3f}')
        ax2.legend()
        
        plt.tight_layout()
        
        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
            print(f"月度图表已保存到: {save_path}")
        
        plt.show()

def main():
    """主函数"""
    optimizer = ContractOptimizer()
    
    print("="*50)
    print("           原合约优化分析器")
    print("="*50)
    
    while True:
        print("\n请选择操作：")
        print("1. 分析单日原合约优化")
        print("2. 分析月度原合约优化")
        print("3. 批量分析所有日期")
        print("0. 退出")
        print("-"*30)
        
        choice = input("请输入选项 (0-3): ").strip()
        
        if choice == '0':
            print("谢谢使用！")
            break
        elif choice == '1':
            date = input("请输入日期 (格式: 2025-05-10): ").strip()
            print(f"\n正在分析 {date} 的原合约优化...")
            
            result = optimizer.analyze_daily_optimization(date)
            if result:
                print(f"\n=== {date} 原合约优化分析结果 ===")
                print(f"平均最优原合约值: {result['avg_optimal_contract']:.3f}")
                print(f"总最优收益: {result['total_optimal_revenue']:.2f}")
                print(f"原合约取值范围: {result['contract_range']}")
                
                plot_choice = input("\n是否绘制分析图表? (y/n): ").strip().lower()
                if plot_choice in ['y', 'yes', '是']:
                    save_choice = input("是否保存图表? (y/n): ").strip().lower()
                    save_path = None
                    if save_choice in ['y', 'yes', '是']:
                        save_path = f"原合约优化分析_{date}.png"
                    
                    optimizer.plot_daily_optimization(result, save_path)
            else:
                print("分析失败，请检查数据文件")
        
        elif choice == '2':
            month_input = input("请输入月份 (格式: 2025-05): ").strip()
            try:
                # 改进月份解析逻辑
                if '-' in month_input and len(month_input.split('-')) == 2:
                    year_str, month_str = month_input.split('-')
                    year = int(year_str)
                    month = int(month_str)
                    
                    # 验证年份和月份的合理性
                    if year < 2020 or year > 2030:
                        print("年份应在2020-2030之间")
                        continue
                    if month < 1 or month > 12:
                        print("月份应在1-12之间")
                        continue
                        
                    print(f"\n正在分析 {year}年{month}月 的原合约优化...")
                    
                    # 重置daily_total_limit以便重新询问
                    optimizer.daily_total_limit = None
                    
                    result = optimizer.analyze_monthly_optimization(year, month)
                    if result:
                        print(f"\n=== {year}年{month}月 原合约优化分析结果 ===")
                        print(f"处理天数: {result['days_count']}天")
                        if result['days_count'] > 0:
                            print(f"月平均原合约值: {np.mean(result['monthly_average']):.3f}")
                            print(f"月平均总合约量: {np.sum(result['monthly_average']):.3f}")
                            print(f"每日总量限制: {result['daily_total_limit']}")
                        
                            plot_choice = input("\n是否绘制月度分析图表? (y/n): ").strip().lower()
                            if plot_choice in ['y', 'yes', '是']:
                                save_choice = input("是否保存图表? (y/n): ").strip().lower()
                                save_path = None
                                if save_choice in ['y', 'yes', '是']:
                                    save_path = f"月度原合约优化分析_{year}年{month}月.png"
                                
                                optimizer.plot_monthly_optimization(result, save_path)
                        else:
                            print("该月份没有找到有效数据")
                    else:
                        print("分析失败，请检查该月份的数据文件是否存在")
                else:
                    print("月份格式错误，请使用 YYYY-MM 格式 (例如: 2025-05)")
            except ValueError:
                print("输入格式错误，请使用 YYYY-MM 格式 (例如: 2025-05)")
            except Exception as e:
                print(f"处理月份输入时出错: {e}")
                print("请检查输入格式并重试")
        
        elif choice == '3':
            print("批量分析功能开发中...")
        
        else:
            print("无效选项，请重新选择")
        
        input("\n按回车键继续...")

if __name__ == "__main__":
    main() 