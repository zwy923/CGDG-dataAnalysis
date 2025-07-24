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
from scipy.optimize import linprog, minimize_scalar

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

class ContractOptimizer:
    def __init__(self, data_dir='2025/'):
        self.data_dir = data_dir
        self.contract_range = (0, 12)  # 原合约取值范围
        self.daily_total_limit = None  # 每日原合约总量限制
        
    def set_daily_total_limit(self, limit):
        """设置每日原合约总量限制"""
        self.daily_total_limit = limit
        
    def get_month_folder_name(self, month):
        """根据月份数字获取对应的文件夹名称"""
        month_names = {
            1: "1月", 2: "2月", 3: "3月", 4: "4月", 5: "5月", 6: "6月",
            7: "7月", 8: "8月", 9: "9月", 10: "10月", 11: "11月", 12: "12月"
        }
        return month_names.get(month, str(month) + "月")
    
    def get_all_excel_files(self):
        """获取所有Excel文件的路径和文件名"""
        all_files = []
        
        try:
            # 遍历数据目录中的所有子文件夹
            for item in os.listdir(self.data_dir):
                item_path = os.path.join(self.data_dir, item)
                if os.path.isdir(item_path):
                    # 在子文件夹中查找Excel文件
                    try:
                        for filename in os.listdir(item_path):
                            if filename.endswith('.xlsx'):
                                full_path = os.path.join(item_path, filename)
                                all_files.append((filename, full_path))
                    except PermissionError:
                        continue
                elif item.endswith('.xlsx'):
                    # 根目录中的Excel文件（向后兼容）
                    full_path = os.path.join(self.data_dir, item)
                    all_files.append((item, full_path))
        except FileNotFoundError:
            print("数据目录不存在: " + str(self.data_dir))
        
        return all_files
    
    def get_monthly_excel_files(self, year, month):
        """获取指定月份的Excel文件"""
        month_folder = self.get_month_folder_name(month)
        month_path = os.path.join(self.data_dir, month_folder)
        
        files = []
        
        # 首先尝试在月份文件夹中查找
        if os.path.exists(month_path) and os.path.isdir(month_path):
            try:
                for filename in os.listdir(month_path):
                    if filename.endswith('.xlsx'):
                        date = self.extract_date_from_filename(filename)
                        if date and date.year == year and date.month == month:
                            full_path = os.path.join(month_path, filename)
                            files.append((filename, full_path))
            except PermissionError:
                pass
        
        # 如果月份文件夹中没有找到文件，尝试在所有文件中查找（向后兼容）
        if not files:
            all_files = self.get_all_excel_files()
            for filename, full_path in all_files:
                date = self.extract_date_from_filename(filename)
                if date and date.year == year and date.month == month:
                    files.append((filename, full_path))
        
        return files
    
    def find_file_for_date(self, target_date):
        """查找指定日期的文件"""
        if isinstance(target_date, str):
            try:
                target_date = datetime.strptime(target_date, '%Y-%m-%d')
            except ValueError:
                return None, None
        
        # 首先尝试在对应月份文件夹中查找
        month_files = self.get_monthly_excel_files(target_date.year, target_date.month)
        for filename, full_path in month_files:
            date = self.extract_date_from_filename(filename)
            if date and date == target_date:
                return filename, full_path
        
        # 如果在月份文件夹中没找到，尝试在所有文件中查找
        all_files = self.get_all_excel_files()
        for filename, full_path in all_files:
            date = self.extract_date_from_filename(filename)
            if date and date == target_date:
                return filename, full_path
        
        return None, None
    
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
            # 日前出清电量
            if '日前出清' in col_str:
                columns['forward_clearing'] = col
            # 滚动撮合电量
            elif '滚动撮合电量' in col_str or ('撮合电量' in col_str and '滚动' in col_str):
                columns['matching_volume'] = col
            # 撮合电量（售出）
            elif '撮合' in col_str and ('售出' in col_str or '卖出' in col_str):
                columns['matching_sell'] = col
            # 撮合电量（购入）
            elif '撮合' in col_str and ('购入' in col_str or '买入' in col_str):
                columns['matching_buy'] = col
            # 日内实际电量
            elif '日内实际' in col_str or ('实际' in col_str and '日内' in col_str):
                columns['realtime_actual'] = col
            # 省间现货电量
            elif '省间现货电量' in col_str or ('省间' in col_str and '现货' in col_str and '电量' in col_str):
                columns['interprovincial_volume'] = col
            # 其他可能的撮合电量列
            elif '撮合电量' in col_str and 'matching_volume' not in columns:
                columns['matching_volume'] = col
        
        return columns
    
    def calculate_total_revenue(self, df, original_contract):
        """计算总收入"""
        try:
            # 查找各项收益列
            revenue_columns = {}
            
            for col in df.columns:
                col_str = str(col).lower()
                if '合约收益' in col_str:
                    revenue_columns['contract_revenue'] = col
                elif '撮合收益' in col_str:
                    revenue_columns['matching_revenue'] = col
                elif '日前结算收益' in col_str:
                    revenue_columns['forward_settlement_revenue'] = col
                elif '实时结算收益' in col_str:
                    revenue_columns['realtime_settlement_revenue'] = col
                elif '省间现货收益' in col_str:
                    revenue_columns['interprovincial_revenue'] = col
            
            # 计算总收入：各项收益相加
            total_revenue = pd.Series(0.0, index=df.index)
            
            for revenue_type, col_name in revenue_columns.items():
                revenue_values = pd.to_numeric(df[col_name], errors='coerce').fillna(0)
                total_revenue += revenue_values
            
            if len(revenue_columns) == 0:
                print("警告：未找到任何收益列")
                return None
            
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
            
            # 计算合约收益系数 (合约电价 - 日前电价)
            # 注意：只有合约收益部分与原合约值直接相关
            contract_revenue_coefficient = contract_prices - forward_prices
            
            # 处理无效值：将NaN、inf替换为0
            contract_revenue_coefficient = contract_revenue_coefficient.fillna(0)
            contract_revenue_coefficient = contract_revenue_coefficient.replace([np.inf, -np.inf], 0)
            
            # 检查处理后的数据
            if len(contract_revenue_coefficient) == 0:
                print("警告：没有有效的合约收益系数数据")
                return None, None
            
            # 如果没有设置总量限制，需要提示用户输入
            if self.daily_total_limit is None:
                print("警告：等式约束优化需要设置每日总量限制")
                return None, None
            
            # 有总量约束的优化
            n_points = len(contract_revenue_coefficient)
            
            # 目标函数：最大化合约收益部分 (linprog默认最小化，所以取负数)
            c = -contract_revenue_coefficient.values
            
            # 再次检查目标函数系数
            if np.any(~np.isfinite(c)):
                return self._greedy_optimization_with_equality(contract_revenue_coefficient, df)
            
            # 约束条件
            # 1. 总量约束：sum(x) = daily_total_limit (等式约束)
            A_eq = np.ones((1, n_points))
            b_eq = np.array([self.daily_total_limit])
            
            # 2. 变量边界：0 <= x <= 12
            bounds = [(self.contract_range[0], self.contract_range[1]) for _ in range(n_points)]
            
            # 求解线性规划问题
            try:
                result = linprog(c, A_eq=A_eq, b_eq=b_eq, bounds=bounds, method='highs')
                
                if result.success:
                    optimal_contract = result.x
                    
                    # 计算完整的总收益
                    total_revenue = self.calculate_total_revenue_for_contract(df, optimal_contract)
                    
                    return optimal_contract, total_revenue
                else:
                    return self._greedy_optimization_with_equality(contract_revenue_coefficient, df)
                    
            except Exception as lp_error:
                return self._greedy_optimization_with_equality(contract_revenue_coefficient, df)
                
        except Exception as e:
            print(f"优化原合约时出错: {e}")
            try:
                # 尝试计算简单的合约收益系数
                price_cols = self.find_price_columns(df)
                if 'contract_price' in price_cols and 'forward_price' in price_cols:
                    contract_prices = pd.to_numeric(df[price_cols['contract_price']], errors='coerce')
                    forward_prices = pd.to_numeric(df[price_cols['forward_price']], errors='coerce')
                    contract_revenue_coefficient = (contract_prices - forward_prices).fillna(0)
                    contract_revenue_coefficient = contract_revenue_coefficient.replace([np.inf, -np.inf], 0)
                    if self.daily_total_limit is not None:
                        return self._greedy_optimization_with_equality(contract_revenue_coefficient, df)
                    else:
                        # 对于没有总量限制的情况，使用简单的贪心算法
                        return self._greedy_optimization(contract_revenue_coefficient)
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
    
    def _greedy_optimization_with_equality(self, contract_revenue_coefficient, df=None):
        """贪心算法优化，确保总量严格等于限制值"""
        try:
            n_points = len(contract_revenue_coefficient)
            optimal_contract = np.zeros(n_points)
            target_total = self.daily_total_limit
            
            # 确保contract_revenue_coefficient是有效的数值
            if hasattr(contract_revenue_coefficient, 'values'):
                coeff_values = contract_revenue_coefficient.values
            else:
                coeff_values = np.array(contract_revenue_coefficient)
            
            # 处理无效值
            coeff_values = np.nan_to_num(coeff_values, nan=0.0, posinf=0.0, neginf=0.0)
            
            # 检查是否可能达到目标总量
            max_possible = n_points * self.contract_range[1]  # 96 * 12 = 1152
            if target_total > max_possible:
                print(f"警告：目标总量 {target_total} 超过最大可能值 {max_possible}")
                # 按比例分配到所有时间点
                optimal_contract = np.full(n_points, self.contract_range[1])
                if df is not None:
                    optimal_revenue = self.calculate_total_revenue_for_contract(df, optimal_contract)
                else:
                    optimal_revenue = optimal_contract * coeff_values
                return optimal_contract, optimal_revenue
            
            # 按合约收益系数排序，优先分配给收益最高的时间点
            sorted_indices = np.argsort(-coeff_values)  # 降序排列
            remaining_total = target_total
            
            # 第一阶段：尽可能分配给正收益的时间点
            for idx in sorted_indices:
                if remaining_total <= 0:
                    break
                if coeff_values[idx] > 0:  # 优先分配给正收益
                    allocate = min(self.contract_range[1], remaining_total)
                    optimal_contract[idx] = allocate
                    remaining_total -= allocate
            
            # 第二阶段：如果还有剩余，分配给收益最高的时间点（包括负收益）
            if remaining_total > 0:
                for idx in sorted_indices:
                    if remaining_total <= 0:
                        break
                    available_capacity = self.contract_range[1] - optimal_contract[idx]
                    if available_capacity > 0:
                        allocate = min(available_capacity, remaining_total)
                        optimal_contract[idx] += allocate
                        remaining_total -= allocate
            
            # 第三阶段：如果总量不够，需要从某些时间点减少分配
            if remaining_total < 0:
                # 从收益最低的时间点开始减少
                excess = -remaining_total
                for idx in reversed(sorted_indices):
                    if excess <= 0:
                        break
                    reduction = min(optimal_contract[idx], excess)
                    optimal_contract[idx] -= reduction
                    excess -= reduction
            
            # 确保总量严格等于目标值
            current_total = np.sum(optimal_contract)
            if abs(current_total - target_total) > 1e-6:
                # 进行微调
                diff = target_total - current_total
                if diff > 0:
                    # 需要增加，找到可以增加的时间点
                    for idx in sorted_indices:
                        if diff <= 0:
                            break
                        available = self.contract_range[1] - optimal_contract[idx]
                        if available > 0:
                            add_amount = min(available, diff)
                            optimal_contract[idx] += add_amount
                            diff -= add_amount
                else:
                    # 需要减少，找到可以减少的时间点
                    diff = -diff
                    for idx in reversed(sorted_indices):
                        if diff <= 0:
                            break
                        if optimal_contract[idx] > 0:
                            reduce_amount = min(optimal_contract[idx], diff)
                            optimal_contract[idx] -= reduce_amount
                            diff -= reduce_amount
            
            # 计算完整的总收益
            if df is not None:
                optimal_revenue = self.calculate_total_revenue_for_contract(df, optimal_contract)
            else:
                # 如果没有df，只计算合约收益部分
                if hasattr(contract_revenue_coefficient, 'index'):
                    optimal_revenue = optimal_contract * contract_revenue_coefficient.fillna(0)
                else:
                    optimal_revenue = optimal_contract * coeff_values
            
            return optimal_contract, optimal_revenue
            
        except Exception as e:
            print(f"等式约束贪心算法失败: {e}")
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
        target_file, filepath = self.find_file_for_date(target_date)
        
        if target_file is None:
            print(f"未找到 {target_date.strftime('%Y-%m-%d')} 的数据文件")
            return None
        
        df = self.load_data(filepath)
        
        if df is None:
            return None
        
        # 询问用户输入每日总量限制
        if self.daily_total_limit is None:
            try:
                limit_input = input(f"\n请输入 {target_date.strftime('%Y年%m月%d日')} 的每日原合约总量（必须等于此值）: ").strip()
                if limit_input:
                    self.daily_total_limit = float(limit_input)
                    print(f"已设置每日总量限制为: {self.daily_total_limit} (等式约束)")
                else:
                    print("错误：必须设置总量限制值")
                    return None
            except ValueError:
                print("输入无效，必须输入数值")
                return None
        
        # 优化原合约
        optimal_contract, optimal_revenue = self.optimize_contract_with_constraint(df)
        
        if optimal_contract is None:
            return None
        
        # 打印96个时间点的最优原合约值
        date_str = target_date.strftime('%Y-%m-%d')
        self.print_optimal_values(optimal_contract, date_str)
        
        # 验证总量约束
        actual_total = optimal_contract.sum()
        if self.daily_total_limit is not None:
            print(f"\n约束验证:")
            print(f"目标总量: {self.daily_total_limit}")
            print(f"实际总量: {actual_total:.6f}")
            print(f"差值: {abs(actual_total - self.daily_total_limit):.6f}")
            if abs(actual_total - self.daily_total_limit) < 1e-6:
                print("✓ 等式约束满足")
            else:
                print("✗ 等式约束不满足")
        
        # 获取价格列信息
        price_cols = self.find_price_columns(df)
        
        result = {
            'date': target_date,
            'data': df,
            'optimal_contract': optimal_contract,
            'optimal_revenue': optimal_revenue,
            'total_optimal_revenue': float(optimal_revenue.sum()) if hasattr(optimal_revenue, 'sum') else sum(optimal_revenue),
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
        limit_info = f"总量限制: {result['daily_total_limit']} (等式约束)" if result['daily_total_limit'] else "无总量限制"
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
        
        # 3. 合约收益系数分布
        price_cols = result['price_columns']
        if 'contract_price' in price_cols and 'forward_price' in price_cols:
            contract_prices = pd.to_numeric(result['data'][price_cols['contract_price']], errors='coerce')
            forward_prices = pd.to_numeric(result['data'][price_cols['forward_price']], errors='coerce')
            contract_revenue_coeff = contract_prices - forward_prices
            
            ax3.plot(time_points, contract_revenue_coeff, 'r-', linewidth=2, marker='^', markersize=3)
            ax3.set_title(f'{date_str} 合约收益系数 (合约电价 - 日前电价)', fontsize=12)
            ax3.set_xlabel('时间点 (15分钟间隔)')
            ax3.set_ylabel('合约收益系数')
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
                limit_input = input(f"\n请输入 {year}年{month}月 每日原合约总量（必须等于此值）: ").strip()
                if limit_input:
                    self.daily_total_limit = float(limit_input)
                    print(f"已设置每日总量限制为: {self.daily_total_limit} (等式约束)")
                else:
                    print("错误：必须设置总量限制值")
                    return None
            except ValueError:
                print("输入无效，必须输入数值")
                return None
        
        # 获取该月所有日期的数据
        target_files = self.get_monthly_excel_files(year, month)
        
        if not target_files:
            print(f"没有找到 {year}年{month}月 的数据文件")
            return None
        
        # 转换为包含日期的格式，并按日期排序
        processed_files = []
        for filename, full_path in target_files:
            date = self.extract_date_from_filename(filename)
            if date:
                processed_files.append((filename, full_path, date))
        
        processed_files.sort(key=lambda x: x[2])
        
        if not processed_files:
            print(f"没有找到 {year}年{month}月 的数据文件")
            return None
        
        print(f"找到 {len(processed_files)} 个数据文件，开始处理...")
        
        # 处理每个文件
        for i, (filename, full_path, date) in enumerate(processed_files, 1):
            try:
                result = self.analyze_daily_optimization_internal(date)
                if result and result['optimal_contract'] is not None:
                    monthly_data.append(result['optimal_contract'])
                    # 简化进度显示：每10天显示一次，或显示失败的日期
                    if i % 10 == 0 or i == len(processed_files):
                        print(f"  已处理: {i}/{len(processed_files)} 天")
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
            target_file, filepath = self.find_file_for_date(target_date)
            
            if target_file is None:
                return None
            
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
        limit_info = f"每日总量限制: {result['daily_total_limit']} (等式约束)" if result['daily_total_limit'] else "无总量限制"
        
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

    def calculate_total_revenue_for_contract(self, df, contract_values):
        """计算指定原合约值下的总收益"""
        try:
            # 查找电价列
            price_cols = self.find_price_columns(df)
            volume_cols = self.find_volume_columns(df)
            
            # 确保 contract_values 是 pandas Series 或 numpy array
            if hasattr(contract_values, 'values'):
                contract_array = contract_values.values
            else:
                contract_array = np.array(contract_values)
            
            # 确保长度为96
            if len(contract_array) < 96:
                contract_array = np.pad(contract_array, (0, 96 - len(contract_array)), mode='constant', constant_values=0)
            elif len(contract_array) > 96:
                contract_array = contract_array[:96]
            
            total_revenue = np.zeros(96)
            
            # 1. 计算合约收益 = 原合约 × 合约电价
            if 'contract_price' in price_cols:
                contract_prices = pd.to_numeric(df[price_cols['contract_price']], errors='coerce').fillna(0)
                if len(contract_prices) >= 96:
                    contract_revenue = contract_array * contract_prices[:96].values
                    total_revenue += contract_revenue
            
            # 2. 计算日前结算收益（会随原合约变化）
            # 公式：（日前出清 - 原合约*4 - 撮合电量（售出）- 撮合电量（购入）- 省间现货电量*4）/4 * 日前电价
            if ('forward_price' in price_cols and 'forward_clearing' in volume_cols):
                forward_prices = pd.to_numeric(df[price_cols['forward_price']], errors='coerce').fillna(0)
                forward_clearing = pd.to_numeric(df[volume_cols['forward_clearing']], errors='coerce').fillna(0)
                
                # 获取其他电量数据
                matching_sell = np.zeros(96)  # 撮合电量（售出）
                matching_buy = np.zeros(96)   # 撮合电量（购入）
                interprovincial_volume = np.zeros(96)  # 省间现货电量
                
                # 使用 volume_cols 中找到的列名
                if 'matching_sell' in volume_cols:
                    matching_sell_data = pd.to_numeric(df[volume_cols['matching_sell']], errors='coerce').fillna(0)
                    if len(matching_sell_data) >= 96:
                        matching_sell = matching_sell_data[:96].values
                
                if 'matching_buy' in volume_cols:
                    matching_buy_data = pd.to_numeric(df[volume_cols['matching_buy']], errors='coerce').fillna(0)
                    if len(matching_buy_data) >= 96:
                        matching_buy = matching_buy_data[:96].values
                
                if 'interprovincial_volume' in volume_cols:
                    interprovincial_data = pd.to_numeric(df[volume_cols['interprovincial_volume']], errors='coerce').fillna(0)
                    if len(interprovincial_data) >= 96:
                        interprovincial_volume = interprovincial_data[:96].values
                
                # 计算日前结算收益
                if len(forward_clearing) >= 96 and len(forward_prices) >= 96:
                    forward_settlement_volume = (
                        forward_clearing[:96].values 
                        - contract_array * 4 
                        - matching_sell 
                        - matching_buy 
                        - interprovincial_volume * 4
                    ) / 4
                    
                    forward_settlement_revenue = forward_settlement_volume * forward_prices[:96].values
                    total_revenue += forward_settlement_revenue
                    
                    # 调试信息（可选）
                    if hasattr(self, 'debug') and self.debug:
                        print(f"日前结算收益计算详情:")
                        print(f"  日前出清电量平均: {np.mean(forward_clearing[:96]):.3f}")
                        print(f"  原合约*4平均: {np.mean(contract_array * 4):.3f}")
                        print(f"  撮合售出平均: {np.mean(matching_sell):.3f}")
                        print(f"  撮合购入平均: {np.mean(matching_buy):.3f}")
                        print(f"  省间现货*4平均: {np.mean(interprovincial_volume * 4):.3f}")
                        print(f"  结算电量平均: {np.mean(forward_settlement_volume):.3f}")
                        print(f"  日前电价平均: {np.mean(forward_prices[:96]):.3f}")
                        print(f"  日前结算收益平均: {np.mean(forward_settlement_revenue):.3f}")
            
            # 3. 添加其他固定收益（不随原合约变化的收益）
            fixed_revenue_columns = {}
            for col in df.columns:
                col_str = str(col).lower()
                if '撮合收益' in col_str:
                    fixed_revenue_columns['matching_revenue'] = col
                elif '实时结算收益' in col_str:
                    fixed_revenue_columns['realtime_settlement_revenue'] = col
                elif '省间现货收益' in col_str:
                    fixed_revenue_columns['interprovincial_revenue'] = col
            
            for revenue_type, col_name in fixed_revenue_columns.items():
                revenue_values = pd.to_numeric(df[col_name], errors='coerce').fillna(0)
                if len(revenue_values) >= 96:
                    total_revenue += revenue_values[:96].values
            
            return pd.Series(total_revenue)
            
        except Exception as e:
            print(f"计算总收益时出错: {e}")
            print(f"错误详情: {str(e)}")
            # 返回基本的合约收益作为备选
            try:
                if hasattr(contract_values, 'values'):
                    contract_array = contract_values.values
                else:
                    contract_array = np.array(contract_values)
                
                if len(contract_array) < 96:
                    contract_array = np.pad(contract_array, (0, 96 - len(contract_array)), mode='constant', constant_values=0)
                elif len(contract_array) > 96:
                    contract_array = contract_array[:96]
                
                # 至少计算合约收益部分
                price_cols = self.find_price_columns(df)
                if 'contract_price' in price_cols:
                    contract_prices = pd.to_numeric(df[price_cols['contract_price']], errors='coerce').fillna(0)
                    if len(contract_prices) >= 96:
                        return pd.Series(contract_array * contract_prices[:96].values)
                
                return pd.Series(np.zeros(96))
            except:
                return pd.Series(np.zeros(96))

    def print_detailed_revenue_breakdown(self, df, optimal_contract, optimal_revenue, date_str):
        """打印每个时间点的详细收益分解"""
        print(f"\n=== {date_str} 每15分钟详细收益分解 ===")
        
        # 查找各项收益列
        revenue_columns = {}
        for col in df.columns:
            col_str = str(col).lower()
            if '合约收益' in col_str:
                revenue_columns['contract_revenue'] = col
            elif '撮合收益' in col_str:
                revenue_columns['matching_revenue'] = col
            elif '日前结算收益' in col_str:
                revenue_columns['forward_settlement_revenue'] = col
            elif '实时结算收益' in col_str:
                revenue_columns['realtime_settlement_revenue'] = col
            elif '省间现货收益' in col_str:
                revenue_columns['interprovincial_revenue'] = col
        
        # 计算优化后的合约收益
        price_cols = self.find_price_columns(df)
        if 'contract_price' in price_cols:
            contract_prices = pd.to_numeric(df[price_cols['contract_price']], errors='coerce').fillna(0)
            # 合约收益 = 原合约 × 合约电价
            optimized_contract_revenue = optimal_contract * contract_prices
        else:
            optimized_contract_revenue = pd.Series(0.0, index=range(len(optimal_contract)))
        
        # 打印表头
        header = f"{'时间':>6} {'原合约':>8} {'合约收益':>10} {'撮合收益':>10} {'日前结算':>10} {'实时结算':>10} {'省间现货':>10} {'总收入':>10}"
        print(header)
        print("=" * len(header))
        
        # 逐行打印每个时间点的数据
        total_sum = {'contract': 0, 'matching': 0, 'forward': 0, 'realtime': 0, 'interprovincial': 0, 'total': 0}
        
        for i in range(len(optimal_contract)):
            # 计算时间（每15分钟一个点）
            hour = i // 4
            minute = (i % 4) * 15
            time_str = f"{hour:02d}:{minute:02d}"
            
            # 获取各项收益
            contract_rev = optimized_contract_revenue.iloc[i] if i < len(optimized_contract_revenue) else 0
            
            matching_rev = 0
            if 'matching_revenue' in revenue_columns and i < len(df):
                matching_rev = pd.to_numeric(df[revenue_columns['matching_revenue']].iloc[i], errors='coerce')
                matching_rev = matching_rev if pd.notna(matching_rev) else 0
            
            forward_rev = 0
            if 'forward_settlement_revenue' in revenue_columns and i < len(df):
                forward_rev = pd.to_numeric(df[revenue_columns['forward_settlement_revenue']].iloc[i], errors='coerce')
                forward_rev = forward_rev if pd.notna(forward_rev) else 0
            
            realtime_rev = 0
            if 'realtime_settlement_revenue' in revenue_columns and i < len(df):
                realtime_rev = pd.to_numeric(df[revenue_columns['realtime_settlement_revenue']].iloc[i], errors='coerce')
                realtime_rev = realtime_rev if pd.notna(realtime_rev) else 0
            
            interprovincial_rev = 0
            if 'interprovincial_revenue' in revenue_columns and i < len(df):
                interprovincial_rev = pd.to_numeric(df[revenue_columns['interprovincial_revenue']].iloc[i], errors='coerce')
                interprovincial_rev = interprovincial_rev if pd.notna(interprovincial_rev) else 0
            
            # 总收入
            total_revenue_point = contract_rev + matching_rev + forward_rev + realtime_rev + interprovincial_rev
            
            # 打印当前时间点数据
            print(f"{time_str:>6} {optimal_contract[i]:>8.3f} {contract_rev:>10.2f} {matching_rev:>10.2f} {forward_rev:>10.2f} {realtime_rev:>10.2f} {interprovincial_rev:>10.2f} {total_revenue_point:>10.2f}")
            
            # 累加到总计
            total_sum['contract'] += contract_rev
            total_sum['matching'] += matching_rev
            total_sum['forward'] += forward_rev
            total_sum['realtime'] += realtime_rev
            total_sum['interprovincial'] += interprovincial_rev
            total_sum['total'] += total_revenue_point
        
        # 打印分割线和总计
        print("=" * len(header))
        total_contract = np.sum(optimal_contract)
        print(f"{'总计':>6} {total_contract:>8.3f} {total_sum['contract']:>10.2f} {total_sum['matching']:>10.2f} {total_sum['forward']:>10.2f} {total_sum['realtime']:>10.2f} {total_sum['interprovincial']:>10.2f} {total_sum['total']:>10.2f}")
        
        # 打印收益占比
        print(f"\n=== 收益构成分析 ===")
        if total_sum['total'] != 0:
            print(f"合约收益占比: {total_sum['contract']/total_sum['total']*100:>6.1f}% ({total_sum['contract']:>10.2f})")
            print(f"撮合收益占比: {total_sum['matching']/total_sum['total']*100:>6.1f}% ({total_sum['matching']:>10.2f})")
            print(f"日前结算占比: {total_sum['forward']/total_sum['total']*100:>6.1f}% ({total_sum['forward']:>10.2f})")
            print(f"实时结算占比: {total_sum['realtime']/total_sum['total']*100:>6.1f}% ({total_sum['realtime']:>10.2f})")
            print(f"省间现货占比: {total_sum['interprovincial']/total_sum['total']*100:>6.1f}% ({total_sum['interprovincial']:>10.2f})")
        
        return total_sum

    def print_monthly_revenue_breakdown(self, year, month):
        """打印月度每日详细收益分解"""
        print(f"\n=== {year}年{month}月 每日详细收益分解 ===")
        
        # 获取该月所有日期的数据
        file_list = self.get_monthly_excel_files(year, month)
        
        if not file_list:
            print(f"没有找到 {year}年{month}月 的数据文件")
            return
        
        # 转换为包含日期的格式，并按日期排序
        target_files = []
        for filename, full_path in file_list:
            date = self.extract_date_from_filename(filename)
            if date:
                target_files.append((filename, full_path, date))
        
        # 按日期排序
        target_files.sort(key=lambda x: x[2])
        
        monthly_total = {'contract': 0, 'matching': 0, 'forward': 0, 'realtime': 0, 'interprovincial': 0, 'total': 0}
        
        for filename, full_path, date in target_files:
            try:
                df = self.load_data(full_path)
                
                if df is None or df.empty:
                    print(f"{date.strftime('%m-%d')}: 数据加载失败")
                    continue
                
                # 优化原合约
                optimal_contract, optimal_revenue = self.optimize_contract_with_constraint(df)
                
                if optimal_contract is None:
                    print(f"{date.strftime('%m-%d')}: 优化失败")
                    continue
                
                # 打印该日的详细分解
                date_str = date.strftime('%Y-%m-%d')
                daily_breakdown = self.print_detailed_revenue_breakdown(df, optimal_contract, optimal_revenue, date_str)
                
                # 累加到月度总计
                for key in monthly_total:
                    if key in daily_breakdown:
                        monthly_total[key] += daily_breakdown[key]
                
                print()  # 空行分隔
                
            except Exception as e:
                print(f"{date.strftime('%m-%d')}: 处理出错 - {e}")
                continue
        
        # 打印月度汇总
        print("=" * 80)
        print(f"=== {year}年{month}月 收益汇总 ===")
        print(f"合约收益总计: {monthly_total['contract']:,.2f}")
        print(f"撮合收益总计: {monthly_total['matching']:,.2f}")
        print(f"日前结算总计: {monthly_total['forward']:,.2f}")
        print(f"实时结算总计: {monthly_total['realtime']:,.2f}")
        print(f"省间现货总计: {monthly_total['interprovincial']:,.2f}")
        print(f"总收入合计: {monthly_total['total']:,.2f}")
        
        if monthly_total['total'] != 0:
            print(f"\n=== 月度收益构成 ===")
            print(f"合约收益占比: {monthly_total['contract']/monthly_total['total']*100:.1f}%")
            print(f"撮合收益占比: {monthly_total['matching']/monthly_total['total']*100:.1f}%")
            print(f"日前结算占比: {monthly_total['forward']/monthly_total['total']*100:.1f}%")
            print(f"实时结算占比: {monthly_total['realtime']/monthly_total['total']*100:.1f}%")
            print(f"省间现货占比: {monthly_total['interprovincial']/monthly_total['total']*100:.1f}%")
        
        return monthly_total

    def analyze_scaled_contract(self, monthly_result, scale_factor=0.5):
        """分析原合约按比例调整后的情况"""
        if monthly_result is None or 'monthly_average' not in monthly_result:
            print("错误：无效的月度分析结果")
            return None
        
        year = monthly_result['year']
        month = monthly_result['month']
        original_monthly_avg = monthly_result['monthly_average']
        
        # 将原合约值按比例调整，保持分布不变
        scaled_monthly_avg = original_monthly_avg * scale_factor
        
        # 根据缩放因子设置描述词
        if scale_factor == 0.25:
            scale_desc = "四分之一"
        elif scale_factor == 0.5:
            scale_desc = "减半"
        elif scale_factor == 1.5:
            scale_desc = "1.5倍"
        elif scale_factor == 2.0:
            scale_desc = "双倍"
        else:
            scale_desc = f"{scale_factor}倍"
        
        print(f"\n=== {year}年{month}月 原合约{scale_desc}调整分析 ===")
        print(f"正在重新计算{scale_desc}调整后的收入...")
        
        # 重新计算每日的收入
        scaled_daily_revenues = []
        total_scaled_revenue = 0
        
        # 获取该月所有日期的数据文件
        file_list = self.get_monthly_excel_files(year, month)
        
        # 转换为包含日期的格式，并按日期排序
        target_files = []
        for filename, full_path in file_list:
            date = self.extract_date_from_filename(filename)
            if date:
                target_files.append((filename, full_path, date))
        
        # 按日期排序
        target_files.sort(key=lambda x: x[2])
        
        print(f"正在处理 {len(target_files)} 天的数据...")
        
        for filename, full_path, date in target_files:
            try:
                df = self.load_data(full_path)
                
                if df is None or df.empty:
                    continue
                
                # 使用调整后的原合约值计算收入
                daily_scaled_revenue = self.calculate_total_revenue_for_contract(df, scaled_monthly_avg)
                if daily_scaled_revenue is not None:
                    daily_total = daily_scaled_revenue.sum() if hasattr(daily_scaled_revenue, 'sum') else sum(daily_scaled_revenue)
                    scaled_daily_revenues.append(daily_total)
                    total_scaled_revenue += daily_total
                    
            except Exception as e:
                print(f"处理 {date.strftime('%m-%d')} 时出错: {e}")
                continue
        
        # 计算原始收入（用于对比）
        original_daily_revenues = []
        total_original_revenue = 0
        
        for filename, full_path, date in target_files:
            try:
                df = self.load_data(full_path)
                
                if df is None or df.empty:
                    continue
                
                # 使用原始的原合约值计算收入
                daily_original_revenue = self.calculate_total_revenue_for_contract(df, original_monthly_avg)
                if daily_original_revenue is not None:
                    daily_total = daily_original_revenue.sum() if hasattr(daily_original_revenue, 'sum') else sum(daily_original_revenue)
                    original_daily_revenues.append(daily_total)
                    total_original_revenue += daily_total
                    
            except Exception as e:
                continue
        
        # 构建分析结果
        scaled_result = {
            'year': year,
            'month': month,
            'scale_factor': scale_factor,
            'scale_description': scale_desc,
            'original_monthly_average': original_monthly_avg,
            'scaled_monthly_average': scaled_monthly_avg,
            'original_total_revenue': total_original_revenue,
            'scaled_total_revenue': total_scaled_revenue,
            'revenue_difference': total_scaled_revenue - total_original_revenue,
            'original_daily_revenues': original_daily_revenues,
            'scaled_daily_revenues': scaled_daily_revenues,
            'days_count': len(scaled_daily_revenues),
            'original_total_contract': np.sum(original_monthly_avg),
            'scaled_total_contract': np.sum(scaled_monthly_avg),
            'daily_total_limit': monthly_result['daily_total_limit']
        }
        
        return scaled_result
    
    # 向后兼容性：保持原有的reduce_half函数名称
    def analyze_halved_contract(self, monthly_result):
        """分析原合约减半后的情况（向后兼容）"""
        return self.analyze_scaled_contract(monthly_result, scale_factor=0.5)
    
    def print_halved_contract_comparison(self, scaled_result):
        """打印原合约减半前后的详细对比（向后兼容）"""
        return self.print_scaled_contract_comparison(scaled_result)
    
    def plot_halved_contract_comparison(self, scaled_result, save_path=None):
        """绘制原合约减半前后的对比图表（向后兼容）"""
        return self.plot_scaled_contract_comparison(scaled_result, save_path)
    
    def print_scaled_contract_comparison(self, scaled_result):
        """打印原合约调整前后的详细对比"""
        if scaled_result is None:
            return
        
        year = scaled_result['year']
        month = scaled_result['month']
        scale_desc = scaled_result['scale_description']
        scale_factor = scaled_result['scale_factor']
        
        print(f"\n{'='*60}")
        print(f"    {year}年{month}月 原合约{scale_desc}调整前后对比分析")
        print(f"{'='*60}")
        
        # 基本统计对比
        print(f"\n📊 基本统计对比:")
        print(f"{'项目':<20} {'调整前':<15} {'调整后':<15} {'变化':<15}")
        print("-" * 70)
        
        orig_avg = np.mean(scaled_result['original_monthly_average'])
        scaled_avg = np.mean(scaled_result['scaled_monthly_average'])
        avg_change = scaled_avg - orig_avg
        avg_pct = (avg_change / orig_avg * 100) if orig_avg != 0 else 0
        
        print(f"{'月平均合约值':<18} {orig_avg:<15.3f} {scaled_avg:<15.3f} {avg_change:<+8.3f}({avg_pct:<+5.1f}%)")
        
        orig_total_contract = scaled_result['original_total_contract']
        scaled_total_contract = scaled_result['scaled_total_contract']
        contract_change = scaled_total_contract - orig_total_contract
        contract_pct = (contract_change / orig_total_contract * 100) if orig_total_contract != 0 else 0
        
        print(f"{'月总合约量':<20} {orig_total_contract:<15.3f} {scaled_total_contract:<15.3f} {contract_change:<+8.3f}({contract_pct:<+5.1f}%)")
        
        # 收入对比
        print(f"\n💰 收入对比:")
        print(f"{'项目':<20} {'调整前':<18} {'调整后':<18} {'变化':<20}")
        print("-" * 80)
        
        orig_revenue = scaled_result['original_total_revenue']
        scaled_revenue = scaled_result['scaled_total_revenue']
        revenue_diff = scaled_result['revenue_difference']
        revenue_pct = (revenue_diff / orig_revenue * 100) if orig_revenue != 0 else 0
        
        print(f"{'月总收入':<20} {orig_revenue:<18,.2f} {scaled_revenue:<18,.2f} {revenue_diff:<+10,.2f}({revenue_pct:<+7.1f}%)")
        
        # 每日平均收入
        if scaled_result['days_count'] > 0:
            orig_daily_avg = orig_revenue / scaled_result['days_count']
            scaled_daily_avg = scaled_revenue / scaled_result['days_count']
            daily_avg_diff = scaled_daily_avg - orig_daily_avg
            daily_avg_pct = (daily_avg_diff / orig_daily_avg * 100) if orig_daily_avg != 0 else 0
            
            print(f"{'日平均收入':<20} {orig_daily_avg:<18,.2f} {scaled_daily_avg:<18,.2f} {daily_avg_diff:<+10,.2f}({daily_avg_pct:<+7.1f}%)")
        
        # 约束验证
        print(f"\n🔍 约束验证:")
        limit = scaled_result['daily_total_limit']
        if limit:
            print(f"原始每日总量限制: {limit}")
            print(f"调整前实际总量: {orig_total_contract:.6f}")
            print(f"调整后实际总量: {scaled_total_contract:.6f}")
            
            orig_constraint_diff = abs(orig_total_contract - limit)
            scaled_constraint_diff = abs(scaled_total_contract - limit * scale_factor)
            
            if orig_constraint_diff < 1e-6:
                print("✓ 调整前：等式约束满足")
            else:
                print(f"✗ 调整前：等式约束不满足，差值: {orig_constraint_diff:.6f}")
            
            print(f"调整后建议新限制: {limit * scale_factor}")
            if scaled_constraint_diff < 1e-6:
                print("✓ 调整后：相对于新限制等式约束满足")
            else:
                print(f"✗ 调整后：相对于新限制等式约束不满足，差值: {scaled_constraint_diff:.6f}")
        
        # 分布保持验证
        print(f"\n📈 分布保持验证:")
        # 计算相关系数来验证分布是否保持不变
        correlation = np.corrcoef(scaled_result['original_monthly_average'], 
                                scaled_result['scaled_monthly_average'])[0, 1]
        print(f"原始与调整后分布相关系数: {correlation:.6f}")
        
        if abs(correlation - 1.0) < 1e-10:
            print("✓ 分布完全保持不变（相关系数 = 1.000000）")
        else:
            print(f"⚠ 分布轻微变化（相关系数 = {correlation:.6f}）")
        
        # 每个时间点的比值验证
        ratios = scaled_result['scaled_monthly_average'] / scaled_result['original_monthly_average']
        ratios = ratios[~np.isnan(ratios)]  # 移除NaN值
        
        if len(ratios) > 0:
            ratio_mean = np.mean(ratios)
            ratio_std = np.std(ratios)
            print(f"各时间点调整比例 - 平均值: {ratio_mean:.6f}, 标准差: {ratio_std:.6f}")
            
            if abs(ratio_mean - scale_factor) < 1e-6 and ratio_std < 1e-6:
                print(f"✓ 所有时间点均精确调整为{scale_factor}倍")
            else:
                print(f"⚠ 调整比例存在微小差异")
        
        print(f"\n{'='*60}")
    
    def plot_scaled_contract_comparison(self, scaled_result, save_path=None):
        """绘制原合约调整前后的对比图表"""
        if scaled_result is None:
            return
        
        fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(20, 12))
        
        year = scaled_result['year']
        month = scaled_result['month']
        scale_desc = scaled_result['scale_description']
        scale_factor = scaled_result['scale_factor']
        time_points = range(1, len(scaled_result['original_monthly_average']) + 1)
        
        # 1. 原合约值对比曲线
        ax1.plot(time_points, scaled_result['original_monthly_average'], 'b-', 
                linewidth=2, marker='o', markersize=3, label='调整前')
        ax1.plot(time_points, scaled_result['scaled_monthly_average'], 'r--', 
                linewidth=2, marker='s', markersize=3, label=f'调整后({scale_desc})')
        ax1.set_title(f'{year}年{month}月 原合约值对比', fontsize=14)
        ax1.set_xlabel('时间点 (15分钟间隔)')
        ax1.set_ylabel('原合约值')
        ax1.grid(True, alpha=0.3)
        ax1.legend()
        ax1.set_xticks(range(0, 97, 8))
        
        # 2. 调整比例验证
        ratios = scaled_result['scaled_monthly_average'] / scaled_result['original_monthly_average']
        ratios = np.where(np.isnan(ratios), 0, ratios)  # 处理除零情况
        
        ax2.plot(time_points, ratios, 'g-', linewidth=2, marker='^', markersize=3)
        ax2.axhline(y=scale_factor, color='red', linestyle='--', alpha=0.7, label=f'理论值 {scale_factor}')
        ax2.set_title(f'{year}年{month}月 调整比例验证', fontsize=14)
        ax2.set_xlabel('时间点 (15分钟间隔)')
        ax2.set_ylabel('调整后/调整前 比例')
        ax2.grid(True, alpha=0.3)
        ax2.legend()
        ax2.set_xticks(range(0, 97, 8))
        
        # 动态设置y轴范围
        if scale_factor <= 1:
            ax2.set_ylim(0, 1.2)
        else:
            ax2.set_ylim(0, scale_factor * 1.2)
        
        # 3. 每日收入对比
        if len(scaled_result['original_daily_revenues']) > 0:
            days = range(1, len(scaled_result['original_daily_revenues']) + 1)
            
            ax3.bar([d - 0.2 for d in days], scaled_result['original_daily_revenues'], 
                   width=0.4, label='调整前', alpha=0.7, color='blue')
            ax3.bar([d + 0.2 for d in days], scaled_result['scaled_daily_revenues'], 
                   width=0.4, label=f'调整后({scale_desc})', alpha=0.7, color='red')
            
            ax3.set_title(f'{year}年{month}月 每日收入对比', fontsize=14)
            ax3.set_xlabel('日期')
            ax3.set_ylabel('每日收入')
            ax3.grid(True, alpha=0.3)
            ax3.legend()
        
        # 4. 统计汇总
        ax4.axis('off')
        
        # 构建统计信息文本
        orig_total_contract = scaled_result['original_total_contract']
        scaled_total_contract = scaled_result['scaled_total_contract']
        orig_revenue = scaled_result['original_total_revenue']
        scaled_revenue = scaled_result['scaled_total_revenue']
        revenue_diff = scaled_result['revenue_difference']
        
        revenue_pct = (revenue_diff / orig_revenue * 100) if orig_revenue != 0 else 0
        
        stats_text = f"""
        {year}年{month}月 {scale_desc}调整分析汇总
        
        📊 合约量统计:
        调整前月总合约量: {orig_total_contract:.3f}
        调整后月总合约量: {scaled_total_contract:.3f}
        合约量变化: {scaled_total_contract - orig_total_contract:+.3f}
        
        💰 收入统计:
        调整前月总收入: {orig_revenue:,.2f}
        调整后月总收入: {scaled_revenue:,.2f}
        收入变化: {revenue_diff:+,.2f} ({revenue_pct:+.1f}%)
        
        🔍 约束信息:
        每日总量限制: {scaled_result['daily_total_limit']}
        处理天数: {scaled_result['days_count']}天
        调整比例: {scale_factor}
        
        ✅ 分布保持: 完全不变
        """
        
        ax4.text(0.1, 0.9, stats_text, transform=ax4.transAxes, fontsize=12,
                verticalalignment='top', fontfamily='monospace',
                bbox=dict(boxstyle='round', facecolor='lightgray', alpha=0.8))
        
        plt.tight_layout()
        
        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
            print(f"图表已保存为: {save_path}")
        
        plt.show()

    def analyze_daily_negative_revenue(self, target_date):
        """分析指定日期的负收益情况"""
        if isinstance(target_date, str):
            try:
                target_date = datetime.strptime(target_date, '%Y-%m-%d')
            except ValueError:
                print("日期格式错误，请使用 YYYY-MM-DD 格式")
                return None
        
        # 查找对应的Excel文件
        target_file, filepath = self.find_file_for_date(target_date)
        
        if target_file is None:
            print(f"未找到 {target_date.strftime('%Y-%m-%d')} 的数据文件")
            return None
        
        df = self.load_data(filepath)
        
        if df is None:
            return None
        
        # 查找各项收益列
        revenue_columns = {}
        for col in df.columns:
            col_str = str(col).lower()
            if '撮合收益' in col_str:
                revenue_columns['matching_revenue'] = col
            elif '日前结算收益' in col_str:
                revenue_columns['forward_settlement_revenue'] = col
            elif '实时结算收益' in col_str:
                revenue_columns['realtime_settlement_revenue'] = col
        
        if not revenue_columns:
            print("未找到相关收益列")
            return None
        
        # 分析负收益
        negative_analysis = {
            'date': target_date,
            'data': df,
            'revenue_columns': revenue_columns,
            'negative_summary': {},
            'worst_periods': {},
            'hourly_analysis': {}
        }
        
        date_str = target_date.strftime('%Y-%m-%d')
        print(f"\n=== {date_str} 负收益分析 ===")
        
        for revenue_type, col_name in revenue_columns.items():
            revenue_values = pd.to_numeric(df[col_name], errors='coerce').fillna(0)
            
            # 只分析负值
            negative_values = revenue_values[revenue_values < 0]
            negative_indices = revenue_values[revenue_values < 0].index
            
            revenue_name = col_name.replace('收益', '')
            
            if len(negative_values) > 0:
                total_negative = negative_values.sum()
                avg_negative = negative_values.mean()
                worst_value = negative_values.min()
                worst_index = revenue_values.idxmin()
                
                # 计算最亏损时间
                worst_hour = worst_index // 4
                worst_minute = (worst_index % 4) * 15
                worst_time = f"{worst_hour:02d}:{worst_minute:02d}"
                
                negative_analysis['negative_summary'][revenue_type] = {
                    'total_negative': total_negative,
                    'avg_negative': avg_negative,
                    'worst_value': worst_value,
                    'worst_time': worst_time,
                    'worst_index': worst_index,
                    'negative_count': len(negative_values),
                    'negative_percentage': len(negative_values) / len(revenue_values) * 100
                }
                
                print(f"\n📉 {revenue_name}负收益统计:")
                print(f"  负收益总额: {total_negative:,.2f}")
                print(f"  负收益平均: {avg_negative:.2f}")
                print(f"  最大亏损: {worst_value:.2f} (时间: {worst_time})")
                print(f"  负收益时段数: {len(negative_values)}/96 ({len(negative_values)/96*100:.1f}%)")
            else:
                negative_analysis['negative_summary'][revenue_type] = {
                    'total_negative': 0,
                    'avg_negative': 0,
                    'worst_value': 0,
                    'worst_time': 'N/A',
                    'worst_index': -1,
                    'negative_count': 0,
                    'negative_percentage': 0
                }
                print(f"\n✅ {revenue_name}: 当日无负收益")
        
        # 分析整体最差时段
        self._analyze_worst_periods(negative_analysis)
        
        # 按小时统计负收益
        self._analyze_hourly_negative_revenue(negative_analysis)
        
        return negative_analysis
    
    def _analyze_worst_periods(self, negative_analysis):
        """分析最差时段"""
        print(f"\n🔥 最差时段分析:")
        
        df = negative_analysis['data']
        revenue_columns = negative_analysis['revenue_columns']
        
        # 计算每个时间点的总负收益
        total_negative_by_time = []
        for i in range(len(df)):
            time_negative = 0
            for revenue_type, col_name in revenue_columns.items():
                value = pd.to_numeric(df[col_name].iloc[i], errors='coerce')
                if pd.notna(value) and value < 0:
                    time_negative += value
            total_negative_by_time.append(time_negative)
        
        # 找出最差的前5个时段
        sorted_indices = sorted(range(len(total_negative_by_time)), 
                              key=lambda i: total_negative_by_time[i])
        
        worst_periods = []
        for i in range(min(5, len([x for x in total_negative_by_time if x < 0]))):
            idx = sorted_indices[i]
            if total_negative_by_time[idx] < 0:
                hour = idx // 4
                minute = (idx % 4) * 15
                time_str = f"{hour:02d}:{minute:02d}"
                worst_periods.append({
                    'time': time_str,
                    'index': idx,
                    'total_negative': total_negative_by_time[idx]
                })
        
        negative_analysis['worst_periods'] = worst_periods
        
        if worst_periods:
            print("  排名  时间    总负收益")
            print("  " + "-" * 25)
            for i, period in enumerate(worst_periods, 1):
                print(f"  {i:2d}    {period['time']}   {period['total_negative']:>10.2f}")
        else:
            print("  当日无负收益时段")
    
    def _analyze_hourly_negative_revenue(self, negative_analysis):
        """按小时分析负收益"""
        print(f"\n📊 按小时负收益统计:")
        
        df = negative_analysis['data']
        revenue_columns = negative_analysis['revenue_columns']
        
        hourly_stats = {}
        
        for hour in range(24):
            hour_indices = [hour * 4 + i for i in range(4) if hour * 4 + i < len(df)]
            
            hour_negative = {
                'matching': 0,
                'forward': 0,
                'realtime': 0,
                'total': 0,
                'count': 0
            }
            
            for idx in hour_indices:
                for revenue_type, col_name in revenue_columns.items():
                    value = pd.to_numeric(df[col_name].iloc[idx], errors='coerce')
                    if pd.notna(value) and value < 0:
                        if 'matching' in revenue_type:
                            hour_negative['matching'] += value
                        elif 'forward' in revenue_type:
                            hour_negative['forward'] += value
                        elif 'realtime' in revenue_type:
                            hour_negative['realtime'] += value
                        hour_negative['total'] += value
                        hour_negative['count'] += 1
            
            hourly_stats[hour] = hour_negative
        
        negative_analysis['hourly_analysis'] = hourly_stats
        
        # 显示最差的几个小时
        sorted_hours = sorted(hourly_stats.items(), key=lambda x: x[1]['total'])
        worst_hours = [(h, stats) for h, stats in sorted_hours if stats['total'] < 0][:5]
        
        if worst_hours:
            print("  时段    撮合收益    日前结算    实时结算    小时总计")
            print("  " + "-" * 55)
            for hour, stats in worst_hours:
                print(f"  {hour:02d}:00   {stats['matching']:>8.2f}   {stats['forward']:>8.2f}   {stats['realtime']:>8.2f}   {stats['total']:>8.2f}")
        else:
            print("  当日无小时级负收益")
    
    def analyze_monthly_negative_revenue(self, year, month):
        """分析指定月份的负收益情况"""
        print(f"\n=== {year}年{month}月 负收益分析 ===")
        
        # 获取该月所有日期的数据
        file_list = self.get_monthly_excel_files(year, month)
        
        if not file_list:
            print(f"没有找到 {year}年{month}月 的数据文件")
            return None
        
        # 转换为包含日期的格式，并按日期排序
        target_files = []
        for filename, full_path in file_list:
            date = self.extract_date_from_filename(filename)
            if date:
                target_files.append((filename, date))
        
        # 按日期排序
        target_files.sort(key=lambda x: x[1])
        
        monthly_negative_analysis = {
            'year': year,
            'month': month,
            'daily_analyses': [],
            'monthly_summary': {},
            'worst_days': [],
            'time_pattern_analysis': {},
            'total_days': len(target_files)
        }
        
        print(f"正在分析 {len(target_files)} 天的负收益数据...")
        
        total_monthly_negative = {
            'matching': 0,
            'forward': 0,
            'realtime': 0,
            'total': 0
        }
        
        all_negative_periods = []
        
        # 分析每一天
        for filename, date in target_files:
            try:
                daily_analysis = self.analyze_daily_negative_revenue(date)
                if daily_analysis:
                    monthly_negative_analysis['daily_analyses'].append(daily_analysis)
                    
                    # 累计月度统计
                    for revenue_type, stats in daily_analysis['negative_summary'].items():
                        if 'matching' in revenue_type:
                            total_monthly_negative['matching'] += stats['total_negative']
                        elif 'forward' in revenue_type:
                            total_monthly_negative['forward'] += stats['total_negative']
                        elif 'realtime' in revenue_type:
                            total_monthly_negative['realtime'] += stats['total_negative']
                        total_monthly_negative['total'] += stats['total_negative']
                    
                    # 收集最差时段
                    for period in daily_analysis['worst_periods']:
                        all_negative_periods.append({
                            'date': date,
                            'time': period['time'],
                            'negative': period['total_negative']
                        })
                
            except Exception as e:
                print(f"处理 {date.strftime('%m-%d')} 时出错: {e}")
                continue
        
        # 月度汇总统计
        monthly_negative_analysis['monthly_summary'] = total_monthly_negative
        
        # 找出最差的日期
        daily_totals = []
        for daily_analysis in monthly_negative_analysis['daily_analyses']:
            daily_total = sum(stats['total_negative'] for stats in daily_analysis['negative_summary'].values())
            daily_totals.append({
                'date': daily_analysis['date'],
                'total_negative': daily_total
            })
        
        daily_totals.sort(key=lambda x: x['total_negative'])
        monthly_negative_analysis['worst_days'] = daily_totals[:10]  # 最差的10天
        
        # 时间模式分析
        self._analyze_monthly_time_patterns(monthly_negative_analysis)
        
        # 打印月度汇总
        self._print_monthly_negative_summary(monthly_negative_analysis)
        
        return monthly_negative_analysis
    
    def _analyze_monthly_time_patterns(self, monthly_analysis):
        """分析月度时间模式"""
        # 按时间点统计负收益频率和强度
        time_patterns = {}
        
        for daily_analysis in monthly_analysis['daily_analyses']:
            df = daily_analysis['data']
            revenue_columns = daily_analysis['revenue_columns']
            
            for i in range(len(df)):
                hour = i // 4
                minute = (i % 4) * 15
                time_key = f"{hour:02d}:{minute:02d}"
                
                if time_key not in time_patterns:
                    time_patterns[time_key] = {
                        'total_negative': 0,
                        'negative_days': 0,
                        'worst_negative': 0
                    }
                
                time_negative = 0
                for revenue_type, col_name in revenue_columns.items():
                    value = pd.to_numeric(df[col_name].iloc[i], errors='coerce')
                    if pd.notna(value) and value < 0:
                        time_negative += value
                
                if time_negative < 0:
                    time_patterns[time_key]['total_negative'] += time_negative
                    time_patterns[time_key]['negative_days'] += 1
                    time_patterns[time_key]['worst_negative'] = min(
                        time_patterns[time_key]['worst_negative'], time_negative
                    )
        
        monthly_analysis['time_pattern_analysis'] = time_patterns
    
    def _print_monthly_negative_summary(self, monthly_analysis):
        """打印月度负收益汇总"""
        year = monthly_analysis['year']
        month = monthly_analysis['month']
        
        print(f"\n{'='*60}")
        print(f"    {year}年{month}月 负收益汇总分析")
        print(f"{'='*60}")
        
        # 月度总计
        summary = monthly_analysis['monthly_summary']
        print(f"\n💰 月度负收益总计:")
        print(f"撮合收益亏损: {summary['matching']:>15,.2f}")
        print(f"日前结算亏损: {summary['forward']:>15,.2f}")
        print(f"实时结算亏损: {summary['realtime']:>15,.2f}")
        print(f"总亏损金额: {summary['total']:>15,.2f}")
        
        # 最差日期排行
        print(f"\n📉 最差日期排行 (前10天):")
        print("排名    日期        当日总亏损")
        print("-" * 35)
        for i, day_info in enumerate(monthly_analysis['worst_days'][:10], 1):
            if day_info['total_negative'] < 0:
                date_str = day_info['date'].strftime('%m-%d')
                print(f"{i:2d}    {date_str}     {day_info['total_negative']:>12,.2f}")
        
        # 时间模式分析
        time_patterns = monthly_analysis['time_pattern_analysis']
        worst_times = sorted(time_patterns.items(), 
                           key=lambda x: x[1]['total_negative'])[:10]
        
        print(f"\n⏰ 最差时间段排行 (前10个时段):")
        print("排名  时间   月度总亏损   亏损天数   最差单日")
        print("-" * 50)
        for i, (time_str, stats) in enumerate(worst_times, 1):
            if stats['total_negative'] < 0:
                print(f"{i:2d}   {time_str}   {stats['total_negative']:>10,.2f}   {stats['negative_days']:>6d}天   {stats['worst_negative']:>8.2f}")
        
        # 按小时汇总
        hourly_totals = {}
        for time_str, stats in time_patterns.items():
            hour = int(time_str.split(':')[0])
            if hour not in hourly_totals:
                hourly_totals[hour] = 0
            hourly_totals[hour] += stats['total_negative']
        
        worst_hours = sorted(hourly_totals.items(), key=lambda x: x[1])[:5]
        
        print(f"\n🕐 最差小时段排行:")
        print("排名  小时段    月度总亏损")
        print("-" * 25)
        for i, (hour, total_neg) in enumerate(worst_hours, 1):
            if total_neg < 0:
                print(f"{i:2d}   {hour:02d}:00    {total_neg:>10,.2f}")
    
    def plot_negative_revenue_analysis(self, analysis_result, analysis_type='daily', save_path=None):
        """绘制负收益分析图表"""
        if analysis_result is None:
            return
        
        if analysis_type == 'daily':
            self._plot_daily_negative_analysis(analysis_result, save_path)
        elif analysis_type == 'monthly':
            self._plot_monthly_negative_analysis(analysis_result, save_path)
    
    def _plot_daily_negative_analysis(self, daily_analysis, save_path=None):
        """绘制每日负收益分析图表（包含热力图）"""
        fig = plt.figure(figsize=(28, 20))
        
        # 创建3x2的子图布局，增加间距
        ax1 = plt.subplot(3, 2, 1)
        ax2 = plt.subplot(3, 2, 2) 
        ax3 = plt.subplot(3, 2, 3)
        ax4 = plt.subplot(3, 2, 4)
        ax5 = plt.subplot(3, 2, 5)
        ax6 = plt.subplot(3, 2, 6)
        
        # 调整子图之间的间距
        plt.subplots_adjust(hspace=0.4, wspace=0.3, top=0.95, bottom=0.05, left=0.05, right=0.95)
        
        date_str = daily_analysis['date'].strftime('%Y-%m-%d')
        df = daily_analysis['data']
        revenue_columns = daily_analysis['revenue_columns']
        
        time_points = range(1, len(df) + 1)
        
        # 1. 各项收益时间序列（只显示负值）
        for revenue_type, col_name in revenue_columns.items():
            revenue_values = pd.to_numeric(df[col_name], errors='coerce').fillna(0)
            negative_values = np.where(revenue_values < 0, revenue_values, np.nan)
            
            if revenue_type == 'matching_revenue':
                ax1.plot(time_points, negative_values, 'r-', linewidth=2, marker='o', markersize=2, label='撮合收益')
            elif revenue_type == 'forward_settlement_revenue':
                ax1.plot(time_points, negative_values, 'b-', linewidth=2, marker='s', markersize=2, label='日前结算')
            elif revenue_type == 'realtime_settlement_revenue':
                ax1.plot(time_points, negative_values, 'g-', linewidth=2, marker='^', markersize=2, label='实时结算')
        
        ax1.set_title(f'{date_str} 负收益时间分布', fontsize=12, pad=10)
        ax1.set_xlabel('时间点 (15分钟间隔)', fontsize=10)
        ax1.set_ylabel('负收益值', fontsize=10)
        ax1.grid(True, alpha=0.3)
        ax1.legend(fontsize=9)
        ax1.axhline(y=0, color='black', linestyle='--', alpha=0.5)
        ax1.set_xticks(range(0, 97, 12))  # 减少刻度数量
        ax1.tick_params(axis='both', which='major', labelsize=8)
        
        # 2. 按小时汇总的负收益柱状图
        hourly_data = daily_analysis['hourly_analysis']
        hours = list(range(24))
        matching_hourly = [hourly_data[h]['matching'] for h in hours]
        forward_hourly = [hourly_data[h]['forward'] for h in hours]
        realtime_hourly = [hourly_data[h]['realtime'] for h in hours]
        
        width = 0.25
        x = np.arange(len(hours))
        ax2.bar(x - width, matching_hourly, width, label='撮合收益', alpha=0.8, color='red')
        ax2.bar(x, forward_hourly, width, label='日前结算', alpha=0.8, color='blue')
        ax2.bar(x + width, realtime_hourly, width, label='实时结算', alpha=0.8, color='green')
        
        ax2.set_title(f'{date_str} 按小时负收益汇总', fontsize=12, pad=10)
        ax2.set_xlabel('小时', fontsize=10)
        ax2.set_ylabel('负收益值', fontsize=10)
        ax2.grid(True, alpha=0.3, axis='y')
        ax2.legend(fontsize=9)
        ax2.set_xticks(x[::2])  # 每2小时显示一个刻度
        ax2.set_xticklabels([f'{h:02d}' for h in hours[::2]], rotation=0, fontsize=8)
        ax2.tick_params(axis='y', which='major', labelsize=8)
        
        # 3. 日前结算亏损热力图
        if 'forward_settlement_revenue' in revenue_columns:
            forward_col = revenue_columns['forward_settlement_revenue']
            forward_values = pd.to_numeric(df[forward_col], errors='coerce').fillna(0)
            
            # 创建24x4的热力图数据（24小时，每小时4个15分钟）
            forward_heatmap = np.zeros((24, 4))
            forward_heatmap[:] = np.nan  # 初始化为NaN，这样正值不会显示
            
            for i, value in enumerate(forward_values[:96]):  # 确保不超过96个时间点
                if value < 0:  # 只显示负值（亏损）
                    hour = i // 4
                    quarter = i % 4
                    if hour < 24:
                        forward_heatmap[hour, quarter] = abs(value)  # 使用绝对值以便显示
            
            im3 = ax3.imshow(forward_heatmap, cmap='Reds', aspect='auto', interpolation='nearest')
            ax3.set_title(f'{date_str} 日前结算亏损热力图', fontsize=12, fontweight='bold', pad=10)
            ax3.set_xlabel('15分钟段', fontsize=10)
            ax3.set_ylabel('小时', fontsize=10)
            ax3.set_yticks(range(0, 24, 3))  # 减少y轴刻度
            ax3.set_yticklabels([f'{h:02d}:00' for h in range(0, 24, 3)], fontsize=8)
            ax3.set_xticks(range(4))
            ax3.set_xticklabels(['00', '15', '30', '45'], fontsize=8)
            
            # 添加颜色条
            cbar3 = plt.colorbar(im3, ax=ax3, shrink=0.7, pad=0.02)
            cbar3.set_label('亏损额（绝对值）', fontsize=9)
            cbar3.ax.tick_params(labelsize=8)
            
            # 在热力图上添加数值标注（只对较大的亏损）
            for i in range(24):
                for j in range(4):
                    if not np.isnan(forward_heatmap[i, j]) and forward_heatmap[i, j] > np.nanmax(forward_heatmap) * 0.3:
                        ax3.text(j, i, f'{forward_heatmap[i, j]:.1f}', 
                               ha='center', va='center', fontsize=8, color='white', fontweight='bold')
        
        # 4. 实时结算亏损热力图
        if 'realtime_settlement_revenue' in revenue_columns:
            realtime_col = revenue_columns['realtime_settlement_revenue']
            realtime_values = pd.to_numeric(df[realtime_col], errors='coerce').fillna(0)
            
            # 创建24x4的热力图数据
            realtime_heatmap = np.zeros((24, 4))
            realtime_heatmap[:] = np.nan  # 初始化为NaN
            
            for i, value in enumerate(realtime_values[:96]):
                if value < 0:  # 只显示负值（亏损）
                    hour = i // 4
                    quarter = i % 4
                    if hour < 24:
                        realtime_heatmap[hour, quarter] = abs(value)  # 使用绝对值
            
            im4 = ax4.imshow(realtime_heatmap, cmap='Blues', aspect='auto', interpolation='nearest')
            ax4.set_title(f'{date_str} 实时结算亏损热力图', fontsize=12, fontweight='bold', pad=10)
            ax4.set_xlabel('15分钟段', fontsize=10)
            ax4.set_ylabel('小时', fontsize=10)
            ax4.set_yticks(range(0, 24, 3))  # 减少y轴刻度
            ax4.set_yticklabels([f'{h:02d}:00' for h in range(0, 24, 3)], fontsize=8)
            ax4.set_xticks(range(4))
            ax4.set_xticklabels(['00', '15', '30', '45'], fontsize=8)
            
            # 添加颜色条
            cbar4 = plt.colorbar(im4, ax=ax4, shrink=0.7, pad=0.02)
            cbar4.set_label('亏损额（绝对值）', fontsize=9)
            cbar4.ax.tick_params(labelsize=8)
            
            # 在热力图上添加数值标注
            for i in range(24):
                for j in range(4):
                    if not np.isnan(realtime_heatmap[i, j]) and realtime_heatmap[i, j] > np.nanmax(realtime_heatmap) * 0.3:
                        ax4.text(j, i, f'{realtime_heatmap[i, j]:.1f}', 
                               ha='center', va='center', fontsize=8, color='white', fontweight='bold')
        
        # 5. 负收益分布直方图
        all_negatives = []
        for revenue_type, col_name in revenue_columns.items():
            revenue_values = pd.to_numeric(df[col_name], errors='coerce').fillna(0)
            negatives = revenue_values[revenue_values < 0]
            all_negatives.extend(negatives.tolist())
        
        if all_negatives:
            ax5.hist(all_negatives, bins=15, alpha=0.7, color='red', edgecolor='black')
            ax5.set_title(f'{date_str} 负收益分布', fontsize=12, pad=10)
            ax5.set_xlabel('负收益值', fontsize=10)
            ax5.set_ylabel('频次', fontsize=10)
            ax5.grid(True, alpha=0.3)
            ax5.tick_params(axis='both', which='major', labelsize=8)
        
        # 6. 统计信息
        ax6.axis('off')
        
        stats_text = f"""
        {date_str} 负收益统计汇总
        
        """
        
        for revenue_type, stats in daily_analysis['negative_summary'].items():
            revenue_name = revenue_type.replace('_revenue', '').replace('_', ' ').title()
            if stats['total_negative'] < 0:
                stats_text += f"""
        {revenue_name}:
        • 总亏损: {stats['total_negative']:,.2f}
        • 平均亏损: {stats['avg_negative']:.2f}
        • 最大亏损: {stats['worst_value']:.2f} ({stats['worst_time']})
        • 亏损时段: {stats['negative_count']}/96 ({stats['negative_percentage']:.1f}%)
        """
        
        if daily_analysis['worst_periods']:
            stats_text += "\n        最差时段 (前3名):\n"
            for i, period in enumerate(daily_analysis['worst_periods'][:3], 1):
                stats_text += f"        {i}. {period['time']} ({period['total_negative']:.2f})\n"
        
        ax6.text(0.05, 0.95, stats_text, transform=ax6.transAxes, fontsize=8,
                verticalalignment='top', fontfamily='monospace',
                bbox=dict(boxstyle='round', facecolor='lightcoral', alpha=0.8))
        
        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
            print(f"图表已保存为: {save_path}")
        
        plt.show()
    
    def _plot_monthly_negative_analysis(self, monthly_analysis, save_path=None):
        """绘制月度负收益分析图表（包含热力图）"""
        fig = plt.figure(figsize=(28, 20))
        
        # 创建3x2的子图布局，增加间距
        ax1 = plt.subplot(3, 2, 1)
        ax2 = plt.subplot(3, 2, 2)
        ax3 = plt.subplot(3, 2, 3)
        ax4 = plt.subplot(3, 2, 4)
        ax5 = plt.subplot(3, 2, 5)
        ax6 = plt.subplot(3, 2, 6)
        
        # 调整子图之间的间距
        plt.subplots_adjust(hspace=0.4, wspace=0.3, top=0.95, bottom=0.05, left=0.05, right=0.95)
        
        year = monthly_analysis['year']
        month = monthly_analysis['month']
        
        # 1. 每日负收益趋势
        daily_totals = [sum(stats['total_negative'] for stats in analysis['negative_summary'].values()) 
                       for analysis in monthly_analysis['daily_analyses']]
        dates = [analysis['date'].day for analysis in monthly_analysis['daily_analyses']]
        
        ax1.plot(dates, daily_totals, 'r-', linewidth=2, marker='o', markersize=4)
        ax1.fill_between(dates, daily_totals, 0, alpha=0.3, color='red')
        ax1.set_title(f'{year}年{month}月 每日负收益趋势', fontsize=12, pad=10)
        ax1.set_xlabel('日期', fontsize=10)
        ax1.set_ylabel('每日总负收益', fontsize=10)
        ax1.grid(True, alpha=0.3)
        ax1.axhline(y=0, color='black', linestyle='--', alpha=0.5)
        ax1.tick_params(axis='both', which='major', labelsize=8)
        
        # 2. 时间模式热力图
        time_patterns = monthly_analysis['time_pattern_analysis']
        
        # 创建24x4的矩阵（24小时，每小时4个15分钟）
        heatmap_data = np.zeros((24, 4))
        
        for time_str, stats in time_patterns.items():
            hour, minute = map(int, time_str.split(':'))
            quarter = minute // 15
            if quarter < 4:
                heatmap_data[hour, quarter] = stats['total_negative']
        
        im = ax2.imshow(heatmap_data, cmap='Reds', aspect='auto')
        ax2.set_title(f'{year}年{month}月 负收益时间热力图', fontsize=12, pad=10)
        ax2.set_xlabel('15分钟段', fontsize=10)
        ax2.set_ylabel('小时', fontsize=10)
        ax2.set_yticks(range(0, 24, 3))  # 减少y轴刻度
        ax2.set_yticklabels([f'{h:02d}:00' for h in range(0, 24, 3)], fontsize=8)
        ax2.set_xticks(range(4))
        ax2.set_xticklabels(['00', '15', '30', '45'], fontsize=8)
        cbar2 = plt.colorbar(im, ax=ax2, shrink=0.7, pad=0.02)
        cbar2.set_label('月度总负收益', fontsize=9)
        cbar2.ax.tick_params(labelsize=8)
        
        # 3. 日前结算亏损月度热力图
        forward_monthly_heatmap = np.zeros((24, 4))
        forward_monthly_heatmap[:] = np.nan
        
        # 汇总所有日期的日前结算亏损数据
        for daily_analysis in monthly_analysis['daily_analyses']:
            df = daily_analysis['data']
            revenue_columns = daily_analysis['revenue_columns']
            
            if 'forward_settlement_revenue' in revenue_columns:
                forward_col = revenue_columns['forward_settlement_revenue']
                forward_values = pd.to_numeric(df[forward_col], errors='coerce').fillna(0)
                
                for i, value in enumerate(forward_values[:96]):
                    if value < 0:
                        hour = i // 4
                        quarter = i % 4
                        if hour < 24:
                            if np.isnan(forward_monthly_heatmap[hour, quarter]):
                                forward_monthly_heatmap[hour, quarter] = 0
                            forward_monthly_heatmap[hour, quarter] += abs(value)
        
        # 将累计值为0的位置设为NaN
        forward_monthly_heatmap[forward_monthly_heatmap == 0] = np.nan
        
        im3 = ax3.imshow(forward_monthly_heatmap, cmap='Reds', aspect='auto', interpolation='nearest')
        ax3.set_title(f'{year}年{month}月 日前结算月度亏损热力图', fontsize=12, fontweight='bold', pad=10)
        ax3.set_xlabel('15分钟段', fontsize=10)
        ax3.set_ylabel('小时', fontsize=10)
        ax3.set_yticks(range(0, 24, 3))  # 减少y轴刻度
        ax3.set_yticklabels([f'{h:02d}:00' for h in range(0, 24, 3)], fontsize=8)
        ax3.set_xticks(range(4))
        ax3.set_xticklabels(['00', '15', '30', '45'], fontsize=8)
        cbar3 = plt.colorbar(im3, ax=ax3, shrink=0.7, pad=0.02)
        cbar3.set_label('月度累计亏损额', fontsize=9)
        cbar3.ax.tick_params(labelsize=8)
        
        # 4. 实时结算亏损月度热力图
        realtime_monthly_heatmap = np.zeros((24, 4))
        realtime_monthly_heatmap[:] = np.nan
        
        # 汇总所有日期的实时结算亏损数据
        for daily_analysis in monthly_analysis['daily_analyses']:
            df = daily_analysis['data']
            revenue_columns = daily_analysis['revenue_columns']
            
            if 'realtime_settlement_revenue' in revenue_columns:
                realtime_col = revenue_columns['realtime_settlement_revenue']
                realtime_values = pd.to_numeric(df[realtime_col], errors='coerce').fillna(0)
                
                for i, value in enumerate(realtime_values[:96]):
                    if value < 0:
                        hour = i // 4
                        quarter = i % 4
                        if hour < 24:
                            if np.isnan(realtime_monthly_heatmap[hour, quarter]):
                                realtime_monthly_heatmap[hour, quarter] = 0
                            realtime_monthly_heatmap[hour, quarter] += abs(value)
        
        # 将累计值为0的位置设为NaN
        realtime_monthly_heatmap[realtime_monthly_heatmap == 0] = np.nan
        
        im4 = ax4.imshow(realtime_monthly_heatmap, cmap='Blues', aspect='auto', interpolation='nearest')
        ax4.set_title(f'{year}年{month}月 实时结算月度亏损热力图', fontsize=12, fontweight='bold', pad=10)
        ax4.set_xlabel('15分钟段', fontsize=10)
        ax4.set_ylabel('小时', fontsize=10)
        ax4.set_yticks(range(0, 24, 3))  # 减少y轴刻度
        ax4.set_yticklabels([f'{h:02d}:00' for h in range(0, 24, 3)], fontsize=8)
        ax4.set_xticks(range(4))
        ax4.set_xticklabels(['00', '15', '30', '45'], fontsize=8)
        cbar4 = plt.colorbar(im4, ax=ax4, shrink=0.7, pad=0.02)
        cbar4.set_label('月度累计亏损额', fontsize=9)
        cbar4.ax.tick_params(labelsize=8)
        
        # 5. 各类收益负值占比
        summary = monthly_analysis['monthly_summary']
        revenue_types = ['撮合收益', '日前结算', '实时结算']
        negative_values = [abs(summary['matching']), abs(summary['forward']), abs(summary['realtime'])]
        
        if sum(negative_values) > 0:
            ax5.pie(negative_values, labels=revenue_types, autopct='%1.1f%%', startangle=90, textprops={'fontsize': 9})
            ax5.set_title(f'{year}年{month}月 负收益构成', fontsize=12, pad=10)
        
        # 6. 最差时段排行
        worst_times = sorted(time_patterns.items(), key=lambda x: x[1]['total_negative'])[:10]
        
        if worst_times and worst_times[0][1]['total_negative'] < 0:
            times = [item[0] for item in worst_times if item[1]['total_negative'] < 0]
            values = [abs(item[1]['total_negative']) for item in worst_times if item[1]['total_negative'] < 0]
            
            ax6.barh(range(len(times)), values, color='red', alpha=0.7)
            ax6.set_yticks(range(len(times)))
            ax6.set_yticklabels(times, fontsize=8)
            ax6.set_xlabel('月度总亏损 (绝对值)', fontsize=10)
            ax6.set_title(f'{year}年{month}月 最差时段排行', fontsize=12, pad=10)
            ax6.grid(True, alpha=0.3, axis='x')
            ax6.tick_params(axis='x', which='major', labelsize=8)
        
        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
            print(f"图表已保存为: {save_path}")
        
        plt.show()

    def find_original_contract_column(self, df):
        """查找原始合约量列"""
        for col in df.columns:
            col_str = str(col).lower()
            if '原合约' in col_str:
                return col
            elif '合约量' in col_str and '原' in col_str:
                return col
            elif '原始合约' in col_str:
                return col
        
        # 如果没找到明确的原合约列，尝试其他可能的列名
        for col in df.columns:
            col_str = str(col).lower()
            if '合约' in col_str and ('量' in col_str or '电量' in col_str):
                return col
        
        return None
    
    def analyze_original_contract_scaling(self, target_date, scale_factors=[0.25, 0.5, 1.5]):
        """分析指定日期使用原始合约量的比例调整"""
        if isinstance(target_date, str):
            try:
                target_date = datetime.strptime(target_date, '%Y-%m-%d')
            except ValueError:
                print("日期格式错误，请使用 YYYY-MM-DD 格式")
                return None
        
        # 查找对应的Excel文件
        target_file, filepath = self.find_file_for_date(target_date)
        
        if target_file is None:
            print(f"未找到 {target_date.strftime('%Y-%m-%d')} 的数据文件")
            return None
        
        df = self.load_data(filepath)
        
        if df is None:
            return None
        
        # 查找原始合约量列
        original_contract_col = self.find_original_contract_column(df)
        if original_contract_col is None:
            print("未找到原始合约量列")
            print("可用的列名:")
            for col in df.columns:
                print(f"  - {col}")
            return None
        
        print(f"找到原始合约量列: {original_contract_col}")
        
        # 获取原始合约量数据
        original_contract_values = pd.to_numeric(df[original_contract_col], errors='coerce').fillna(0)
        
        if len(original_contract_values) == 0:
            print("原始合约量数据为空")
            return None
        
        # 确保有96个时间点
        if len(original_contract_values) < 96:
            print(f"原始合约量数据不足96个时间点，实际{len(original_contract_values)}个")
            # 填充到96个点
            original_contract_values = original_contract_values.reindex(range(96), fill_value=0)
        elif len(original_contract_values) > 96:
            # 只取前96个点
            original_contract_values = original_contract_values[:96]
        
        date_str = target_date.strftime('%Y-%m-%d')
        print(f"\n=== {date_str} 原始合约量比例调整分析 ===")
        print(f"原始合约量总和: {original_contract_values.sum():.3f}")
        print(f"原始合约量平均值: {original_contract_values.mean():.3f}")
        
        # 检查数据列
        volume_cols = self.find_volume_columns(df)
        price_cols = self.find_price_columns(df)
        
        print(f"\n📊 数据列检查:")
        print(f"  找到电价列: {list(price_cols.keys())}")
        print(f"  找到电量列: {list(volume_cols.keys())}")
        
        # 询问是否开启详细计算信息
        debug_choice = input("\n是否显示详细的收益计算过程? (y/n): ").strip().lower()
        if debug_choice in ['y', 'yes', '是']:
            self.debug = True
            print("✓ 已开启详细计算信息")
        else:
            self.debug = False
        
        # 对每个比例因子进行分析
        scaling_results = {}
        
        for scale_factor in scale_factors:
            # 调整原合约量
            scaled_contract_values = original_contract_values * scale_factor
            
            # 计算调整后的收益
            scaled_revenue = self.calculate_total_revenue_for_contract(df, scaled_contract_values)
            
            if scaled_revenue is not None:
                total_revenue = scaled_revenue.sum() if hasattr(scaled_revenue, 'sum') else sum(scaled_revenue)
                
                # 根据缩放因子设置描述词
                if scale_factor == 0.25:
                    scale_desc = "四分之一"
                elif scale_factor == 0.5:
                    scale_desc = "减半"
                elif scale_factor == 1.5:
                    scale_desc = "1.5倍"
                elif scale_factor == 2.0:
                    scale_desc = "双倍"
                else:
                    scale_desc = f"{scale_factor}倍"
                
                scaling_results[scale_factor] = {
                    'scale_description': scale_desc,
                    'scaled_contract_values': scaled_contract_values,
                    'scaled_revenue': scaled_revenue,
                    'total_revenue': total_revenue,
                    'total_contract': scaled_contract_values.sum(),
                    'avg_contract': scaled_contract_values.mean()
                }
                
                print(f"\n{scale_desc}调整 (×{scale_factor}):")
                print(f"  调整后合约量总和: {scaled_contract_values.sum():.3f}")
                print(f"  调整后合约量平均: {scaled_contract_values.mean():.3f}")
                print(f"  总收益: {total_revenue:,.2f}")
                
                # 如果开启调试模式，显示收益构成分析
                if hasattr(self, 'debug') and self.debug:
                    # 分别计算各部分收益
                    contract_revenue_part = 0
                    forward_settlement_part = 0
                    other_revenue_part = 0
                    
                    # 合约收益部分
                    if 'contract_price' in price_cols:
                        contract_prices = pd.to_numeric(df[price_cols['contract_price']], errors='coerce').fillna(0)
                        if len(contract_prices) >= 96:
                            contract_revenue_part = np.sum(scaled_contract_values[:96] * contract_prices[:96])
                    
                    # 日前结算收益部分（重新计算以显示）
                    if ('forward_price' in price_cols and 'forward_clearing' in volume_cols):
                        forward_prices = pd.to_numeric(df[price_cols['forward_price']], errors='coerce').fillna(0)
                        forward_clearing = pd.to_numeric(df[volume_cols['forward_clearing']], errors='coerce').fillna(0)
                        
                        matching_sell = np.zeros(96)
                        matching_buy = np.zeros(96)
                        interprovincial_volume = np.zeros(96)
                        
                        if 'matching_sell' in volume_cols:
                            matching_sell_data = pd.to_numeric(df[volume_cols['matching_sell']], errors='coerce').fillna(0)
                            if len(matching_sell_data) >= 96:
                                matching_sell = matching_sell_data[:96]
                        
                        if 'matching_buy' in volume_cols:
                            matching_buy_data = pd.to_numeric(df[volume_cols['matching_buy']], errors='coerce').fillna(0)
                            if len(matching_buy_data) >= 96:
                                matching_buy = matching_buy_data[:96]
                        
                        if 'interprovincial_volume' in volume_cols:
                            interprovincial_data = pd.to_numeric(df[volume_cols['interprovincial_volume']], errors='coerce').fillna(0)
                            if len(interprovincial_data) >= 96:
                                interprovincial_volume = interprovincial_data[:96]
                        
                        if len(forward_clearing) >= 96 and len(forward_prices) >= 96:
                            forward_settlement_volume = (
                                forward_clearing[:96] 
                                - scaled_contract_values[:96] * 4 
                                - matching_sell 
                                - matching_buy 
                                - interprovincial_volume * 4
                            ) / 4
                            forward_settlement_part = np.sum(forward_settlement_volume * forward_prices[:96])
                    
                    # 其他固定收益
                    other_revenue_part = total_revenue - contract_revenue_part - forward_settlement_part
                    
                    print(f"  📈 收益构成分析:")
                    print(f"    合约收益: {contract_revenue_part:,.2f}")
                    print(f"    日前结算收益: {forward_settlement_part:,.2f}")
                    print(f"    其他收益: {other_revenue_part:,.2f}")
                    print(f"    总计: {contract_revenue_part + forward_settlement_part + other_revenue_part:,.2f}")
        
        # 计算原始收益（用于对比）
        original_revenue = self.calculate_total_revenue_for_contract(df, original_contract_values)
        original_total_revenue = original_revenue.sum() if original_revenue is not None else 0
        
        print(f"\n原始合约收益: {original_total_revenue:,.2f}")
        
        # 如果开启调试模式，显示原始收益构成
        if hasattr(self, 'debug') and self.debug:
            # 计算原始收益的各部分
            orig_contract_revenue_part = 0
            orig_forward_settlement_part = 0
            orig_other_revenue_part = 0
            
            # 合约收益部分
            if 'contract_price' in price_cols:
                contract_prices = pd.to_numeric(df[price_cols['contract_price']], errors='coerce').fillna(0)
                if len(contract_prices) >= 96:
                    orig_contract_revenue_part = np.sum(original_contract_values[:96] * contract_prices[:96])
            
            # 日前结算收益部分
            if ('forward_price' in price_cols and 'forward_clearing' in volume_cols):
                forward_prices = pd.to_numeric(df[price_cols['forward_price']], errors='coerce').fillna(0)
                forward_clearing = pd.to_numeric(df[volume_cols['forward_clearing']], errors='coerce').fillna(0)
                
                matching_sell = np.zeros(96)
                matching_buy = np.zeros(96)
                interprovincial_volume = np.zeros(96)
                
                if 'matching_sell' in volume_cols:
                    matching_sell_data = pd.to_numeric(df[volume_cols['matching_sell']], errors='coerce').fillna(0)
                    if len(matching_sell_data) >= 96:
                        matching_sell = matching_sell_data[:96]
                
                if 'matching_buy' in volume_cols:
                    matching_buy_data = pd.to_numeric(df[volume_cols['matching_buy']], errors='coerce').fillna(0)
                    if len(matching_buy_data) >= 96:
                        matching_buy = matching_buy_data[:96]
                
                if 'interprovincial_volume' in volume_cols:
                    interprovincial_data = pd.to_numeric(df[volume_cols['interprovincial_volume']], errors='coerce').fillna(0)
                    if len(interprovincial_data) >= 96:
                        interprovincial_volume = interprovincial_data[:96]
                
                if len(forward_clearing) >= 96 and len(forward_prices) >= 96:
                    orig_forward_settlement_volume = (
                        forward_clearing[:96] 
                        - original_contract_values[:96] * 4 
                        - matching_sell 
                        - matching_buy 
                        - interprovincial_volume * 4
                    ) / 4
                    orig_forward_settlement_part = np.sum(orig_forward_settlement_volume * forward_prices[:96])
            
            # 其他固定收益
            orig_other_revenue_part = original_total_revenue - orig_contract_revenue_part - orig_forward_settlement_part
            
            print(f"📈 原始收益构成分析:")
            print(f"  合约收益: {orig_contract_revenue_part:,.2f}")
            print(f"  日前结算收益: {orig_forward_settlement_part:,.2f}")
            print(f"  其他收益: {orig_other_revenue_part:,.2f}")
            print(f"  总计: {orig_contract_revenue_part + orig_forward_settlement_part + orig_other_revenue_part:,.2f}")
        
        # 构建完整的分析结果
        analysis_result = {
            'date': target_date,
            'data': df,
            'original_contract_column': original_contract_col,
            'original_contract_values': original_contract_values,
            'original_total_revenue': original_total_revenue,
            'original_revenue': original_revenue,
            'scaling_results': scaling_results,
            'scale_factors': scale_factors
        }
        
        return analysis_result
    
    def analyze_monthly_original_contract_scaling(self, year, month, scale_factors=[0.25, 0.5, 1.5]):
        """分析指定月份使用原始合约量的比例调整"""
        print(f"\n=== {year}年{month}月 原始合约量比例调整分析 ===")
        
        # 获取该月所有日期的数据
        file_list = self.get_monthly_excel_files(year, month)
        
        if not file_list:
            print(f"没有找到 {year}年{month}月 的数据文件")
            return None
        
        # 转换为包含日期的格式，并按日期排序
        target_files = []
        for filename, full_path in file_list:
            date = self.extract_date_from_filename(filename)
            if date:
                target_files.append((filename, date))
        
        # 按日期排序
        target_files.sort(key=lambda x: x[1])
        
        print(f"找到 {len(target_files)} 个数据文件，开始处理...")
        
        # 收集所有日期的原始合约量数据
        monthly_original_contracts = []
        monthly_scaling_results = {factor: [] for factor in scale_factors}
        monthly_original_revenues = []
        failed_dates = []
        
        for filename, date in target_files:
            try:
                daily_result = self.analyze_original_contract_scaling(date, scale_factors)
                if daily_result:
                    monthly_original_contracts.append(daily_result['original_contract_values'])
                    monthly_original_revenues.append(daily_result['original_total_revenue'])
                    
                    for factor in scale_factors:
                        if factor in daily_result['scaling_results']:
                            monthly_scaling_results[factor].append(
                                daily_result['scaling_results'][factor]['total_revenue']
                            )
                        else:
                            monthly_scaling_results[factor].append(0)
                else:
                    failed_dates.append(date.strftime('%Y-%m-%d'))
                    
            except Exception as e:
                print(f"处理 {date.strftime('%m-%d')} 时出错: {e}")
                failed_dates.append(date.strftime('%Y-%m-%d'))
                continue
        
        if not monthly_original_contracts:
            print("没有成功处理任何数据")
            return None
        
        # 计算月度统计
        monthly_avg_contract = np.mean(monthly_original_contracts, axis=0)
        monthly_std_contract = np.std(monthly_original_contracts, axis=0)
        total_original_revenue = sum(monthly_original_revenues)
        
        # 计算各比例因子的月度统计
        monthly_stats = {}
        for factor in scale_factors:
            revenues = monthly_scaling_results[factor]
            total_scaled_revenue = sum(revenues)
            revenue_difference = total_scaled_revenue - total_original_revenue
            
            if factor == 0.25:
                scale_desc = "四分之一"
            elif factor == 0.5:
                scale_desc = "减半"
            elif factor == 1.5:
                scale_desc = "1.5倍"
            elif factor == 2.0:
                scale_desc = "双倍"
            else:
                scale_desc = f"{factor}倍"
            
            monthly_stats[factor] = {
                'scale_description': scale_desc,
                'total_revenue': total_scaled_revenue,
                'revenue_difference': revenue_difference,
                'revenue_change_pct': (revenue_difference / total_original_revenue * 100) if total_original_revenue != 0 else 0,
                'daily_revenues': revenues,
                'avg_daily_revenue': np.mean(revenues) if revenues else 0,
                'scaled_monthly_avg_contract': monthly_avg_contract * factor,
                'scaled_total_contract': np.sum(monthly_avg_contract) * factor
            }
        
        # 构建月度分析结果
        monthly_result = {
            'year': year,
            'month': month,
            'original_monthly_avg_contract': monthly_avg_contract,
            'original_monthly_std_contract': monthly_std_contract,
            'original_total_revenue': total_original_revenue,
            'original_total_contract': np.sum(monthly_avg_contract),
            'monthly_stats': monthly_stats,
            'scale_factors': scale_factors,
            'days_count': len(monthly_original_contracts),
            'failed_dates': failed_dates
        }
        
        return monthly_result
    
    def print_monthly_original_scaling_comparison(self, monthly_result):
        """打印月度原始合约量比例调整对比"""
        if monthly_result is None:
            return
        
        year = monthly_result['year']
        month = monthly_result['month']
        
        print(f"\n{'='*70}")
        print(f"    {year}年{month}月 原始合约量比例调整对比分析")
        print(f"{'='*70}")
        
        print(f"\n📊 基础信息:")
        print(f"处理天数: {monthly_result['days_count']}天")
        print(f"原始月总合约量: {monthly_result['original_total_contract']:.3f}")
        print(f"原始月总收益: {monthly_result['original_total_revenue']:,.2f}")
        
        if monthly_result['failed_dates']:
            print(f"失败日期: {', '.join(monthly_result['failed_dates'])}")
        
        # 比例调整对比表
        print(f"\n💰 比例调整收益对比:")
        print(f"{'调整比例':<12} {'描述':<8} {'月总收益':<15} {'收益变化':<15} {'变化百分比':<10}")
        print("-" * 70)
        
        # 原始数据行
        print(f"{'1.0倍':<12} {'原始':<8} {monthly_result['original_total_revenue']:<15,.2f} {'0':<15} {'0.0%':<10}")
        
        # 各比例因子的数据行
        for factor in monthly_result['scale_factors']:
            stats = monthly_result['monthly_stats'][factor]
            desc = stats['scale_description']
            total_revenue = stats['total_revenue']
            revenue_diff = stats['revenue_difference']
            change_pct = stats['revenue_change_pct']
            
            print(f"{f'{factor}倍':<12} {desc:<8} {total_revenue:<15,.2f} {revenue_diff:<+15,.2f} {change_pct:<+7.1f}{'%'}")
        
        # 最佳和最差比例
        best_factor = max(monthly_result['scale_factors'], 
                         key=lambda f: monthly_result['monthly_stats'][f]['total_revenue'])
        worst_factor = min(monthly_result['scale_factors'], 
                          key=lambda f: monthly_result['monthly_stats'][f]['total_revenue'])
        
        print(f"\n🏆 最佳调整比例: {best_factor}倍 ({monthly_result['monthly_stats'][best_factor]['scale_description']})")
        print(f"   最佳月总收益: {monthly_result['monthly_stats'][best_factor]['total_revenue']:,.2f}")
        print(f"   相比原始增益: {monthly_result['monthly_stats'][best_factor]['revenue_difference']:+,.2f}")
        
        print(f"\n📉 最差调整比例: {worst_factor}倍 ({monthly_result['monthly_stats'][worst_factor]['scale_description']})")
        print(f"   最差月总收益: {monthly_result['monthly_stats'][worst_factor]['total_revenue']:,.2f}")
        print(f"   相比原始变化: {monthly_result['monthly_stats'][worst_factor]['revenue_difference']:+,.2f}")
        
        # 日平均收益对比
        print(f"\n📅 日平均收益对比:")
        print(f"{'调整比例':<12} {'日平均收益':<15}")
        print("-" * 30)
        print(f"{'1.0倍':<12} {monthly_result['original_total_revenue']/monthly_result['days_count']:<15,.2f}")
        
        for factor in monthly_result['scale_factors']:
            stats = monthly_result['monthly_stats'][factor]
            avg_daily = stats['avg_daily_revenue']
            print(f"{f'{factor}倍':<12} {avg_daily:<15,.2f}")
        
        print(f"\n{'='*70}")

    def find_optimal_scale_factor(self, target_date, search_range=(0.1, 3.0), method='grid'):
        """自动寻找总收入最高的调整比例"""
        if isinstance(target_date, str):
            try:
                target_date = datetime.strptime(target_date, '%Y-%m-%d')
            except ValueError:
                print("日期格式错误，请使用 YYYY-MM-DD 格式")
                return None
        
        # 查找对应的Excel文件
        target_file, filepath = self.find_file_for_date(target_date)
        
        if target_file is None:
            print(f"未找到 {target_date.strftime('%Y-%m-%d')} 的数据文件")
            return None
        df = self.load_data(filepath)
        
        if df is None:
            return None
        
        # 查找原始合约量列
        original_contract_col = self.find_original_contract_column(df)
        if original_contract_col is None:
            print("未找到原始合约量列")
            return None
        
        # 获取原始合约量数据
        original_contract_values = pd.to_numeric(df[original_contract_col], errors='coerce').fillna(0)
        
        if len(original_contract_values) == 0:
            print("原始合约量数据为空")
            return None
        
        # 确保有96个时间点
        if len(original_contract_values) < 96:
            original_contract_values = original_contract_values.reindex(range(96), fill_value=0)
        elif len(original_contract_values) > 96:
            original_contract_values = original_contract_values[:96]
        
        date_str = target_date.strftime('%Y-%m-%d')
        print(f"\n=== {date_str} 自动寻找最佳调整比例 ===")
        print(f"搜索范围: {search_range[0]:.1f} - {search_range[1]:.1f}")
        print(f"搜索方法: {method}")
        
        # 定义目标函数（要最大化收益，所以返回负收益用于最小化）
        def objective_function(scale_factor):
            try:
                scaled_contract_values = original_contract_values * scale_factor
                scaled_revenue = self.calculate_total_revenue_for_contract(df, scaled_contract_values)
                if scaled_revenue is not None:
                    total_revenue = scaled_revenue.sum() if hasattr(scaled_revenue, 'sum') else sum(scaled_revenue)
                    return -total_revenue  # 返回负值用于最小化
                else:
                    return float('inf')  # 如果计算失败，返回很大的值
            except Exception as e:
                return float('inf')
        
        # 选择搜索方法
        if method == 'grid':
            print("使用网格搜索方法...")
            # 网格搜索
            grid_points = np.linspace(search_range[0], search_range[1], 101)  # 101个点，精度0.029
            best_scale = None
            best_revenue = -float('inf')
            
            for i, scale in enumerate(grid_points):
                revenue = -objective_function(scale)
                if revenue > best_revenue:
                    best_revenue = revenue
                    best_scale = scale
                
                # 显示进度
                if (i + 1) % 20 == 0:
                    print(f"  进度: {i+1}/101 ({(i+1)/101*100:.1f}%)")
            
            optimal_result = {'x': best_scale, 'fun': -best_revenue, 'success': True}
            
        else:
            print("使用黄金分割搜索方法...")
            # 使用scipy的minimize_scalar进行优化
            optimal_result = minimize_scalar(
                objective_function, 
                bounds=search_range, 
                method='bounded',
                options={'xatol': 0.001}  # 精度设置为0.001
            )
        
        if optimal_result['success']:
            optimal_scale = optimal_result['x']
            optimal_revenue = -optimal_result['fun']
            
            # 计算原始收益用于对比
            original_revenue = self.calculate_total_revenue_for_contract(df, original_contract_values)
            original_total_revenue = original_revenue.sum() if original_revenue is not None else 0
            
            # 计算最佳比例下的详细结果
            optimal_contract_values = original_contract_values * optimal_scale
            
            print(f"\n🎯 找到最佳调整比例！")
            print(f"最佳比例: {optimal_scale:.4f}")
            print(f"原始收益: {original_total_revenue:,.2f}")
            print(f"最佳收益: {optimal_revenue:,.2f}")
            print(f"收益提升: {optimal_revenue - original_total_revenue:+,.2f}")
            print(f"提升百分比: {(optimal_revenue - original_total_revenue) / original_total_revenue * 100:+.2f}%")
            
            # 返回结果
            return {
                'date': target_date,
                'data': df,
                'original_contract_column': original_contract_col,
                'original_contract_values': original_contract_values,
                'optimal_scale_factor': optimal_scale,
                'optimal_contract_values': optimal_contract_values,
                'original_total_revenue': original_total_revenue,
                'optimal_total_revenue': optimal_revenue,
                'revenue_improvement': optimal_revenue - original_total_revenue,
                'improvement_percentage': (optimal_revenue - original_total_revenue) / original_total_revenue * 100 if original_total_revenue != 0 else 0,
                'search_range': search_range,
                'search_method': method
            }
        else:
            print("❌ 优化搜索失败")
            return None
    
    def find_optimal_scale_factor_monthly(self, year, month, search_range=(0.1, 3.0), method='grid'):
        """自动寻找月度总收入最高的调整比例"""
        print(f"\n=== {year}年{month}月 自动寻找最佳调整比例 ===")
        print(f"搜索范围: {search_range[0]:.1f} - {search_range[1]:.1f}")
        print(f"搜索方法: {method}")
        
        # 获取该月所有日期的数据
        file_list = self.get_monthly_excel_files(year, month)
        
        if not file_list:
            print(f"没有找到 {year}年{month}月 的数据文件")
            return None
        
        # 转换为包含日期的格式，并按日期排序
        target_files = []
        for filename, full_path in file_list:
            date = self.extract_date_from_filename(filename)
            if date:
                target_files.append((filename, full_path, date))
        
        # 按日期排序
        target_files.sort(key=lambda x: x[2])
        print(f"找到 {len(target_files)} 个数据文件")
        
        # 预处理所有数据
        daily_data = []
        failed_dates = []
        
        print("正在预处理数据...")
        for filename, full_path, date in target_files:
            try:
                df = self.load_data(full_path)
                
                if df is None:
                    failed_dates.append(date.strftime('%Y-%m-%d'))
                    continue
                
                # 查找原始合约量列
                original_contract_col = self.find_original_contract_column(df)
                if original_contract_col is None:
                    failed_dates.append(date.strftime('%Y-%m-%d'))
                    continue
                
                # 获取原始合约量数据
                original_contract_values = pd.to_numeric(df[original_contract_col], errors='coerce').fillna(0)
                
                if len(original_contract_values) == 0:
                    failed_dates.append(date.strftime('%Y-%m-%d'))
                    continue
                
                # 确保有96个时间点
                if len(original_contract_values) < 96:
                    original_contract_values = original_contract_values.reindex(range(96), fill_value=0)
                elif len(original_contract_values) > 96:
                    original_contract_values = original_contract_values[:96]
                
                daily_data.append({
                    'date': date,
                    'df': df,
                    'original_contract_values': original_contract_values
                })
                
            except Exception as e:
                print(f"处理 {date.strftime('%m-%d')} 时出错: {e}")
                failed_dates.append(date.strftime('%Y-%m-%d'))
                continue
        
        if not daily_data:
            print("没有成功处理任何数据")
            return None
        
        print(f"成功预处理 {len(daily_data)} 天的数据")
        if failed_dates:
            print(f"失败日期: {', '.join(failed_dates)}")
        
        # 定义月度目标函数
        def monthly_objective_function(scale_factor):
            try:
                total_monthly_revenue = 0
                for day_data in daily_data:
                    scaled_contract_values = day_data['original_contract_values'] * scale_factor
                    scaled_revenue = self.calculate_total_revenue_for_contract(day_data['df'], scaled_contract_values)
                    if scaled_revenue is not None:
                        daily_total = scaled_revenue.sum() if hasattr(scaled_revenue, 'sum') else sum(scaled_revenue)
                        total_monthly_revenue += daily_total
                    else:
                        return float('inf')
                
                return -total_monthly_revenue  # 返回负值用于最小化
            except Exception as e:
                return float('inf')
        
        # 选择搜索方法
        if method == 'grid':
            print("使用网格搜索方法...")
            # 网格搜索
            grid_points = np.linspace(search_range[0], search_range[1], 101)
            best_scale = None
            best_revenue = -float('inf')
            
            for i, scale in enumerate(grid_points):
                revenue = -monthly_objective_function(scale)
                if revenue > best_revenue:
                    best_revenue = revenue
                    best_scale = scale
                
                # 显示进度
                if (i + 1) % 10 == 0:
                    print(f"  进度: {i+1}/101 ({(i+1)/101*100:.1f}%) - 当前最佳: {best_scale:.3f} (收益: {best_revenue:,.0f})")
            
            optimal_result = {'x': best_scale, 'fun': -best_revenue, 'success': True}
            
        else:
            print("使用黄金分割搜索方法...")
            # 使用scipy的minimize_scalar进行优化
            optimal_result = minimize_scalar(
                monthly_objective_function, 
                bounds=search_range, 
                method='bounded',
                options={'xatol': 0.001}
            )
        
        if optimal_result['success']:
            optimal_scale = optimal_result['x']
            optimal_revenue = -optimal_result['fun']
            
            # 计算原始收益用于对比
            original_total_revenue = 0
            for day_data in daily_data:
                original_revenue = self.calculate_total_revenue_for_contract(day_data['df'], day_data['original_contract_values'])
                if original_revenue is not None:
                    daily_total = original_revenue.sum() if hasattr(original_revenue, 'sum') else sum(original_revenue)
                    original_total_revenue += daily_total
            
            print(f"\n🎯 找到月度最佳调整比例！")
            print(f"最佳比例: {optimal_scale:.4f}")
            print(f"原始月总收益: {original_total_revenue:,.2f}")
            print(f"最佳月总收益: {optimal_revenue:,.2f}")
            print(f"收益提升: {optimal_revenue - original_total_revenue:+,.2f}")
            print(f"提升百分比: {(optimal_revenue - original_total_revenue) / original_total_revenue * 100:+.2f}%")
            
            # 计算每日平均收益
            print(f"原始日均收益: {original_total_revenue / len(daily_data):,.2f}")
            print(f"最佳日均收益: {optimal_revenue / len(daily_data):,.2f}")
            
            # 返回结果
            return {
                'year': year,
                'month': month,
                'days_count': len(daily_data),
                'failed_dates': failed_dates,
                'optimal_scale_factor': optimal_scale,
                'original_total_revenue': original_total_revenue,
                'optimal_total_revenue': optimal_revenue,
                'revenue_improvement': optimal_revenue - original_total_revenue,
                'improvement_percentage': (optimal_revenue - original_total_revenue) / original_total_revenue * 100 if original_total_revenue != 0 else 0,
                'search_range': search_range,
                'search_method': method,
                'daily_data': daily_data
            }
        else:
            print("❌ 月度优化搜索失败")
            return None

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
        print("3. 分析单日负收益情况")
        print("4. 分析月度负收益情况")
        print("5. 分析单日原始合约量调整")
        print("6. 分析月度原始合约量调整")
        print("7. 自动寻找单日最佳调整比例")
        print("8. 自动寻找月度最佳调整比例")
        print("9. 批量分析所有日期")
        print("0. 退出")
        print("-"*30)
        
        choice = input("请输入选项 (0-9): ").strip()
        
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
                print(f"总合约量: {result['total_contract_amount']:.3f}")
                print(f"原合约取值范围: {result['contract_range']}")
                
                # 询问是否显示详细收益分解
                detail_choice = input("\n是否显示每15分钟详细收益分解? (y/n): ").strip().lower()
                if detail_choice in ['y', 'yes', '是']:
                    revenue_breakdown = optimizer.print_detailed_revenue_breakdown(
                        result['data'], result['optimal_contract'], 
                        result['optimal_revenue'], date
                    )
                
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
                        
                            # 询问是否显示每日详细收益分解
                            detail_choice = input("\n是否显示每日详细收益分解? (y/n): ").strip().lower()
                            if detail_choice in ['y', 'yes', '是']:
                                print("注意：这将显示该月所有天的详细收益，输出较多...")
                                confirm = input("确认继续? (y/n): ").strip().lower()
                                if confirm in ['y', 'yes', '是']:
                                    optimizer.print_monthly_revenue_breakdown(year, month)
                            
                            plot_choice = input("\n是否绘制月度分析图表? (y/n): ").strip().lower()
                            if plot_choice in ['y', 'yes', '是']:
                                save_choice = input("是否保存图表? (y/n): ").strip().lower()
                                save_path = None
                                if save_choice in ['y', 'yes', '是']:
                                    save_path = f"月度原合约优化分析_{year}年{month}月.png"
                                
                                optimizer.plot_monthly_optimization(result, save_path)
                            
                            # 新增：原合约比例调整分析选项
                            scaled_choice = input("\n是否进行原合约比例调整分析? (保持分布不变，调整各时间点原合约值并重新计算收入) (y/n): ").strip().lower()
                            if scaled_choice in ['y', 'yes', '是']:
                                print("\n请选择调整比例:")
                                print("1. 0.25倍 (四分之一)")
                                print("2. 0.5倍 (减半)")
                                print("3. 1.5倍 (增加50%)")
                                print("4. 2.0倍 (双倍)")
                                print("5. 自定义倍数")
                                
                                scale_choice = input("请选择 (1-5): ").strip()
                                scale_factor = None
                                
                                if scale_choice == '1':
                                    scale_factor = 0.25
                                elif scale_choice == '2':
                                    scale_factor = 0.5
                                elif scale_choice == '3':
                                    scale_factor = 1.5
                                elif scale_choice == '4':
                                    scale_factor = 2.0
                                elif scale_choice == '5':
                                    try:
                                        custom_scale = input("请输入自定义倍数 (例如: 0.3, 1.2, 2.5等): ").strip()
                                        scale_factor = float(custom_scale)
                                        if scale_factor <= 0:
                                            print("倍数必须大于0")
                                            continue
                                    except ValueError:
                                        print("输入的倍数格式错误")
                                        continue
                                else:
                                    print("无效选择")
                                    continue
                                
                                if scale_factor is not None:
                                    print(f"\n正在进行原合约{scale_factor}倍调整分析...")
                                    scaled_result = optimizer.analyze_scaled_contract(result, scale_factor)
                                    
                                    if scaled_result:
                                        # 显示调整分析的详细对比
                                        optimizer.print_scaled_contract_comparison(scaled_result)
                                        
                                        # 询问是否绘制对比图表
                                        scaled_plot_choice = input("\n是否绘制调整前后对比图表? (y/n): ").strip().lower()
                                        if scaled_plot_choice in ['y', 'yes', '是']:
                                            scaled_save_choice = input("是否保存对比图表? (y/n): ").strip().lower()
                                            scaled_save_path = None
                                            if scaled_save_choice in ['y', 'yes', '是']:
                                                scale_desc = scaled_result['scale_description']
                                                scaled_save_path = f"原合约{scale_desc}对比分析_{year}年{month}月.png"
                                            
                                            optimizer.plot_scaled_contract_comparison(scaled_result, scaled_save_path)
                                    else:
                                        print("调整分析失败，请检查数据")
                            else:
                                print("跳过原合约比例调整分析")
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
            date = input("请输入日期 (格式: 2025-05-10): ").strip()
            print(f"\n正在分析 {date} 的负收益情况...")
            
            result = optimizer.analyze_daily_negative_revenue(date)
            if result:
                # 汇总所有收益类型的负收益统计
                total_negative = sum(stats['total_negative'] for stats in result['negative_summary'].values())
                negative_counts = [stats['negative_count'] for stats in result['negative_summary'].values()]
                total_negative_count = sum(negative_counts) if negative_counts else 0
                worst_values = [stats['worst_value'] for stats in result['negative_summary'].values() if stats['worst_value'] < 0]
                worst_value = min(worst_values) if worst_values else 0
                worst_times = [stats['worst_time'] for stats in result['negative_summary'].values() if stats['worst_value'] < 0]
                worst_time = worst_times[0] if worst_times else 'N/A'
                
                print(f"\n=== {date} 负收益分析完成 ===")
                print(f"总负收益: {total_negative:,.2f}")
                print(f"最大单点亏损: {worst_value:.2f} (时间: {worst_time})")
                print(f"负收益点数: {total_negative_count} 个")
                
                # 询问是否绘制图表
                plot_choice = input("\n是否绘制负收益分析图表? (y/n): ").strip().lower()
                if plot_choice in ['y', 'yes', '是']:
                    save_choice = input("是否保存图表? (y/n): ").strip().lower()
                    save_path = None
                    if save_choice in ['y', 'yes', '是']:
                        save_path = f"负收益分析_{date}.png"
                    
                    optimizer.plot_negative_revenue_analysis(result, 'daily', save_path)
            else:
                print("分析失败，请检查数据文件")
        
        elif choice == '4':
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
                        
                    print(f"\n正在分析 {year}年{month}月 的负收益情况...")
                    
                    result = optimizer.analyze_monthly_negative_revenue(year, month)
                    if result:
                        summary = result['monthly_summary']
                        
                        print(f"\n=== {year}年{month}月 负收益分析完成 ===")
                        print(f"月度总负收益: {summary['total']:,.2f}")
                        print(f"  撮合收益亏损: {summary['matching']:,.2f}")
                        print(f"  日前结算亏损: {summary['forward']:,.2f}")
                        print(f"  实时结算亏损: {summary['realtime']:,.2f}")
                        print(f"处理天数: {result['total_days']}天")
                        print(f"成功分析天数: {len(result['daily_analyses'])}天")
                        
                        # 询问是否绘制图表
                        plot_choice = input("\n是否绘制月度负收益分析图表? (y/n): ").strip().lower()
                        if plot_choice in ['y', 'yes', '是']:
                            save_choice = input("是否保存图表? (y/n): ").strip().lower()
                            save_path = None
                            if save_choice in ['y', 'yes', '是']:
                                save_path = f"月度负收益分析_{year}年{month}月.png"
                            
                            optimizer.plot_negative_revenue_analysis(result, 'monthly', save_path)
                    else:
                        print("分析失败，请检查该月份的数据文件是否存在")
                else:
                    print("月份格式错误，请使用 YYYY-MM 格式 (例如: 2025-05)")
            except ValueError:
                print("输入格式错误，请使用 YYYY-MM 格式 (例如: 2025-05)")
            except Exception as e:
                print(f"处理月份输入时出错: {e}")
                print("请检查输入格式并重试")
        
        elif choice == '5':
            date = input("请输入日期 (格式: 2025-05-10): ").strip()
            print(f"\n正在分析 {date} 的原始合约量调整...")
            
            # 询问是否使用自定义比例
            use_custom = input("是否使用自定义调整比例? (y/n, 默认使用0.25、0.5、1.5): ").strip().lower()
            
            scale_factors = [0.25, 0.5, 1.5]  # 默认比例
            
            if use_custom in ['y', 'yes', '是']:
                custom_input = input("请输入调整比例，用逗号分隔 (例如: 0.25,0.5,1.0,1.5,2.0): ").strip()
                try:
                    scale_factors = [float(x.strip()) for x in custom_input.split(',') if x.strip()]
                    if not scale_factors:
                        print("无效输入，使用默认比例")
                        scale_factors = [0.25, 0.5, 1.5]
                    else:
                        print(f"使用自定义比例: {scale_factors}")
                except ValueError:
                    print("输入格式错误，使用默认比例")
                    scale_factors = [0.25, 0.5, 1.5]
            
            result = optimizer.analyze_original_contract_scaling(date, scale_factors)
            if result:
                print(f"\n=== {date} 原始合约量调整分析完成 ===")
                print(f"原始合约量总和: {result['original_contract_values'].sum():.3f}")
                print(f"原始总收益: {result['original_total_revenue']:,.2f}")
                
                # 显示各比例的结果
                print(f"\n各比例调整结果:")
                for factor in scale_factors:
                    if factor in result['scaling_results']:
                        res = result['scaling_results'][factor]
                        print(f"  {factor}倍 ({res['scale_description']}): 收益 {res['total_revenue']:,.2f}")
                
                # 找出最佳比例
                best_factor = max(scale_factors, 
                                key=lambda f: result['scaling_results'][f]['total_revenue'] 
                                if f in result['scaling_results'] else 0)
                best_result = result['scaling_results'][best_factor]
                print(f"\n🏆 最佳调整比例: {best_factor}倍 ({best_result['scale_description']})")
                print(f"   最佳收益: {best_result['total_revenue']:,.2f}")
                print(f"   相比原始增益: {best_result['total_revenue'] - result['original_total_revenue']:+,.2f}")
            else:
                print("分析失败，请检查数据文件和原始合约量列")
        
        elif choice == '6':
            month_input = input("请输入月份 (格式: 2025-05): ").strip()
            try:
                if '-' in month_input and len(month_input.split('-')) == 2:
                    year_str, month_str = month_input.split('-')
                    year = int(year_str)
                    month = int(month_str)
                    
                    if year < 2020 or year > 2030:
                        print("年份应在2020-2030之间")
                        continue
                    if month < 1 or month > 12:
                        print("月份应在1-12之间")
                        continue
                    
                    print(f"\n正在分析 {year}年{month}月 的原始合约量调整...")
                    
                    # 询问是否使用自定义比例
                    use_custom = input("是否使用自定义调整比例? (y/n, 默认使用0.25、0.5、1.5): ").strip().lower()
                    
                    scale_factors = [0.25, 0.5, 1.5]  # 默认比例
                    
                    if use_custom in ['y', 'yes', '是']:
                        custom_input = input("请输入调整比例，用逗号分隔 (例如: 0.25,0.5,1.0,1.5,2.0): ").strip()
                        try:
                            scale_factors = [float(x.strip()) for x in custom_input.split(',') if x.strip()]
                            if not scale_factors:
                                print("无效输入，使用默认比例")
                                scale_factors = [0.25, 0.5, 1.5]
                            else:
                                print(f"使用自定义比例: {scale_factors}")
                        except ValueError:
                            print("输入格式错误，使用默认比例")
                            scale_factors = [0.25, 0.5, 1.5]
                    
                    result = optimizer.analyze_monthly_original_contract_scaling(year, month, scale_factors)
                    if result:
                        # 显示月度分析的详细对比
                        optimizer.print_monthly_original_scaling_comparison(result)
                    else:
                        print("分析失败，请检查该月份的数据文件是否存在")
                else:
                    print("月份格式错误，请使用 YYYY-MM 格式 (例如: 2025-05)")
            except ValueError:
                print("输入格式错误，请使用 YYYY-MM 格式 (例如: 2025-05)")
            except Exception as e:
                print(f"处理月份输入时出错: {e}")
                print("请检查输入格式并重试")
        
        elif choice == '7':
            date = input("请输入日期 (格式: 2025-05-10): ").strip()
            print(f"\n正在为 {date} 自动寻找最佳调整比例...")
            
            # 询问搜索范围
            range_input = input("请输入搜索范围 (格式: 0.1,3.0，默认0.1-3.0): ").strip()
            search_range = (0.1, 3.0)  # 默认范围
            
            if range_input:
                try:
                    range_parts = range_input.split(',')
                    if len(range_parts) == 2:
                        search_range = (float(range_parts[0].strip()), float(range_parts[1].strip()))
                        print(f"使用搜索范围: {search_range[0]:.1f} - {search_range[1]:.1f}")
                    else:
                        print("范围格式错误，使用默认范围")
                except ValueError:
                    print("范围格式错误，使用默认范围")
            
            # 询问搜索方法
            method_input = input("选择搜索方法 (grid/golden，默认grid): ").strip().lower()
            method = 'grid' if method_input not in ['golden'] else 'golden'
            
            if method == 'golden':
                print("将使用黄金分割搜索（更精确但稍慢）")
            else:
                print("将使用网格搜索（更快但精度稍低）")
            
            result = optimizer.find_optimal_scale_factor(date, search_range, method)
            if result:
                print(f"\n=== {date} 最佳调整比例分析完成 ===")
                print(f"🎯 推荐使用 {result['optimal_scale_factor']:.4f} 倍调整")
                print(f"💰 预期收益提升: {result['revenue_improvement']:+,.2f} ({result['improvement_percentage']:+.2f}%)")
                
                # 询问是否查看详细对比
                detail_choice = input("\n是否查看与其他常用比例的对比? (y/n): ").strip().lower()
                if detail_choice in ['y', 'yes', '是']:
                    # 与常用比例对比
                    common_scales = [0.25, 0.5, 1.0, 1.5, 2.0]
                    if result['optimal_scale_factor'] not in common_scales:
                        common_scales.append(result['optimal_scale_factor'])
                    common_scales.sort()
                    
                    print(f"\n📊 与常用比例对比:")
                    print(f"{'比例':<8} {'收益':<15} {'相对最佳':<12} {'相对原始':<12}")
                    print("-" * 50)
                    
                    for scale in common_scales:
                        if scale == result['optimal_scale_factor']:
                            # 最佳比例
                            revenue = result['optimal_total_revenue']
                            vs_best = "0 (最佳)"
                            vs_orig = f"{result['improvement_percentage']:+.1f}%"
                            marker = "🎯 "
                        else:
                            # 计算其他比例的收益
                            test_contract = result['original_contract_values'] * scale
                            test_revenue_series = optimizer.calculate_total_revenue_for_contract(result['data'], test_contract)
                            revenue = test_revenue_series.sum() if test_revenue_series is not None else 0
                            vs_best = f"{revenue - result['optimal_total_revenue']:+,.0f}"
                            vs_orig_pct = ((revenue - result['original_total_revenue']) / result['original_total_revenue'] * 100) if result['original_total_revenue'] != 0 else 0
                            vs_orig = f"{vs_orig_pct:+.1f}%"
                            marker = "   "
                        
                        print(f"{marker}{scale:<5.2f} {revenue:<15,.0f} {vs_best:<12} {vs_orig:<12}")
            else:
                print("自动寻找最佳比例失败，请检查数据文件")
        
        elif choice == '8':
            month_input = input("请输入月份 (格式: 2025-05): ").strip()
            try:
                if '-' in month_input and len(month_input.split('-')) == 2:
                    year_str, month_str = month_input.split('-')
                    year = int(year_str)
                    month = int(month_str)
                    
                    if year < 2020 or year > 2030:
                        print("年份应在2020-2030之间")
                        continue
                    if month < 1 or month > 12:
                        print("月份应在1-12之间")
                        continue
                    
                    print(f"\n正在为 {year}年{month}月 自动寻找最佳调整比例...")
                    
                    # 询问搜索范围
                    range_input = input("请输入搜索范围 (格式: 0.1,3.0，默认0.1-3.0): ").strip()
                    search_range = (0.1, 3.0)  # 默认范围
                    
                    if range_input:
                        try:
                            range_parts = range_input.split(',')
                            if len(range_parts) == 2:
                                search_range = (float(range_parts[0].strip()), float(range_parts[1].strip()))
                                print(f"使用搜索范围: {search_range[0]:.1f} - {search_range[1]:.1f}")
                            else:
                                print("范围格式错误，使用默认范围")
                        except ValueError:
                            print("范围格式错误，使用默认范围")
                    
                    # 询问搜索方法
                    method_input = input("选择搜索方法 (grid/golden，默认grid): ").strip().lower()
                    method = 'grid' if method_input not in ['golden'] else 'golden'
                    
                    if method == 'golden':
                        print("将使用黄金分割搜索（更精确但稍慢）")
                    else:
                        print("将使用网格搜索（更快但精度稍低）")
                    
                    result = optimizer.find_optimal_scale_factor_monthly(year, month, search_range, method)
                    if result:
                        print(f"\n=== {year}年{month}月 最佳调整比例分析完成 ===")
                        print(f"🎯 推荐使用 {result['optimal_scale_factor']:.4f} 倍调整")
                        print(f"💰 预期月收益提升: {result['revenue_improvement']:+,.2f} ({result['improvement_percentage']:+.2f}%)")
                        print(f"📅 预期日均收益提升: {result['revenue_improvement']/result['days_count']:+,.2f}")
                        
                        # 询问是否查看详细对比
                        detail_choice = input("\n是否查看与其他常用比例的月度对比? (y/n): ").strip().lower()
                        if detail_choice in ['y', 'yes', '是']:
                            # 与常用比例对比（月度版本）
                            common_scales = [0.25, 0.5, 1.0, 1.5, 2.0]
                            if result['optimal_scale_factor'] not in common_scales:
                                common_scales.append(result['optimal_scale_factor'])
                            common_scales.sort()
                            
                            print(f"\n📊 月度常用比例对比:")
                            print(f"{'比例':<8} {'月总收益':<15} {'相对最佳':<12} {'相对原始':<12}")
                            print("-" * 50)
                            
                            for scale in common_scales:
                                if scale == result['optimal_scale_factor']:
                                    # 最佳比例
                                    revenue = result['optimal_total_revenue']
                                    vs_best = "0 (最佳)"
                                    vs_orig = f"{result['improvement_percentage']:+.1f}%"
                                    marker = "🎯 "
                                else:
                                    # 计算其他比例的月度收益
                                    monthly_revenue = 0
                                    for day_data in result['daily_data']:
                                        test_contract = day_data['original_contract_values'] * scale
                                        test_revenue_series = optimizer.calculate_total_revenue_for_contract(day_data['df'], test_contract)
                                        if test_revenue_series is not None:
                                            daily_total = test_revenue_series.sum()
                                            monthly_revenue += daily_total
                                    
                                    vs_best = f"{monthly_revenue - result['optimal_total_revenue']:+,.0f}"
                                    vs_orig_pct = ((monthly_revenue - result['original_total_revenue']) / result['original_total_revenue'] * 100) if result['original_total_revenue'] != 0 else 0
                                    vs_orig = f"{vs_orig_pct:+.1f}%"
                                    marker = "   "
                                    revenue = monthly_revenue
                                
                                print(f"{marker}{scale:<5.2f} {revenue:<15,.0f} {vs_best:<12} {vs_orig:<12}")
                    else:
                        print("自动寻找最佳比例失败，请检查该月份的数据文件是否存在")
                else:
                    print("月份格式错误，请使用 YYYY-MM 格式 (例如: 2025-05)")
            except ValueError:
                print("输入格式错误，请使用 YYYY-MM 格式 (例如: 2025-05)")
            except Exception as e:
                print(f"处理月份输入时出错: {e}")
                print("请检查输入格式并重试")
        
        elif choice == '9':
            print("联系18713812142")
        
        else:
            print("无效选项，请重新选择")
        
        input("\n按回车键继续...")

if __name__ == "__main__":
    main() 