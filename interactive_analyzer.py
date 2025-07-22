#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
电力收益分析器 - 交互式界面
"""

import os
import sys
import numpy as np
import matplotlib.pyplot as plt
from power_analyzer import PowerSettlementAnalyzer

def main_menu():
    """主菜单"""
    while True:
        print("\n" + "="*40)
        print("           日前结算收益")
        print("="*40)
        print("\n请选择查询方式：")
        print("1. 显示总体汇总")
        print("2. 查询特定日期")
        print("3. 查询特定月份")
        print("4. 查询特定年份")
        print("5. 显示详细汇总")
        print("0. 退出")
        print("-"*40)
        
        choice = input("请输入选项 (0-8): ").strip()
        
        if choice == '0':
            print("谢谢使用：）")
            break
        elif choice in ['1', '2', '3', '4', '5', '6', '7', '8']:
            process_choice(choice)
        else:
            print("无效选项，请重新选择")

def plot_contract_curve(optimal_contract, date_str, save_path=None):
    """绘制原合约曲线"""
    try:
        plt.figure(figsize=(12, 6))
        time_points = range(1, len(optimal_contract) + 1)
        plt.plot(time_points, optimal_contract, 'b-', linewidth=2, marker='o', markersize=4)
        plt.title(f'{date_str} 最优原合约曲线 (96个时间点)', fontsize=14)
        plt.xlabel('时间点 (15分钟间隔)', fontsize=12)
        plt.ylabel('原合约值', fontsize=12)
        plt.grid(True, alpha=0.3)
        plt.xticks(range(0, 97, 8))  # 每2小时一个刻度
        
        # 添加统计信息
        avg_value = np.mean(optimal_contract)
        max_value = np.max(optimal_contract)
        min_value = np.min(optimal_contract)
        plt.text(0.02, 0.98, f'平均值: {avg_value:.3f}\n最大值: {max_value:.3f}\n最小值: {min_value:.3f}', 
                transform=plt.gca().transAxes, verticalalignment='top',
                bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.8))
        
        plt.tight_layout()
        
        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
            print(f"图表已保存到: {save_path}")
        
        plt.show()
        
    except Exception as e:
        print(f"绘制图表时出错: {e}")

def plot_monthly_average_curve(monthly_avg, year, month, save_path=None):
    """绘制月度平均原合约曲线"""
    try:
        plt.figure(figsize=(12, 6))
        time_points = range(1, len(monthly_avg) + 1)
        plt.plot(time_points, monthly_avg, 'r-', linewidth=2, marker='s', markersize=4)
        plt.title(f'{year}年{month}月 平均原合约曲线 (96个时间点)', fontsize=14)
        plt.xlabel('时间点 (15分钟间隔)', fontsize=12)
        plt.ylabel('平均原合约值', fontsize=12)
        plt.grid(True, alpha=0.3)
        plt.xticks(range(0, 97, 8))  # 每2小时一个刻度
        
        # 添加统计信息
        avg_value = np.mean(monthly_avg)
        max_value = np.max(monthly_avg)
        min_value = np.min(monthly_avg)
        plt.text(0.02, 0.98, f'月平均值: {avg_value:.3f}\n月最大值: {max_value:.3f}\n月最小值: {min_value:.3f}', 
                transform=plt.gca().transAxes, verticalalignment='top',
                bbox=dict(boxstyle='round', facecolor='lightcoral', alpha=0.8))
        
        plt.tight_layout()
        
        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
            print(f"月度图表已保存到: {save_path}")
        
        plt.show()
        
    except Exception as e:
        print(f"绘制月度图表时出错: {e}")

def process_choice(choice):
    """处理用户选择"""
    print("\n正在加载数据...")
    analyzer = PowerSettlementAnalyzer()
    analyzer.load_all_data()
    
    if not analyzer.data:
        print("没有找到任何有效数据")
        return
    
    print("\n" + "-"*50)
    
    if choice == '1':
        # 显示总体汇总
        total_sum = sum(analyzer.data.values())
        total_days = len(analyzer.data)
        print(f"总计: {total_sum:.2f} (共{total_days}天)")
        
    elif choice == '2':
        # 查询特定日期
        date = input("请输入日期 (格式: 2025-05-10): ").strip()
        result = analyzer.query_by_date(date)
        if result is not None:
            print(f"{date} 的日前结算收益负值总和: {result:.2f}")
            
    elif choice == '3':
        # 查询特定月份
        month_input = input("请输入月份 (格式: 2025-05): ").strip()
        try:
            year, month = map(int, month_input.split('-'))
            result, count = analyzer.query_by_month(year, month)
            print(f"{year}年{month}月的日前结算收益负值总和: {result:.2f}")
            print(f"包含 {count} 天的数据")
        except ValueError:
            print("月份格式错误，请使用 YYYY-MM 格式")
            
    elif choice == '4':
        # 查询特定年份
        try:
            year = int(input("请输入年份 (格式: 2025): ").strip())
            result, count = analyzer.query_by_year(year)
            print(f"{year}年的日前结算收益负值总和: {result:.2f}")
            print(f"包含 {count} 天的数据")
        except ValueError:
            print("年份格式错误，请输入数字")
            
    elif choice == '5':
        # 显示详细汇总
        print("=== 数据汇总 ===")
        
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
            
    elif choice == '6':
        # 列出所有日期
        dates = analyzer.get_all_dates()
        print(f"可用日期列表 (共 {len(dates)} 天):")
        for date in dates:
            value = analyzer.data[date]
            print(f"{date.strftime('%Y-%m-%d')}: {value:.2f}")
    
    elif choice == '7':
        # 原合约优化分析
        from contract_optimizer import ContractOptimizer
        
        date = input("请输入要分析的日期 (格式: 2025-05-10): ").strip()
        print(f"\n正在分析 {date} 的原合约优化...")
        
        # 使用新的原合约优化分析器
        optimizer = ContractOptimizer()
        result = optimizer.analyze_daily_optimization(date)
        
        if result:
            print(f"\n=== {date} 原合约优化分析结果 ===")
            print(f"总合约量: {result['total_contract_amount']:.3f}")
            print(f"平均最优原合约值: {result['avg_optimal_contract']:.3f}")
            print(f"总最优收益: {result['total_optimal_revenue']:.2f}")
            if result['daily_total_limit']:
                print(f"每日总量限制: {result['daily_total_limit']:.3f}")
            print(f"96个时间点的最优原合约值已计算完成")
            
            # 询问是否绘制图表
            plot_choice = input("\n是否绘制原合约分析图表? (y/n): ").strip().lower()
            if plot_choice in ['y', 'yes', '是']:
                save_choice = input("是否保存图表? (y/n): ").strip().lower()
                save_path = None
                if save_choice in ['y', 'yes', '是']:
                    save_path = f"原合约优化分析_{date}.png"
                
                optimizer.plot_daily_optimization(result, save_path)
        else:
            print("分析失败，请检查数据文件")
    
    elif choice == '8':
        # 生成月度原合约曲线
        from contract_optimizer import ContractOptimizer
        
        month_input = input("请输入月份 (格式: 2025-05): ").strip()
        try:
            year, month = map(int, month_input.split('-'))
            print(f"\n正在分析 {year}年{month}月 的原合约优化...")
            
            # 使用新的原合约优化分析器
            optimizer = ContractOptimizer()
            result = optimizer.analyze_monthly_optimization(year, month)
            
            if result:
                print(f"\n=== {year}年{month}月 原合约优化分析结果 ===")
                print(f"处理天数: {result['days_count']}天")
                print(f"月平均原合约值: {np.mean(result['monthly_average']):.3f}")
                print(f"平均总合约量: {np.sum(result['monthly_average']):.3f}")
                if result['daily_total_limit']:
                    print(f"每日总量限制: {result['daily_total_limit']:.3f}")
                
                # 询问是否绘制图表
                plot_choice = input("\n是否绘制月度分析图表? (y/n): ").strip().lower()
                if plot_choice in ['y', 'yes', '是']:
                    save_choice = input("是否保存图表? (y/n): ").strip().lower()
                    save_path = None
                    if save_choice in ['y', 'yes', '是']:
                        save_path = f"月度原合约优化分析_{year}年{month}月.png"
                    
                    optimizer.plot_monthly_optimization(result, save_path)
            else:
                print("分析失败，请检查该月份的数据")
        except ValueError:
            print("月份格式错误，请使用 YYYY-MM 格式")
    
    input("\n按回车键返回主菜单...")

if __name__ == "__main__":
    try:
        main_menu()
    except KeyboardInterrupt:
        print("\n\n用户中断，程序退出")
    except Exception as e:
        print(f"\n程序运行出错: {e}")
        input("按回车键退出...") 