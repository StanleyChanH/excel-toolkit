#!/usr/bin/env python3
"""
创建示例数据文件用于测试
"""

import pandas as pd
import os

def create_sample_files():
    """创建示例Excel和CSV文件"""

    # 创建示例数据
    data1 = pd.DataFrame({
        '姓名': ['张三', '李四', '王五', '赵六'],
        '部门': ['销售部', '技术部', '销售部', '财务部'],
        '工资': [8000, 12000, 9000, 7000],
        '入职日期': ['2023-01-15', '2022-06-20', '2023-03-10', '2021-12-01']
    })

    data2 = pd.DataFrame({
        '姓名': ['钱七', '孙八', '周九', '吴十'],
        '部门': ['技术部', '财务部', '销售部', '技术部'],
        '工资': [11000, 7500, 8500, 13000],
        '入职日期': ['2023-02-20', '2022-11-15', '2023-04-05', '2021-08-12']
    })

    # 确保示例文件夹存在
    os.makedirs('examples', exist_ok=True)

    # 保存示例文件
    data1.to_excel('examples/sample_data1.xlsx', index=False)
    data2.to_excel('examples/sample_data2.xlsx', index=False)

    # 创建多Sheet示例文件
    with pd.ExcelWriter('examples/multi_sheet_data.xlsx', engine='openpyxl') as writer:
        data1.to_excel(writer, sheet_name='一月', index=False)
        data2.to_excel(writer, sheet_name='二月', index=False)

    # 创建CSV示例文件
    data1.to_csv('examples/sample_data1.csv', index=False, encoding='utf-8-sig')

    print("✅ 示例文件已创建完成！")
    print("📁 文件位置：examples/")
    print("📄 包含文件：")
    print("  - sample_data1.xlsx (单Sheet示例)")
    print("  - sample_data2.xlsx (单Sheet示例)")
    print("  - multi_sheet_data.xlsx (多Sheet示例)")
    print("  - sample_data1.csv (CSV示例)")

if __name__ == '__main__':
    create_sample_files()