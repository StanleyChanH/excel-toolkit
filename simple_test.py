#!/usr/bin/env python3
"""
简化的Excel工具箱测试脚本
使用现有示例文件进行基本功能测试
"""

import requests
import os

def test_basic_functionality():
    """测试基本功能"""
    base_url = "http://127.0.0.1:5000"

    print("🔧 Excel工具箱基本功能测试")
    print("=" * 50)

    # 检查服务器是否运行
    try:
        response = requests.get(base_url, timeout=5)
        if response.status_code == 200:
            print("✅ 服务器正常运行")
        else:
            print(f"❌ 服务器响应异常: {response.status_code}")
            return False
    except requests.exceptions.RequestException as e:
        print(f"❌ 无法连接到服务器: {e}")
        return False

    # 检查示例文件是否存在
    required_files = [
        'examples/sample_data1.xlsx',
        'examples/sample_data2.xlsx',
        'examples/multi_sheet_data.xlsx',
        'examples/sample_data1.csv'
    ]

    missing_files = []
    for file_path in required_files:
        if not os.path.exists(file_path):
            missing_files.append(file_path)

    if missing_files:
        print(f"❌ 缺少示例文件: {', '.join(missing_files)}")
        print("请先运行: uv run python create_samples.py")
        return False

    print("✅ 所有示例文件存在")

    # 测试合并文件功能
    try:
        with open('examples/sample_data1.xlsx', 'rb') as f1, open('examples/sample_data2.xlsx', 'rb') as f2:
            files = [
                ('files', ('sample_data1.xlsx', f1, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')),
                ('files', ('sample_data2.xlsx', f2, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'))
            ]

            data = {
                'keep_headers': 'true',
                'add_source_column': 'true'
            }

            response = requests.post(f"{base_url}/api/merge-files", files=files, data=data, timeout=30)

            if response.status_code == 200:
                print("✅ 合并文件功能正常")
                print(f"   返回文件大小: {len(response.content)} bytes")
            else:
                print(f"❌ 合并文件功能异常: {response.status_code}")
                if response.headers.get('content-type', '').startswith('application/json'):
                    print(f"   错误信息: {response.json().get('error', '未知错误')}")
                return False
    except Exception as e:
        print(f"❌ 合并文件测试失败: {e}")
        return False

    # 测试错误处理
    try:
        response = requests.post(f"{base_url}/api/merge-files", data={}, timeout=10)
        if response.status_code == 400:
            error_data = response.json()
            if 'error' in error_data:
                print("✅ 错误处理机制正常")
                print(f"   错误信息: {error_data['error']}")
            else:
                print("⚠️ 错误处理不完整")
        else:
            print(f"❌ 错误处理异常: {response.status_code}")
            return False
    except Exception as e:
        print(f"❌ 错误处理测试失败: {e}")
        return False

    print("\n🎉 基本功能测试完成！")
    print("✅ Web界面可访问")
    print("✅ 示例文件存在")
    print("✅ 合并文件功能正常")
    print("✅ 错误处理机制正常")
    print("\n📋 测试总结:")
    print("Excel工具箱的核心功能运行正常，可以处理文件上传、")
    print("数据处理和错误响应。建议在浏览器中访问")
    print("http://127.0.0.1:5000 进行完整的手动测试。")

    return True

if __name__ == "__main__":
    success = test_basic_functionality()
    exit(0 if success else 1)