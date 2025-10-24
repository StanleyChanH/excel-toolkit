#!/usr/bin/env python3
"""
Excel工具箱启动脚本
简单的启动入口，提供更好的用户体验
"""

import os
import sys
import webbrowser
import time
from threading import Timer

def open_browser():
    """延迟2秒后自动打开浏览器"""
    time.sleep(2)
    webbrowser.open('http://127.0.0.1:5000')

def main():
    """主函数"""
    print("=" * 60)
    print("🔧 Excel批量操作工具箱")
    print("=" * 60)
    print("正在启动服务器...")
    print("服务器地址: http://127.0.0.1:5000")
    print("按 Ctrl+C 停止服务器")
    print("=" * 60)
    print()

    # 询问是否自动打开浏览器
    try:
        auto_open = input("是否自动打开浏览器？(Y/n): ").strip().lower()
        if auto_open in ['', 'y', 'yes', '是']:
            print("正在启动浏览器...")
            Timer(2, open_browser).start()
    except KeyboardInterrupt:
        print("\n启动已取消")
        return

    print("\n🚀 启动Flask应用...")

    # 导入并运行Flask应用
    from app import app

    try:
        app.run(debug=True, host='0.0.0.0', port=5000)
    except KeyboardInterrupt:
        print("\n👋 感谢使用Excel工具箱！")
    except Exception as e:
        print(f"\n❌ 启动失败: {e}")
        print("请检查端口5000是否被占用，或查看上方错误信息")
        sys.exit(1)

if __name__ == '__main__':
    main()