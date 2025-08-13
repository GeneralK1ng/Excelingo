#!/usr/bin/env python3
"""
XLSX翻译工具主程序
"""
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from src.gui.main_window import MainWindow

def main():
    app = MainWindow()
    app.run()

if __name__ == "__main__":
    main()