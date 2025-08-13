#!/usr/bin/env python3
"""
XLSX翻译工具主程序
"""
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from src.gui.modern_window import ModernWindow

def main():
    app = ModernWindow()
    app.run()

if __name__ == "__main__":
    main()