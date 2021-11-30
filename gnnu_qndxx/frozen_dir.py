# -*- coding: utf-8 -*-
# @Time    : 2021/9/26 9:01
# @Author  : xuzhifeng
# @File    : frozen_dir.py
import sys
import os


def app_path():
    """Returns the base application path."""
    if hasattr(sys, 'frozen'):
        # Handles PyInstaller
        return os.path.dirname(sys.executable)  # 使用pyinstaller打包后的exe目录
    return os.path.dirname(__file__)  # 没打包前的py目录

# C:\Users\66483\Desktop\全院学生信息.xlsx