#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
打包脚本：使用 PyInstaller 将 Streamlit 应用打包为可执行文件

用法：
    python build.py          # 打包当前平台
    python build.py --clean  # 清理后重新打包
"""

import subprocess
import sys
import os
import shutil
import platform

# 修复 Windows CI 环境下中文输出编码问题
if sys.stdout.encoding and sys.stdout.encoding.lower() not in ('utf-8', 'utf8'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')


def get_streamlit_path():
    """获取 streamlit 包的安装路径"""
    import streamlit
    return os.path.dirname(streamlit.__file__)


def build():
    clean = '--clean' in sys.argv

    if clean:
        for d in ['build', 'dist']:
            if os.path.exists(d):
                shutil.rmtree(d)
                print(f"已清理 {d}/")
        for f in os.listdir('.'):
            if f.endswith('.spec'):
                os.remove(f)
                print(f"已清理 {f}")

    streamlit_path = get_streamlit_path()
    app_name = "发票合并工具"

    # PyInstaller 参数
    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--name', app_name,
        '--onedir',
        # 添加 app.py 和 merge.py 作为数据文件
        '--add-data', f'app.py{os.pathsep}.',
        '--add-data', f'merge.py{os.pathsep}.',
        # 添加 streamlit 整个包（包含静态资源）
        '--add-data', f'{streamlit_path}{os.pathsep}streamlit',
        # 隐藏导入
        '--hidden-import', 'streamlit',
        '--hidden-import', 'streamlit.web.cli',
        '--hidden-import', 'streamlit.runtime.scriptrunner',
        '--hidden-import', 'pandas',
        '--hidden-import', 'openpyxl',
        '--hidden-import', 'openpyxl.utils.dataframe',
        '--hidden-import', 'openpyxl.cell',
        '--hidden-import', 'openpyxl.styles',
        '--hidden-import', 'openpyxl.workbook',
        '--hidden-import', 'openpyxl.worksheet',
        '--hidden-import', 'PIL',
        '--hidden-import', 'pyarrow',
        '--hidden-import', 'packaging',
        '--hidden-import', 'packaging.version',
        '--hidden-import', 'packaging.specifiers',
        '--hidden-import', 'packaging.requirements',
        '--hidden-import', 'importlib_metadata',
        # 不弹确认
        '--noconfirm',
        # 入口
        'run_app.py',
    ]

    if platform.system() == 'Darwin':
        cmd.extend(['--osx-bundle-identifier', 'com.invoice-merge-tool'])

    print(f"正在打包 {app_name}...")
    print(f"平台: {platform.system()} {platform.machine()}")
    print(f"Streamlit 路径: {streamlit_path}")
    print()

    result = subprocess.run(cmd)

    if result.returncode == 0:
        dist_path = os.path.join('dist', app_name)
        print()
        print("=" * 50)
        print(f"✅ 打包成功！")
        print(f"输出目录: {os.path.abspath(dist_path)}")
        if platform.system() == 'Windows':
            print(f"运行: dist\\{app_name}\\{app_name}.exe")
        else:
            print(f"运行: dist/{app_name}/{app_name}")
        print("=" * 50)
    else:
        print()
        print("❌ 打包失败，请检查上方错误信息")
        sys.exit(1)


if __name__ == '__main__':
    build()
