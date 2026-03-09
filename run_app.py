#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
启动器：用于 PyInstaller 打包后启动 Streamlit 应用
"""

import sys
import os

# 在打包环境下，需要在 streamlit 导入前 patch importlib.metadata
if getattr(sys, 'frozen', False):
    import importlib.metadata as _meta

    _original_version = _meta.version

    def _patched_version(name):
        try:
            return _original_version(name)
        except _meta.PackageNotFoundError:
            # 为已知包提供版本号
            _fallback = {
                'streamlit': '1.45.0',
                'altair': '5.0.0',
                'pandas': '2.0.0',
                'numpy': '1.24.0',
                'pillow': '10.0.0',
                'protobuf': '4.0.0',
                'pyarrow': '12.0.0',
                'rich': '13.0.0',
                'toml': '0.10.2',
                'packaging': '23.0',
                'click': '8.0.0',
                'tornado': '6.0',
                'watchdog': '3.0.0',
            }
            if name.lower() in _fallback:
                return _fallback[name.lower()]
            raise

    _meta.version = _patched_version


def get_app_path():
    """获取 app.py 的路径（兼容打包后和开发环境）"""
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, 'app.py')
    else:
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'app.py')


if __name__ == '__main__':
    import webbrowser
    import threading
    from streamlit.web import cli as stcli

    app_path = get_app_path()

    # 自动打开浏览器
    def open_browser():
        import time
        time.sleep(2)
        webbrowser.open('http://localhost:8501')

    threading.Thread(target=open_browser, daemon=True).start()

    sys.argv = [
        'streamlit', 'run', app_path,
        '--server.headless', 'true',
        '--browser.gatherUsageStats', 'false',
        '--global.developmentMode', 'false',
        '--server.port', '8501',
    ]

    stcli.main()
