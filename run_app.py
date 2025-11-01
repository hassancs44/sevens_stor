# -*- coding: utf-8 -*-
import os
import sys
import time
import webbrowser
from pathlib import Path

def resource_path(relpath: str) -> str:
    """يدعم التشغيل من PyInstaller (onefile/onedir)"""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
    return str(base / relpath)

def main():
    app_path = resource_path("app_xlsx.py")

    # تهيئة ستريمليت
    os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"        # سيرفر بدون واجهة
    os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"  # أوقف dev-mode
    os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
    os.environ["STREAMLIT_THEME_BASE"] = "light"

    # افتح المتصفح تلقائيًا على المنفذ الافتراضي 8501
    url = "http://localhost:8501"
    try:
        time.sleep(0.6)
        webbrowser.open(url)
    except Exception:
        pass

    # شغّل ستريمليت عبر CLI الداخلي بدون تحديد المنفذ
    from streamlit.web.cli import main as stcli
    sys.argv = ["streamlit", "run", app_path, "--server.headless", "true"]
    sys.exit(stcli())

if __name__ == "__main__":
    main()
