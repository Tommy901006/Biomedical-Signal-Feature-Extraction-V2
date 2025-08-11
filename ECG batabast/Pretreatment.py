import os
import sys
import subprocess

# 讓工作目錄設為這個檔案所在位置（確保能找到 .py / .xlsx）
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# 依序要執行的腳本（需要就改順序或檔名）
scripts = ["SR.py", "SB.py", "GSVT.py"]

for s in scripts:
    print(f"\n=== Running {s} ===")
    # 用同一個 Python 來跑，失敗不會中斷後面腳本；想要中斷就把 check=True 打開
    try:
        subprocess.run([sys.executable, s], check=False)
        print(f"--- Finished {s} ---")
    except Exception as e:
        print(f"!!! {s} failed: {e}")

print("\nAll done.")
