import os
import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
from config import INPUT_DIR, TEMPLATE_PATH, DEFAULT_EMPLOYEE_NAME, update_config  # `config.ini` を使用

def run_main():
    csv_file = csv_entry.get()
    template_file = template_entry.get()
    employee_name = name_entry.get().strip()
    
    if not csv_file or not template_file or not employee_name:
        messagebox.showerror("エラー", "必要な情報を入力してください。")
        return
    
    # `config.py` ではなく `config.ini` を更新
    update_config(employee_name, template_file)
    
    try:
        subprocess.run(["python", "main.py", "--name", employee_name, "--file", csv_file, "--template", template_file], check=True)
        messagebox.showinfo("完了", "勤怠表の作成が完了しました！")
        open_output_folder()
    except subprocess.CalledProcessError:
        messagebox.showerror("エラー", "処理中にエラーが発生しました。")
