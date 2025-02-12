import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
from config import TEMPLATE_PATH, CONFIG_FILE, INPUT_DIR
import configparser

# 設定ファイルの初期化
config = configparser.ConfigParser()

# `config.ini` が存在しない場合、新規作成
if not os.path.exists(CONFIG_FILE):
    config["Settings"] = {
        "DEFAULT_EMPLOYEE_NAME": "氏名",
        "TEMPLATE_PATH": "templates/勤怠表雛形_2025年版.xlsx"
    }
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        config.write(f)

# `config.ini` を読み込む
config.read(CONFIG_FILE, encoding="utf-8")

# `[Settings]` セクションがない場合、追加
if "Settings" not in config:
    config["Settings"] = {
        "DEFAULT_EMPLOYEE_NAME": "氏名",
        "TEMPLATE_PATH": "templates/勤怠表雛形_2025年版.xlsx"
    }
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        config.write(f)

def select_csv():
    """CSVファイルを選択する"""
    file_path = filedialog.askopenfilename(initialdir=INPUT_DIR, title="CSVファイルを選択", filetypes=[("CSV files", "*.csv")])
    if file_path:
        csv_entry.delete(0, tk.END)
        csv_entry.insert(0, file_path)

def select_template():
    """テンプレートファイルを選択する"""
    file_path = filedialog.askopenfilename(initialdir=os.path.dirname(TEMPLATE_PATH), title="テンプレートファイルを選択", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        template_entry.delete(0, tk.END)
        template_entry.insert(0, file_path)

def update_config(name, template):
    """config.ini に氏名とテンプレートファイルのパスを保存"""
    config["Settings"]["DEFAULT_EMPLOYEE_NAME"] = name
    config["Settings"]["TEMPLATE_PATH"] = template
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        config.write(f)

def open_output_folder():
    """出力フォルダを開く"""
    output_dir = os.path.abspath("output")
    if os.path.exists(output_dir):
        os.startfile(output_dir)  # Windows
    else:
        messagebox.showerror("エラー", "出力フォルダが見つかりません。")

def run_main():
    csv_file = csv_entry.get()
    template_file = template_entry.get()
    employee_name = name_entry.get().strip()
    
    if not csv_file:
        messagebox.showerror("エラー", "CSVファイルを選択してください。")
        return
    
    if not template_file:
        messagebox.showerror("エラー", "テンプレートファイルを選択してください。")
        return
    
    if not employee_name:
        messagebox.showerror("エラー", "氏名を入力してください。")
        return
    
    # config.ini を更新
    update_config(employee_name, template_file)
    
    try:
        # `.exe` 実行時は `main.exe` を実行
        if getattr(sys, 'frozen', False):
            main_executable = os.path.join(os.path.dirname(sys.executable), "main.exe")
        else:
            main_executable = "python main.py"
        
        result = subprocess.run(
            [main_executable, "--name", employee_name, "--file", csv_file, "--template", template_file],
            capture_output=True, text=True, check=True
        )
        
        messagebox.showinfo("完了", "勤怠表の作成が完了しました！")
        open_output_folder()

    except subprocess.CalledProcessError as e:
        error_message = f"エラー発生:\n{e.stderr}\n{e.stdout}"
        messagebox.showerror("エラー", error_message)
        print(error_message)  # コマンドプロンプトでもエラーを確認


# GUIウィンドウの作成
root = tk.Tk()
root.title("勤怠表作成ツール")
root.geometry("500x400")
root.configure(bg="#f5f5f5")

font_style = ("Arial", 12)
button_style = {"font": font_style, "bg": "#4CAF50", "fg": "white", "activebackground": "#45a049", "padx": 10, "pady": 5}

tk.Label(root, text="CSVファイル:", font=font_style, bg="#f5f5f5").pack(pady=(10, 0))

csv_entry = tk.Entry(root, width=50, font=font_style)
csv_entry.pack()
tk.Button(root, text="参照", command=select_csv, **button_style).pack(pady=5)

tk.Label(root, text="テンプレートファイル:", font=font_style, bg="#f5f5f5").pack(pady=(10, 0))

template_entry = tk.Entry(root, width=50, font=font_style)
template_entry.insert(0, TEMPLATE_PATH)
template_entry.pack()
tk.Button(root, text="参照", command=select_template, **button_style).pack(pady=5)

tk.Label(root, text="氏名:", font=font_style, bg="#f5f5f5").pack(pady=(10, 0))
name_entry = tk.Entry(root, width=50, font=font_style)
name_entry.insert(0, config["Settings"].get("DEFAULT_EMPLOYEE_NAME", ""))
name_entry.pack()

tk.Button(root, text="実行", command=run_main, **button_style).pack(pady=15)

root.mainloop()
