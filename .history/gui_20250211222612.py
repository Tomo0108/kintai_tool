import os
import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
from config import INPUT_DIR, TEMPLATE_PATH, CONFIG_FILE, DEFAULT_EMPLOYEE_NAME, update_config
import configparser

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
        subprocess.run(["python", "main.py", "--name", employee_name, "--file", csv_file, "--template", template_file], check=True)
        messagebox.showinfo("完了", "勤怠表の作成が完了しました！")
    except subprocess.CalledProcessError:
        messagebox.showerror("エラー", "処理中にエラーが発生しました。")

# GUIウィンドウの作成
root = tk.Tk()
root.title("勤怠表作成ツール")
root.geometry("4x300")
root.configure(bg="#f5f5f5")

font_style = ("Arial", 12)
button_style = {"font": font_style, "bg": "#4CAF50", "fg": "white", "activebackground": "#45a049", "padx": 10, "pady": 5}

tk.Label(root, text="CSVファイル:", font=font_style, bg="#f5f5f5").pack(pady=(10, 0))
csv_entry = tk.Entry(root, width=50, font=font_style)
csv_entry.pack()
tk.Button(root, text="参照", command=select_csv, **button_style).pack(pady=5)

tk.Label(root, text="テンプレートファイル:", font=font_style, bg="#f5f5f5").pack(pady=(10, 0))
template_entry = tk.Entry(root, width=50, font=font_style)
template_entry.insert(0, TEMPLATE_PATH)  # 設定ファイルから取得
template_entry.pack()
tk.Button(root, text="参照", command=select_template, **button_style).pack(pady=5)

tk.Label(root, text="氏名:", font=font_style, bg="#f5f5f5").pack(pady=(10, 0))
name_entry = tk.Entry(root, width=50, font=font_style)
name_entry.insert(0, DEFAULT_EMPLOYEE_NAME)  # 設定ファイルから取得
name_entry.pack()

tk.Button(root, text="実行", command=run_main, **button_style).pack(pady=15)

root.mainloop()
