import os
import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import sys
import traceback
from pathlib import Path
from ttkbootstrap import Style, ttk
import ttkbootstrap as ttk
from config import (get_input_dir, get_template_path, 
                   get_default_employee_name, update_config)
from main import process_attendance

def select_csv():
    """CSVファイルを選択する"""
    file_path = filedialog.askopenfilename(
        initialdir=get_input_dir(), 
        title="CSVファイルを選択", 
        filetypes=[("CSV files", "*.csv")])
    if file_path:
        csv_entry.delete(0, tk.END)
        csv_entry.insert(0, file_path)

def select_template():
    """テンプレートファイルを選択する"""
    file_path = filedialog.askopenfilename(
        initialdir=os.path.dirname(get_template_path()), 
        title="テンプレートファイルを選択", 
        filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        template_entry.delete(0, tk.END)
        template_entry.insert(0, file_path)

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
    
    # config.py を更新
    update_config(employee_name, template_file)
    
    try:
        subprocess.run(["python", "main.py", "--name", employee_name, "--file", csv_file, "--template", template_file], check=True)
        messagebox.showinfo("完了", "勤怠表の作成が完了しました！")
        open_output_folder()
    except subprocess.CalledProcessError:
        messagebox.showerror("エラー", "処理中にエラーが発生しました。")

class AttendanceConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("勤怠表作成ツール")
        self.root.geometry("500x400")
        self.root.configure(bg="#f5f5f5")

        self.font_style = ("Arial", 12)
        self.button_style = {"font": self.font_style, "bg": "#4CAF50", "fg": "white", "activebackground": "#45a049", "padx": 10, "pady": 5}

        self.create_widgets()

    def create_widgets(self):
        tk.Label(self.root, text="CSVファイル:", font=self.font_style, bg="#f5f5f5").pack(pady=(10, 0))
        self.csv_entry = tk.Entry(self.root, width=50, font=self.font_style)
        self.csv_entry.pack()
        tk.Button(self.root, text="参照", command=select_csv, **self.button_style).pack(pady=5)

        tk.Label(self.root, text="テンプレートファイル:", font=self.font_style, bg="#f5f5f5").pack(pady=(10, 0))
        self.template_entry = tk.Entry(self.root, width=50, font=self.font_style)
        self.template_entry.insert(0, get_template_path())  # config.iniから取得
        self.template_entry.pack()
        tk.Button(self.root, text="参照", command=select_template, **self.button_style).pack(pady=5)

        tk.Label(self.root, text="氏名:", font=self.font_style, bg="#f5f5f5").pack(pady=(10, 0))
        self.name_entry = tk.Entry(self.root, width=50, font=self.font_style)
        self.name_entry.insert(0, get_default_employee_name())  # config.iniから取得
        self.name_entry.pack()

        self.status_label = tk.Label(self.root, text="", font=self.font_style, bg="#f5f5f5")
        self.status_label.pack(pady=(10, 0))

        tk.Button(self.root, text="実行", command=self.convert, **self.button_style).pack(pady=15)

    def convert(self):
        """変換処理を実行する"""
        csv_path = self.csv_entry.get()
        template_path = self.template_entry.get()
        name = self.name_entry.get()

        if not csv_path or not template_path or not name:
            messagebox.showerror("エラー", "すべての項目を入力してください。")
            return

        try:
            self.status_label.config(text="変換中...", foreground="blue")
            self.root.update()
            
            process_attendance(
                csv_path=csv_path,
                template_path=template_path,
                employee_name=name
            )
            
            self.status_label.config(
                text="✅ 変換が完了しました！", 
                foreground="green"
            )
            messagebox.showinfo("完了", "変換が完了しました！")
            
            # 設定を保存
            update_config(name, template_path)
            
        except Exception as e:
            error_msg = f"エラーが発生しました:\n{str(e)}\n\n"
            error_msg += "詳細:\n" + traceback.format_exc()
            print(error_msg)  # コンソールにエラーを出力
            
            self.status_label.config(
                text="❌ エラーが発生しました", 
                foreground="red"
            )
            messagebox.showerror("エラー", error_msg)

def main():
    try:
        root = ttk.Window()
        app = AttendanceConverterGUI(root)
        root.mainloop()
    except Exception as e:
        error_msg = f"起動時エラー:\n{str(e)}\n\n"
        error_msg += "詳細:\n" + traceback.format_exc()
        print(error_msg)  # コンソールにエラーを出力
        messagebox.showerror("致命的なエラー", error_msg)
        sys.exit(1)

if __name__ == "__main__":
    main()