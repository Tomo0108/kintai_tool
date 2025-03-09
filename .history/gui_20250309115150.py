import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys
import traceback
from pathlib import Path
from tkinter.font import Font
from config import (get_input_dir, get_template_path, get_output_dir,
                   get_default_employee_name, update_config)
from main import process_attendance
from utils import open_folder

class AttendanceConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("勤怠データ変換ツール")
        
        # スタイル設定
        self.style = ttk.Style()
        self.style.configure('Main.TFrame', background='#f0f0f0')
        self.style.configure('Title.TLabel', 
                           font=('Yu Gothic UI', 20, 'bold'),
                           background='#f0f0f0',
                           foreground='#1a73e8')
        self.style.configure('Custom.TButton',
                           font=('Yu Gothic UI', 10),
                           padding=10)
        self.style.configure('Status.TLabel',
                           font=('Yu Gothic UI', 10),
                           background='#f0f0f0')
        self.style.configure('Field.TLabel',
                           font=('Yu Gothic UI', 10),
                           background='#f0f0f0')
        
        # メインフレーム
        self.main_frame = ttk.Frame(root, padding=30, style='Main.TFrame')
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        
        # タイトル
        title_label = ttk.Label(
            self.main_frame, 
            text="勤怠データ変換ツール", 
            style='Title.TLabel'
        )
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 30))
        
        # CSVファイル選択
        ttk.Label(self.main_frame, text="CSVファイル:", style='Field.TLabel').grid(
            row=1, column=0, sticky="w", pady=10
        )
        self.csv_entry = ttk.Entry(self.main_frame, width=50, font=('Yu Gothic UI', 10))
        self.csv_entry.grid(row=1, column=1, padx=5, pady=10)
        ttk.Button(
            self.main_frame, 
            text="参照", 
            command=self.select_csv,
            style='Custom.TButton'
        ).grid(row=1, column=2, padx=(10, 0), pady=10)
        
        # テンプレート選択
        ttk.Label(self.main_frame, text="テンプレート:", style='Field.TLabel').grid(
            row=2, column=0, sticky="w", pady=10
        )
        self.template_entry = ttk.Entry(self.main_frame, width=50, font=('Yu Gothic UI', 10))
        self.template_entry.grid(row=2, column=1, padx=5, pady=10)
        self.template_entry.insert(0, get_template_path())
        ttk.Button(
            self.main_frame, 
            text="参照", 
            command=self.select_template,
            style='Custom.TButton'
        ).grid(row=2, column=2, padx=(10, 0), pady=10)
        
        # 従業員名入力
        ttk.Label(self.main_frame, text="従業員名:", style='Field.TLabel').grid(
            row=3, column=0, sticky="w", pady=10
        )
        self.name_entry = ttk.Entry(self.main_frame, width=50, font=('Yu Gothic UI', 10))
        self.name_entry.grid(row=3, column=1, padx=5, pady=10)
        self.name_entry.insert(0, get_default_employee_name())
        
        # ボタンフレーム
        button_frame = ttk.Frame(self.main_frame, style='Main.TFrame')
        button_frame.grid(row=4, column=0, columnspan=3, pady=(30,20))
        
        # 変換ボタン
        convert_button = ttk.Button(
            button_frame,
            text="変換実行",
            command=self.convert,
            style='Custom.TButton',
            width=15
        )
        convert_button.pack(side=tk.LEFT, padx=10)

        # フォルダを開くボタン
        open_folder_button = ttk.Button(
            button_frame,
            text="出力フォルダを開く",
            command=self.open_output_folder,
            style='Custom.TButton',
            width=15
        )
        open_folder_button.pack(side=tk.LEFT, padx=10)
        
        # ステータス表示
        self.status_label = ttk.Label(
            self.main_frame, 
            text="", 
            style='Status.TLabel'
        )
        self.status_label.grid(row=5, column=0, columnspan=3, pady=(10, 0))
        
        # グリッド設定
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(1, weight=1)

    def select_csv(self):
        """CSVファイルを選択する"""
        file_path = filedialog.askopenfilename(
            initialdir=get_input_dir(),
            title="CSVファイルを選択",
            filetypes=[("CSV files", "*.csv")]
        )
        if file_path:
            self.csv_entry.delete(0, tk.END)
            self.csv_entry.insert(0, file_path)

    def select_template(self):
        """テンプレートファイルを選択する"""
        file_path = filedialog.askopenfilename(
            initialdir=os.path.dirname(get_template_path()),
            title="テンプレートファイルを選択",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file_path:
            self.template_entry.delete(0, tk.END)
            self.template_entry.insert(0, file_path)
            update_config(self.name_entry.get(), file_path)

    def convert(self):
        """変換処理を実行する"""
        csv_path = self.csv_entry.get()
        template_path = self.template_entry.get()
        name = self.name_entry.get()

        if not csv_path or not template_path or not name:
            messagebox.showerror("エラー", "すべての項目を入力してください。")
            return

        try:
            self.status_label.config(text="変換中...", foreground="#1a73e8")
            self.root.update()
            
            process_attendance(
                csv_path=csv_path,
                template_path=template_path,
                employee_name=name
            )
            
            self.status_label.config(
                text="✅ 変換が完了しました！", 
                foreground="#34a853"
            )
            # 変換完了時にoutputフォルダを開く
            open_folder(get_output_dir())
            messagebox.showinfo("完了", "変換が完了しました！")
            
            # 設定を保存
            update_config(name, template_path)
            
        except Exception as e:
            error_msg = f"エラーが発生しました:\n{str(e)}\n\n"
            error_msg += "詳細:\n" + traceback.format_exc()
            print(error_msg)
            
            self.status_label.config(
                text="❌ エラーが発生しました", 
                foreground="#ea4335"
            )
            messagebox.showerror("エラー", error_msg)

    def open_output_folder(self):
        """出力フォルダを開く"""
        if open_folder(get_output_dir()):
            self.status_label.config(
                text="✅ 出力フォルダを開きました", 
                foreground="#34a853"
            )
        else:
            self.status_label.config(
                text="❌ フォルダを開けませんでした", 
                foreground="#ea4335"
            )

def main():
    try:
        root = tk.Tk()
        app = AttendanceConverterGUI(root)
        root.mainloop()
    except Exception as e:
        error_msg = f"起動時エラー:\n{str(e)}\n\n"
        error_msg += "詳細:\n" + traceback.format_exc()
        print(error_msg)
        messagebox.showerror("致命的なエラー", error_msg)
        sys.exit(1)

if __name__ == "__main__":
    main()
