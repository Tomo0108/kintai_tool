import os
os.environ['TK_SILENCE_DEPRECATION'] = '1'  # tkの警告を抑制

# ログ設定を最初に行う
from config import setup_logging
setup_logging()

# Tkウィンドウを非表示にする
import tkinter as tk
root = tk.Tk()
root.withdraw()
root.update()  # 確実に非表示にするためにupdateを呼び出す

import customtkinter as ctk
from tkinter import filedialog, messagebox
import sys
from pathlib import Path  # 標準ライブラリからインポート
import subprocess
import logging
from datetime import datetime
from config import (get_input_dir, get_template_path, 
                   get_default_employee_name, update_config,
                   get_output_dir, get_app_path, ensure_config)
from main import process_attendance
import configparser

def setup_required_files():
    """必要なファイルとディレクトリを設定"""
    try:
        app_path = get_app_path()
        logging.info(f"Setting up required directories at: {app_path}")
        
        # 必要なディレクトリを作成
        required_dirs = ['input', 'output', 'templates', 'logs']
        for dir_name in required_dirs:
            dir_path = app_path / dir_name
            dir_path.mkdir(exist_ok=True)
            logging.info(f"Created directory: {dir_path}")
        
        # テンプレートファイルのコピー
        template_dir = app_path / 'templates'
        template_file = template_dir / '勤怠表雛形_2025年版.xlsx'
        if not template_file.exists():
            source_template = Path(sys._MEIPASS) / 'templates' / '勤怠表雛形_2025年版.xlsx'
            if source_template.exists():
                import shutil
                shutil.copy2(source_template, template_file)
                logging.info(f"Copied template file to: {template_file}")
        
        # config.iniの生成
        config_file = app_path / 'config.ini'
        if not config_file.exists():
            config = configparser.ConfigParser()
            config['DEFAULT'] = {
                'employee_name': '小島知将',
                'template_path': str(template_file)
            }
            config['PATHS'] = {
                'input_dir': str(app_path / 'input'),
                'output_dir': str(app_path / 'output')
            }
            
            with open(config_file, 'w', encoding='utf-8') as configfile:
                config.write(configfile)
            logging.info(f"Created config file: {config_file}")
        
        return True
    except Exception as e:
        logging.error(f"Error in setup_required_files: {e}")
        raise

class AttendanceApp:
    def __init__(self):
        try:
            logging.info("Initializing AttendanceApp")
            
            # テーマとカラーの設定
            ctk.set_appearance_mode("system")
            ctk.set_default_color_theme("blue")
            
            # メインウィンドウ
            self.root = ctk.CTk()
            self.root.title("勤怠表自動作成ツール")
            self.root.geometry("700x600")
            
            logging.info("Setting up UI")
            self.setup_ui()
            logging.info("UI setup complete")
            
        except Exception as e:
            logging.error(f"Error in AttendanceApp initialization: {e}", exc_info=True)
            raise
    
    def setup_ui(self):
        try:
            # タイトル
            title_label = ctk.CTkLabel(self.root, text="勤怠表自動作成ツール", 
                                     font=ctk.CTkFont(size=24, weight="bold"))
            title_label.pack(pady=20)
            
            # メインフレーム
            main_frame = ctk.CTkFrame(self.root)
            main_frame.pack(padx=20, pady=10, fill="both", expand=True)
            
            # CSVファイル選択
            csv_label = ctk.CTkLabel(main_frame, text="CSVファイル:")
            csv_label.pack(padx=20, pady=(20,5), anchor="w")
            
            csv_frame = ctk.CTkFrame(main_frame)
            csv_frame.pack(padx=20, pady=(0,10), fill="x")
            
            self.csv_entry = ctk.CTkEntry(csv_frame, placeholder_text="CSVファイルを選択してください")
            self.csv_entry.pack(side="left", padx=(0,10), fill="x", expand=True)
            
            csv_button = ctk.CTkButton(csv_frame, text="参照", command=self.select_csv, width=100)
            csv_button.pack(side="right")
            
            # テンプレート選択
            template_label = ctk.CTkLabel(main_frame, text="テンプレート:")
            template_label.pack(padx=20, pady=(20,5), anchor="w")
            
            template_frame = ctk.CTkFrame(main_frame)
            template_frame.pack(padx=20, pady=(0,10), fill="x")
            
            self.template_entry = ctk.CTkEntry(template_frame)
            self.template_entry.pack(side="left", padx=(0,10), fill="x", expand=True)
            self.template_entry.insert(0, get_template_path())
            
            template_button = ctk.CTkButton(template_frame, text="参照", 
                                          command=self.select_template, width=100)
            template_button.pack(side="right")
            
            # 従業員名
            name_label = ctk.CTkLabel(main_frame, text="従業員名:")
            name_label.pack(padx=20, pady=(20,5), anchor="w")
            
            self.name_entry = ctk.CTkEntry(main_frame, placeholder_text="従業員名を入力してください")
            self.name_entry.pack(padx=20, pady=(0,10), fill="x")
            self.name_entry.insert(0, get_default_employee_name())
            
            # ステータス表示
            self.status_label = ctk.CTkLabel(main_frame, text="")
            self.status_label.pack(pady=20)
            
            # 操作ボタン
            button_frame = ctk.CTkFrame(main_frame)
            button_frame.pack(pady=20)
            
            convert_button = ctk.CTkButton(button_frame, text="変換実行", 
                                         command=self.convert_files, width=150)
            convert_button.pack(side="left", padx=10)
            
            reset_button = ctk.CTkButton(button_frame, text="リセット", 
                                       command=self.reset_form, width=150)
            reset_button.pack(side="left", padx=10)
            
            output_button = ctk.CTkButton(button_frame, text="出力フォルダを開く", 
                                        command=self.open_output, width=150)
            output_button.pack(side="left", padx=10)
            
            # バージョン情報
            version_label = ctk.CTkLabel(self.root, text="Version 1.0.0")
            version_label.pack(pady=10)
            
        except Exception as e:
            logging.error(f"Error in setup_ui: {e}", exc_info=True)
            raise
    
    def select_csv(self):
        try:
            filename = filedialog.askopenfilename(
                title="CSVファイルを選択",
                filetypes=[("CSV files", "*.csv")],
                initialdir=get_input_dir()
            )
            if filename:
                self.csv_entry.delete(0, "end")
                self.csv_entry.insert(0, filename)
                self.status_label.configure(text="CSVファイルを選択しました")
        except Exception as e:
            logging.error(f"Error in select_csv: {e}", exc_info=True)
            messagebox.showerror("エラー", f"CSVファイル選択中にエラーが発生しました: {e}")
    
    def select_template(self):
        try:
            filename = filedialog.askopenfilename(
                title="テンプレートファイルを選択",
                filetypes=[("Excel files", "*.xlsx")],
                initialdir=str(Path(get_template_path()).parent)
            )
            if filename:
                self.template_entry.delete(0, "end")
                self.template_entry.insert(0, filename)
                update_config(self.name_entry.get(), filename)
                self.status_label.configure(text="テンプレートを選択しました")
        except Exception as e:
            logging.error(f"Error in select_template: {e}", exc_info=True)
            messagebox.showerror("エラー", f"テンプレート選択中にエラーが発生しました: {e}")
    
    def convert_files(self):
        try:
            self.status_label.configure(text="変換処理中...")
            self.root.update()

            csv_path = self.csv_entry.get()
            if not csv_path:
                messagebox.showwarning("警告", "CSVファイルを選択してください")
                self.status_label.configure(text="CSVファイルが選択されていません")
                return

            employee_name = self.name_entry.get()
            if not employee_name:
                messagebox.showwarning("警告", "従業員名を入力してください")
                self.status_label.configure(text="従業員名が入力されていません")
                return

            # 出力ディレクトリのパスをログ出力
            output_dir = get_output_dir()
            logging.info(f"Output directory: {output_dir}")
            logging.info(f"Output directory exists: {Path(output_dir).exists()}")
            
            # CSVファイルのパスをログ出力
            csv_file = Path(csv_path)
            logging.info(f"CSV file: {csv_file}")
            
            # テンプレートのパスをログ出力
            template_path = self.template_entry.get()
            logging.info(f"Template path: {template_path}")
            
            # 処理実行
            output_path = process_attendance(
                csv_path=str(csv_file),
                template_path=template_path,
                employee_name=employee_name
            )
            
            logging.info(f"Output path: {output_path}")
            
            # 設定を保存
            update_config(employee_name, template_path)
            
            self.status_label.configure(text="変換が完了しました")
            messagebox.showinfo("完了", f"ファイルを保存しました: {output_path}")
            
        except Exception as e:
            logging.error(f"Error in convert_files: {e}", exc_info=True)
            self.status_label.configure(text="エラーが発生しました")
            messagebox.showerror("エラー", f"変換中にエラーが発生しました: {e}")
    
    def open_output(self):
        subprocess.run(['open', get_output_dir()])

    def reset_form(self):
        self.csv_entry.delete(0, "end")
        self.template_entry.delete(0, "end")
        self.template_entry.insert(0, get_template_path())
        self.name_entry.delete(0, "end")
        self.name_entry.insert(0, get_default_employee_name())
        self.status_label.configure(text="")
    
    def run(self):
        try:
            logging.info("Starting mainloop")
            self.root.mainloop()
        except Exception as e:
            logging.error(f"Error in mainloop: {e}", exc_info=True)
            raise

def main():
    try:
        logging.info("Starting application...")
        setup_required_files()
        app = AttendanceApp()
        app.run()
    except Exception as e:
        logging.critical(f"Critical error in main: {e}", exc_info=True)
        messagebox.showerror("エラー", f"アプリケーションの起動に失敗しました: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()