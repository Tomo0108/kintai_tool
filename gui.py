import os
os.environ['TK_SILENCE_DEPRECATION'] = '1'

import customtkinter as ctk
from tkinter import filedialog, messagebox
import sys
import subprocess
from pathlib import Path
from config import (get_input_dir, get_template_path, 
                   get_default_employee_name, update_config,
                   get_output_dir)
from main import process_attendance

def main():
    # テーマとカラーの設定
    ctk.set_appearance_mode("system")  # システムのテーマに従う
    ctk.set_default_color_theme("blue")  # ブルーテーマを使用
    
    # メインウィンドウ
    root = ctk.CTk()
    root.title("勤怠表自動作成ツール")
    root.geometry("700x600")
    
    # タイトル
    title_label = ctk.CTkLabel(root, text="勤怠表自動作成ツール", 
                              font=ctk.CTkFont(size=24, weight="bold"))
    title_label.pack(pady=20)
    
    # メインフレーム
    main_frame = ctk.CTkFrame(root)
    main_frame.pack(padx=20, pady=10, fill="both", expand=True)
    
    # CSVファイル選択
    csv_label = ctk.CTkLabel(main_frame, text="CSVファイル:")
    csv_label.pack(padx=20, pady=(20,5), anchor="w")
    
    csv_frame = ctk.CTkFrame(main_frame)
    csv_frame.pack(padx=20, pady=(0,10), fill="x")
    
    csv_entry = ctk.CTkEntry(csv_frame, placeholder_text="CSVファイルを選択してください")
    csv_entry.pack(side="left", padx=(0,10), fill="x", expand=True)
    
    def select_csv():
        file_path = filedialog.askopenfilename(
            initialdir=get_input_dir(),
            title="CSVファイルを選択",
            filetypes=[("CSV files", "*.csv")]
        )
        if file_path:
            csv_entry.delete(0, "end")
            csv_entry.insert(0, file_path)
    
    csv_button = ctk.CTkButton(csv_frame, text="参照", command=select_csv, width=100)
    csv_button.pack(side="right")
    
    # テンプレート選択
    template_label = ctk.CTkLabel(main_frame, text="テンプレート:")
    template_label.pack(padx=20, pady=(20,5), anchor="w")
    
    template_frame = ctk.CTkFrame(main_frame)
    template_frame.pack(padx=20, pady=(0,10), fill="x")
    
    template_entry = ctk.CTkEntry(template_frame)
    template_entry.pack(side="left", padx=(0,10), fill="x", expand=True)
    template_entry.insert(0, get_template_path())
    
    def select_template():
        file_path = filedialog.askopenfilename(
            initialdir=str(Path(get_template_path()).parent),
            title="テンプレートファイルを選択",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file_path:
            template_entry.delete(0, "end")
            template_entry.insert(0, file_path)
            update_config(name_entry.get(), file_path)
    
    template_button = ctk.CTkButton(template_frame, text="参照", command=select_template, width=100)
    template_button.pack(side="right")
    
    # 従業員名
    name_label = ctk.CTkLabel(main_frame, text="従業員名:")
    name_label.pack(padx=20, pady=(20,5), anchor="w")
    
    name_entry = ctk.CTkEntry(main_frame, placeholder_text="従業員名を入力してください")
    name_entry.pack(padx=20, pady=(0,10), fill="x")
    name_entry.insert(0, get_default_employee_name())
    
    # 出力先
    output_label = ctk.CTkLabel(main_frame, text="出力先:")
    output_label.pack(padx=20, pady=(20,5), anchor="w")
    
    output_frame = ctk.CTkFrame(main_frame)
    output_frame.pack(padx=20, pady=(0,10), fill="x")
    
    output_entry = ctk.CTkEntry(output_frame, state="readonly")
    output_entry.pack(side="left", padx=(0,10), fill="x", expand=True)
    output_entry.insert(0, str(Path(get_output_dir())))
    
    def open_output():
        subprocess.run(['open', get_output_dir()])
    
    output_button = ctk.CTkButton(output_frame, text="フォルダを開く", 
                                 command=open_output, width=100)
    output_button.pack(side="right")
    
    # ステータス表示
    status_label = ctk.CTkLabel(main_frame, text="")
    status_label.pack(pady=20)
    
    # 操作ボタン
    button_frame = ctk.CTkFrame(main_frame)
    button_frame.pack(pady=20)
    
    def convert():
        if not csv_entry.get() or not template_entry.get() or not name_entry.get():
            messagebox.showerror("エラー", "必須項目を入力してください。")
            return
        
        try:
            status_label.configure(text="変換中...", text_color="blue")
            root.update()
            
            output_path = process_attendance(
                csv_path=csv_entry.get(),
                template_path=template_entry.get(),
                employee_name=name_entry.get()
            )
            
            status_label.configure(text="✅ 変換が完了しました！", text_color="green")
            messagebox.showinfo("完了", f"変換が完了しました！\n保存先: {output_path}")
            
            update_config(name_entry.get(), template_entry.get())
            
        except Exception as e:
            error_msg = f"処理中にエラーが発生しました:\n{str(e)}"
            status_label.configure(text="❌ エラーが発生しました", text_color="red")
            messagebox.showerror("エラー", error_msg)
    
    def reset_form():
        csv_entry.delete(0, "end")
        template_entry.delete(0, "end")
        template_entry.insert(0, get_template_path())
        name_entry.delete(0, "end")
        name_entry.insert(0, get_default_employee_name())
        status_label.configure(text="")
    
    convert_button = ctk.CTkButton(button_frame, text="変換実行", 
                                  command=convert, width=150)
    convert_button.pack(side="left", padx=10)
    
    reset_button = ctk.CTkButton(button_frame, text="リセット", 
                                command=reset_form, width=150)
    reset_button.pack(side="left", padx=10)
    
    # バージョン情報
    version_label = ctk.CTkLabel(root, text="Version 1.0.0")
    version_label.pack(pady=10)
    
    root.mainloop()

if __name__ == "__main__":
    main()