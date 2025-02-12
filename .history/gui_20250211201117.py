import os
import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
from config import INPUT_DIR, update_config, DEFAULT_EMPLOYEE_NAME

def select_csv():
    file_path = filedialog.askopenfilename(initialdir=INPUT_DIR, title="CSVファイルを選択", filetypes=[("CSV files", "*.csv")])
    if file_path:
        csv_entry.delete(0, tk.END)
        csv_entry.insert(0, file_path)

def run_main():
    csv_file = csv_entry.get()
    employee_name = name_entry.get().strip()
    
    if not csv_file:
        messagebox.showerror("エラー", "CSVファイルを選択してください。")
        return
    
    if not employee_name:
        messagebox.showerror("エラー", "氏名を入力してください。")
        return
    
    # config.ini を更新
    update_config(employee_name)

    try:
        subprocess.run(["python", "main.py", "--name", employee_name, "--file", csv_file], check=True)
        messagebox.showinfo("完了", "勤怠表の作成が完了しました！")
    except subprocess.CalledProcessError:
        messagebox.showerror("エラー", "処理中にエラーが発生しました。")

# GUIウィンドウの作成
root = tk.Tk()
root.title("勤怠表作成ツール")
root.geometry("400x200")

tk.Label(root, text="CSVファイル:").pack()
csv_entry = tk.Entry(root, width=40)
csv_entry.pack()
tk.Button(root, text="参照", command=select_csv).pack()

tk.Label(root, text="氏名:").pack()
name_entry = tk.Entry(root, width=40)
name_entry.insert(0, DEFAULT_EMPLOYEE_NAME)  # デフォルトの氏名をセット
name_entry.pack()

tk.Button(root, text="実行", command=run_main).pack()

root.mainloop()
