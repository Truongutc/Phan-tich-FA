import tkinter as tk
from tkinter import messagebox
import subprocess
import threading
import sys
import os

def run_pipeline():
    ticker = entry.get().strip().upper()
    if not ticker:
        messagebox.showwarning("Cảnh báo", "Vui lòng nhập mã cổ phiếu!")
        return
        
    btn_run.config(state=tk.DISABLED)
    lbl_status.config(text=f"Đang phân tích {ticker}... Vui lòng đợi...", fg="#2563eb")
    
    def work():
        try:
            # Run the python script in the project directory
            proj_dir = os.path.dirname(os.path.abspath(__file__))
            script_path = os.path.join(proj_dir, "run_analysis.py")
            
            result = subprocess.run(
                [sys.executable, script_path, ticker], 
                cwd=proj_dir,
                capture_output=True, 
                text=True, 
                encoding="utf-8"
            )
            
            if result.returncode == 0:
                lbl_status.config(text=f"Phân tích thành công mã {ticker}!", fg="#10b981")
                messagebox.showinfo("Thành công", f"Đã hoàn thành phân tích mã {ticker}!")
            else:
                lbl_status.config(text="Có lỗi xảy ra khi phân tích!", fg="#ef4444")
                error_msg = result.stderr or result.stdout or "Không rõ lỗi."
                messagebox.showerror("Lỗi", f"Không thể phân tích {ticker}.\nChi tiết:\n{error_msg}")
        except Exception as e:
            lbl_status.config(text="Lỗi hệ thống!", fg="#ef4444")
            messagebox.showerror("Lỗi", f"Lỗi hệ thống: {e}")
        finally:
            btn_run.config(state=tk.NORMAL)
            entry.delete(0, tk.END)
            
    threading.Thread(target=work, daemon=True).start()

# Build GUI
root = tk.Tk()
root.title("AIC FA - Công cụ Phân tích Cổ phiếu")
root.geometry("450x250")
root.configure(bg="#f8fafc")

# Set window icon if exists
icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app_icon.png")
if os.path.exists(icon_path):
    try:
        img = tk.PhotoImage(file=icon_path)
        root.tk.call('wm', 'iconphoto', root._w, img)
    except:
        pass

# Center screen
root.eval('tk::PlaceWindow . center')

lbl_title = tk.Label(root, text="Hệ thống Phân tích Cổ phiếu AIC", font=("Helvetica", 14, "bold"), bg="#f8fafc", fg="#1e293b")
lbl_title.pack(pady=20)

frame = tk.Frame(root, bg="#f8fafc")
frame.pack(pady=10)

lbl_prompt = tk.Label(frame, text="Nhập mã cổ phiếu:", font=("Helvetica", 10), bg="#f8fafc", fg="#475569")
lbl_prompt.pack(side=tk.LEFT, padx=5)

entry = tk.Entry(frame, font=("Helvetica", 12), width=15, justify="center")
entry.pack(side=tk.LEFT, padx=5)
entry.focus()

btn_run = tk.Button(root, text="Bắt đầu phân tích", font=("Helvetica", 10, "bold"), bg="#2563eb", fg="white", activebackground="#1d4ed8", activeforeground="white", command=run_pipeline, padx=15, pady=5)
btn_run.pack(pady=15)

lbl_status = tk.Label(root, text="Sẵn sàng", font=("Helvetica", 9, "italic"), bg="#f8fafc", fg="#64748b")
lbl_status.pack(pady=5)

# Bind enter key
entry.bind("<Return>", lambda event: run_pipeline())

root.mainloop()
