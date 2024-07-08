import tkinter as tk 
from tkinter import ttk
import os 
from tkinter import filedialog
import subprocess
import pandas as pd
from tkinter import messagebox as msgbox
import sys

root = tk.Tk()
style = ttk.Style(root)
root.title("HoverTech Excel Converter")

dir_path = os.path.dirname(os.path.realpath(__file__))
root.tk.call('source', os.path.join(dir_path, 'forest-dark.tcl'))
root.tk.call('source', os.path.join(dir_path, 'forest-light.tcl'))
style.theme_use("forest-dark")

def generatetheme():
    try:
      def get_path(filename):
        if hasattr(sys, "_MEIPASS"): #True if it's an exe
            return os.path.join(sys._MEIPASS, filename)
        else:
            return filename
      tcl_folder = get_path("forest-dark.tcl")
      root.tk.call('source', tcl_folder)
      root.tk.call("set_theme", "dark")
    except:
       pass

frame = ttk.Frame(root)
frame.pack(fill='none', expand=False)

widgets_frame = ttk.LabelFrame(frame, text="Insert File")
widgets_frame.grid(row=0, column=0, padx=20, pady=10, sticky="nw")

treeFrame = ttk.Frame(frame, width=500, height=700)
treeFrame.grid(row=0, column=1, padx=10, pady=10, sticky="nw")
treeFrame.grid_propagate(False)  # Prevents the frame from resizing to fit its children
treeFrame.columnconfigure(0, weight=1)
treeFrame.rowconfigure(0, weight=1)

treeScrollY = ttk.Scrollbar(treeFrame)
treeScrollY.grid(row=0, column=1, sticky='ns')
treeScrollX = ttk.Scrollbar(treeFrame, orient="horizontal")
treeScrollX.grid(row=1, column=0, sticky='ew')

treeview = ttk.Treeview(treeFrame, yscrollcommand=treeScrollY.set, xscrollcommand=treeScrollX.set)
treeview.grid(row=0, column=0, sticky='nsew')
treeScrollY.config(command=treeview.yview)
treeScrollX.config(command=treeview.xview)


file_path = ""
output_file_path = ""

def open_file_dialog():
    global file_path
    file_path = filedialog.askopenfilename(title="Select a File", filetypes=[("Excel", "*.xlsx"), ("All files", "*.*")])
    selected_file_label.config(text=f"Selected File: {file_path}")
    
    try:
        df = pd.read_excel(file_path)
    except FileNotFoundError:
        msgbox.showerror("Woah!", "File not found")
        return
    except Exception as e:
        msgbox.showerror("Woah!", f"Please choose a Microsoft Excel Workbook (.XLSX) file: {e}" )
        return
         
    treeview.delete(*treeview.get_children())
    
    treeview["column"] = list(df.columns)
    treeview["show"] = "headings"

    for col in treeview["column"]:
        treeview.heading(col, text=col)

    df_rows = df.head(50).to_numpy().tolist()
    for row in df_rows:
        treeview.insert("", "end", values=row)

    
def save_file_dialog():
        global output_file_path
        output_file_path = filedialog.asksaveasfilename(title="Save the output Excel file", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        output_selected_file_label.config(text=f"Selected File: {output_file_path}")

        if not output_file_path:
            msgbox.showerror("Whoops!", "No save file selected!")
        return

open_file_button = ttk.Button(widgets_frame, text="Choose File", command=open_file_dialog)
open_file_button.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

selected_file_label = ttk.Label(widgets_frame, text="Selected File:")
selected_file_label.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

save_file_button = ttk.Button(widgets_frame, text="Save File", command=save_file_dialog)
save_file_button.grid(row=2, column=0, padx=10, pady=10, sticky="ew")

output_selected_file_label = ttk.Label(widgets_frame, text="Selected File:")
output_selected_file_label.grid(row=3, column=0, padx=10, pady=10, sticky="ew")

def run_pandasvs():
    if not file_path:
        msgbox.showerror("Whoops!", "No file selected")
        return
    if not output_file_path:
        msgbox.showerror("Whoops!", "No output file selected")
        return
    try:
        script_path = os.path.join(dir_path, "PandasVS.py")
        subprocess.check_call(["python3", script_path, file_path, output_file_path])
        msgbox.showinfo("Yay!", "Success! Check your files")
        return
    except Exception as e:
        msgbox.showerror("What", f"Failed to convert: {e}")
        return 

convert_button = ttk.Button(widgets_frame, text="Convert", command=run_pandasvs)
convert_button.grid(row=5, column=0, padx=10, pady=10, sticky="ew")

seperator = ttk.Separator(widgets_frame)
seperator.grid(row=4, column=0, padx=5, pady=(10, 10), sticky="ew")

def toggle_mode():
    if mode_switch.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")

mode_switch = ttk.Checkbutton(widgets_frame, text="Mode", style="Switch", command=toggle_mode)
mode_switch.grid(row=6, column=0, padx=5, pady=10, sticky="sw")

root.mainloop()
