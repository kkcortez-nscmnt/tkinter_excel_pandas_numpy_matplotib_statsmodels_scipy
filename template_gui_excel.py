import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

#---------------------------------------------------
#funções

def File_dialog():
    filename = filedialog.askopenfilename(initialdir="/", 
                                          title='Select a file', 
                                          filetype=(('xlsx files', '*.xlsx'),("All Files", "*.*")))
    label_file["text"] = filename
    return None

def Load_excel_data():
    file_path = label_file['text']
    try:
        excel_filename = r"{}".format(file_path)
        df = pd.read_excel(excel_filename)
    except ValueError:
        tk.messagebox.showerror("Information", "Arquivo invalido")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Information", f"No such file as {file_path}")
        return None
    
    clear_data()
    tv1["column"] = list(df.columns)
    tv1["show"] = "headings"
    for column in tv1['columns']:
        tv1.heading(column, text=column)
    df_rows = df.to_numpy().tolist()
    for row in df_rows:
        tv1.insert("", "end", values=row)
    return None

def clear_data():
    tv1.delete(*tv1.get_children())


#---------------------------------------------------
#GUI
root = tk.Tk()
root.geometry('500x400')
root.pack_propagate(False)
root.resizable(0, 0)
#--------------------------------------------------
#widgets
frame1 = tk.LabelFrame(root,  text="Excel Data")
file_frame = tk.LabelFrame(root, text='Openfile')

button1 = tk.Button(file_frame, text="Browse a File", command=lambda:File_dialog()) 
button2 = tk.Button(file_frame, text="Load File", command=lambda:Load_excel_data()) 

label_file = ttk.Label(file_frame, text="No file Selected")

tv1 = ttk.Treeview(frame1)

treescrolly = tk.Scrollbar(frame1, orient='vertical', command=tv1.yview)
treescrollx = tk.Scrollbar(frame1, orient='horizontal', command=tv1.xview)
tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set)
#---------------------------------------------------
#layout
frame1.place(height=250, width=500)
file_frame.place(height=100, width=400,rely=0.65, relx=0)

button1.place(rely=0.65, relx=0.50)
button2.place(rely=0.65, relx=0.30)

label_file.place(rely=0 , relx=0)

tv1.place(relheight=1, relwidth=1)

treescrollx.pack(side='bottom',fill='x')
treescrolly.pack(side='right', fill='y')
#---------------------------------------------------------------
#loop janela principal
root.mainloop()