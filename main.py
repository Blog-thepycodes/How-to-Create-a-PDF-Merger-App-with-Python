import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import TkinterDnD, DND_FILES
import PyPDF2


# Initialize main window
root = TkinterDnD.Tk()
root.title("PDF Merger - The Pycodes")
root.geometry("500x500")


pdf_list = []


# Select PDFs function
def select_pdfs():
   files = filedialog.askopenfilenames(
       title="Select PDF files to merge",
       filetypes=[("PDF Files", "*.pdf")]
   )
   add_files(files)


# Handle Drag-and-Drop
def handle_drop(event):
   files = root.tk.splitlist(event.data)
   add_files(files)


# Add files to list
def add_files(files):
   for file in files:
       if file not in pdf_list:
           pdf_list.append(file)
           files_listbox.insert(tk.END, file)


# Move selected file up
def move_up():
   selected_index = files_listbox.curselection()
   if not selected_index:
       return
   index = selected_index[0]
   if index > 0:
       pdf_list[index], pdf_list[index - 1] = pdf_list[index - 1], pdf_list[index]
       update_listbox_selection(index - 1)


# Move selected file down
def move_down():
   selected_index = files_listbox.curselection()
   if not selected_index:
       return
   index = selected_index[0]
   if index < len(pdf_list) - 1:
       pdf_list[index], pdf_list[index + 1] = pdf_list[index + 1], pdf_list[index]
       update_listbox_selection(index + 1)


# Remove selected file
def remove_selected():
   selected_index = files_listbox.curselection()
   if not selected_index:
       return
   index = selected_index[0]
   pdf_list.pop(index)
   files_listbox.delete(index)


# Clear all files
def clear_all():
   pdf_list.clear()
   files_listbox.delete(0, tk.END)


# Update Listbox Selection
def update_listbox_selection(new_index):
   files_listbox.delete(0, tk.END)
   for file in pdf_list:
       files_listbox.insert(tk.END, file)
   files_listbox.select_set(new_index)


# Merge PDFs
def merge_pdfs():
   if not pdf_list:
       messagebox.showerror("Error", "No PDF files selected!")
       return


   output_path = filedialog.asksaveasfilename(
       title="Save Merged PDF",
       defaultextension=".pdf",
       filetypes=[("PDF Files", "*.pdf")]
   )
   if not output_path:
       return  # User canceled the save dialog


   pdf_writer = PyPDF2.PdfWriter()
   try:
       progress = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate", maximum=len(pdf_list))
       progress.pack(pady=10)
       for idx, pdf_file in enumerate(pdf_list):
           pdf_reader = PyPDF2.PdfReader(pdf_file)
           for page_num in range(len(pdf_reader.pages)):
               pdf_writer.add_page(pdf_reader.pages[page_num])
           progress["value"] = idx + 1
           root.update_idletasks()


       with open(output_path, "wb") as output_pdf:
           pdf_writer.write(output_pdf)
       messagebox.showinfo("Success", f"Merged PDF saved as: {output_path}")
   except Exception as e:
       messagebox.showerror("Error", f"Failed to merge PDFs: {e}")
   finally:
       progress.pack_forget()  # Hide the progress bar after merging


# Set up UI components
ttk.Button(root, text="Select PDFs", command=select_pdfs).pack(pady=10)
ttk.Button(root, text="Merge PDFs", command=merge_pdfs).pack(pady=5)


files_listbox = tk.Listbox(root, selectmode=tk.SINGLE, width=50, height=15)
files_listbox.pack(pady=10)


scrollbar = ttk.Scrollbar(root, orient="vertical", command=files_listbox.yview)
scrollbar.pack(side="right", fill="y")
files_listbox.config(yscrollcommand=scrollbar.set)


# Reorder and clear buttons
ttk.Button(root, text="Move Up", command=move_up).pack(pady=2)
ttk.Button(root, text="Move Down", command=move_down).pack(pady=2)
ttk.Button(root, text="Remove Selected", command=remove_selected).pack(pady=2)
ttk.Button(root, text="Clear All", command=clear_all).pack(pady=2)


# Drag and Drop Bindings
root.drop_target_register(DND_FILES)
root.dnd_bind('<<Drop>>', handle_drop)


# Start the application
root.mainloop()
