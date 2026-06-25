import tkinter as tk
from tkinter import filedialog, messagebox
from converter import generate_xml


def browse_input():
    path = filedialog.askopenfilename(
        title="Select Vyapar Daybook Excel",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    input_var.set(path)


def browse_output():
    path = filedialog.askdirectory(title="Select Output Folder")
    output_var.set(path)


def run_generate():
    try:
        generate_xml(
            input_file=input_var.get(),
            output_folder=output_var.get(),
            generate_masters=masters_var.get(),
            generate_vouchers=vouchers_var.get(),
            sales=sales_var.get(),
            receipts=receipts_var.get()
        )
        messagebox.showinfo("Success", "Tally XML files generated successfully.")
    except Exception as e:
        messagebox.showerror("Error", str(e))


root = tk.Tk()
root.title("Vyapar to Tally XML Generator")
root.geometry("650x350")

input_var = tk.StringVar()
output_var = tk.StringVar()

masters_var = tk.BooleanVar(value=True)
vouchers_var = tk.BooleanVar(value=True)
sales_var = tk.BooleanVar(value=True)
receipts_var = tk.BooleanVar(value=True)

tk.Label(root, text="Vyapar Daybook Excel").pack(anchor="w", padx=20, pady=(20, 5))
tk.Entry(root, textvariable=input_var, width=70).pack(padx=20)
tk.Button(root, text="Browse Input File", command=browse_input).pack(pady=5)

tk.Label(root, text="Output Folder").pack(anchor="w", padx=20, pady=(10, 5))
tk.Entry(root, textvariable=output_var, width=70).pack(padx=20)
tk.Button(root, text="Browse Output Folder", command=browse_output).pack(pady=5)

tk.Checkbutton(root, text="Generate Masters XML", variable=masters_var).pack(anchor="w", padx=20)
tk.Checkbutton(root, text="Generate Vouchers XML", variable=vouchers_var).pack(anchor="w", padx=20)
tk.Checkbutton(root, text="Sales Vouchers", variable=sales_var).pack(anchor="w", padx=20)
tk.Checkbutton(root, text="Receipt Vouchers", variable=receipts_var).pack(anchor="w", padx=20)

tk.Button(root, text="Generate XML", command=run_generate, height=2, width=25).pack(pady=20)

root.mainloop()