import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from converter import generate_xml
from vyapar_daybook_to_tally_xml import MAX_DATA_AGE_DAYS


def browse_input():
    path = filedialog.askopenfilename(
        title="Select Vyapar Daybook Excel",
        filetypes=[("Excel Files", "*.xlsx *.xls")],
    )
    if path:
        input_var.set(path)
        if not output_var.get():
            output_var.set(path.rsplit("/", 1)[0] if "/" in path else path.rsplit("\\", 1)[0])


def browse_output():
    path = filedialog.askdirectory(title="Select Output Folder")
    if path:
        output_var.set(path)


def selected_age_days():
    try:
        value = int(max_age_var.get())
    except ValueError as exc:
        raise ValueError("Allowed data age must be a whole number of days.") from exc
    if value < 0:
        raise ValueError("Allowed data age cannot be negative. Use 0 to disable the check.")
    return value


def run_generate():
    try:
        result = generate_xml(
            input_file=input_var.get().strip(),
            output_folder=output_var.get().strip(),
            generate_masters=masters_var.get(),
            generate_vouchers=vouchers_var.get(),
            generate_combined=combined_var.get(),
            sales=sales_var.get(),
            receipts=receipts_var.get(),
            purchases=purchases_var.get(),
            payments=payments_var.get(),
            notes=notes_var.get(),
            max_data_age_days=selected_age_days(),
        )
        files = "\n".join(result.get("files", []))
        message = "Tally XML files generated successfully."
        if files:
            message += f"\n\nCreated files:\n{files}"
        messagebox.showinfo("Success", message)
    except Exception as e:
        messagebox.showerror("Error", str(e))


root = tk.Tk()
root.title("Vyapar to Tally XML Generator")
root.geometry("760x470")
root.minsize(720, 440)

input_var = tk.StringVar()
output_var = tk.StringVar()
max_age_var = tk.StringVar(value=str(MAX_DATA_AGE_DAYS))

masters_var = tk.BooleanVar(value=True)
vouchers_var = tk.BooleanVar(value=True)
combined_var = tk.BooleanVar(value=True)
sales_var = tk.BooleanVar(value=True)
receipts_var = tk.BooleanVar(value=True)
purchases_var = tk.BooleanVar(value=True)
payments_var = tk.BooleanVar(value=True)
notes_var = tk.BooleanVar(value=True)

content = ttk.Frame(root, padding=20)
content.pack(fill="both", expand=True)
content.columnconfigure(1, weight=1)

ttk.Label(content, text="Vyapar Daybook Excel").grid(row=0, column=0, sticky="w", pady=(0, 6))
ttk.Entry(content, textvariable=input_var).grid(row=0, column=1, sticky="ew", padx=(12, 8), pady=(0, 6))
ttk.Button(content, text="Browse", command=browse_input).grid(row=0, column=2, sticky="ew", pady=(0, 6))

ttk.Label(content, text="Output Folder").grid(row=1, column=0, sticky="w", pady=6)
ttk.Entry(content, textvariable=output_var).grid(row=1, column=1, sticky="ew", padx=(12, 8), pady=6)
ttk.Button(content, text="Browse", command=browse_output).grid(row=1, column=2, sticky="ew", pady=6)

ttk.Separator(content).grid(row=2, column=0, columnspan=3, sticky="ew", pady=14)

xml_frame = ttk.LabelFrame(content, text="XML Files", padding=12)
xml_frame.grid(row=3, column=0, columnspan=3, sticky="ew")
xml_frame.columnconfigure((0, 1, 2), weight=1)
ttk.Checkbutton(xml_frame, text="Masters", variable=masters_var).grid(row=0, column=0, sticky="w")
ttk.Checkbutton(xml_frame, text="Vouchers", variable=vouchers_var).grid(row=0, column=1, sticky="w")
ttk.Checkbutton(xml_frame, text="Combined", variable=combined_var).grid(row=0, column=2, sticky="w")

voucher_frame = ttk.LabelFrame(content, text="Voucher Types", padding=12)
voucher_frame.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(14, 0))
voucher_frame.columnconfigure((0, 1, 2), weight=1)
ttk.Checkbutton(voucher_frame, text="Sales", variable=sales_var).grid(row=0, column=0, sticky="w")
ttk.Checkbutton(voucher_frame, text="Receipts", variable=receipts_var).grid(row=0, column=1, sticky="w")
ttk.Checkbutton(voucher_frame, text="Purchases", variable=purchases_var).grid(row=0, column=2, sticky="w")
ttk.Checkbutton(voucher_frame, text="Payments", variable=payments_var).grid(row=1, column=0, sticky="w", pady=(8, 0))
ttk.Checkbutton(voucher_frame, text="Credit/Debit Notes", variable=notes_var).grid(row=1, column=1, sticky="w", pady=(8, 0))

rules_frame = ttk.LabelFrame(content, text="Import Rules", padding=12)
rules_frame.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(14, 0))
rules_frame.columnconfigure(1, weight=1)
ttk.Label(rules_frame, text="Allowed data age in days").grid(row=0, column=0, sticky="w")
ttk.Entry(rules_frame, textvariable=max_age_var, width=8).grid(row=0, column=1, sticky="w", padx=(12, 0))

ttk.Button(content, text="Generate XML", command=run_generate).grid(
    row=6,
    column=0,
    columnspan=3,
    pady=(24, 0),
    ipadx=28,
    ipady=8,
)

root.mainloop()
