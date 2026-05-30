import sys
import tabulate
import datetime as dt
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from utils import load_holidays, sub_days, sub_court_days, open_doc, find_and_replace

calendar = []

# Load holidays into utils
try:
    load_holidays()
except FileNotFoundError:
    messagebox.showwarning(
            "Missing File",
            "holidays.csv was not found."
        )

# Generate calendar
def generate_calendar():
    output_box.config(state=tk.NORMAL)
    output_box.delete("1.0", tk.END)

    trial_date_str = trial_date_entry.get().strip()
    msc_date_str = msc_date_entry.get().strip()

    if not trial_date_str:
        messagebox.showerror("Input Error", "Please enter a trial date.")
        return

    try:
        trial_date = dt.datetime.strptime(trial_date_str, "%m/%d/%y")
    except ValueError:
        messagebox.showerror(
            "Date Format Error",
            "Trial date must be in MM/DD/YY format."
        )
        return

    msc_date = None

    if msc_date_str:
        try:
            msc_date = dt.datetime.strptime(msc_date_str, "%m/%d/%y")
        except ValueError:
            messagebox.showerror(
                "Date Format Error",
                "Settlement/Issue Conference date must be in MM/DD/YY format."
            )
            return

    deadlines = []

    if msc_date is not None:
        deadlines.append([
            "Settlement/Issue Conference Statement Due",
            sub_days(msc_date, 5)
        ])

    deadlines.append([
        "Pre-Trial Report Due to Insurance Carrier",
        sub_days(trial_date, 45)
    ])

    deadlines.append([
        "LDFS Motion for Summary Judgment via Mail",
        sub_days(trial_date, 110)
    ])

    deadlines.append([
        "LDFS Motion for Summary Judgment via Hand Delivery",
        sub_days(trial_date, 105)
    ])

    deadlines.append([
        "LD to Serve Written Discovery Requests (Soft Deadline)",
        sub_days(trial_date, 100)
    ])

    deadlines.append([
        "LD to Serve Written Discovery Requests (Hard Deadline) via Mail",
        sub_days(trial_date, 65)
    ])

    deadlines.append([
        "LD to Serve Written Discovery Requests (Hard Deadline) via Hand Delivery",
        sub_days(trial_date, 60)
    ])

    deadlines.append([
        "LD to Request Expert Disclosure",
        sub_days(trial_date, 70)
    ])

    deadlines.append([
        "LD to Exchange Expert Witness",
        sub_days(trial_date, 50)
    ])

    deadlines.append([
        "LDFS Discovery Motions via Mail",
        sub_days(sub_court_days(sub_days(trial_date, 15), 16), 5)
    ])

    deadlines.append([
        "LDFS Discovery Motions via Hand Delivery",
        sub_court_days(sub_days(trial_date, 15), 16)
    ])

    deadlines.append([
        "LDFS Motions Re Experts via Mail",
        sub_days(sub_court_days(sub_days(trial_date, 10), 16), 5)
    ])

    deadlines.append([
        "LDFS Motions Re Experts via Hand Delivery",
        sub_court_days(sub_days(trial_date, 10), 16)
    ])
    
    deadlines.append([
        "LD to Post Jury Fees",
        "TBD"
    ])

    deadlines.append([
        "Discovery Cut-off/LD to hear Motion for Summary Judgement",
        sub_days(trial_date, 30)
    ])

    deadlines.append([
        "LD to Issue Subpenas Duces Tecum",
        sub_days(trial_date, 20)
    ])

    deadlines.append([
        "LD to Serve §1987 Notice to Appear with Documents via Mail",
        sub_days(trial_date, 25)
    ])

    deadlines.append([
        "LD to Serve §1987 Notice to Appear with Documents via Hand Delivery",
        sub_days(trial_date, 20)
    ])

    deadlines.append([
        "LD to Serve §1987 Notice to Appear without Documents via Mail",
        sub_days(trial_date, 15)
    ])

    deadlines.append([
        "LD to Serve §1987 Notice to Appear without Documents via Hand Delivery",
        sub_days(trial_date, 10)
    ])

    deadlines.append([
        "LD to Conduct Expert Discovery/LD to Hear Discovery Motions",
        sub_days(trial_date, 15)
    ])

    deadlines.append([
        "LD to Hear Motions Re Experts/LD to Serve C.C.P.§998",
        sub_days(trial_date, 10)
    ])
    
    deadlines.append([
        """LDFS Trial Documents: Motions in Limine/Jury Instructions,"""
        """\nLists of Exhibits and Witnesses and Voir Dire""",
        "TBD"
    ])

    for deadline in deadlines:
        calendar.append(deadline)
        if type(deadline[1]) is dt.datetime:
            deadline[1] = deadline[1].strftime("%m/%d/%y")

    table = tabulate.tabulate(deadlines, headers=["Event", "Date"], tablefmt="grid")

    output_box.insert(tk.END, table)
    output_box.config(state=tk.DISABLED)

# Clear all fields
def clear_fields():
    calendar.clear()
    case_number_entry.delete(0, tk.END)
    trial_date_entry.delete(0, tk.END)
    msc_date_entry.delete(0, tk.END)
    output_box.config(state=tk.NORMAL)
    output_box.delete("1.0", tk.END)
    output_box.config(state=tk.DISABLED)

# Export output to TXT file
def export_output_txt(file_path, content):
    content = output_box.get("1.0", tk.END).strip()
    with open(file_path, "w", encoding="utf-8") as f:
        f.write(content)
    return True

# Export output to DOCX file based on template
def export_output_docx(file_path):
    doc = open_doc()
    msc = 1 if calendar[0][0].startswith("Settlement/Issue Conference Statement Due") else 0
    for i in range(0, len(calendar) - msc):
        find_and_replace(doc, "Date" + str(i - msc + 1), calendar[i][1])
    doc.save(file_path)
    return True

# Save As dialog for TXT or DOCX
def save_as_output():
    case_number = case_number_entry.get().strip()
    trial_date = trial_date_entry.get().strip().replace('/', '-')
    content = output_box.get("1.0", tk.END).strip()
    
    if not content:
        messagebox.showwarning("No Output", "There is no output to export.")
        return

    default_name = "Trial Calendar - " + trial_date
    if case_number:
        default_name = case_number + ' ' + default_name

    file_path = filedialog.asksaveasfilename(
        initialdir=sys.path[0],
        initialfile=default_name,
        defaultextension=".docx",
        filetypes=[("Word Files", "*.docx"), ("Text Files", "*.txt")],
        title="Save Calendar Output"
    )

    if not file_path:
        messagebox.showinfo("Export Canceled", "Export operation was canceled.")
        return

    lower_path = file_path.lower()
    if lower_path.endswith('.docx'):
        success = export_output_docx(file_path)
    else:
        success = export_output_txt(file_path, content)

    if success:
        messagebox.showinfo("Export Successful", "Output exported successfully.")
    else:
        messagebox.showerror("Export Failed", "An error occurred during export.")

# GUI Setup
root = tk.Tk()
root.title("Trial Calendar Generator")
root.geometry("1000x1000")

# Input Frame
input_frame = tk.Frame(root)
input_frame.pack(pady=10)

#Case Number
case_number_label = tk.Label(
    input_frame,
    text="Case Number (OPTIONAL):"
)
case_number_label.grid(row=0, column=0, padx=5, pady=5, sticky="e")

case_number_entry = tk.Entry(input_frame, width=20)
case_number_entry.grid(row=0, column=1, padx=5, pady=5)

# Trial Date
trial_date_label = tk.Label(
    input_frame,
    text="Trial Date (MM/DD/YY):"
)
trial_date_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")

trial_date_entry = tk.Entry(input_frame, width=20)
trial_date_entry.grid(row=1, column=1, padx=5, pady=5)

# MSC Date
msc_date_label = tk.Label(
    input_frame,
    text="MSC Date (OPTIONAL):"
)
msc_date_label.grid(row=2, column=0, padx=5, pady=5, sticky="e")

msc_date_entry = tk.Entry(input_frame, width=20)
msc_date_entry.grid(row=2, column=1, padx=5, pady=5)

# Button Frame
button_frame = tk.Frame(root)
button_frame.pack(pady=10)

generate_button = tk.Button(
    button_frame,
    text="Generate",
    width=10,
    command=generate_calendar
)
generate_button.grid(row=0, column=0, padx=10)

clear_button = tk.Button(
    button_frame,
    text="Clear",
    width=10,
    command=clear_fields
)
clear_button.grid(row=0, column=1, padx=10)

save_as_button = tk.Button(
    button_frame,
    text="Save As...",
    width=10,
    command=save_as_output
)
save_as_button.grid(row=0, column=2, padx=10)

# Output Box
output_box = scrolledtext.ScrolledText(
    root,
    wrap=tk.WORD,
    width=120,
    height=30,
    font=("Courier New", 10)
)
output_box.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
output_box.config(state=tk.DISABLED)

# Run App
root.mainloop()