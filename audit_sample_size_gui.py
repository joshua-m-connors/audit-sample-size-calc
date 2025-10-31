import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import math, csv, sys, ctypes
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

# --- DPI Awareness for Windows ---
if sys.platform.startswith("win"):
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass


# --- Detect system dark mode ---
def system_uses_dark_mode():
    try:
        if sys.platform.startswith("win"):
            import winreg
            reg = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
            key = winreg.OpenKey(reg, r"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize")
            val, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            return val == 0
    except Exception:
        return False
    return False


IS_DARK = system_uses_dark_mode()


# --- Tooltip helper ---
class CreateToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip = None
        widget.bind("<Enter>", self.show)
        widget.bind("<Leave>", self.hide)

    def show(self, _=None):
        if self.tip or not self.text:
            return
        x, y, _, _ = self.widget.bbox("insert") or (0, 0, 0, 0)
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        self.tip = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.geometry(f"+{x}+{y}")
        bg = "#333333" if IS_DARK else "#ffffe0"
        fg = "#ffffff" if IS_DARK else "#000000"
        label = tk.Label(
            tw, text=self.text, justify="left", background=bg, foreground=fg,
            relief="solid", borderwidth=1, wraplength=300, font=("Segoe UI", 9)
        )
        label.pack(ipadx=4, ipady=2)

    def hide(self, _=None):
        if self.tip:
            self.tip.destroy()
            self.tip = None


# --- Core Audit Sampling Formula ---
def calculate_sample_size(pop, conf, expected_dev=0.0, tolerable_dev=0.05):
    if tolerable_dev <= 0:
        tolerable_dev = 0.05
    conf_p = conf / 100.0
    n = math.log(1 - conf_p) / math.log(1 - tolerable_dev)
    if expected_dev > 0:
        n *= (1 + expected_dev * 2)
    if pop < 1000:
        n *= (pop / 1000)
    return max(5, math.ceil(n))


# --- Calculation Logic ---
def on_calculate():
    global last_results
    try:
        total_pop = int(pop_entry.get())
        conf = int(conf_combo.get().replace("%", ""))
        user_exp = float(exp_dev_entry.get())
        tol = float(tol_dev_entry.get())

        base_req = calculate_sample_size(total_pop, conf, user_exp, tol)
        note = ""
        adj_total = base_req
        addl_roll = 0
        observed_rate = 0
        planned_exp = user_exp
        req_total = base_req
        interim_sample_sz = 0
        exceptions = 0
        roll_pop = 0

        if rollforward_var.get():
            roll_pop = int(rollforward_entry.get())
            issues_found = (issues_var.get() == "Yes")

            interim_sample_sz = int(interim_sample_entry.get())
            exceptions = int(exceptions_count_entry.get()) if issues_found else 0

            observed_rate = (exceptions / interim_sample_sz) if interim_sample_sz > 0 else 0.0
            planned_exp = max(user_exp, observed_rate)

            req_total = calculate_sample_size(total_pop, conf, planned_exp, tol)
            roll_pop = max(0, min(roll_pop, total_pop))
            prop_min = math.ceil(req_total * (roll_pop / total_pop)) if total_pop > 0 else 0

            addl_roll = max(0, req_total - interim_sample_sz, prop_min)
            adj_total = interim_sample_sz + addl_roll

            if issues_found and exceptions > 0:
                note = (
                    "Note: Interim exceptions increased the expected deviation used for planning. "
                    "Consider the nature and cause of deviations and whether additional procedures "
                    "are required (AU-C 530 / PCAOB AS 2315)."
                )

        base_result.set(f"Base sample size (no rollforward): {base_req}")
        replan_result.set(f"Replanned full-year required sample: {req_total}")
        addl_result.set(f"Additional rollforward sample required: {addl_roll}")
        adj_total_result.set(f"Adjusted total sample (interim + rollforward): {adj_total}")
        note_result.set(note)

        last_results = {
            "Population": total_pop,
            "Confidence": conf,
            "User Expected Deviation": user_exp,
            "Observed Interim Deviation": round(observed_rate, 4) if rollforward_var.get() else "N/A",
            "Planned Expected Deviation": round(planned_exp, 4) if rollforward_var.get() else user_exp,
            "Tolerable Deviation": tol,
            "Base Required Sample": base_req,
            "Rollforward Enabled": rollforward_var.get(),
            "Issues Found": issues_var.get() if rollforward_var.get() else "N/A",
            "Interim Sample Size": interim_sample_sz if rollforward_var.get() else "N/A",
            "Interim Exceptions": exceptions if rollforward_var.get() else "N/A",
            "Rollforward Population": roll_pop if rollforward_var.get() else "N/A",
            "Replanned Full-Year Required Sample": req_total,
            "Additional Rollforward Sample": addl_roll,
            "Adjusted Total Sample": adj_total,
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }

    except Exception as e:
        messagebox.showerror("Error", f"Invalid input.\n\n{e}")


# --- Rollforward Field Control ---
def toggle_rollforward_fields():
    if rollforward_var.get():
        rollforward_frame.grid(row=2, column=0, pady=10, sticky="nsew")
    else:
        rollforward_frame.grid_remove()


def toggle_exceptions_field(event=None):
    if issues_var.get() == "Yes":
        exceptions_label.grid(row=2, column=0, sticky="w", pady=3)
        exceptions_count_entry.grid(row=2, column=1, sticky="ew", pady=3)
    else:
        exceptions_label.grid_remove()
        exceptions_count_entry.grid_remove()


def clear_fields():
    pop_entry.delete(0, tk.END); pop_entry.insert(0, "5000")
    exp_dev_entry.delete(0, tk.END); exp_dev_entry.insert(0, "0.00")
    tol_dev_entry.delete(0, tk.END); tol_dev_entry.insert(0, "0.05")
    conf_combo.set("90%")
    rollforward_var.set(False)
    toggle_rollforward_fields()
    for v in [base_result, replan_result, addl_result, adj_total_result, note_result]:
        v.set("")


# --- Export Results ---
def export_results():
    if not last_results:
        messagebox.showinfo("No Results", "Please calculate a sample first.")
        return
    choice = messagebox.askquestion("Export", "Export as PDF?\n(Click 'No' for CSV)")
    try:
        if choice == "yes":
            path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
            if not path:
                return
            generate_pdf(path, last_results)
        else:
            path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
            if not path:
                return
            with open(path, "w", newline="") as f:
                writer = csv.writer(f)
                for k, v in last_results.items():
                    writer.writerow([k, v])
        messagebox.showinfo("Saved", f"Results exported to:\n{path}")
    except Exception as e:
        messagebox.showerror("Export Error", str(e))


def generate_pdf(path, data):
    doc = SimpleDocTemplate(path, pagesize=letter)
    styles = getSampleStyleSheet()
    elements = [Paragraph("Audit Sample Size Summary", styles["Title"]), Spacer(1, 12)]
    rows = [
        ["Category", "Value"],
        ["Population Size", data["Population"]],
        ["Confidence Level", f"{data['Confidence']}%"],
        ["Expected Deviation", data["Planned Expected Deviation"]],
        ["Tolerable Deviation", data["Tolerable Deviation"]],
        ["Base Sample Size", data["Base Required Sample"]],
        ["Replanned Required Sample", data["Replanned Full-Year Required Sample"]],
        ["Additional Rollforward Sample", data["Additional Rollforward Sample"]],
        ["Adjusted Total Sample", data["Adjusted Total Sample"]],
    ]
    tbl = Table(rows, colWidths=[220, 220])
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1B4F72")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.lightgrey]),
    ]))
    elements.append(tbl)
    elements.append(Spacer(1, 12))
    elements.append(Paragraph("Generated: " + data["Timestamp"], styles["Normal"]))
    doc.build(elements)


# --- GUI Setup ---
bg = "#2B2B2B" if IS_DARK else "#F4F6F7"
fg = "#EAECEE" if IS_DARK else "#000"
accent = "#00BCD4" if IS_DARK else "#1B4F72"

root = tk.Tk()
root.title("Audit Sample Size Calculator")
root.configure(bg=bg)
root.minsize(620, 680)

style = ttk.Style(root)
style.theme_use("clam")
style.configure("TLabel", background=bg, foreground=fg, font=("Segoe UI", 10))
style.configure("TButton", font=("Segoe UI", 10, "bold"))
style.configure("TLabelframe", background=bg, foreground=fg, font=("Segoe UI", 10, "bold"))
style.configure("TCheckbutton", background=bg, foreground=fg, font=("Segoe UI", 10))

tk.Label(root, text="Audit Sample Size Calculator", bg=accent, fg="white",
         font=("Segoe UI Semibold", 15), pady=12).grid(row=0, column=0, sticky="ew")

# --- Input Frame ---
frame = ttk.Frame(root, padding=20)
frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)

ttk.Label(frame, text="Population Size:").grid(row=0, column=0, sticky="w", pady=5)
pop_entry = ttk.Entry(frame); pop_entry.insert(0, "5000")
pop_entry.grid(row=0, column=1, sticky="ew", pady=5)

ttk.Label(frame, text="Confidence Level:").grid(row=1, column=0, sticky="w", pady=5)
conf_combo = ttk.Combobox(frame, values=["90%", "95%", "99%"], state="readonly")
conf_combo.set("90%")
conf_combo.grid(row=1, column=1, sticky="ew", pady=5)

ttk.Label(frame, text="Expected Deviation Rate:").grid(row=2, column=0, sticky="w", pady=5)
exp_dev_entry = ttk.Entry(frame); exp_dev_entry.insert(0, "0.00")
exp_dev_entry.grid(row=2, column=1, sticky="ew", pady=5)

ttk.Label(frame, text="Tolerable Deviation Rate:").grid(row=3, column=0, sticky="w", pady=5)
tol_dev_entry = ttk.Entry(frame); tol_dev_entry.insert(0, "0.05")
tol_dev_entry.grid(row=3, column=1, sticky="ew", pady=5)

rollforward_var = tk.BooleanVar()
ttk.Checkbutton(frame, text="Include Rollforward Testing", variable=rollforward_var,
                command=toggle_rollforward_fields).grid(row=4, column=0, columnspan=2, sticky="w", pady=10)
frame.columnconfigure(1, weight=1)

# --- Rollforward Frame ---
rollforward_frame = ttk.LabelFrame(root, text="Rollforward Details", padding=10)
ttk.Label(rollforward_frame, text="Issues found in interim testing?").grid(row=0, column=0, sticky="w", pady=3)
issues_var = tk.StringVar(value="No")
issues_combo = ttk.Combobox(rollforward_frame, textvariable=issues_var, values=["Yes", "No"], state="readonly")
issues_combo.grid(row=0, column=1, sticky="ew", pady=3)
issues_combo.bind("<<ComboboxSelected>>", toggle_exceptions_field)

ttk.Label(rollforward_frame, text="Interim Sample Size Tested:").grid(row=1, column=0, sticky="w", pady=3)
interim_sample_entry = ttk.Entry(rollforward_frame); interim_sample_entry.insert(0, "45")
interim_sample_entry.grid(row=1, column=1, sticky="ew", pady=3)

exceptions_label = ttk.Label(rollforward_frame, text="Exceptions Found at Interim (Count):")
exceptions_count_entry = ttk.Entry(rollforward_frame); exceptions_count_entry.insert(0, "0")

ttk.Label(rollforward_frame, text="Rollforward Population Size:").grid(row=3, column=0, sticky="w", pady=3)
rollforward_entry = ttk.Entry(rollforward_frame); rollforward_entry.insert(0, "1000")
rollforward_entry.grid(row=3, column=1, sticky="ew", pady=3)
rollforward_frame.columnconfigure(1, weight=1)
toggle_exceptions_field()
rollforward_frame.grid_remove()

# --- Buttons ---
button_frame = ttk.Frame(root, padding=15)
button_frame.grid(row=3, column=0, sticky="ew")
for i in range(3):
    button_frame.columnconfigure(i, weight=1)
ttk.Button(button_frame, text="Calculate", command=on_calculate).grid(row=0, column=0, padx=10, sticky="ew")
ttk.Button(button_frame, text="Clear", command=clear_fields).grid(row=0, column=1, padx=10, sticky="ew")
ttk.Button(button_frame, text="Export", command=export_results).grid(row=0, column=2, padx=10, sticky="ew")

# --- Results Summary ---
result_frame = ttk.LabelFrame(root, text="Results Summary", padding=10)
result_frame.grid(row=4, column=0, sticky="ew", padx=20, pady=10)
result_frame.columnconfigure(0, weight=1)
base_result = tk.StringVar(value="")
replan_result = tk.StringVar(value="")
addl_result = tk.StringVar(value="")
adj_total_result = tk.StringVar(value="")
note_result = tk.StringVar(value="")

ttk.Label(result_frame, textvariable=base_result, font=("Segoe UI", 10), foreground=fg, background=bg).grid(sticky="w", pady=2)
ttk.Label(result_frame, textvariable=replan_result, font=("Segoe UI", 10), foreground=fg, background=bg).grid(sticky="w", pady=2)
ttk.Label(result_frame, textvariable=addl_result, font=("Segoe UI", 10), foreground=fg, background=bg).grid(sticky="w", pady=2)
ttk.Label(result_frame, textvariable=adj_total_result, font=("Segoe UI", 10, "bold"), foreground=accent, background=bg).grid(sticky="w", pady=4)
ttk.Label(result_frame, textvariable=note_result, font=("Segoe UI", 9, "italic"),
          foreground="#5A6D7A", background=bg, wraplength=560, justify="left").grid(sticky="w", pady=4)

# --- Help Text ---
help_text = (
    "Typical Audit Sampling Assumptions:\n"
    "• Confidence Level: 90% (≈10% risk of overreliance)\n"
    "• Expected Deviation: 0% (no known errors)\n"
    "• Tolerable Deviation: 5% (maximum acceptable error rate)\n"
    "These assumptions produce a sample size of approximately 44 items under a binomial model.\n\n"
    "Rollforward Testing:\n"
    "If interim testing is performed, this calculator uses your observed interim exceptions to "
    "update the expected deviation and re-plan the total required sample for the full period. "
    "It then ensures at least proportional coverage of the rollforward population. If exceptions "
    "were identified, professional judgment should be applied to determine if additional procedures "
    "such as inquiry, re-testing, or control redesign are warranted (AU-C 530 / PCAOB AS 2315)."
)
ttk.Label(root, text=help_text, background=bg, foreground=fg, justify="left",
          font=("Segoe UI", 9, "italic"), wraplength=560).grid(row=5, column=0, sticky="w", padx=20, pady=5)

last_results = {}
root.mainloop()
