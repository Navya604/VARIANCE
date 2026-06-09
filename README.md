# VARIANCE
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys

workflow_path = None
lv_path = None


def get_output_folder():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def select_workflow():
    global workflow_path

    workflow_path = filedialog.askopenfilename(
        title="Select Workflow File",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if workflow_path:
        workflow_label.config(
            text=f"Workflow Selected:\n{os.path.basename(workflow_path)}"
        )


def select_lv():
    global lv_path

    lv_path = filedialog.askopenfilename(
        title="Select LV Report",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if lv_path:
        lv_label.config(
            text=f"LV Report Selected:\n{os.path.basename(lv_path)}"
        )


def run_extraction():

    if not workflow_path or not lv_path:
        messagebox.showerror(
            "Missing Files",
            "Please select Workflow and LV Report."
        )
        return

    try:

        # ---------------------------
        # LOAD FILES
        # ---------------------------

        lv = pd.read_excel(
            lv_path,
            sheet_name="NAV Large Variance Report",
            header=9,
            engine="openpyxl"
        )

        wf = pd.read_excel(
            workflow_path,
            sheet_name="Workflow",
            engine="openpyxl"
        )

        # ---------------------------
        # CLEAN HEADERS
        # ---------------------------

        lv.columns = lv.columns.str.strip()
        wf.columns = wf.columns.str.strip()

        # ---------------------------
        # FORMAT UCN
        # ---------------------------

        lv["Fund UCN"] = (
            lv["Fund UCN"]
            .astype(str)
            .str.replace(".0", "", regex=False)
            .str.strip()
        )

        wf["Fund UCN"] = (
            wf["Fund UCN"]
            .astype(str)
            .str.replace(".0", "", regex=False)
            .str.strip()
        )

        # ---------------------------
        # KEEP ONLY WORKFLOW FIELDS
        # ---------------------------

        wf_small = wf[
            [
                "Fund UCN",
                "CCY",
                "Reg"
            ]
        ]

        # ---------------------------
        # MERGE
        # ---------------------------

        merged = lv.merge(
            wf_small,
            on="Fund UCN",
            how="left"
        )

        # ---------------------------
        # REGION
        # ---------------------------

        merged["Region"] = (
            merged["Reg"]
            .astype(str)
            .str.upper()
        )

        # ---------------------------
        # FILTER REGION
        # ---------------------------

        merged = merged[
            merged["Region"].isin(
                ["LATAM", "NAHF"]
            )
        ]

        # ---------------------------
        # FILTER VARIANCE
        # ---------------------------

        merged = merged[
            (merged["NAV % Change"] >= 31)
            |
            (merged["NAV % Change"] <= -21)
        ]

        # ---------------------------
        # OUTPUT
        # ---------------------------

        output_df = merged[
            [
                "Family SPN",
                "Family UCN",
                "Family Name",
                "Fund SPN",
                "Fund UCN",
                "Fund Name",
                "Primary Credit Exec",
                "Sup Credit Exec",
                "Credit Contact",
                "CCY",
                "Region",
                "Starting NAV (In Thousands)",
                "Starting NAV Date",
                "Ending NAV (In Thousands)",
                "Ending NAV Date",
                "NAV % Change"
            ]
        ]

        # ---------------------------
        # SAVE
        # ---------------------------

        output_folder = get_output_folder()

        output_path = os.path.join(
            output_folder,
            "Monthly_Variance_Report.xlsx"
        )

        output_df.to_excel(
            output_path,
            index=False
        )

        messagebox.showinfo(
            "Success",
            f"Report Created Successfully!\n\n{output_path}"
        )

    except Exception as e:
        messagebox.showerror(
            "Error",
            str(e)
        )


# ==================================
# GUI
# ==================================

root = tk.Tk()
root.title("Monthly Variance Report")
root.geometry("700x450")
root.configure(bg="#F4F6F8")

title = tk.Label(
    root,
    text="Monthly Variance Report Automation",
    font=("Segoe UI", 20, "bold"),
    bg="#F4F6F8",
    fg="#1F2937"
)
title.pack(pady=20)

workflow_btn = tk.Button(
    root,
    text="Select Workflow File",
    command=select_workflow,
    font=("Segoe UI", 12, "bold"),
    bg="#2563EB",
    fg="white",
    width=25,
    height=2
)
workflow_btn.pack(pady=10)

workflow_label = tk.Label(
    root,
    text="No Workflow Selected",
    bg="#F4F6F8"
)
workflow_label.pack()

lv_btn = tk.Button(
    root,
    text="Select LV Report",
    command=select_lv,
    font=("Segoe UI", 12, "bold"),
    bg="#2563EB",
    fg="white",
    width=25,
    height=2
)
lv_btn.pack(pady=10)

lv_label = tk.Label(
    root,
    text="No LV Report Selected",
    bg="#F4F6F8"
)
lv_label.pack()

run_btn = tk.Button(
    root,
    text="Run Extraction",
    command=run_extraction,
    font=("Segoe UI", 14, "bold"),
    bg="#16A34A",
    fg="white",
    width=25,
    height=2
)
run_btn.pack(pady=25)

root.mainloop()






# ----------------------------------------
# FIX ID FORMATS
# ----------------------------------------

lv["Fund UCN"] = (
    lv["Fund UCN"]
    .astype(str)
    .str.replace(".0", "", regex=False)
    .str.strip()
    .str.zfill(10)
)

wf["Fund UCN"] = (
    wf["Fund UCN"]
    .astype(str)
    .str.replace(".0", "", regex=False)
    .str.strip()
    .str.zfill(10)
)

wf["IA UCN"] = (
    wf["IA UCN"]
    .astype(str)
    .str.replace(".0", "", regex=False)
    .str.strip()
    .str.zfill(10)
)

lv["Fund SPN"] = (
    lv["Fund SPN"]
    .astype(str)
    .str.replace(".0", "", regex=False)
    .str.strip()
)

wf["Fund SPN"] = (
    wf["Fund SPN"]
    .astype(str)
    .str.replace(".0", "", regex=False)
    .str.strip()
)

lv["Family UCN"] = (
    lv["Family UCN"]
    .astype(str)
    .str.replace(".0", "", regex=False)
    .str.strip()
    .str.zfill(10)
)

lv["Family SPN"] = (
    lv["Family SPN"]
    .astype(str)
    .str.replace(".0", "", regex=False)
    .str.strip()
)







def select_wcr():
    global wcr_path

    wcr_path = filedialog.askopenfilename(
        title="Select WCR Report",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if wcr_path:
        messagebox.showinfo(
            "Selected",
            f"WCR File Selected:\n{wcr_path}"
        )




# ---------------------------------
# WCR MERGE
# ---------------------------------

wcr["UCN (Client)"] = (
    wcr["UCN (Client)"]
    .astype(str)
    .str.replace(".0", "", regex=False)
    .str.strip()
)

merged["Fund UCN"] = (
    merged["Fund UCN"]
    .astype(str)
    .str.replace(".0", "", regex=False)
    .str.strip()
)

merged = merged.merge(
    wcr[
        [
            "UCN (Client)",
            "Credit Contact Name (Out Oblg)",
            "Prim Credit Officer",
            "Supervisory Credit Officer"
        ]
    ],
    left_on="Fund UCN",
    right_on="UCN (Client)",
    how="left"
)





tk.Button(
    root,
    text="Select WCR Report",
    command=select_wcr,
    **button_style
).pack(pady=10)
        







# ==========================================
# MODERN GUI
# ==========================================

root = tk.Tk()
root.title("Monthly NAV Variance Analyzer")
root.geometry("700x550")
root.configure(bg="#1E293B")
root.resizable(False, False)

# ------------------------------------------
# TITLE
# ------------------------------------------

title_label = tk.Label(
    root,
    text="Monthly NAV Variance Analyzer",
    font=("Segoe UI", 22, "bold"),
    bg="#1E293B",
    fg="white"
)

title_label.pack(pady=(25, 10))

subtitle_label = tk.Label(
    root,
    text="CRMO Monthly Variance Reporting Tool",
    font=("Segoe UI", 10),
    bg="#1E293B",
    fg="#CBD5E1"
)

subtitle_label.pack(pady=(0, 25))

# ------------------------------------------
# BUTTON STYLE
# ------------------------------------------

button_style = {
    "font": ("Segoe UI", 14, "bold"),
    "bg": "#2563EB",
    "fg": "white",
    "activebackground": "#1D4ED8",
    "activeforeground": "white",
    "width": 28,
    "height": 2,
    "bd": 0,
    "cursor": "hand2"
}

# ------------------------------------------
# WORKFLOW BUTTON
# ------------------------------------------

workflow_btn = tk.Button(
    root,
    text="Select Workflow",
    command=select_workflow,
    **button_style
)

workflow_btn.pack(pady=10)

# ------------------------------------------
# LV BUTTON
# ------------------------------------------

lv_btn = tk.Button(
    root,
    text="Select LV Report",
    command=select_lv,
    **button_style
)

lv_btn.pack(pady=10)

# ------------------------------------------
# WCR BUTTON
# ------------------------------------------

wcr_btn = tk.Button(
    root,
    text="Select WCR Report",
    command=select_wcr,
    **button_style
)

wcr_btn.pack(pady=10)

# ------------------------------------------
# RUN BUTTON
# ------------------------------------------

run_btn = tk.Button(
    root,
    text="Run Extraction",
    command=run_extraction,
    font=("Segoe UI", 15, "bold"),
    bg="#16A34A",
    fg="white",
    activebackground="#15803D",
    activeforeground="white",
    width=28,
    height=2,
    bd=0,
    cursor="hand2"
)

run_btn.pack(pady=30)

# ------------------------------------------
# STATUS LABEL
# ------------------------------------------

status_label = tk.Label(
    root,
    text="Ready",
    font=("Segoe UI", 11, "italic"),
    bg="#1E293B",
    fg="#94A3B8"
)

status_label.pack(pady=5)

# ------------------------------------------
# FOOTER
# ------------------------------------------

footer_label = tk.Label(
    root,
    text="JPMC CRMO Automation | Monthly Variance Reporting",
    font=("Segoe UI", 9),
    bg="#1E293B",
    fg="#64748B"
)

footer_label.pack(side="bottom", pady=15)

root.mainloop()
