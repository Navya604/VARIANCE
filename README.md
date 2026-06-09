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
