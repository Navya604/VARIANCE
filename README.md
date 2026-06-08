# VARIANCE
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys

# =====================================================
# SAVE OUTPUT IN SAME FOLDER AS EXE
# =====================================================

def get_output_folder():

    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)

    return os.path.dirname(os.path.abspath(__file__))


workflow_path = None
lv_path = None

# =====================================================
# SELECT WORKFLOW
# =====================================================

def select_workflow():

    global workflow_path

    workflow_path = filedialog.askopenfilename(
        title="Select Workflow.xlsx",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if workflow_path:
        messagebox.showinfo(
            "Selected",
            f"Workflow Selected:\n{workflow_path}"
        )

# =====================================================
# SELECT LV REPORT
# =====================================================

def select_lv():

    global lv_path

    lv_path = filedialog.askopenfilename(
        title="Select NAV Large Variance Report",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if lv_path:
        messagebox.showinfo(
            "Selected",
            f"LV File Selected:\n{lv_path}"
        )

# =====================================================
# RUN EXTRACTION
# =====================================================

def run_extraction():

    global workflow_path, lv_path

    if not workflow_path or not lv_path:

        messagebox.showerror(
            "Error",
            "Please select BOTH Workflow and LV files."
        )

        return

    try:

        # -----------------------------------------
        # LOAD LV
        # -----------------------------------------

        lv = pd.read_excel(
            lv_path,
            sheet_name="NAV Large Variance Report",
            header=9,
            engine="openpyxl"
        )

        # -----------------------------------------
        # LOAD WORKFLOW
        # -----------------------------------------

        wf = pd.read_excel(
            workflow_path,
            sheet_name="Workflow",
            engine="openpyxl"
        )

        # -----------------------------------------
        # CLEAN COLUMN NAMES
        # -----------------------------------------

        lv.columns = lv.columns.str.strip()
        wf.columns = wf.columns.str.strip()

        # -----------------------------------------
        # CHECK REQUIRED COLUMNS
        # -----------------------------------------

        if "Fund UCN" not in lv.columns:
            messagebox.showerror(
                "Error",
                "'Fund UCN' missing in LV file."
            )
            return

        if "Fund UCN" not in wf.columns:
            messagebox.showerror(
                "Error",
                "'Fund UCN' missing in Workflow."
            )
            return

        # -----------------------------------------
        # FIX UCN FORMAT
        # -----------------------------------------

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

        # -----------------------------------------
        # MERGE WORKFLOW
        # -----------------------------------------

        merged = lv.merge(
            wf,
            on="Fund UCN",
            how="left"
        )

        # -----------------------------------------
        # REGION
        # -----------------------------------------

        merged["Region"] = (
            merged["Reg"]
            .astype(str)
            .str.upper()
        )

        # -----------------------------------------
        # FILTER REGION
        # -----------------------------------------

        region_filtered = merged[
            merged["Region"].isin(["LATAM", "NAHF"])
        ]

        # -----------------------------------------
        # CHECK NAV CHANGE COLUMN
        # -----------------------------------------

        if "NAV % Change" not in region_filtered.columns:

            messagebox.showerror(
                "Error",
                "'NAV % Change' missing."
            )

            return

        # -----------------------------------------
        # YOUR CURRENT THRESHOLD
        # -----------------------------------------

        final_data = region_filtered[
            (region_filtered["NAV % Change"] >= 31)
            |
            (region_filtered["NAV % Change"] <= -21)
        ]

        # -----------------------------------------
        # OUTPUT COLUMNS
        # -----------------------------------------

        output_df = pd.DataFrame()

        output_df["IA UCN"] = final_data["IA UCN"]
        output_df["IA Name"] = final_data["IA Name"]

        output_df["FUND SPN"] = final_data["FUND SPN"]
        output_df["Fund UCN"] = final_data["Fund UCN"]
        output_df["Fund Name"] = final_data["Fund Name"]

        output_df["CRU"] = final_data["CRU"]

        output_df["Region"] = final_data["Region"]

        output_df["Prim Credit Exec"] = final_data["Prim Credit Exec"]
        output_df["Sup Credit Exec"] = final_data["Sup Credit Exec"]
        output_df["Credit Contact"] = final_data["Credit Contact"]

        output_df["Starting NAV Currency"] = final_data["Starting NAV Currency"]

        output_df["Starting NAV (Thousands)"] = final_data["Starting NAV (In Thousands)"]
        output_df["Starting NAV Date"] = final_data["Starting NAV Date"]

        output_df["Ending NAV (Thousands)"] = final_data["Ending NAV (In Thousands)"]
        output_df["Ending NAV Date"] = final_data["Ending NAV Date"]

        output_df["NAV % Change"] = final_data["NAV % Change"]

        # -----------------------------------------
        # SAVE OUTPUT
        # -----------------------------------------

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
            f"Monthly Variance Report Generated!\n\nSaved at:\n{output_path}"
        )

    except Exception as e:

        messagebox.showerror(
            "Processing Error",
            str(e)
        )

# =====================================================
# GUI
# =====================================================

root = tk.Tk()

root.title("Monthly Variance Report Tool")
root.geometry("520x360")
root.config(bg="gray25")

button_style = {
    "font": ("Arial", 16, "bold"),
    "bg": "black",
    "fg": "white",
    "activebackground": "gray20",
    "activeforeground": "white",
    "width": 20,
    "height": 2
}

tk.Button(
    root,
    text="Select Workflow",
    command=select_workflow,
    **button_style
).pack(pady=20)

tk.Button(
    root,
    text="Select LV Report",
    command=select_lv,
    **button_style
).pack(pady=10)

tk.Button(
    root,
    text="Run Extraction",
    command=run_extraction,
    **button_style
).pack(pady=20)

root.mainloop()
