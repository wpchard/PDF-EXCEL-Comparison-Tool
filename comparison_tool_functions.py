#functions
# Import libraries:
import pandas as pd
import numpy as np
import time
from tkinter import Tk
from tkinter.filedialog import asksaveasfilename
from comparison_tool_helper_functions import ( convert_pdf_to_excel,get_input_path_for_excel, determine_status, fix_single_column, remove_row_number_column, normalize_serial, normalize_scan_description, date_column, old_date_col, EXPECTED_COLUMNS, merge_key)

#prints header and gets file paths of the two documents to compare. handles both pdf and excel files.
def get_file_paths():
    print("FF3 TFFF Comparison Tool - Version 3.2\n")

    print("Choose two files. Pick the older file first or the program will not execute correctly.")

   
    tfff_old = get_file_location("OLD")
    tfff_new = get_file_location("NEW")    
    
    if tfff_old.lower().endswith(".pdf"):
        tfff_old = convert_pdf_to_excel(tfff_old)
    if tfff_new.lower().endswith(".pdf"):
        tfff_new = convert_pdf_to_excel(tfff_new)

    #display selected file paths
    print("\nOld file:", tfff_old)
    print("New file:", tfff_new)
    print("Reading Excel files...\n")
    return tfff_old, tfff_new


#function to find the directory to save the output file. also adds the name of the file that includes the date of creation.
def get_output_path():
    Tk().withdraw()
    print("Select location to save output Excel file:")
    date = time.strftime("%m-%d-%Y")

    return asksaveasfilename(
        defaultextension=".xlsx",
        initialfile=f"tfff_wt_updated_info_{date}.xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )

# This function reads all sheets from the given Excel file, cleans and standardizes the data, and combines it into a single DataFrame. It handles various formatting issues commonly found in TFFF reports, such as single-column imports, row number columns, and inconsistent column counts. It also applies serial number cleaning to ensure consistent formatting across all entries.
def read_and_clean_sheets(file_path):

    excel_file = pd.ExcelFile(file_path)

    dataframes = []

    for sheet_name in excel_file.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        df = (df.dropna(how='all').dropna(axis=1, how='all'))
        df.columns = range(len(df.columns))
        df = fix_single_column(df, sheet_name)
        df = remove_row_number_column(df, sheet_name)
        df = assign_columns(df, sheet_name)

        if 'SERIAL NO.' in df.columns:
            df = clean_serial_column(df, 'SERIAL NO.')

        if len(df) > 0:
            dataframes.append(df)
    if not dataframes:
        raise ValueError(f"No data found in {file_path}")
    combined = pd.concat(dataframes, ignore_index=True)
    return combined

def assign_columns(df, sheet_name):

    # Remove extra columns
    if len(df.columns) > 8:

        print(
            f"  Sheet '{sheet_name}': "
            f"Removing extra columns"
        )

        df = df.iloc[:, :8]

    # Pad missing columns
    while len(df.columns) < 8:
        df[len(df.columns)] = np.nan

    # Assign expected columns
    df.columns = EXPECTED_COLUMNS

    return df

#this function applies the normalization to the entire serial number column, removes invalid entries, and ensures we only keep rows with valid numeric serial numbers. It also removes any rows where the serial number is missing or empty after cleaning.
def clean_serial_column(df, column_name):
    df = df.copy()
    df = df[df[column_name].notna()]
    
    # Use the module-level normalize_serial function
    df[column_name] = df[column_name].apply(normalize_serial)
    
    df = df[df[column_name].notna()]
    df = df[df[column_name] != '']
    df = df[df[column_name].str.match(r'^\d+$', na=False)]
    
    return df

#create the merged dataframe by cleaning some columns and comparing the old column date and new column date
def merge_and_compare(tfff_old, tfff_new):
    # Keep ALL columns from old file for removed row tracking
    old_cols = tfff_old.copy()

    # Rename date column to avoid conflict
    old_cols.rename(columns={date_column: f"{date_column}_OLD"}, inplace=True)

    # Add suffix to other columns that exist in both files (except merge key)
    old_cols_to_suffix = [ c for c in old_cols.columns if c != merge_key and c != f"{date_column}_OLD" and c in tfff_new.columns]
    old_cols.rename(columns={c: f"{c}_OLD" for c in old_cols_to_suffix}, inplace=True)

    # Merge the new and old dataframes on the merge key (SERIAL NO.) using an outer join to keep all records from both files. This allows us to identify added, removed, and changed rows based on the presence of data in the new and old columns.
    tfff_merged = pd.merge(tfff_new,old_cols, on=merge_key,how="outer",indicator=True)

  #convert dates of all date values to datetime formatting 
    for col in [date_column, "DATE RECEIVED", old_date_col]:
        if col in tfff_merged.columns:
            tfff_merged[col] = pd.to_datetime(tfff_merged[col],format="%m/%d/%y",errors="coerce")

    # get rows that were removed from new to show in merged
    deleted_mask = tfff_merged["_merge"] == "right_only"

    for col in EXPECTED_COLUMNS:
        if col == merge_key:
            continue
        old_col = f"{col}_OLD"

        if old_col in tfff_merged.columns:
            tfff_merged.loc[deleted_mask,col] = tfff_merged.loc[deleted_mask,old_col]

    #get status
    tfff_merged["STATUS"] = tfff_merged.apply(determine_status,axis=1)

    #drop no longer needed helper columns
    tfff_merged = tfff_merged.drop(columns=[c for c in tfff_merged.columns if c.endswith("_OLD") or c == "_merge"],errors="ignore")

    # fix issues with blank or incomplete last scan descriptions
    tfff_merged = normalize_scan_description(tfff_merged)

    #Sort the sheet by date received, then serial NO, and convert date received to date format with no time visible.
    if "DATE RECEIVED" in tfff_merged.columns:
        tfff_merged = tfff_merged.sort_values(by = ["DATE RECEIVED","SERIAL NO."])
        tfff_merged["DATE RECEIVED"] = (tfff_merged["DATE RECEIVED"].dt.strftime("%m-%d-%Y"))

    # format other datetime columns
    for col in tfff_merged.columns:
        if col != "DATE RECEIVED" and pd.api.types.is_datetime64_any_dtype(tfff_merged[col]):
            tfff_merged[col] = tfff_merged[col].dt.strftime("%m-%d-%Y")

    return tfff_merged

#creates a summary of the changes in the final file and prompts the user to save the report in their desired location.
def export_report(df, output_path):

    if "STATUS" not in df.columns:
        raise ValueError("STATUS column missing from dataframe")

    # get the status totals and whcich columns the summary will be under
    status_summary = (df["STATUS"].value_counts().reset_index())
    status_summary.columns = ["DATE RECEIVED", "SERIAL NO."]

    #calculate totals 
    total_rows = len(df)
    total_changed = df["STATUS"].ne("Unchanged").sum()

    #build rows for summary and spacing 
    blank_row = pd.DataFrame([["", ""]], columns=["DATE RECEIVED", "SERIAL NO."])
    summary_title = pd.DataFrame([["SUMMARY", ""]], columns=["DATE RECEIVED", "SERIAL NO."])
    total_changed_row = pd.DataFrame([["TOTAL CHANGED", total_changed]], columns=["DATE RECEIVED", "SERIAL NO."])
    total_rows_row = pd.DataFrame([["TOTAL ROWS", total_rows]], columns=["DATE RECEIVED", "SERIAL NO."])

    #Combine everything and export 
    final_df = pd.concat([df,blank_row,summary_title,status_summary,blank_row,total_changed_row,total_rows_row], ignore_index=True)
    final_df.to_excel(output_path, sheet_name="Showing Update Info", index=False)
    print(f"\nReport exported to: {output_path}")

