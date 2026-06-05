#TFFF Comparison Tool - Version 3.3
# Reads TFFF reports directly from Excel and/ or pdf files and compares them to identify and summarize changes.

from functions_main import (export_report, get_output_path, get_file_paths, read_and_clean_sheets, merge_and_compare)

def main():
    print("FF3 TFFF Comparison Tool - Version 3.3\n")
    print("Choose two files. Pick the older file first or the program will not execute correctly.")
    
    #gets file paths for the two documents to compare.
    tfff_old, tfff_new = get_file_paths()
    
    #read and clean sheets to add information to the columns and standardize the data for comparison. 
    # This will handle multi-sheet files and ensure all data is in a consistent format for merging and analysis.
    tfff_old = read_and_clean_sheets(tfff_old)
    tfff_new = read_and_clean_sheets(tfff_new)
    
    # merge the two dataframes and compare them to identify changes, additions, and removals. 
    # The result will be a new dataframe that includes a STATUS column indicating the type of change for each row.
    tfff_merged = merge_and_compare(tfff_old, tfff_new)
    
    #ask user where to save the output file and store it as excel_file_name. If user cancels, exit the program.        
    excel_file_name = get_output_path()
    if not excel_file_name:
        print("No save location selected. Exiting.")
        exit()
    #save the merged dataframe and add summary of the changes to the new excel file.
    export_report(tfff_merged, excel_file_name)
      
    input("\nProcessing complete. Press Enter to exit...")
    
if __name__ == "__main__":
    main()