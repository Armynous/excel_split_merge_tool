import os
import glob
import pandas as pd

class SplitFile:

    def __init__(self):
        pass

# -------------------------------------------------------------------------- #

    def splitLatestExcelFile(self, output: int) -> None:
        excel_files = glob.glob(os.path.join("data", "*.xlsx"))

        if not excel_files:
            print("No Excel files found in the 'data' directory.")
            return
        
        latest_excel_file = max(excel_files, key=os.path.getmtime)  # Get the latest modified Excel file
        file = pd.read_excel(latest_excel_file)

        rows_per_file = file.shape[0] // output
        remainder_rows = file.shape[0] % output

        output_dir = "result"
        if not os.path.exists(output_dir):
            os.mkdir(output_dir)  # Create the "result" directory if it doesn't exist

        start_idx = 0

        for i in range(output):

            end_idx = start_idx + rows_per_file

            if i < remainder_rows:
                end_idx += 1

            data = file[start_idx:end_idx]

            # Create and save the split file in the "result" directory
            file_name = os.path.splitext(os.path.basename(latest_excel_file))[0]
            output_file_path = os.path.join(output_dir, f"{file_name}_split_{i+1}.xlsx")
            data.to_excel(output_file_path)

            print(f"Saved {output_file_path}")

            start_idx = end_idx

# -------------------------------------------------------------------------- #

splitter = SplitFile()
splitter.splitLatestExcelFile(7)  # Replace with desired output count
