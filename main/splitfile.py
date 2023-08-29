import os
import glob

from tqdm import tqdm
import pandas as pd

class SplitFile:

    def __init__(self):
        pass

# -------------------------------------------------------------------------- #

    def splitExcelFile(self, file_path:str, n_output_file:int) -> None:

        excel_files = glob.glob(os.path.join("data", "*.xlsx"))

        if not excel_files:
            print("No Excel files found in the 'data' directory.")
            return
        
        latest_excel_file = max(excel_files, key=os.path.getmtime)  
        file = pd.read_excel(latest_excel_file)

        rows_per_file = file.shape[0] // output
        remainder_rows = file.shape[0] % output

        output_dir = "OUTPUT"
        if not os.path.exists(output_dir):
            os.mkdir(output_dir)  

        start_idx = 0

        for i in range(output):

            end_idx = start_idx + rows_per_file

            if i < remainder_rows:
                end_idx += 1

            data = file[start_idx:end_idx]

            file_name = os.path.splitext(os.path.basename(latest_excel_file))[0]
            output_file_path = os.path.join(output_dir, f"{file_name}_split_{i+1}.xlsx")
            data.to_excel(output_file_path)

            print(f"Saved {output_file_path}")

        print("Split file Successful")

# -------------------------------------------------------------------------- #

    def mergeFile(self):

        output_file = glob.glob(os.path.join("merge_file", "*.xlsx"))

        excel_lst = []

        for file in tqdm(range(len(output_file))):

            output = pd.read_excel(output_file[file])

            excel_lst.append(output)

        combined = pd.concat(excel_lst, axis=0)

        output_file_path = os.path.join('merge_file', 'combined_file.xlsx')

        combined.to_excel(output_file_path, index=False)

        print("Merge file Successful")

# -------------------------------------------------------------------------- #

    def read_excel(self):
        pass
    
# -------------------------------------------------------------------------- #

    def read_config(self):
        pass

# -------------------------------------------------------------------------- #