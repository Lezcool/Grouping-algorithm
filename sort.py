import sys
import pandas as pd
import numpy as np

def sort(file_path,group_size):
    file = pd.read_excel(file_path)
    df2 = file.drop(file[file["# Responses"] == 0.0].index)
    New_list = []
    new_row = None
    for row in range(df2.shape[0]):
        if pd.isnull(df2.iloc[row,0]):
            new_row.append(df2.iloc[row,7])
        else:
            if new_row != None:
                New_list.append(new_row)
            new_row = [df2.iloc[row,0]]
            
    results_df = pd.DataFrame(New_list)
    results_df.columns = ["Name","Study line","Type"]
    results_df["Group"] = -1
    results_df["type number"] = -1

    type_list = list(results_df["Type"].unique())

    for row,type in enumerate(results_df["Type"]):
        results_df.iloc[row,4] = type_list.index(type)+1

    results_sorted = results_df.sort_values(by="type number")
    i = 0
    group_size = round(results_sorted.shape[0]/int(group_size))
    for row in range(results_sorted.shape[0]):
        i = i + 1 if i <= group_size else 1
        results_sorted.iloc[row,3] = i
    final_results = results_sorted.sort_values(by="Group")
    return final_results.to_excel("results.xlsx")

if __name__ == "__main__":
    file_path = sys.argv[0]
    group_size = sys.argv[1]
    sort(file_path,group_size)
