import os
import pandas as pd
import xlsxwriter


directory_name = 'TOY2_Remaining1\TOY2_Remaining1'
final_file = 'TOY2_M1+R1+R0_Final ForSeg_AfterDup_Finalized.csv'
segmented_file = 'TOY2_Subsample@0.2_Point Cloud Segmentation Based on Seed 1.csv'
output_file = 'TOY2_Report.xlsx'

cwd = os.getcwd()
directory = os.path.join(cwd, directory_name)

files = os.listdir(directory)
csv_files = [file for file in files if file.endswith('.csv')]
dataframes = [pd.read_csv(os.path.join(directory, file)) for file in csv_files]
    
merged = pd.concat(dataframes)
merged['Tree ID'] = range(1, len(merged) + 1)

merged[['X', 'Y', 'Z', 'DBH']] = merged[['X', 'Y', 'Z', 'DBH']].astype(float)

def create_merged_duplicate(merged, adjustment):
    merged_dup = merged.copy()
    for column, value in adjustment.items():
        merged_dup[column] = merged_dup[column] + value
    return merged_dup

# Adjustments for each duplicate
adjustments = [
    {'X': 0.001},  
    {'X': -0.001}, 
    {'Y': 0.001}, 
    {'Y': -0.001}
]

# Creating duplicates with adjustments
merged_dup1 = create_merged_duplicate(merged, adjustments[0])
merged_dup2 = create_merged_duplicate(merged, adjustments[1])
merged_dup3 = create_merged_duplicate(merged, adjustments[2])
merged_dup4 = create_merged_duplicate(merged, adjustments[3])

merged = pd.concat([merged, merged_dup1, merged_dup2, merged_dup3, merged_dup4], ignore_index=True)
merged[['X','Y','Z','DBH']] = merged[['X','Y','Z','DBH']].round(3)

merged = merged.rename(columns={'X': 'Trunk X', 'Y': 'Trunk Y', 'Z': 'Trunk Z', 'DBH': 'Trunk Diameter'})

final = pd.read_csv(final_file)
final = final.rename(columns={'TreeID': 'Diameter_ID', 'TreeLocationX': 'Trunk X', 'TreeLocationY': 'Trunk Y', 'TreeLocationZ': 'Trunk Z', 'DBH': 'Trunk Diameter'})

segmented = pd.read_csv(segmented_file)
segmented = segmented.rename(columns={'TreeID': 'Tree_Seg_Tree_ID', 'TreeLocationX': 'Canopy X', 'TreeLocationY': 'Canopy Y', 'TreeLocationZ': 'Canopy Z'})

final_segmented = pd.merge(final, segmented, left_on='Diameter_ID', right_on='OldID', how='inner')
final_segmented = final_segmented.drop(columns=['OldID'])

final_segmented[['Trunk X', 'Trunk Y', 'Trunk Z', 'Trunk Diameter']] = final_segmented[['Trunk X', 'Trunk Y', 'Trunk Z', 'Trunk Diameter']].astype(float)
final_segmented[['Trunk X', 'Trunk Y', 'Trunk Z', 'Trunk Diameter']] = final_segmented[['Trunk X', 'Trunk Y', 'Trunk Z', 'Trunk Diameter']].round(3)

temp_df = pd.merge(final_segmented, merged, on=['Trunk X', 'Trunk Y', 'Trunk Z', 'Trunk Diameter'], how="left", indicator=True)

temp_df['Origin'] = 'M0'
temp_df.loc[temp_df['_merge'] == 'both', 'Origin'] = temp_df.apply(lambda row: 'R1' if row['Trunk Diameter'] != 0 else 'R0', axis=1)
temp_df.loc[(temp_df['_merge'] == 'left_only') & (temp_df['Trunk Diameter'] != 0), 'Origin'] = 'M1'

final_segmented['Origin'] = temp_df['Origin']

print(final_segmented['Origin'].value_counts())

with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    final_segmented.to_excel(writer, index=False, sheet_name='Final_Segmented')
    
    workbook = writer.book
    worksheet = writer.sheets['Final_Segmented']

    for i, col in enumerate(final_segmented.columns):
        worksheet.set_column(i, i, len(col) + 2)
    