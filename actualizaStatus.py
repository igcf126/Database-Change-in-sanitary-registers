import pandas as pd

# Load the first Excel file
file1 = r'C:\Users\Dell\Dropbox\SEGUIMIENTO DEPÓSITOS\depósitos.xlsx'  # Replace with the actual path
df1 = pd.read_excel(file1, sheet_name='2021')  # Change sheet_name as needed
df1_2 = pd.read_excel(file1, sheet_name='2022')  # Change sheet_name as needed
df1_3 = pd.read_excel(file1, sheet_name='2023')  # Change sheet_name as needed

# Load the second Excel fileim
file2 = r'C:\Users\Dell\OneDrive\Artes Ian\Base de datos\Base de datos 20ago23.xlsx'  # Replace with the actual path
df2 = pd.read_excel(file2, sheet_name='Sheet1')  # Change sheet_name as needed

# Filter data from the first Excel where status is not "Devuelto"
filtered_df1 = df1[df1['Status actual según portal'] != 'Devuelto']
filtered_df1_2 = df1_2[df1_2['Status actual según portal'] != 'Devuelto']
filtered_df1_3 = df1_3[df1_3['Status'] != 'Devuelto']

# Merge the dataframes with the second Excel using "expediente" and "Request ID"
merged_df = pd.merge(filtered_df1, df2, left_on='expediente', right_on='Request ID', how='inner')
merged_df_2 = pd.merge(filtered_df1_2, df2, left_on='expediente', right_on='Request ID', how='inner')
merged_df_3 = pd.merge(filtered_df1_3, df2, left_on='expediente', right_on='Request ID', how='inner')

# Filter rows where status in the second Excel is "Devuelto"
discrepancies = merged_df[merged_df['Estado'] == 'Devuelto']
discrepancies2 = merged_df_2[merged_df_2['Estado'] == 'Devuelto']
discrepancies3 = merged_df_3[merged_df_3['Estado'] == 'Devuelto']

# Concatenate the three discrepancies DataFrames into a single DataFrame
all_discrepancies = pd.concat([discrepancies, discrepancies2, discrepancies3])

# Print or save the discrepancies to a new Excel file
all_discrepancies.to_excel('all_discrepancies.xlsx', index=False)
