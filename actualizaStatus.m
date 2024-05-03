clc, clear
% Load the first Excel files
file1 = "C:\Users\Dell\Dropbox\SEGUIMIENTO DEPÓSITOS\depósitos (Katiuska F. Franco's conflicted copy 2023-10-03).xlsx";  % Replace with the actual path
df1 = readtable(file1, 'Sheet', '2021');  % Change sheet name as needed
df1_2 = readtable(file1, 'Sheet', '2022');  % Change sheet name as needed
df1_3 = readtable(file1, 'Sheet', '2023');  % Change sheet name as needed
df1_4 = readtable(file1, 'Sheet', '2024');  % Change sheet name as needed

% Load the second Excel file
file2 = 'C:\Users\Dell\OneDrive\Artes Ian\Base de datos\Base de datos 02may24.xlsx';  % Replace with the actual path
df2 = readtable(file2, 'Sheet', 'data 02may24');  % Change sheet name as needed

%% 
df1 = renamevars(df1, 'Var5', 'Expediente');
df1 = renamevars(df1, 'Var6', 'Status');
df1 = renamevars(df1, 'Var4', 'Nombre');
df1 = renamevars(df1, 'Var2', 'TITULAR');
df1 = renamevars(df1, 'Var3', 'Fecha');

df1_3 = renamevars(df1_3, 'Var5', 'Expediente');
df1_3 = renamevars(df1_3, 'Var6', 'Status');
df1_3 = renamevars(df1_3, 'Var4', 'Nombre');
df1_3 = renamevars(df1_3, 'Var2', 'TITULAR');
df1_3 = renamevars(df1_3, 'Var3', 'Fecha');


% Filter data from the first Excel where status is not "Devuelto"
filtered_df1 = df1(df1.Status ~= "Devuelto", :);
filtered_df1_2 = df1_2(df1_2.StatusActualSeg_nPortal ~= "Devuelto", :);
filtered_df1_3 = df1_3(df1_3.Status ~= "Devuelto", :);
filtered_df1_4 = df1_4(df1_4.Status ~= "Devuelto", :);


%%
% Merge the dataframes with the second Excel using "expediente" and "Request ID"
merged_df = innerjoin(filtered_df1, df2, 'LeftKeys', 'Expediente', 'RightKeys', 'Solicitud');
merged_df_2 = innerjoin(filtered_df1_2, df2, 'LeftKeys', 'Expediente', 'RightKeys', 'Solicitud');
merged_df_3 = innerjoin(filtered_df1_3, df2, 'LeftKeys', 'Expediente', 'RightKeys', 'Solicitud');
merged_df_4 = innerjoin(filtered_df1_4, df2, 'LeftKeys', 'Expediente', 'RightKeys', 'Solicitud');


%%
% Filter rows where status in the second Excel is "Devuelto"
discrepancies = merged_df(ismember(merged_df.Estado, "Devuelta"), {'Expediente', 'Estado', 'Nombre', 'TITULAR', 'Fecha'});
discrepancies2 = merged_df_2(ismember(merged_df_2.Estado, "Devuelta"), {'Expediente', 'Estado', 'Nombre', 'TITULAR', 'Fecha'}); % var5 es expediente, var6 estado y var4 nombre
discrepancies3 = merged_df_3(ismember(merged_df_3.Estado, "Devuelta"), {'Expediente', 'Estado', 'Nombre', 'TITULAR', 'Fecha'});
discrepancies4 = merged_df_4(ismember(merged_df_4.Estado, "Devuelta"), {'Expediente', 'Estado', 'Nombre', 'TITULAR', 'Fecha'});

%discrepancies2 = renamevars(discrepancies2, 'Var5', 'Expediente');
%discrepancies2 = renamevars(discrepancies2, 'Var6', 'Estado');
%discrepancies2 = renamevars(discrepancies2, 'Var4', 'Nombre');
%discrepancies2 = renamevars(discrepancies2, 'Var2', 'TITULAR');


%%
% Concatenate the three discrepancies tables into a single table
all_discrepancies = [discrepancies; discrepancies2; discrepancies3; discrepancies4];

% Save the discrepancies to a new Excel file
writetable(all_discrepancies, 'all_discrepancies.xlsx');
