import pandas as pd
import os
import datetime

# List of Excel files
excel_files = "Consolidated result Resources.xlsx"

fact_df = pd.read_excel(excel_files)

# Target list file
target_file = "Data_Model.xlsx" 

# Load the target list from Target list file (Sheet: DIM_Employee)
root_df = pd.read_excel(target_file, sheet_name="DIM_Employee")

#df_r = fact_df.describe()
#print(df_r)

#to apply filter to exclude row_name "Total"
fact_df = fact_df[(fact_df["First"] != "Total") & ((fact_df['First'].notna()) | (fact_df['Last'].notna()))]
#print(df["First"])
#df.to_excel("filtered_data.xlsx")

df_filtered = fact_df[['First', 'Last']]
df_filtered['FullName'] = df_filtered['First'] + ' ' + df_filtered['Last']
#to make grouping of data
df_group = df_filtered['FullName'].value_counts()
#print(df_group)
df_filtered = df_filtered[['First', 'Last', 'FullName']].drop_duplicates()

#df_filtered['user_id'] = range(1, len(df_filtered) + 1)
#print(df_filtered)

#to make grouping of data
#df_group = df.groupby(['Source', "Type"])["1-Jan-25"].aggregate(["min", "max", "mean"])

#print(df_group)

# Використовуємо merge для додавання тільки унікальних рядків з df2
df_merged = root_df.merge(df_filtered, how='outer', indicator=True).query('_merge == "right_only"').drop('_merge', axis=1)
# Залишаємо тільки ті рядки з df2, яких немає в df1
df_new_rows = df_filtered[~df_filtered.isin(root_df.to_dict(orient='list')).all(axis=1)]

# Додаємо унікальні рядки з df2 до df1
df_result = pd.concat([root_df, df_merged])

#Створюємо функцію генерування id
df_result = df_result.reset_index(drop=True)
for i in range(1, len(df_result)):
    if pd.isna(df_result.loc[i, 'EmployeeID']):
        max_id = df_result['EmployeeID'].max()
        df_result.loc[i, 'EmployeeID'] = max_id + 1
print(df_result)

# Робимо backup Data model.xlsx file та додаємо дату та час
# Rename file of the specified path
time = datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
os.rename("Data_Model.xlsx",f'Data_Model {time}.xlsx')

df_result.to_excel(f"Data_Model_DIM_Employee.xlsx",sheet_name="DIM_Employee")

# Додаємо ці нові рядки до df1
df_result = pd.concat([root_df, df_new_rows])
df_result.to_excel(f'IsIn result {excel_files}.xlsx')

# df to be mergered to dictionary
# df_q = {"count_row":df["Source"]["1-Jan-25"].count(),
#         "unique_row":df["Source"]["1-Jan-25"].value_counts(),
#         "distinct_rows":None}
# print(df_q)

#df_r.to_excel(f'Statistic result {excel_files}.xlsx')



