# Writes all files in a directory with modified date to excel file
import os, time, datetime, shutil
import pandas as pd

folder="I:\\Quality Control\\After the Fact Documentation\\ATF Sort"
file_data = []
column_labels = ["Directory","Customer", "File", "Modified Date"]

shutil.copy(folder + "\\ATF_FIFO.xlsx", folder + "\\ATF_FIFO_prev.xlsx")

for path, dirs, files in os.walk(folder):
    for filename in files:
        try:
            if path == "I:/Quality Control/After the Fact Documentation/ATF Sort\2 Duplicates":
                pass
            else:
                pathname = os.path.join(path)
                
                # Clean the customer name
                path_base, customer = path.split("I:\\Quality Control\\After the Fact Documentation\\ATF Sort")
                if customer.startswith('\\'):
                    customer = customer[1:]
                customer = customer.replace('\\Sorted','')
                customer = customer.replace('\\Product Lined','')
                customer = customer.replace('2 Duplicates','Duplicate')
                customer = customer.rstrip()
                
                filename_with_path = os.path.join(path,filename)
                file_stats = os.stat(filename_with_path)
                mtime = file_stats.st_mtime
                mod_timestamp = datetime.datetime.fromtimestamp(mtime)
                local_time = datetime.datetime.strftime(mod_timestamp, "%Y-%m-%d")
                file_data.append([pathname,customer,filename,local_time])
        except:
            pass

df = pd.DataFrame.from_records(file_data, columns=column_labels)
df_sorted = df.sort_values(by = "Modified Date", ascending = True)
df_sorted.index.name = "ID"

writer = pd.ExcelWriter('I:/Quality Control/After the Fact Documentation/ATF Sort/ATF_FIFO.xlsx', engine='xlsxwriter')
df_sorted.to_excel(writer, sheet_name='Sheet1', index=True)

workbook = writer.book
worksheet = writer.sheets['Sheet1']

worksheet.set_column('A:A', 75)
worksheet.set_column('B:B', 25)
worksheet.set_column('C:C', 50)
worksheet.set_column('D:D', 15)
worksheet.autofilter('A1:D1')

try:
    workbook.close()
except:
    print("Unable to write to Excel workbook. Ensure it is closed and then try again.")

print(df_sorted.head(1))
