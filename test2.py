import pandas as pd

excel1= r"C:\Users\Admin\Downloads\Timetable Workbook - SUTT Task 1.xlsx"

S= ['S1','S2','S3','S4','S5','S6']

df=pd.read_excel(excel1, sheet_name=S)


df=pd.read_excel(excel1, header=1)

data={}
topic_data=[]
current_entry= None


df.dropna(how="all",inplace=True)


for _, row in df.iterrows():
        row_data=row.dropna().to_dict()
        first_col_value=row.iloc[0]

       
        if pd.isnull(first_col_value):
            if current_entry is not None:
                current_entry.setdefault("Sections",[]).append(row_data)
        else:

                current_entry=row_data
                topic_data.append(current_entry)

data['S']= topic_data




import json
with open("output.json","w") as json_file:
      json.dump(data,json_file,indent=5)

print("JSON data has been written to output.json")
    