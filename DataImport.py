import xlsxreader as xlr
import pandas as pd

datafram1 = pd.read_excel('C:/Users/Harsh Rathore/Downloads/Students(new).xlsx')

surnames = []

height=[]
for  name in datafram1['Name']:
    match name:
        case "Harsh": 
            surnames.append('Jack')
            height.append(190)
        case "John":
            surnames.append('Smith')
            height.append(110)

        case "Vivek":
            surnames.append('swami')
            height.append(170)

datafram1["Surname"]=surnames
datafram1["Height"]=height
name=datafram1.loc[1]
a = datafram1["Height"].tolist
print(type(a))
with  pd.ExcelWriter('C:/Users/Harsh Rathore/Downloads/Students(new).xlsx',mode='a',if_sheet_exists="replace",engine='openpyxl') as writer:

    datafram1.to_excel(writer,sheet_name='Sheet1',index=False)
print(datafram1)
writer.close    

