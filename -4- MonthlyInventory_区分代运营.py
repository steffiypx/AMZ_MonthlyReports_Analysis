import pandas as pd
import os
from openpyxl import load_workbook

# 注意调整【EF/HH】
path = r"C:\Users\Steffi\Desktop\EF Power BI_Data\EF 2022支持数据\4_Monthly_Storage_Fee_月仓储费"
AllmInventory = pd.DataFrame()

for root, dirs, files in os.walk(path):
    for file in files:
        # 读取csv数据需要特别明确encoding为"ISO-8859-1"
        file_path = os.path.join(root, file)
        mInventory = pd.read_csv(file_path, encoding='ISO-8859-1')

        # 【适用HH，不适用EF】修改列名：['Country code']&['country-code']，统一为['country_code']
        mInventory.rename(columns={'Country code':'country_code','country-code':'country_code'}, inplace= True)

        # 【适用HH，不适用EF】修改列名:['ASIN'] & ['asin.1']，统一为['ASIN']
        mInventory.rename(columns={'ASIN':'ASIN', 'asin.1':'ASIN','asin':'ASIN','ï»¿ASIN':'ASIN'}, inplace=True)

        # 修改EU Inventory中的country_code列下的字符，使GB --> UK
        if file.split('_')[1] == 'EU':
            mInventory['country_code'] = mInventory['country_code'].str.replace('GB', 'UK')

        # 添加一列【截取month】
        截取month = file.rstrip('.csv').split('_')[2]
        mInventory.insert(0, '截取month', 截取month)

        #各个国家/地区每个月份的mInventory合并为AllmInventory
        AllmInventory = pd.concat([AllmInventory, mInventory])

# 在本表区分代运营
# Step1: 新建['Client']列
AllmInventory['Client'] = ''


# Step2: 根据《代运营信息》用dict建立键值对，ASIN为key，客户简称为value
代运营sku信息 = pd.read_excel('代运营信息.xlsx', sheet_name= "代运营SKU信息")
ASIN_Client = 代运营sku信息.set_index(['ASIN'])['客户简称'].to_dict()

# Step3.1: 遍历AllmInventory['ASIN']列，找出对应客户简称
for a in AllmInventory['ASIN']:
    if a in ASIN_Client.keys():
        AllmInventory.loc[AllmInventory['ASIN'] == a, 'Client'] = ASIN_Client[a]


# 输出结果，注意调整【EF/HH】
# 向已有excel(第-1-步)添加新表sheet，注意调整【EF/HH】
FilePath = 'EF PowerBI.xlsx'
ExcelWorkbook = load_workbook(FilePath)
writer = pd.ExcelWriter(FilePath, engine = 'openpyxl')
writer.book = ExcelWorkbook
AllmInventory.to_excel(writer, sheet_name= 'MonthlyInventory', index = False)

writer.close()

print('=====MX的数据可能发生偏移=====')
print("手动调整['Client']列空白处为【自营】")