import os
import pandas as pd
from openpyxl import load_workbook

# 注意调整【EF/HH】
path = r"C:\Users\Steffi\Desktop\HH Power BI_Data\HH后台原始数据\3_Inventory_Ledger_库存分类库"
All_AFN = pd.DataFrame()

for root, dirs, files in os.walk(path):
    for file in files:
        # 读取csv数据需要特别明确encoding为"ISO-8859-1"
        file_path = os.path.join(root, file)
        AFN = pd.read_csv(file_path, encoding='ISO-8859-1')

        # 添加一列【截取month】
        截取month = file.rstrip('.csv').split('_')[2]
        AFN.insert(0, '截取month', 截取month)

        # 修改列名
        AFN.rename(columns={'Merchant SKU':'MSKU'},inplace=True)
        AFN.rename(columns={'Fulfilment network SKU (FNSKU)': 'FNSKU'}, inplace=True)
        AFN.rename(columns={'disposition': 'Disposition'}, inplace=True)
        AFN.rename(columns={'Customer shipments': 'Customer Shipments'}, inplace=True)
        AFN.rename(columns={'Customer returns': 'Customer Returns'}, inplace=True)
        AFN.rename(columns={'Vendor returns': 'Vendor Returns'}, inplace=True)
        AFN.rename(columns={'Warehouse transfer in/out': 'Warehouse Transfer In/Out'}, inplace=True)
        AFN.rename(columns={'Other': 'Other Events'}, inplace=True)
        AFN.rename(columns={'snapshot-date': 'Date'}, inplace=True)
        AFN.rename(columns={'In transit between warehouses': 'In Transit Between Warehouses'}, inplace=True)

        #各个国家/地区每个月份的mInventory合并为AllmInventory
        All_AFN = pd.concat([All_AFN,AFN])


# 在本表区分代运营
# Step1: 新建['客户简称']列
All_AFN['Client'] = ''


# Step2: 根据《代运营信息》用dict建立键值对，MSKU为key，客户简称为value
代运营sku信息 = pd.read_excel('代运营信息.xlsx', sheet_name= "代运营SKU信息")
MSKU_Client = 代运营sku信息.set_index(['MSKU'])['客户简称'].to_dict()

# Step3: 遍历All_AFN['MSKU']列，找出对应Client
for m in All_AFN['MSKU']:
    if m in MSKU_Client.keys():
        All_AFN.loc[All_AFN['MSKU'] == m, 'Client'] = MSKU_Client[m]


# Step4: ['Location']列替换GB为UK
for l in All_AFN['Location']:
    if l == 'GB':
        All_AFN['Location'] = All_AFN['Location'].str.replace('GB','UK')


# 向已有excel(第-1-步)添加新表sheet，注意调整【EF/HH】
FilePath = 'HH PowerBI.xlsx'
ExcelWorkbook = load_workbook(FilePath)
writer = pd.ExcelWriter(FilePath, engine = 'openpyxl')
writer.book = ExcelWorkbook
All_AFN.to_excel(writer, sheet_name= 'AFN', index = False)

writer.close()

# 提示
print("手动调整['Client']列空白处为【自营】")