import os
import pandas as pd
from openpyxl import load_workbook

# 注意调整【EF/HH】
path = r"C:\Users\Steffi\Desktop\HH Power BI_Data\HH后台原始数据\5_LongTerm_Storage_Fee_长期仓储费"
All_LT_Inventory = pd.DataFrame()

for root, dirs, files in os.walk(path):
    for file in files:
        # 读取csv数据需要特别明确encoding为"ISO-8859-1"
        file_path = os.path.join(root, file)
        LT_Inventory = pd.read_csv(file_path, encoding='ISO-8859-1')

        # 修改EU Inventory中的country_code列下的字符，使GB --> UK
        if file.split('_')[1] == 'EU':
            LT_Inventory['country'] = LT_Inventory['country'].str.replace('GB', 'UK')

        # 添加一列【截取month】
        截取month = file.rstrip('.csv').split('_')[2]
        LT_Inventory.insert(0, '截取month', 截取month)

        # 统一列名：替换列名：sku-->Merchant SKU, fnsku--> FNSKU, asin--> ASIN, country--> Country, currency--> Currency
        colNameDict = {'sku': 'Merchant SKU', 'fnsku': 'FNSKU', 'asin': 'ASIN', 'country': 'Country',
                       'currency': 'Currency'}
        LT_Inventory.rename(columns=colNameDict, inplace=True)

        #各个国家/地区每个月份的mInventory合并为AllmInventory
        All_LT_Inventory = pd.concat([All_LT_Inventory,LT_Inventory])

# 在本表区分代运营
# Step1: 新建['Client']列
All_LT_Inventory['Client'] = ''

# Step2: 根据《代运营信息》用dict建立键值对，MSKU为key，客户简称为value
代运营sku信息 = pd.read_excel('代运营信息.xlsx', sheet_name= "代运营SKU信息")
MSKU_Client = 代运营sku信息.set_index(['MSKU'])['客户简称'].to_dict()

# Step3: 遍历All_LT_Inventory['Merchant SKU']列，找出对应客户简称
for a in All_LT_Inventory['Merchant SKU']:
    if a in MSKU_Client.keys():
        All_LT_Inventory.loc[All_LT_Inventory['Merchant SKU'] == a, 'Client'] = MSKU_Client[a]


# 向已有excel(第-1-步)添加新表sheet，注意调整【EF/HH】
FilePath = 'HH PowerBI.xlsx'
ExcelWorkbook = load_workbook(FilePath)
writer = pd.ExcelWriter(FilePath, engine = 'openpyxl')
writer.book = ExcelWorkbook
All_LT_Inventory.to_excel(writer, sheet_name= 'LongTermInventory', index = False)

writer.close()
print("手动调整['Client']列空白处为【自营】")