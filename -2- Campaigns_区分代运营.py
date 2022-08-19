import pandas as pd
import os
from openpyxl import load_workbook

# 注意调整【EF/HH】
path = r"C:\Users\Steffi\Desktop\EF Power BI_Data\EF 2022支持数据\2_Campaign_广告费"
AllCampaign = pd.DataFrame()

# 对path文件夹下各个csv文件统一修改：新增列，修改列名，合并为一张表AllCampaign
for root, dirs, files in os.walk(path):
    for file in files:
        if file.startswith('Campaigns_'):
            # 读取csv数据
            file_path = os.path.join(root, file)
            campaign= pd.read_csv(file_path)

            # 新增三列的列名：
            country= file[:-4].split('_')[1]
            currency= file[:-4].split('_')[2]
            截取month = file[:-4].split('_')[3]

            # 添加列
            campaign.insert(0, '截取month', 截取month)
            campaign.insert(0, 'currency', currency)
            campaign.insert(0, 'country', country)

            # 更改列名
            a = 'Spend'+'(' + currency +')'
            b = 'Spend'
            c = 'Budget'+'(' + currency +')'
            d = 'Budget'
            e = 'Sales'+'(' + currency +')'
            f = 'Sales'
            g = 'VCPM'+'(' + currency +')'
            h = 'VCPM'
            i = 'NTB Sales'+'(' + currency +')'
            j = 'NTB Sales'
            campaign.rename(columns={a:b}, inplace = True)
            campaign.rename(columns={c:d}, inplace=True)
            campaign.rename(columns={e:f}, inplace=True)
            campaign.rename(columns={g:h}, inplace=True)
            campaign.rename(columns={i:j}, inplace=True)

            # 合并各国每月campaign到AllCampaign中
            AllCampaign = pd.concat([AllCampaign,campaign])


# 用dict创建《代运营sku信息表中》对应的键值对
代运营广告组合 = pd.read_excel('代运营信息.xlsx', sheet_name= '代运营广告组合')
Campaign_Client = 代运营广告组合.set_index(['广告组合'])['客户简称'].to_dict()

# 在AllCampaign中新建一个空列，名为["客户简称"]
AllCampaign['Client'] = ""

# 如果AllCampaign['广告组合']有代运营的，在['客户简称']列中加入该客户名字，，HH的key为"Portfolio"，EF的key为"广告组合"
for c in AllCampaign['Portfolio']:
    if c in Campaign_Client.keys():
        AllCampaign.loc[AllCampaign['Portfolio'] == c,'Client'] = Campaign_Client[c]

# 向已有excel(第-1-步)添加新表sheet，注意调整【EF/HH】
FilePath = 'EF PowerBI.xlsx'
ExcelWorkbook = load_workbook(FilePath)
writer = pd.ExcelWriter(FilePath, engine = 'openpyxl')
writer.book = ExcelWorkbook
AllCampaign.to_excel(writer, sheet_name= 'Campaign', index = False)

writer.close()


# 提示
print("手动调整['Client']列空白处为【自营】")