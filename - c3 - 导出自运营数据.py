import pandas as pd

# 重要提示：确认以下两个excel【空文档】在对应文件夹中
# EF_自营_PowerBI.xlsx
# HH_自营_PowerBI.xlsx

Input_EF_Path = 'EF PowerBI.xlsx'
Input_HH_Path = 'HH PowerBI.xlsx'

# 读取HH与EF【Sales数据】的自营部分
EFSales = pd.read_excel(Input_EF_Path, sheet_name = 'Sales')
HHSales = pd.read_excel(Input_HH_Path, sheet_name= 'Sales')

Self_EFSales = EFSales[EFSales['Client']=="自营"]
Self_HHSales = HHSales[HHSales['Client']=="自营"]

# 读取HH与EF【Campaign数据】的自营部分
EFCampaign = pd.read_excel(Input_EF_Path, sheet_name= 'Campaign')
HHCampaign = pd.read_excel(Input_HH_Path, sheet_name= 'Campaign')

Self_EFCampaign = EFCampaign[EFCampaign['Client']=='自营']
Self_HHCampaign = HHCampaign[HHCampaign['Client']=='自营']

# 读取HH与EF【AFN数据】的自营部分
EFAFN = pd.read_excel(Input_EF_Path, sheet_name='AFN')
HHAFN = pd.read_excel(Input_HH_Path, sheet_name='AFN')

Self_EFAFN = EFAFN[EFAFN['Client']=='自营']
Self_HHAFN = HHAFN[HHAFN['Client']=='自营']


# 读取HH与EF【MonthlyInventory数据】的自营部分
EF_MonthlyInventory = pd.read_excel(Input_EF_Path, sheet_name= 'MonthlyInventory')
HH_MonthlyInventory = pd.read_excel(Input_HH_Path, sheet_name ='MonthlyInventory')

Self_EFMonthlyInventory = EF_MonthlyInventory[EF_MonthlyInventory['Client']=='自营']
Self_HHMonthlyInventory = HH_MonthlyInventory[HH_MonthlyInventory['Client']=='自营']


# 读取HH与EF【LongTermInventory数据】的自营部分
EF_LongTermInventory = pd.read_excel(Input_EF_Path, sheet_name = 'LongTermInventory')
HH_LongTermInventory = pd.read_excel(Input_HH_Path, sheet_name = 'LongTermInventory')

Self_EFLongTermInventory = EF_LongTermInventory[EF_LongTermInventory['Client']=='自营']
Self_HHLongTermInventory = HH_LongTermInventory[HH_LongTermInventory['Client']=='自营']


# 将HH与EF【Sales】，【Campaign】，【AFN】，【MonthlyInventory】，【LongTermInventory】分别汇总到【HH_自营_PowerBI.xlsx】与【EF_自营_PowerBI.xlsx】中
EF_FilePath = 'EF_自营_PowerBI.xlsx'
HH_FilePath = 'HH_自营_PowerBI.xlsx'

with pd.ExcelWriter(EF_FilePath) as writer:
    Self_EFSales.to_excel(writer, sheet_name= 'Sales', index = False)
    Self_EFCampaign.to_excel(writer, sheet_name= 'Campaign', index = False)
    Self_EFAFN.to_excel(writer, sheet_name= 'AFN', index = False)
    Self_EFMonthlyInventory.to_excel(writer, sheet_name= 'MonthlyInventory', index = False)
    Self_EFLongTermInventory.to_excel(writer, sheet_name = 'LongTermInventory', index = False)

with pd.ExcelWriter(HH_FilePath) as writer:
    Self_HHSales.to_excel(writer, sheet_name= 'Sales', index = False)
    Self_HHCampaign.to_excel(writer, sheet_name= 'Campaign', index = False)
    Self_HHAFN.to_excel(writer, sheet_name= 'AFN', index = False)
    Self_HHMonthlyInventory.to_excel(writer, sheet_name= 'MonthlyInventory', index = False)
    Self_HHLongTermInventory.to_excel(writer, sheet_name = 'LongTermInventory', index = False)