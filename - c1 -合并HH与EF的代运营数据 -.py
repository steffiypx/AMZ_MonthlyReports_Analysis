import pandas as pd

Input_EF_Path = 'EF PowerBI.xlsx'
Input_HH_Path = 'HH PowerBI.xlsx'

# 读取代运营客户的简称【注意表格内不能有空行】
Info = pd.read_excel('代运营信息.xlsx', sheet_name= "代运营SKU信息")
Client = list(Info['客户简称'].unique())

# 合并HH与EF的"代运营客户"Sales数据
EFSales = pd.read_excel(Input_EF_Path, sheet_name = 'Sales')
HHSales = pd.read_excel(Input_HH_Path, sheet_name= 'Sales')

Client_EFSales = EFSales[EFSales['Client'].isin(Client)]
Client_HHSales = HHSales[HHSales['Client'].isin(Client)]

ClientSales = pd.DataFrame()
ClientSales = pd.concat([ClientSales, Client_EFSales, Client_HHSales], ignore_index= True)



# 合并HH与EF的“代运营客户”Campaign数据
EFCampaign = pd.read_excel(Input_EF_Path, sheet_name= 'Campaign')
HHCampaign = pd.read_excel(Input_HH_Path, sheet_name= 'Campaign')

Client_EFCampaign = EFCampaign[EFCampaign['Client'].isin(Client)]
Client_HHCampaign = HHCampaign[HHCampaign['Client'].isin(Client)]

ClientCampaign = pd.DataFrame()
ClientCampaign = pd.concat([ClientCampaign, Client_EFCampaign, Client_HHCampaign], ignore_index= True)


# 合并HH与EF的“代运营客户”AFN数据
EFAFN = pd.read_excel(Input_EF_Path, sheet_name='AFN')
HHAFN = pd.read_excel(Input_HH_Path, sheet_name='AFN')

Client_EFAFN = EFAFN[EFAFN['Client'].isin(Client)]
Client_HHAFN =HHAFN[HHAFN['Client'].isin(Client)]

ClientAFN = pd.DataFrame()
ClientAFN = pd.concat([ClientAFN, Client_EFAFN, Client_HHAFN], ignore_index= True)

# 合并HH与EF的“代运营客户”MonthlyInventory数据
EF_MonthlyInventory = pd.read_excel(Input_EF_Path, sheet_name= 'MonthlyInventory')
HH_MonthlyInventory = pd.read_excel(Input_HH_Path, sheet_name ='MonthlyInventory')

Client_EF_MonthlyInventory = EF_MonthlyInventory[EF_MonthlyInventory['Client'].isin(Client)]
Client_HH_MonthlyInventory =HH_MonthlyInventory[HH_MonthlyInventory['Client'].isin(Client)]

Client_MonthlyInventory = pd.DataFrame()
Client_MonthlyInventory = pd.concat([Client_MonthlyInventory, Client_EF_MonthlyInventory, Client_HH_MonthlyInventory], ignore_index= True)

# 合并HH与EF的“代运营客户”LongTermInventory数据
EF_LongTermInventory = pd.read_excel(Input_EF_Path, sheet_name = 'LongTermInventory')
HH_LongTermInventory = pd.read_excel(Input_HH_Path, sheet_name = 'LongTermInventory')

Client_EF_LongTermInventory = EF_LongTermInventory[EF_LongTermInventory['Client'].isin(Client)]
Client_HH_LongTermInventory =HH_LongTermInventory[HH_LongTermInventory['Client'].isin(Client)]

Client_LongTermInventory = pd.DataFrame()
Client_LongTermInventory = pd.concat([Client_LongTermInventory, Client_EF_LongTermInventory, Client_HH_LongTermInventory], ignore_index= True)


# 将客户的Sales，Campaign，AFN，MonthlyInventory，LongTermInventory汇总到Client PowerBI.xlsx中
# 确认【Client PowerBI.xlsx】已在文件夹中
FilePath = 'Client PowerBI.xlsx'

with pd.ExcelWriter(FilePath) as writer:
    ClientSales.to_excel(writer, sheet_name= 'Sales', index = False)
    ClientCampaign.to_excel(writer, sheet_name= 'Campaign', index = False)
    ClientAFN.to_excel(writer, sheet_name= 'AFN', index = False)
    Client_MonthlyInventory.to_excel(writer, sheet_name= 'MonthlyInventory', index = False)
    Client_LongTermInventory.to_excel(writer, sheet_name = 'LongTermInventory', index = False)
