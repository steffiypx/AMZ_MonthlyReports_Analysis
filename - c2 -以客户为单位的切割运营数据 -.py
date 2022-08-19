import pandas as pd

FilePath = 'Client PowerBI.xlsx'

# 读取代运营客户的简称
Info = pd.read_excel('代运营信息.xlsx', sheet_name= "代运营SKU信息")
Client = list(Info['客户简称'].unique())

# 读取代运营数据
ClientSales = pd.read_excel(FilePath, sheet_name = 'Sales')
ClientCampaign = pd.read_excel(FilePath, sheet_name = 'Campaign')
ClientAFN = pd.read_excel(FilePath, sheet_name = 'AFN')
ClientMonthlyInventory = pd.read_excel(FilePath, sheet_name = 'MonthlyInventory')
ClientLongTermInventory = pd.read_excel(FilePath, sheet_name = 'LongTermInventory')

for c in Client:
    c_Sales = ClientSales[ClientSales['Client'] == f'{c}']
    c_Campaign = ClientCampaign[ClientCampaign['Client'] == f'{c}']
    c_AFN = ClientAFN[ClientAFN['Client'] == f'{c}']
    c_MonthlyInventory = ClientMonthlyInventory[ClientMonthlyInventory['Client'] == f'{c}']
    c_LongTermInventory = ClientLongTermInventory[ClientLongTermInventory['Client'] == f'{c}']

    # 将具体客户c的Sales，Campaign，AFN，MonthlyInventory，LongTermInventory汇总到AmazonData_c.xlsx中
    OutputPath = f'月度对账单_{c}.xlsx'

    with pd.ExcelWriter(OutputPath) as writer:
        c_Sales.to_excel(writer, sheet_name='1_订单流水', index=False)
        c_Campaign.to_excel(writer, sheet_name='2_广告费', index=False)
        c_AFN.to_excel(writer, sheet_name='3_月末库存盘点', index=False)
        c_MonthlyInventory.to_excel(writer, sheet_name='4_月度仓储费', index=False)
        c_LongTermInventory.to_excel(writer, sheet_name='5_长期仓储费', index=False)
