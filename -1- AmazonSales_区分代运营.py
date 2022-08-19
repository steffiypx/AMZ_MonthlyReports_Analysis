import pandas as pd

# 根据需要调整【EF/HH】
path = r'C:\Business Analysis\HH Power BI_Sales by Country.xlsx'
AllData = pd.read_excel(path, None)
代运营sku信息 = pd.read_excel('代运营信息.xlsx', sheet_name= "代运营SKU信息")

# 获取Power BI统计表中所有国家的【SalesSheets】
SalesSheets = list()
for Sheet in list(AllData.keys()):
    if len(Sheet) == 2:
        SalesSheets.append(Sheet)
    if Sheet == 'IN-GATI':
        SalesSheets.append(Sheet)
    if Sheet == 'IN-FTZ':
        SalesSheets.append(Sheet)

# 合并所有国家的Sales表到【SalesData】中
SalesData = pd.DataFrame()
for c in SalesSheets:
    newData = AllData[c]
    SalesData = pd.concat([SalesData, newData])

# 添加2列：Quantity Factor, SKU SalesRefund Quantity
# Step1: 各国Sales表中表示Order与Refund的汇总：typeOrderRefund
typeOrder = ["Order","Bestellung","Pedido","Commande","Ordine","注文","拲暥","Pedido","Bestelling","Zamówienie","Sipariş"]
typeRefund = ["Refund","Ersattung","Reembolso","Remboursement","Rimborso","返金","曉嬥","Terugbetaling","Zwrot kosztów","Återbetalning"]
typeOrderRefund = typeRefund+typeOrder

# Step2: 增加Quantity Factor列与SKU OrderRefund Quantity列
SalesData['Quantity Factor'] = SalesData['type'].apply(lambda x : 1 if x in typeOrderRefund else 0)
SalesData['SKU OrderRefund Quantity'] = SalesData['Quantity Factor'] * SalesData['quantity']

# 添加1列：['ASIN']
# Step1: 代运营优惠券的筛选条件
ASIN = list(代运营sku信息.loc[:, 'ASIN'].dropna())

# Step2: 如果description列包含代运营ASIN，在新增['加工列:CouponASIN']列添加该ASIN，其他为空值
des = SalesData['description'].str.rsplit(r" ").str.get(-1)
SalesData['ASIN'] = des.apply(lambda x: x if x in ASIN else None)

# 添加2列：['MSKU'],['FNSKU']
# Step1: 读取代运营MSKU与FNSKU并汇总成allSKU
MSKU = list(代运营sku信息.loc[:, 'MSKU'].dropna())
FNSKU = list(代运营sku信息.loc[:, 'FNSKU'].dropna())

# Step2: 如果sku列包含代运营的MSKU与FNSKU，在新增['加工列：allSKU']列添加该MSKU或者FNSKU，其他为空值
SalesData['MSKU'] = SalesData['sku'].apply(lambda x: x if x in MSKU else None)
SalesData['FNSKU'] = SalesData['sku'].apply(lambda x: x if x in FNSKU else None)

# 使order id列找到对应的MSKU或者FNSKU
# Step1: 创建dict的键值对，order id为key,SKU为value，并去掉value为空的部分
orderID_SKU = SalesData.set_index(['order id'])['sku'].to_dict()
for key, value in dict(orderID_SKU).items():
    if value is None:
        del orderID_SKU

# Step2: 当键值对的value在MSKU或者FNSKU时，在对应列填补value
for o in orderID_SKU.keys():
    if orderID_SKU[o] in MSKU:
        SalesData.loc[SalesData['order id'] == o, 'MSKU'] = orderID_SKU[o]
    elif orderID_SKU[o] in FNSKU:
        SalesData.loc[SalesData['order id'] ==o,'FNSKU'] = orderID_SKU[o]

# 添加1列：['代运营客户']
# Step1: 创建空列
SalesData['Client'] = ''

# Step2: 用dict创建《代运营sku信息表中》对应的键值对
ASIN_Client = 代运营sku信息.set_index(['ASIN'])['客户简称'].to_dict()
MSKU_Client = 代运营sku信息.set_index(['MSKU'])['客户简称'].to_dict()
FNSKU_Client = 代运营sku信息.set_index(['FNSKU'])['客户简称'].to_dict()

# Step3: ['ASIN'],['MSKU'],['FNSKU']三列对客户简称对应，并在['代运营客户']列填充客户名字
for a in SalesData['ASIN']:
    if a in ASIN:
        SalesData.loc[SalesData['ASIN'] == a, 'Client'] = ASIN_Client[a]

for m in SalesData['MSKU']:
    if m in MSKU:
        SalesData.loc[SalesData['MSKU'] == m, 'Client'] = MSKU_Client[m]

for f in SalesData['FNSKU']:
    if f in FNSKU:
        SalesData.loc[SalesData['FNSKU'] == f, 'Client'] = FNSKU_Client[f]

# 输出结果，注意调整【EF/HH】
SalesData.to_excel('HH PowerBI.xlsx', sheet_name = 'Sales',index = False)

# 提示
print("手动调整['Client']列空白处为【自营】")