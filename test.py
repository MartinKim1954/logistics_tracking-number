import csv
import pandas as pd

df = pd.read_csv(
    '/Users/martinkim/GITHUB/00_Automated System/02_Logistics-team/logistics_tracking-number/PO_SKU_LIST_20200904150707.csv')
# print(df.shape)
# print(df.shape[0])
# goyang_list = []
# a = df['물류센터'] == '고양1'
# b = df['발주수량'] >= 10
# is_goyang_1_and_above_10 = df[a & b]
# print(is_goyang_1_and_above_10)
goyang1_order_quantity = []
goyang1_order_number = []
is_goyang1 = df[df['물류센터'] == '고양1']
print(type(is_goyang1.shape[1]))
for i in range(is_goyang1.shape[1]):
    goyang1_order_number.append(is_goyang1.iloc[i].loc['발주번호'])
    goyang1_order_quantity.append(is_goyang1.iloc[i].loc['발주수량'])
print("고양1 발주번호: ", goyang1_order_number)
print("고양1 발주수량: ", goyang1_order_quantity)

# for row_index, row in csv_test.iterrows():
#     # print(row_index)
#     # print(row)
#     goyang_list.append(row.loc['물류센터'])
# print(goyang_list)
