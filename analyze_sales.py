import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']  # 設置中文字體
plt.rcParams['axes.unicode_minus'] = False  # 解決負號顯示問題

# 讀取數據
df = pd.read_excel('CPE銷售明細檔20240901-20240930_新北.xlsx')

# 1. 計算各服務中心的銷售總量
service_center_sales = df.groupby('服務中心名稱').agg({
    '數量': 'sum',
    '金額': 'sum',
    '未稅金額': 'sum'
}).reset_index()

# 計算平均單價
service_center_sales['平均單價'] = service_center_sales['金額'] / service_center_sales['數量']

# 2. 計算Z-score來識別異常值
service_center_sales['數量_zscore'] = stats.zscore(service_center_sales['數量'])
service_center_sales['金額_zscore'] = stats.zscore(service_center_sales['金額'])
service_center_sales['平均單價_zscore'] = stats.zscore(service_center_sales['平均單價'])

# 標記異常值（Z-score > 2 或 < -2）
service_center_sales['數量異常'] = abs(service_center_sales['數量_zscore']) > 2
service_center_sales['金額異常'] = abs(service_center_sales['金額_zscore']) > 2
service_center_sales['平均單價異常'] = abs(service_center_sales['平均單價_zscore']) > 2

# 3. 創建視覺化
# 銷售數量比較
plt.figure(figsize=(15, 8))
plt.subplot(2, 1, 1)
sns.barplot(data=service_center_sales.sort_values('數量', ascending=False), 
            x='服務中心名稱', y='數量')
plt.xticks(rotation=45, ha='right')
plt.title('各服務中心銷售數量比較')

# 銷售金額比較
plt.subplot(2, 1, 2)
sns.barplot(data=service_center_sales.sort_values('金額', ascending=False), 
            x='服務中心名稱', y='金額')
plt.xticks(rotation=45, ha='right')
plt.title('各服務中心銷售金額比較')

plt.tight_layout()
plt.savefig('service_center_analysis.png')

# 4. 輸出分析結果
print("\n=== 異常值分析結果 ===")
print("\n數量異常的服務中心：")
print(service_center_sales[service_center_sales['數量異常']][['服務中心名稱', '數量', '數量_zscore']])

print("\n金額異常的服務中心：")
print(service_center_sales[service_center_sales['金額異常']][['服務中心名稱', '金額', '金額_zscore']])

print("\n平均單價異常的服務中心：")
print(service_center_sales[service_center_sales['平均單價異常']][['服務中心名稱', '平均單價', '平均單價_zscore']])

# 5. 計算基本統計信息
print("\n=== 基本統計信息 ===")
print(service_center_sales[['數量', '金額', '平均單價']].describe())

# 6. 計算各服務中心的排名
print("\n=== 服務中心排名 ===")
print("\n銷售數量排名：")
print(service_center_sales.sort_values('數量', ascending=False)[['服務中心名稱', '數量']].head())

print("\n銷售金額排名：")
print(service_center_sales.sort_values('金額', ascending=False)[['服務中心名稱', '金額']].head()) 