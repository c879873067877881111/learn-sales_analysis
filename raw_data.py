import pandas as pd

df = pd.read_csv('data.csv', encoding='utf-8')

# 查看數據概覽
print(df.info())
print("=======================")
print(df.describe())
print("=======================")
# 查看前/後幾行
print(df.head())
print("=======================")
print(df.tail())
# 查看數據維度
print(df.shape)
print("=======================")
# 查看唯一值
print(df['訂單編號'].unique())
print("=======================")
print(df['訂單編號'].nunique())
print("=======================")
# 查看缺失值
print(df.isna().sum())
# 檢查重複值
print(df.duplicated().sum())

# 使用統計方法檢測異常值
Q1 = df['商品利潤'].quantile(0.25)
Q3 = df['商品利潤'].quantile(0.75)
IQR = Q3 - Q1
lower_bound = Q1 - 1.5 * IQR
upper_bound = Q3 + 1.5 * IQR

# 過濾異常值
df_filtered = df[(df['商品利潤'] >= lower_bound) & (df['商品利潤'] <= upper_bound)]
print(df_filtered)

