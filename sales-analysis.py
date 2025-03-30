import pandas as pd

# 讀取並打印數據概要
def load_data(file_path='Data.csv'):
    df = pd.read_csv(file_path)
    # 將日期列轉換為datetime格式
    df['訂單日期'] = pd.to_datetime(df['訂單日期'])
    df['出貨日期'] = pd.to_datetime(df['出貨日期'])
    print(f"總記錄數: {len(df)}")
    print(f"日期範圍: {df['訂單日期'].min().strftime('%Y-%m-%d')} 至 {df['訂單日期'].max().strftime('%Y-%m-%d')}")
    print(f"列數: {len(df.columns)}")
    print(f"列名: {', '.join(df.columns)}")
    
    return df

# 處理效率分析
def order_processing_analysis(df):
    """分析訂單處理時間效率並顯示結果"""
    # 計算處理天數
    df['處理天數'] = (df['出貨日期'] - df['訂單日期']).dt.days
    
    # 計算各時間區間的訂單比例
    total_orders = len(df)
    same_day = len(df[df['處理天數'] == 0])
    within_week = len(df[(df['處理天數'] > 0) & (df['處理天數'] <= 7)])
    more_than_week = len(df[df['處理天數'] > 7])
    
    # 建立效率分析DataFrame
    efficiency_data = {
        '指標': ['平均處理天數', '當日出貨比例', '一週內出貨比例', '超過一週出貨比例'],
        '數值': [
            df['處理天數'].mean(),
            same_day / total_orders * 100,
            within_week / total_orders * 100,
            more_than_week / total_orders * 100
        ]
    }
    efficiency_df = pd.DataFrame(efficiency_data)
    efficiency_df['數值'] = efficiency_df['數值'].round(2)
    
    # 按配送方式統計平均處理時間
    shipping_time = df.groupby('配送方式')['處理天數'].agg(['mean', 'count']).reset_index()
    shipping_time['比例'] = shipping_time['count'] / total_orders * 100
    shipping_time = shipping_time.rename(columns={'mean': '平均處理天數'})
    shipping_time = shipping_time.round(2)
    
    print(efficiency_df)
    print("\n配送方式分析:")
    print(shipping_time)
    
    return efficiency_df, shipping_time, df

# 金額分析
def order_amount_analysis(df):
    """分析訂單金額分佈並顯示結果"""
    # 根據訂單編號合併訂單項目
    order_df = df.groupby('訂單編號').agg({
        '銷售金額': 'sum',
        '商品利潤': 'sum',
        '訂單編號': 'count',
        '銷售數量': 'sum'
    }).rename(columns={'訂單編號': '商品種類數'})
    
    # 計算平均訂單指標
    avg_metrics = {
        '指標': ['平均訂單金額', '平均訂單數量', '平均商品種類數'],
        '數值': [
            order_df['銷售金額'].mean(),
            order_df['銷售數量'].mean(),
            order_df['商品種類數'].mean()
        ]
    }
    avg_df = pd.DataFrame(avg_metrics)
    avg_df['數值'] = avg_df['數值'].round(2)
    
    # 訂單金額分佈
    order_df['金額區間'] = pd.cut(
        order_df['銷售金額'],
        bins=[0, 100, 500, 1000, float('inf')],
        labels=['小額($0-$100)', '中額($100-$500)', '大額($500-$1000)', '特大($1000以上)']
    )
    
    amount_dist = order_df.groupby('金額區間', observed=False).size().reset_index(name='訂單數')
    amount_dist['比例'] = amount_dist['訂單數'] / len(order_df) * 100
    amount_dist['比例'] = amount_dist['比例'].round(2)
    print(avg_df)
    print("\n訂單金額分佈:")
    print(amount_dist)

    return avg_df, amount_dist, order_df

# 利潤率分析
def profit_margin_analysis(df, order_df):
    """分析利潤率和盈利能力並顯示結果"""
    # 計算總體利潤率
    total_sales = df['銷售金額'].sum()
    total_profit = df['商品利潤'].sum()
    overall_margin = total_profit / total_sales * 100
    
    # 計算有盈利和無盈利的訂單比例
    profitable_orders = len(order_df[order_df['商品利潤'] > 0])
    unprofitable_orders = len(order_df[order_df['商品利潤'] <= 0])
    total_orders = len(order_df)
    
    # 建立利潤分析DataFrame
    profit_data = {
        '指標': ['總銷售額', '總利潤', '總體利潤率', '有盈利訂單比例', '無盈利訂單比例'],
        '數值': [
            total_sales,
            total_profit,
            overall_margin,
            profitable_orders / total_orders * 100,
            unprofitable_orders / total_orders * 100
        ]
    }
    profit_df = pd.DataFrame(profit_data)
    profit_df.loc[profit_df['指標'].isin(['總體利潤率', '有盈利訂單比例', '無盈利訂單比例']), '數值'] = profit_df.loc[profit_df['指標'].isin(['總體利潤率', '有盈利訂單比例', '無盈利訂單比例']), '數值'].round(2)
    print(profit_df)
    
    return profit_df

# 產品分析
def product_analysis(df):
    """產品銷售和利潤分析並顯示結果"""
    # 計算每個產品的銷售和利潤指標
    product_df = df.groupby('產品名稱').agg({
        '銷售金額': 'sum',
        '商品利潤': 'sum',
        '銷售數量': 'sum',
        '訂單編號': 'count'
    }).rename(columns={'訂單編號': '訂單數'})
    
    # 計算利潤率
    product_df['利潤率'] = product_df['商品利潤'] / product_df['銷售金額'] * 100
    product_df['單位平均價格'] = product_df['銷售金額'] / product_df['銷售數量']
    
    # 根據利潤排名
    top_profit_products = product_df.sort_values('商品利潤', ascending=False).head(10).reset_index()
    
    # 根據利潤率排名(僅考慮銷售額大於1000的產品)
    top_margin_products = product_df[product_df['銷售金額'] > 1000].sort_values('利潤率', ascending=False).head(10).reset_index()
    
    # 最低利潤率的產品
    bottom_margin_products = product_df[product_df['銷售金額'] > 1000].sort_values('利潤率').head(5).reset_index()

    print("\n利潤最高的產品(Top 10):")
    print(top_profit_products[['產品名稱', '商品利潤', '銷售金額', '利潤率', '銷售數量']])
        
    print("\n利潤率最高的產品(銷售額>$1000的Top 10):")
    print(top_margin_products[['產品名稱', '利潤率', '商品利潤', '銷售金額']])
        
    print("\n利潤率最低的產品(銷售額>$1000的前5名):")
    print(bottom_margin_products[['產品名稱', '利潤率', '商品利潤', '銷售金額']])

    
    return top_profit_products, top_margin_products, bottom_margin_products, product_df

# 季節和時間趨勢分析
def time_series_analysis(df):
    """銷售時間趨勢分析並顯示結果"""
    # 按月度分析銷售
    df['年月'] = df['訂單日期'].dt.strftime('%Y-%m')
    df['月份'] = df['訂單日期'].dt.month
    df['年'] = df['訂單日期'].dt.year
    
    # 月度銷售趨勢
    monthly_sales = df.groupby('年月').agg({
        '銷售金額': 'sum',
        '商品利潤': 'sum',
        '訂單編號': pd.Series.nunique
    }).rename(columns={'訂單編號': '訂單數'}).reset_index()
    monthly_sales = monthly_sales.sort_values('年月')
    
    # 季節性分析(每月平均)
    seasonal_sales = df.groupby('月份').agg({
        '銷售金額': 'sum',
        '商品利潤': 'sum',
        '訂單編號': pd.Series.nunique
    }).rename(columns={'訂單編號': '訂單數'}).reset_index()
    
    # 找出銷售旺季
    top_months = seasonal_sales.sort_values('銷售金額', ascending=False).head(3)
    month_names = {
        1: '一月', 2: '二月', 3: '三月', 4: '四月', 5: '五月', 6: '六月',
        7: '七月', 8: '八月', 9: '九月', 10: '十月', 11: '十一月', 12: '十二月'
    }
    top_months['月份名稱'] = top_months['月份'].map(month_names)
    
    # 年度比較
    yearly_sales = df.groupby('年').agg({
        '銷售金額': 'sum',
        '商品利潤': 'sum',
        '訂單編號': pd.Series.nunique
    }).rename(columns={'訂單編號': '訂單數'}).reset_index()
    
    # 計算年增長率
    yearly_sales['銷售額增長率'] = yearly_sales['銷售金額'].pct_change() * 100
    yearly_sales['利潤增長率'] = yearly_sales['商品利潤'].pct_change() * 100
    
    print("\n銷售旺季:")
    print(top_months[['月份名稱', '銷售金額', '訂單數']])
        
    print("\n年度銷售比較:")
    print(yearly_sales)
    
    return monthly_sales, seasonal_sales, top_months, yearly_sales

# 客戶分析
def customer_analysis(df):
    """客戶行為和購買模式分析並顯示結果"""
    # 客戶銷售統計
    customer_df = df.groupby('客戶編號').agg({
        '銷售金額': 'sum',
        '商品利潤': 'sum',
        '訂單編號': pd.Series.nunique,
        '產品名稱': pd.Series.nunique
    }).rename(columns={'訂單編號': '訂單數', '產品名稱': '購買產品種類'}).reset_index()
    
    # 計算每個客戶的平均訂單金額
    customer_df['平均訂單金額'] = customer_df['銷售金額'] / customer_df['訂單數']
    
    # 排序客戶(按照銷售金額)
    top_customers = customer_df.sort_values('銷售金額', ascending=False).reset_index(drop=True)
    
    # 客戶分級 (基於訂單數)
    customer_df['客戶分級'] = pd.cut(
        customer_df['訂單數'],
        bins=[0, 500, 1000, float('inf')],
        labels=['低頻客戶', '中頻客戶', '高頻客戶']
    )
    
    customer_segments = customer_df.groupby('客戶分級', observed=False).size().reset_index(name='客戶數')
    customer_segments['比例'] = customer_segments['客戶數'] / len(customer_df) * 100
    customer_segments['比例'] = customer_segments['比例'].round(2)
    
    # 計算客戶終身價值(CLV) - (總利潤)
    customer_df['客戶終身價值'] = customer_df['商品利潤']
    
    print("\n客戶銷售排名(Top 10):")
    print(top_customers.head(10)[['客戶編號', '銷售金額', '商品利潤', '訂單數', '平均訂單金額']])
        
    print("\n客戶分級:")
    print(customer_segments)
    
    return top_customers, customer_segments, customer_df


# 執行所有分析並產生報告
def generate_report(file_path='Data.csv'):
    """生成完整的銷售分析報告並返回結果"""
    
    # 讀取數據
    df = load_data(file_path)
    
    # 執行各類分析
    efficiency_df, shipping_time, df_with_processing = order_processing_analysis(df)
    avg_df, amount_dist, order_df = order_amount_analysis(df)
    profit_df = profit_margin_analysis(df, order_df)
    top_profit_products, top_margin_products, bottom_margin_products, product_df = product_analysis(df)
    monthly_sales, seasonal_sales, top_months, yearly_sales = time_series_analysis(df)
    top_customers, customer_segments, customer_df = customer_analysis(df)
    
    # 返回所有分析結果
    results = {
        'efficiency': efficiency_df,
        'shipping': shipping_time,
        'order_avg': avg_df,
        'order_dist': amount_dist,
        'profit': profit_df,
        'top_products': top_profit_products,
        'top_margin': top_margin_products,
        'monthly_sales': monthly_sales,
        'seasonal_sales': seasonal_sales,
        'top_months': top_months,
        'yearly_sales': yearly_sales,
        'customers': top_customers,
        'customer_segments': customer_segments,
    }
    print(" ")

    if results:
        print("已完成分析")
    else:
        print("分析過程。發生錯誤")
    return results


# 示例：根據特定需求執行部分分析
def analyze_product_performance(file_path='Data.csv'):
    """專注於產品表現的分析"""
    print("開始產品表現專項分析...\n")
    
    df = load_data(file_path)
    
    # 進行產品相關分析
    top_profit_products, top_margin_products, bottom_margin_products, product_df = product_analysis(df)
    
    
    # 執行自定義產品分析，計算各產品類別的總銷售額和利潤
    category_performance = df.groupby('產品類別').agg({
        '銷售金額': 'sum',
        '商品利潤': 'sum',
        '訂單編號': pd.Series.nunique
    }).rename(columns={'訂單編號': '訂單數'}).reset_index()
    
    # 計算利潤率
    category_performance['利潤率'] = category_performance['商品利潤'] / category_performance['銷售金額'] * 100
    category_performance = category_performance.sort_values('銷售金額', ascending=False)
    
    print("\n===================== 產品類別分析 =====================")
    print(category_performance)
    print("======================================================\n")
    
    return {
        'top_profit_products': top_profit_products,
        'top_margin_products': top_margin_products,
        'bottom_margin_products': bottom_margin_products,
        'category_performance': category_performance
    }


if __name__ == "__main__":
    results = generate_report()