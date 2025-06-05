import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats
from datetime import datetime
import os
import re

plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']
plt.rcParams['axes.unicode_minus'] = False

def load_and_prepare_data():
    df = pd.read_excel('CPE銷售明細檔20240901-20240930_新北.xlsx')
    return df

def analyze_service_centers(df):
    # 基本銷售分析
    service_center_sales = df.groupby('服務中心名稱').agg({
        '數量': 'sum',
        '金額': 'sum',
        '未稅金額': 'sum'
    }).reset_index()
    
    service_center_sales['平均單價'] = service_center_sales['金額'] / service_center_sales['數量']
    
    # 計算Z-scores
    for col in ['數量', '金額', '平均單價']:
        service_center_sales[f'{col}_zscore'] = stats.zscore(service_center_sales[col])
        service_center_sales[f'{col}_突出'] = service_center_sales[f'{col}_zscore'] > 2
    
    return service_center_sales

def analyze_products(df):
    # 商品分析
    product_analysis = df.groupby('商品名稱').agg({
        '數量': 'sum',
        '金額': 'sum',
        '商品屬性': 'first'  # 保留商品屬性
    }).reset_index()
    product_analysis['平均單價'] = product_analysis['金額'] / product_analysis['數量']
    product_analysis = product_analysis.sort_values('金額', ascending=False)
    
    # 商品類別分析（使用商品屬性）
    category_analysis = df.groupby('商品屬性').agg({
        '數量': 'sum',
        '金額': 'sum'
    }).reset_index()
    category_analysis['平均單價'] = category_analysis['金額'] / category_analysis['數量']
    category_analysis = category_analysis.sort_values('金額', ascending=False)
    
    return product_analysis, category_analysis

def analyze_daily_trends(df):
    # 日期趨勢分析
    df['日期'] = pd.to_datetime(df[['年', '月', '日']].astype(str).agg('-'.join, axis=1))
    daily_sales = df.groupby('日期').agg({
        '數量': 'sum',
        '金額': 'sum'
    }).reset_index()
    
    # 找出最高銷售日
    max_sales_day = daily_sales.loc[daily_sales['金額'].idxmax()]
    
    # 分析最高銷售日的商品
    max_day_products = df[df['日期'] == max_sales_day['日期']].groupby('商品名稱').agg({
        '數量': 'sum',
        '金額': 'sum',
        '商品屬性': 'first'
    }).reset_index()
    max_day_products = max_day_products.sort_values('金額', ascending=False)
    
    return daily_sales, max_sales_day, max_day_products

def create_visualizations(service_center_sales, product_analysis, category_analysis, daily_sales):
    # 創建圖表目錄
    os.makedirs('report_figures', exist_ok=True)
    
    # 1. 服務中心銷售分析圖
    plt.figure(figsize=(15, 8))
    plt.subplot(2, 1, 1)
    sns.barplot(data=service_center_sales.sort_values('數量', ascending=False), 
                x='服務中心名稱', y='數量')
    plt.xticks(rotation=45, ha='right')
    plt.title('各服務中心銷售數量比較')
    
    plt.subplot(2, 1, 2)
    sns.barplot(data=service_center_sales.sort_values('金額', ascending=False), 
                x='服務中心名稱', y='金額')
    plt.xticks(rotation=45, ha='right')
    plt.title('各服務中心銷售金額比較')
    plt.tight_layout()
    plt.savefig('report_figures/service_center_analysis.png')
    plt.close()
    
    # 2. 商品類別分析圖
    plt.figure(figsize=(12, 6))
    sns.barplot(data=category_analysis, x='商品屬性', y='金額')
    plt.title('各商品類別銷售金額比較')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig('report_figures/category_analysis.png')
    plt.close()
    
    # 3. 商品類別數量分析圖
    plt.figure(figsize=(12, 6))
    sns.barplot(data=category_analysis, x='商品屬性', y='數量')
    plt.title('各商品類別銷售數量比較')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig('report_figures/category_quantity_analysis.png')
    plt.close()
    
    # 4. 日銷售趨勢圖
    plt.figure(figsize=(15, 6))
    plt.plot(daily_sales['日期'], daily_sales['金額'], marker='o')
    plt.title('每日銷售金額趨勢')
    plt.xticks(rotation=45)
    plt.grid(True)
    plt.tight_layout()
    plt.savefig('report_figures/daily_sales_trend.png')
    plt.close()

def generate_html_report(service_center_sales, product_analysis, category_analysis, daily_sales, max_sales_day, max_day_products, df):
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>銷售分析報告</title>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 40px; }}
            .container {{ max-width: 1200px; margin: 0 auto; }}
            table {{ border-collapse: collapse; width: 100%; margin: 20px 0; }}
            th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
            th {{ background-color: #f2f2f2; }}
            img {{ max-width: 100%; height: auto; margin: 20px 0; }}
            .section {{ margin: 40px 0; }}
            h1, h2 {{ color: #333; }}
            .highlight {{ background-color: #fff3cd; }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>銷售分析報告</h1>
            <p>報告生成時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
            
            <div class="section">
                <h2>1. 服務中心分析</h2>
                <img src="report_figures/service_center_analysis.png" alt="服務中心分析">
                <h3>表現突出的服務中心：</h3>
                <table>
                    <tr><th>服務中心</th><th>數量</th><th>Z-score</th></tr>
                    {service_center_sales[service_center_sales['數量_突出']].apply(lambda row: f"<tr><td>{row['服務中心名稱']}</td><td>{row['數量']}</td><td>{row['數量_zscore']:.2f}</td></tr>", axis=1).str.cat(sep='')}
                </table>
                
                <h3>銷售金額TOP5：</h3>
                <table>
                    <tr><th>服務中心</th><th>金額</th></tr>
                    {service_center_sales.nlargest(5, '金額').apply(lambda row: f"<tr><td>{row['服務中心名稱']}</td><td>{row['金額']:,.0f}</td></tr>", axis=1).str.cat(sep='')}
                </table>
            </div>
            
            <div class="section">
                <h2>2. 商品分析</h2>
                <h3>各商品類別銷售情況：</h3>
                <img src="report_figures/category_analysis.png" alt="商品類別金額分析">
                <img src="report_figures/category_quantity_analysis.png" alt="商品類別數量分析">
                <table>
                    <tr><th>商品類別</th><th>銷售金額</th><th>銷售數量</th><th>平均單價</th></tr>
                    {category_analysis.apply(lambda row: f"<tr><td>{row['商品屬性']}</td><td>{row['金額']:,.0f}</td><td>{row['數量']}</td><td>{row['平均單價']:,.0f}</td></tr>", axis=1).str.cat(sep='')}
                </table>
                
                <h3>銷售額TOP20商品：</h3>
                <table>
                    <tr><th>商品名稱</th><th>商品類別</th><th>銷售金額</th><th>銷售數量</th><th>平均單價</th></tr>
                    {product_analysis.head(20).apply(lambda row: f"<tr><td>{row['商品名稱']}</td><td>{row['商品屬性']}</td><td>{row['金額']:,.0f}</td><td>{row['數量']}</td><td>{row['平均單價']:,.0f}</td></tr>", axis=1).str.cat(sep='')}
                </table>
            </div>
            
            <div class="section">
                <h2>3. 銷售趨勢分析</h2>
                <img src="report_figures/daily_sales_trend.png" alt="日銷售趨勢">
                <h3>最高銷售日分析：</h3>
                <p>日期：{max_sales_day['日期'].strftime('%Y-%m-%d')}</p>
                <p>銷售金額：{max_sales_day['金額']:,.0f}</p>
                <p>銷售數量：{max_sales_day['數量']}</p>
                
                <h4>當日各門市銷售金額排名：</h4>
                <table>
                    <tr><th>服務中心</th><th>銷售金額</th><th>銷售數量</th></tr>
                    {df[df['日期'] == max_sales_day['日期']].groupby('服務中心名稱').agg({
                        '金額': 'sum',
                        '數量': 'sum'
                    }).sort_values('金額', ascending=False).apply(lambda row: f"<tr><td>{row.name}</td><td>{row['金額']:,.0f}</td><td>{row['數量']:,}</td></tr>", axis=1).str.cat(sep='')}
                </table>
                
                <h4>當日熱門商品TOP5：</h4>
                <table>
                    <tr><th>商品名稱</th><th>商品類別</th><th>銷售金額</th><th>銷售數量</th></tr>
                    {max_day_products.head(5).apply(lambda row: f"<tr><td>{row['商品名稱']}</td><td>{row['商品屬性']}</td><td>{row['金額']:,.0f}</td><td>{row['數量']}</td></tr>", axis=1).str.cat(sep='')}
                </table>
            </div>
            
            <div class="section">
                <h2>4. 統計摘要</h2>
                <h3>4.1 服務中心統計（基於40個服務中心）</h3>
                <table>
                    <tr><th>指標</th><th>銷售數量</th><th>銷售金額</th><th>平均單價</th></tr>
                    {service_center_sales[['數量', '金額', '平均單價']].describe().apply(lambda row: f"<tr><td>{row.name}</td><td>{row['數量']:,.1f}</td><td>{row['金額']:,.1f}</td><td>{row['平均單價']:,.1f}</td></tr>", axis=1).str.cat(sep='')}
                </table>
                
                <h3>4.2 商品類別統計</h3>
                <table>
                    <tr><th>商品類別</th><th>銷售數量</th><th>銷售金額</th><th>平均單價</th></tr>
                    {category_analysis.apply(lambda row: f"<tr><td>{row['商品屬性']}</td><td>{row['數量']:,}</td><td>{row['金額']:,.0f}</td><td>{row['平均單價']:,.0f}</td></tr>", axis=1).str.cat(sep='')}
                </table>
                
                <h3>4.3 總體銷售統計</h3>
                <table>
                    <tr><th>指標</th><th>數值</th></tr>
                    <tr><td>總銷售數量</td><td>{category_analysis['數量'].sum():,}</td></tr>
                    <tr><td>總銷售金額</td><td>{category_analysis['金額'].sum():,.0f}</td></tr>
                    <tr><td>平均單價</td><td>{category_analysis['金額'].sum() / category_analysis['數量'].sum():,.0f}</td></tr>
                    <tr><td>商品種類數</td><td>{len(product_analysis):,}</td></tr>
                    <tr><td>服務中心數量</td><td>{len(service_center_sales):,}</td></tr>
                </table>
            </div>
        </div>
    </body>
    </html>
    """
    
    with open('sales_analysis_report.html', 'w', encoding='utf-8') as f:
        f.write(html_content)

def main():
    # 載入數據
    df = load_and_prepare_data()
    
    # 執行各項分析
    service_center_sales = analyze_service_centers(df)
    product_analysis, category_analysis = analyze_products(df)
    daily_sales, max_sales_day, max_day_products = analyze_daily_trends(df)
    
    # 創建視覺化
    create_visualizations(service_center_sales, product_analysis, category_analysis, daily_sales)
    
    # 生成HTML報告
    generate_html_report(service_center_sales, product_analysis, category_analysis, daily_sales, max_sales_day, max_day_products, df)
    
    print("分析報告已生成完成！請查看 sales_analysis_report.html")

if __name__ == "__main__":
    main() 