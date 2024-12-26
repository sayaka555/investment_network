import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as font_manager
import numpy as np

# ファイルパス設定
# clustered_nodes_path = "kcore_clustered_nodes.csv"  # クラスタリング結果
clustered_nodes_path = "filtered_kcore_clustered_nodes.csv"  # クラスタリング結果
investment_info_path = "data/CREI_資金調達情報_出資元_2022_04_21.xlsx"  # 投資情報
financials_path = "data/CREI_決算情報_2022_04_21.xlsx"  # 売上情報

# フォント設定
# font_path = "ipaexg.ttf"
# font_prop = font_manager.FontProperties(fname=font_path)
# plt.rcParams['font.family'] = font_prop.get_name()
# plt.rcParams['axes.unicode_minus'] = False

# データのロード
clustered_nodes = pd.read_csv(clustered_nodes_path)
investment_info = pd.read_excel(investment_info_path, engine='openpyxl')
financials = pd.read_excel(financials_path, engine='openpyxl')

# 必要なデータを文字列型に変換
investment_info['企業ID'] = investment_info['企業ID'].astype(str)
investment_info['出資元・企業ID'] = investment_info['出資元・企業ID'].astype(str)
financials['企業ID'] = financials['企業ID'].astype(str)

# クラスタリング結果を結合
investment_info = pd.merge(
    investment_info,
    clustered_nodes,
    left_on="出資元・企業名",  # 投資元企業の名前
    right_on="企業名",  # クラスタリング結果の企業名
    how="inner"
)

# クラスタIDを0から再マッピング
unique_clusters = sorted(investment_info['クラスタID'].unique())
cluster_id_map = {old_id: new_id for new_id, old_id in enumerate(unique_clusters)}
investment_info['新クラスタID'] = investment_info['クラスタID'].map(cluster_id_map)

# クラスタごとに売上統計情報を計算
cluster_sales_stats = {}

for cluster_id in sorted(investment_info['新クラスタID'].unique()):
    # 該当クラスタの投資元が投資している企業のIDを取得
    cluster_investments = investment_info[investment_info['新クラスタID'] == cluster_id]
    target_company_ids = cluster_investments['企業ID'].unique()
    
    # 売上データを取得
    target_financials = financials[financials['企業ID'].isin(target_company_ids)]
    
    # 売上統計情報を計算
    if not target_financials.empty:
        cluster_sales_stats[cluster_id] = {
            '平均売上': target_financials['売上'].mean(),
            '中央値売上': target_financials['売上'].median(),
            '最大売上': target_financials['売上'].max(),
            '最小売上': target_financials['売上'].min(),
            '合計売上': target_financials['売上'].sum()
        }

# 統計情報をソートして表示
sorted_stats = sorted(cluster_sales_stats.items())
for cluster_id, stats in sorted_stats:
    print(f"Cluster {cluster_id}:")
    print(stats)
    print()

# 売上分布をクラスタごとに描画
for cluster_id, stats in sorted_stats:
    cluster_investments = investment_info[investment_info['新クラスタID'] == cluster_id]
    target_company_ids = cluster_investments['企業ID'].unique()
    target_financials = financials[financials['企業ID'].isin(target_company_ids)]
    
    if not target_financials.empty:
        plt.figure(figsize=(10, 5))
        
        # 売上データを取得しログスケールに変換（0以上のデータのみ）
        sales_data = target_financials['売上'].dropna()
        log_sales_data = sales_data[sales_data > 0].apply(np.log10)
        
        # ヒストグラム描画
        plt.hist(log_sales_data, bins=20, alpha=0.7, color='blue')
        plt.title(f"Cluster {cluster_id} Sales Distribution")
        plt.xlabel("Log10(Sales)")
        plt.ylabel("Number of Companies")
        plt.grid(axis='y', alpha=0.75)
        
        # 保存と表示
        filename = f"cluster_{cluster_id}_sales_distribution.png"
        plt.savefig(filename, bbox_inches='tight')  # 保存時に余白を削除
        plt.close()
