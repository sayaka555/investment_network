import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as font_manager

# ファイルパス設定
# clustered_nodes_path = "kcore_clustered_nodes.csv"  # クラスタリング結果
clustered_nodes_path = "filtered_kcore_clustered_nodes.csv"  # クラスタリング結果
investment_info_path = "data/CREI_資金調達情報_出資元_2022_04_21.xlsx"  # 投資情報
company_info_path = "data/CREI_企業一覧_2022_04_21.xlsx"  # 企業一覧

# フォント設定
font_path = "ipaexg.ttf"
font_prop = font_manager.FontProperties(fname=font_path)
plt.rcParams['font.family'] = font_prop.get_name()
plt.rcParams['axes.unicode_minus'] = False
print("font set")

# データのロード
clustered_nodes = pd.read_csv(clustered_nodes_path)
investment_info = pd.read_excel(investment_info_path, engine='openpyxl')
company_info = pd.read_excel(company_info_path, engine='openpyxl')

# 必要なデータを文字列型に変換
investment_info['企業ID'] = investment_info['企業ID'].astype(str)
company_info['企業ID'] = company_info['企業ID'].astype(str)

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

# クラスタごとに上場区分の統計情報を計算
listing_stats = {}

for cluster_id in sorted(investment_info['新クラスタID'].unique()):
    cluster_investments = investment_info[investment_info['新クラスタID'] == cluster_id]
    target_company_ids = cluster_investments['企業ID'].unique()
    
    # 対応する企業情報を取得
    target_company_info = company_info[company_info['企業ID'].isin(target_company_ids)]
    
    if not target_company_info.empty:
        listing_stats[cluster_id] = target_company_info['上場区分'].value_counts().to_dict()

# 統計情報の表示
for cluster_id, stats in listing_stats.items():
    print(f"Cluster {cluster_id}:")
    for market, count in stats.items():
        print(f"  {market}: {count}")
    print()

# クラスタごとの上場区分の分布を視覚化
for cluster_id, stats in listing_stats.items():
    markets = list(stats.keys())
    counts = list(stats.values())
    
    plt.figure(figsize=(8, 5))
    plt.bar(markets, counts, color='skyblue', alpha=0.8)
    plt.title(f"Cluster {cluster_id} - 上場区分分布", fontproperties=font_prop)
    plt.xlabel("市場", fontproperties=font_prop)
    plt.ylabel("企業数", fontproperties=font_prop)
    plt.xticks(rotation=45, ha='right', fontproperties=font_prop)
    plt.yticks(fontproperties=font_prop)
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    
    # グラフの保存と表示
    filename = f"cluster_{cluster_id}_listing_status.png"
    plt.savefig(filename, bbox_inches='tight')
    print(f"Cluster {cluster_id} の上場区分分布のグラフを {filename} に保存しました")
    plt.show()
