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


# 都道府県リストの作成（日本の都道府県名）
prefectures = [
    '北海道', '青森県', '岩手県', '宮城県', '秋田県', '山形県', '福島県',
    '茨城県', '栃木県', '群馬県', '埼玉県', '千葉県', '東京都', '神奈川県',
    '新潟県', '富山県', '石川県', '福井県', '山梨県', '長野県',
    '岐阜県', '静岡県', '愛知県', '三重県',
    '滋賀県', '京都府', '大阪府', '兵庫県', '奈良県', '和歌山県',
    '鳥取県', '島根県', '岡山県', '広島県', '山口県',
    '徳島県', '香川県', '愛媛県', '高知県',
    '福岡県', '佐賀県', '長崎県', '熊本県', '大分県', '宮崎県', '鹿児島県', '沖縄県'
]

# 都道府県から地域へのマッピング
prefecture_to_region = {
    '北海道': '北海道地方',
    '青森県': '東北地方', '岩手県': '東北地方', '宮城県': '東北地方', '秋田県': '東北地方', '山形県': '東北地方', '福島県': '東北地方',
    '茨城県': '関東地方', '栃木県': '関東地方', '群馬県': '関東地方', '埼玉県': '関東地方', '千葉県': '関東地方', '東京都': '関東地方', '神奈川県': '関東地方',
    '新潟県': '中部地方', '富山県': '中部地方', '石川県': '中部地方', '福井県': '中部地方', '山梨県': '中部地方', '長野県': '中部地方',
    '岐阜県': '中部地方', '静岡県': '中部地方', '愛知県': '中部地方',
    '三重県': '近畿地方', '滋賀県': '近畿地方', '京都府': '近畿地方', '大阪府': '近畿地方', '兵庫県': '近畿地方', '奈良県': '近畿地方', '和歌山県': '近畿地方',
    '鳥取県': '中国地方', '島根県': '中国地方', '岡山県': '中国地方', '広島県': '中国地方', '山口県': '中国地方',
    '徳島県': '四国地方', '香川県': '四国地方', '愛媛県': '四国地方', '高知県': '四国地方',
    '福岡県': '九州・沖縄地方', '佐賀県': '九州・沖縄地方', '長崎県': '九州・沖縄地方', '熊本県': '九州・沖縄地方', '大分県': '九州・沖縄地方',
    '宮崎県': '九州・沖縄地方', '鹿児島県': '九州・沖縄地方', '沖縄県': '九州・沖縄地方'
}

# 住所から都道府県を抽出する関数
def extract_prefecture(address):
    for pref in prefectures:
        if isinstance(address, str) and pref in address:
            return pref
    return 'その他'

# 都道府県から地域を取得する関数
def get_region(prefecture):
    return prefecture_to_region.get(prefecture, 'その他')


# クラスタごとに地域の統計情報を計算
cluster_regions = {}

for cluster_id in sorted(investment_info['新クラスタID'].unique()):
    cluster_investments = investment_info[investment_info['新クラスタID'] == cluster_id]
    target_company_ids = cluster_investments['企業ID'].unique()
    
    # 対応する企業情報を取得
    target_company_info = company_info[company_info['企業ID'].isin(target_company_ids)]
    
    # 住所から都道府県を抽出し、地域を取得
    target_company_info['都道府県'] = target_company_info['住所'].apply(extract_prefecture)
    target_company_info['地域'] = target_company_info['都道府県'].apply(get_region)
    
    if not target_company_info.empty:
        region_counts = target_company_info['地域'].value_counts().to_dict()
        cluster_regions[cluster_id] = region_counts

# 結果の表示
for cluster_id, regions in cluster_regions.items():
    print(f"Cluster {cluster_id}:")
    for region, count in regions.items():
        print(f"  {region}: {count}")
    print()

# クラスタごとの地域分布を視覚化
for cluster_id, regions in cluster_regions.items():
    regions_list = list(regions.keys())
    counts = list(regions.values())
    
    plt.figure(figsize=(8, 6))
    plt.bar(regions_list, counts, color='skyblue', alpha=0.8)
    plt.title(f"Cluster {cluster_id} の地域分布", fontproperties=font_prop)
    plt.xlabel("地域", fontproperties=font_prop)
    plt.ylabel("企業数", fontproperties=font_prop)
    plt.xticks(rotation=45, fontproperties=font_prop)
    plt.yticks(fontproperties=font_prop)
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    
    # 保存と表示
    filename = f"cluster_{cluster_id}_regions.png"
    plt.savefig(filename, bbox_inches='tight')
    print(f"Cluster {cluster_id} の地域分布を {filename} に保存しました")
    plt.show()



# 全ての会社の立地場所の分布を、棒グラフで可視化してください
print("distribution of all companies")
company_info['都道府県'] = company_info['住所'].apply(extract_prefecture)
company_info['地域'] = company_info['都道府県'].apply(get_region)
region_counts = company_info['地域'].value_counts().to_dict()
print(region_counts)

plt.figure(figsize=(8, 6))
plt.bar(region_counts.keys(), region_counts.values(), color='skyblue', alpha=0.8)
plt.title("全企業の地域分布", fontproperties=font_prop)
plt.xlabel("地域", fontproperties=font_prop)
plt.ylabel("企業数", fontproperties=font_prop)
plt.xticks(rotation=45, fontproperties=font_prop)
plt.yticks(fontproperties=font_prop)
plt.grid(axis='y', linestyle='--', alpha=0.7)
plt.show()