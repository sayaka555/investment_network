import pandas as pd
import folium  # 地図表示用
import random
# ファイルパス設定
# clustered_nodes_path = "kcore_clustered_nodes.csv"  # クラスタリング結果
clustered_nodes_path = "updated_kcore_clustered_nodes.csv"  # クラスタリング結果
geocode_path = "data/VC_address_geocode.csv"  # 企業の地理情報

# データのロード
clustered_nodes = pd.read_csv(clustered_nodes_path)
geocode_df = pd.read_csv(geocode_path)

# クラスタごとの色を割り当て
clusters = sorted(clustered_nodes['クラスタID'].unique())
cluster_colors = {cluster: f"#{''.join(random.choices('0123456789ABCDEF', k=6))}" for cluster in clusters}

# Foliumマップ作成
m = folium.Map(location=[35.6895, 139.6917], zoom_start=11)  # 東京を中心に地図を初期化

# 企業ごとに地図上にプロット
for _, row in clustered_nodes.iterrows():
    company_name = row['企業名']
    cluster_id = row['クラスタID']
    node_color = cluster_colors[cluster_id]

    # 対応する地理情報を取得
    node_geocode = geocode_df[geocode_df['company_name'] == company_name]
    if not node_geocode.empty:
        lon = node_geocode['lon'].iloc[0]
        lat = node_geocode['lat'].iloc[0]
        folium.CircleMarker(
            location=[lat, lon],
            radius=5,
            color=node_color,
            fill=True,
            fill_color=node_color,
            fill_opacity=0.7,
            popup=folium.Popup(f"{company_name} (Cluster {cluster_id})", max_width=300)
        ).add_to(m)

# 地図を保存または表示
map_file = "clustered_nodes_map.html"
m.save(map_file)
print(f"Map has been saved as {map_file}. Open it in your browser to view.")
