import networkx as nx
import matplotlib.pyplot as plt
import pandas as pd
import random
from matplotlib import font_manager
from networkx.algorithms.link_prediction import jaccard_coefficient
from community.community_louvain import best_partition

# データ読み込み
file_path = "data/CREI_資金調達情報_出資元_2022_04_21.xlsx"
df = pd.read_excel(file_path, engine='openpyxl')

# フォント設定
font_path = "ipaexg.ttf"
try:
    font_prop = font_manager.FontProperties(fname=font_path)
    plt.rcParams['font.family'] = font_prop.get_name()
    plt.rcParams['axes.unicode_minus'] = False
    print("Font successfully set.")
except Exception as e:
    print(f"Error setting font: {e}")

# グラフ作成
G = nx.Graph()
for funding_id in df['資金調達ID'].unique():
    subset = df[df['資金調達ID'] == funding_id]
    investor_companies = subset['出資元・企業名'].unique()
    for i in range(len(investor_companies)):
        for j in range(i + 1, len(investor_companies)):
            G.add_edge(investor_companies[i], investor_companies[j], weight=1)


# ジャカード係数でエッジを追加（閾値調整でクラスタ数を減らす）
threshold = 0.7  # 閾値をさらに高く設定
edges_to_add = []
for u, v, p in jaccard_coefficient(G):
    if p >= threshold:
        edges_to_add.append((u, v, p))
G.add_weighted_edges_from(edges_to_add)

# Kコア分割でノード数を制限
k = 6  # K値をさらに増加させて小さなクラスタを排除
G = nx.k_core(G, k)

# Louvain法でクラスタリング（解像度パラメータ調整）
resolution = 0.8  # 解像度を調整してクラスタ数を10個程度に近づける
partition = best_partition(G, weight='weight', resolution=resolution)

# クラスタ数の確認
clusters = set(partition.values())
print(f"Number of clusters: {len(clusters)}")

# モジュラリティ計算
modularity = nx.algorithms.community.quality.modularity(
    G, 
    [{n for n in partition if partition[n] == c} for c in clusters]
)
print(f"Modularity of the network: {modularity}")

# クラスタサイズ表示
cluster_sizes = {cluster: sum(1 for n in partition if partition[n] == cluster) for cluster in clusters}
print("\nCluster sizes:")
for cluster, size in cluster_sizes.items():
    print(f"Cluster {cluster}: {size} nodes")


# レイアウト生成
pos = nx.spring_layout(G, k=0.1, seed=42)

# クラスタごとの色設定
clusters = set(partition.values())
cluster_colors = {cluster: f"#{''.join(random.choices('0123456789ABCDEF', k=6))}" for cluster in clusters}

# 可視化
plt.figure(figsize=(15, 10))

# クラスタごとにノードとエッジを描画
for cluster in clusters:
    nodes_in_cluster = [node for node in G.nodes() if partition[node] == cluster]
    edges_in_cluster = [(u, v) for u, v in G.edges() if partition[u] == cluster and partition[v] == cluster]
    color = cluster_colors[cluster]
    nx.draw_networkx_nodes(
        G, pos, nodelist=nodes_in_cluster,
        node_color=color, node_size=300, alpha=0.8,
    )
    nx.draw_networkx_edges(G, pos, edgelist=edges_in_cluster, edge_color=color, alpha=0.5)

nx.draw_networkx_edges(G, pos, alpha=0.2, edge_color="grey")

# 小規模クラスタの削除
min_cluster_size = 20  # 最小クラスタサイズを設定
clusters_to_keep = {cluster for cluster, size in cluster_sizes.items() if size >= min_cluster_size}
filtered_partition = {node: cluster for node, cluster in partition.items() if cluster in clusters_to_keep}

# 更新後のクラスタサイズ表示
updated_clusters = set(filtered_partition.values())
updated_cluster_sizes = {cluster: sum(1 for n in filtered_partition if filtered_partition[n] == cluster) for cluster in updated_clusters}
print("\nUpdated Cluster sizes:")
for cluster, size in updated_cluster_sizes.items():
    print(f"Cluster {cluster}: {size} nodes")

# クラスタごとの企業リストを出力
clustered_nodes = pd.DataFrame({'企業名': list(filtered_partition.keys()), 'クラスタID': list(filtered_partition.values())})
clustered_nodes.to_csv("updated_kcore_clustered_nodes.csv", index=False)
