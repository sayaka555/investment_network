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
font_prop = font_manager.FontProperties(fname=font_path)
plt.rcParams['font.family'] = font_prop.get_name()
plt.rcParams['axes.unicode_minus'] = False
print("font set")

# グラフ作成
G = nx.Graph()
for funding_id in df['資金調達ID'].unique():
    subset = df[df['資金調達ID'] == funding_id]
    investor_companies = subset['出資元・企業名'].unique()
    for i in range(len(investor_companies)):
        for j in range(i + 1, len(investor_companies)):
            G.add_edge(investor_companies[i], investor_companies[j], weight=1)

# ジャカード係数でエッジを追加（計算コスト削減）
edges_to_add = []
for u, v, p in jaccard_coefficient(G):
    if p >= 0.3:
        edges_to_add.append((u, v, p))
G.add_weighted_edges_from(edges_to_add)

# Kコア分割でノード数を制限
k = 3
G = nx.k_core(G, k)

# Louvain法でクラスタリング
partition = best_partition(G, weight='weight')

# モジュラリティ計算
modularity = nx.algorithms.community.quality.modularity(G, [{n for n in partition if partition[n] == c} for c in set(partition.values())])
print(f"Modularity of the network: {modularity}")

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
        node_color=color, node_size=300, alpha=0.8
    )
    nx.draw_networkx_edges(G, pos, edgelist=edges_in_cluster, edge_color=color, alpha=0.5)

nx.draw_networkx_edges(G, pos, alpha=0.2, edge_color="grey")

# 中心性が高い上位10ノードのみラベル表示
degree_centrality = nx.degree_centrality(G)
top_labels = sorted(degree_centrality, key=degree_centrality.get, reverse=True)[:10]
nx.draw_networkx_labels(G, pos, labels={node: node for node in top_labels}, font_size=8, font_weight="bold", font_family="sans-serif")

plt.title("Investor Network (K-Core, Colorful Edges and Clusters)")
plt.axis("off")
plt.show()

# クラスタサイズ表示
cluster_sizes = {cluster: sum(1 for n in partition if partition[n] == cluster) for cluster in clusters}
print("\nCluster sizes:")
for cluster, size in cluster_sizes.items():
    print(f"Cluster {cluster}: {size} nodes")

# 中心性指標の計算
eigenvector_centrality = nx.eigenvector_centrality(G, max_iter=500)
betweenness_centrality = nx.betweenness_centrality(G)
closeness_centrality = nx.closeness_centrality(G)

# 中心性指標の可視化
plt.figure(figsize=(10, 6))
plt.bar(range(len(degree_centrality)), sorted(degree_centrality.values(), reverse=True), label="Degree Centrality")
plt.bar(range(len(betweenness_centrality)), sorted(betweenness_centrality.values(), reverse=True), label="Betweenness Centrality", alpha=0.7)
plt.bar(range(len(closeness_centrality)), sorted(closeness_centrality.values(), reverse=True), label="Closeness Centrality", alpha=0.5)
plt.title("Centrality Measures")
plt.xlabel("Nodes (sorted)")
plt.ylabel("Centrality Value")
plt.legend()
plt.show()

# クラスタごとの企業リストを出力
clustered_nodes = pd.DataFrame({'企業名': list(partition.keys()), 'クラスタID': list(partition.values())})
clustered_nodes.to_csv("kcore_clustered_nodes.csv", index=False)

# クラスタごとに表示
for cluster in sorted(clusters):
    cluster_df = clustered_nodes[clustered_nodes['クラスタID'] == cluster]
    print(f"\nクラスタ {cluster} の企業リスト:")
    print(cluster_df['企業名'].to_string(index=False))
