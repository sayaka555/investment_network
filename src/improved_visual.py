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
print("Font set")

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

# クラスタ統計情報の計算
cluster_stats = []
for cluster in set(partition.values()):
    nodes_in_cluster = [node for node in G.nodes() if partition[node] == cluster]
    subgraph = G.subgraph(nodes_in_cluster)  # クラスタの部分グラフを抽出
    
    # ノード数とエッジ数
    num_nodes = subgraph.number_of_nodes()
    num_edges = subgraph.number_of_edges()
    
    # 中心性指標
    degree_centrality = nx.degree_centrality(subgraph)
    betweenness_centrality = nx.betweenness_centrality(subgraph)
    closeness_centrality = nx.closeness_centrality(subgraph)
    eigenvector_centrality = nx.eigenvector_centrality(subgraph, max_iter=500)
    
    cluster_stats.append({
        'Cluster': cluster,
        'Nodes': num_nodes,
        'Edges': num_edges,
        'Average Degree Centrality': sum(degree_centrality.values()) / num_nodes,
        'Average Betweenness Centrality': sum(betweenness_centrality.values()) / num_nodes,
        'Average Closeness Centrality': sum(closeness_centrality.values()) / num_nodes,
        'Average Eigenvector Centrality': sum(eigenvector_centrality.values()) / num_nodes,
    })

# クラスタ統計情報をデータフレームに変換して表示
cluster_stats_df = pd.DataFrame(cluster_stats)
print(cluster_stats_df)

# CSVに保存
cluster_stats_df.to_csv("cluster_stats.csv", index=False)
print("Cluster statistics saved as cluster_stats.csv")

# 可視化
plt.figure(figsize=(10, 6))
plt.bar(cluster_stats_df['Cluster'], cluster_stats_df['Nodes'], color='skyblue', alpha=0.8)
plt.title("Number of Nodes per Cluster", fontproperties=font_prop)
plt.xlabel("Cluster", fontproperties=font_prop)
plt.ylabel("Number of Nodes", fontproperties=font_prop)
plt.xticks(fontproperties=font_prop)
plt.yticks(fontproperties=font_prop)
plt.grid(axis='y', linestyle='--', alpha=0.7)
plt.savefig("nodes_per_cluster.png", bbox_inches='tight')
plt.show()

plt.figure(figsize=(10, 6))
plt.bar(cluster_stats_df['Cluster'], cluster_stats_df['Edges'], color='orange', alpha=0.8)
plt.title("Number of Edges per Cluster", fontproperties=font_prop)
plt.xlabel("Cluster", fontproperties=font_prop)
plt.ylabel("Number of Edges", fontproperties=font_prop)
plt.xticks(fontproperties=font_prop)
plt.yticks(fontproperties=font_prop)
plt.grid(axis='y', linestyle='--', alpha=0.7)
plt.savefig("edges_per_cluster.png", bbox_inches='tight')
plt.show()
