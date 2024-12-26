import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# File paths
clustered_nodes_path = "kcore_clustered_nodes.csv"  # Clustered result
investment_info_path = "data/CREI_資金調達情報_出資元_2022_04_21.xlsx"  # Investment data
company_info_path = "data/CREI_企業一覧_2022_04_21.xlsx"  # Company list

# Load data
clustered_nodes = pd.read_csv(clustered_nodes_path)
investment_info = pd.read_excel(investment_info_path, engine='openpyxl')
company_info = pd.read_excel(company_info_path, engine='openpyxl')

# Convert necessary columns to string
investment_info['企業ID'] = investment_info['企業ID'].astype(str)
company_info['企業ID'] = company_info['企業ID'].astype(str)

# Merge clustering results with investment information
investment_info = pd.merge(
    investment_info,
    clustered_nodes,
    left_on="出資元・企業名",  # Investor company name
    right_on="企業名",  # Clustered company name
    how="inner"
)

# Remap cluster IDs to start from 0
unique_clusters = sorted(investment_info['クラスタID'].unique())
cluster_id_map = {old_id: new_id for new_id, old_id in enumerate(unique_clusters)}
investment_info['New Cluster ID'] = investment_info['クラスタID'].map(cluster_id_map)

# Calculate statistics per cluster
cluster_stats = {}

for cluster_id in sorted(investment_info['New Cluster ID'].unique()):
    cluster_investments = investment_info[investment_info['New Cluster ID'] == cluster_id]
    target_company_ids = cluster_investments['企業ID'].unique()
    
    # Retrieve corresponding company information
    target_company_info = company_info[company_info['企業ID'].isin(target_company_ids)]
    
    if not target_company_info.empty:
        cluster_stats[cluster_id] = {
            'Listing Status': target_company_info['上場区分'].value_counts().to_dict(),
            'Avg Employees': target_company_info['従業員数'].mean(),
            'Max Employees': target_company_info['従業員数'].max(),
            'Min Employees': target_company_info['従業員数'].min(),
            'Total Funding (M JPY)': target_company_info['合計資金調達額（百万円）'].sum(),
            'Avg Valuation': target_company_info['評価額'].mean(),
            'Max Valuation': target_company_info['評価額'].max(),
            'Min Valuation': target_company_info['評価額'].min()
        }

# Display sorted results
sorted_stats = sorted(cluster_stats.items())
for cluster_id, stats in sorted_stats:
    print(f"Cluster {cluster_id}:")
    for key, value in stats.items():
        print(f"  {key}: {value}")
    print()

# Visualize statistics with bar plots
for metric, metric_label in [
    ('Avg Employees', 'Average Number of Employees'),
    ('Total Funding (M JPY)', 'Total Funding (Million JPY)'),
    ('Avg Valuation', 'Average Valuation')
]:
    metric_values = [stats[metric] for cluster_id, stats in sorted_stats if metric in stats]
    cluster_ids = [cluster_id for cluster_id, stats in sorted_stats if metric in stats]
    
    plt.figure(figsize=(10, 6))
    plt.bar(cluster_ids, metric_values, color='skyblue', alpha=0.8)
    plt.title(f"{metric_label} by Cluster")
    plt.xlabel("Cluster ID")
    plt.ylabel(metric_label)
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    
    # Save and display
    filename = f"{metric.replace(' ', '_')}_by_cluster.png"
    plt.savefig(filename, bbox_inches='tight')
    print(f"{metric_label} graph saved as {filename}")
    plt.show()
