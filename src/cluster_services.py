import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import pandas as pd
from janome.tokenizer import Tokenizer
from sklearn.feature_extraction.text import TfidfVectorizer
from wordcloud import WordCloud

# ファイルパス設定
# clustered_nodes_path = "kcore_clustered_nodes.csv"  # クラスタリング結果
clustered_nodes_path = "filtered_kcore_clustered_nodes.csv"  # クラスタリング結果
investment_info_path = "data/CREI_資金調達情報_出資元_2022_04_21.xlsx"  # 投資情報
services_path = "data/CREI_サービス情報_2022_04_21.xlsx"  # サービス情報

font_path = "ipaexg.ttf"  # フォントファイルのパス
font_prop = fm.FontProperties(fname=font_path)
plt.rcParams['font.family'] = font_prop.get_name()

# データのロード
clustered_nodes = pd.read_csv(clustered_nodes_path)
investment_info = pd.read_excel(investment_info_path, engine='openpyxl')
services_df = pd.read_excel(services_path, engine='openpyxl')

# 必要なデータを文字列型に変換
investment_info['企業ID'] = investment_info['企業ID'].astype(str)
services_df['企業ID'] = services_df['企業ID'].astype(str)

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

# クラスタごとにサービス内容を収集
cluster_services = {}
for cluster_id in sorted(investment_info['新クラスタID'].unique()):
    cluster_investments = investment_info[investment_info['新クラスタID'] == cluster_id]
    target_company_ids = cluster_investments['企業ID'].unique()
    
    # 対応するサービス情報を取得
    cluster_service_texts = services_df[services_df['企業ID'].isin(target_company_ids)]['サービス内容'].dropna().tolist()  # NaNを除外
    cluster_service_texts = [str(text) for text in cluster_service_texts if isinstance(text, str)]  # 文字列以外を除外
    cluster_services[cluster_id] = " ".join(cluster_service_texts)

# 形態素解析用のトークナイザー
tokenizer = Tokenizer()

def tokenize(text):
    return [token.surface for token in tokenizer.tokenize(text) if token.part_of_speech.startswith("名詞")]

# ストップワードの設定
stop_words = ["サービス", "情報", "企業", "提供", "事業", "利用", "活動", "関連", "支援", "プラットフォーム"]

# TF-IDF計算
vectorizer = TfidfVectorizer(
    tokenizer=tokenize,
    stop_words=stop_words,
    max_df=0.8,  # 頻出単語の除外（80%以上のクラスタに出現する単語）
    min_df=0.05, # 稀な単語を除外（5%以上のクラスタに出現する単語）
    use_idf=True,
    smooth_idf=True,
    token_pattern=None
)
tfidf_matrix = vectorizer.fit_transform(cluster_services.values())
terms = vectorizer.get_feature_names_out()

# Word Cloudの作成と上位5単語の頻度可視化
for cluster_id, cluster_text in cluster_services.items():
    cluster_vector = tfidf_matrix[cluster_id].toarray().flatten()
    tfidf_scores = {terms[i]: cluster_vector[i] for i in range(len(terms)) if cluster_vector[i] > 0}
    
    # Word Cloudの生成
    wordcloud = WordCloud(font_path="ipaexg.ttf", background_color="white", width=800, height=400, regexp=r'\b\w+\b').generate_from_frequencies(tfidf_scores)
    
    # Word Cloudの表示
    plt.figure(figsize=(10, 5))
    plt.imshow(wordcloud, interpolation='bilinear')
    plt.axis("off")
    plt.title(f"Cluster {cluster_id} Word Cloud")
    plt.savefig(f"cluster_{cluster_id}_wordcloud.png", bbox_inches='tight')
    print(f"Word Cloud saved for Cluster {cluster_id} as cluster_{cluster_id}_wordcloud.png")
    plt.show()

    # 上位5単語の頻度可視化
    top_5_words = sorted(tfidf_scores.items(), key=lambda x: x[1], reverse=True)[:5]
    words, scores = zip(*top_5_words)
    plt.figure(figsize=(8, 4))
    plt.bar(words, scores, color="skyblue")
    plt.title(f"Top 5 Words in Cluster {cluster_id}", fontproperties=font_prop)
    plt.xlabel("Words", fontproperties=font_prop)
    plt.ylabel("TF-IDF Score", fontproperties=font_prop)
    plt.xticks(fontproperties=font_prop)
    plt.yticks(fontproperties=font_prop)
    plt.savefig(f"cluster_{cluster_id}_top5words.png", bbox_inches='tight')
    print(f"Top 5 Words Bar Chart saved for Cluster {cluster_id} as cluster_{cluster_id}_top5words.png")
    plt.show()
