import streamlit as st
from datetime import datetime

# 作成したモジュールから関数をインポート
from data_processor import load_data, process_data
from html_generator import generate_html

# CSV出力機能をインポート
try:
    from gross_profit_metrics_exporter import create_gross_profit_metrics_export_interface
    CSV_EXPORT_AVAILABLE = True
except ImportError:
    CSV_EXPORT_AVAILABLE = False

# --- ページ設定 ---
st.set_page_config(
    page_title="粗利達成率 HTMLレポートジェネレーター",
    page_icon="📄",
    layout="wide"
)

# --- メイン画面 ---
st.title("📄 粗利達成率 インタラクティブレポート生成ツール")
st.markdown("粗利の目標と実績ファイルをアップロードすると、インタラクティブなHTMLレポートをダウンロードできます。")

# CSV出力機能が利用可能な場合は説明を追加
if CSV_EXPORT_AVAILABLE:
    st.markdown("💡 **NEW!** データ処理後にCSVメトリクス出力機能が利用できます（ポータル統合用）")

st.markdown("---")

# --- サイドバー ---
st.sidebar.header("📁 データアップロード")

target_file = st.sidebar.file_uploader(
    "粗利目標ファイル",
    type=['xlsx', 'xls', 'csv'],
    help="診療科別の目標値を含むExcel/CSVファイル"
)

actual_file = st.sidebar.file_uploader(
    "粗利実績ファイル",
    type=['xlsx', 'xls', 'csv'],
    help="診療科別の月次実績値を含むExcel/CSVファイル"
)

# Google Analytics IDの入力欄をサイドバーに追加
st.sidebar.markdown("---")
st.sidebar.header("⚙️ レポート設定")
google_analytics_id = st.sidebar.text_input(
    "Google Analytics ID (任意)",
    help="HTMLにトラッキングコードを埋め込む場合に入力します。例: G-K6XTL1DM13"
)

# --- メイン処理 ---
if target_file and actual_file:
    # 1. データの読み込みと処理
    with st.spinner("データを読み込み、集計しています..."):
        target_df = load_data(target_file)
        actual_df = load_data(actual_file)
        
        summary_df, chart_df = process_data(target_df, actual_df, today=datetime.now())

    # 2. 処理結果の確認
    if not summary_df.empty and not chart_df.empty:
        st.success("✅ データ処理が完了しました。")
        
        # セッションステートにデータを保存（CSV出力用）
        st.session_state['gross_profit_summary_df'] = summary_df
        st.session_state['gross_profit_chart_df'] = chart_df
        
        # タブで機能を分割
        if CSV_EXPORT_AVAILABLE:
            tab1, tab2 = st.tabs(["📊 HTMLレポート生成", "📋 CSVメトリクス出力"])
        else:
            tab1 = st.container()
        
        # HTMLレポート生成タブ
        with tab1 if CSV_EXPORT_AVAILABLE else tab1:
            st.markdown("#### ▼ プレビュー（トップページ表示予定のカード）")
            
            # 画面にサマリーデータを表示して確認
            st.dataframe(summary_df, use_container_width=True)
            st.markdown("---")

            # 3. HTMLファイルの生成
            with st.spinner("インタラクティブHTMLを生成しています..."):
                final_html = generate_html(
                    summary_df, 
                    chart_df, 
                    google_analytics_id=google_analytics_id
                )
            
            st.success("✅ HTMLレポートの準備ができました。")
            
            # 4. ダウンロードボタンの表示
            st.download_button(
                label="📥 HTMLレポートをダウンロード",
                data=final_html,
                file_name=f"index.html",
                mime="text/html"
            )
            
            # データ概要表示
            col1, col2, col3 = st.columns(3)
            with col1:
                total_depts = len(summary_df)
                st.metric("総診療科数", total_depts)
            with col2:
                achieved_depts = len(summary_df[summary_df['直近月達成率'] >= 100])
                st.metric("目標達成診療科", achieved_depts)
            with col3:
                avg_achievement = summary_df['直近月達成率'].mean()
                st.metric("平均達成率", f"{avg_achievement:.1f}%")
        
        # CSVメトリクス出力タブ
        if CSV_EXPORT_AVAILABLE:
            with tab2:
                st.markdown("""
                ### 📊 ポータル統合用メトリクス出力
                
                粗利分析の主要指標をCSV形式で出力し、診療科別ポータルwebページで統合表示するためのデータを作成します。
                """)
                
                # メトリクス出力インターフェースを表示
                create_gross_profit_metrics_export_interface()

    else:
        st.error("データの処理に失敗しました。ファイルの形式が正しいか、中身が空でないか確認してください。")

else:
    st.info("👆 サイドバーから粗利目標ファイルと粗利実績ファイルをアップロードしてください。")
    
    # データが未アップロードの場合の説明
    with st.expander("📚 使用方法とファイル形式"):
        st.markdown("""
        ### 📁 必要なファイル
        
        **1. 粗利目標ファイル**
        - 診療科名と目標値を含むExcel/CSVファイル
        - 形式例：
        ```
        診療科名,目標粗利
        外科,50000000
        内科,30000000
        小児科,20000000
        ```
        
        **2. 粗利実績ファイル**  
        - 診療科名と月次実績値を含むExcel/CSVファイル
        - 形式例：
        ```
        診療科名,2024-01-01,2024-02-01,2024-03-01,...
        外科,45000000,52000000,48000000,...
        内科,32000000,28000000,35000000,...
        小児科,18000000,22000000,19000000,...
        ```
        
        ### 🎯 出力される機能
        
        **HTMLレポート**
        - インタラクティブな診療科別ダッシュボード
        - 達成率推移グラフ
        - トレンド分析とパフォーマンス評価
        
        **CSVメトリクス出力** *(NEW!)*
        - ポータル統合用の標準化データ
        - 診療科別達成率・全体比率・昨年度比較等
        - 他のアプリ（手術分析・入退院分析）と統合表示可能
        
        ### 💡 活用例
        
        - 月次診療科会議での業績報告
        - 経営陣向けダッシュボード
        - 診療科別改善計画策定
        - 全社ポータルでの統合表示
        """)