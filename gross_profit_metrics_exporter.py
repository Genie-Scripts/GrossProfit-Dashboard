# gross_profit_metrics_exporter.py
"""
粗利分析メトリクスCSV出力モジュール
ポータル統合用の標準化されたCSVデータを出力
"""

import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
from typing import Dict, Any, List, Optional, Tuple
import logging
import io
import numpy as np

# 既存のデータ処理モジュールからインポート
from data_processor import process_data

logger = logging.getLogger(__name__)

class GrossProfitMetricsExporter:
    """粗利分析メトリクス出力クラス"""
    
    def __init__(self):
        self.app_name = "粗利分析"
        self.version = "1.0"
        
    def export_metrics_csv(
        self,
        summary_df: pd.DataFrame,
        chart_df: pd.DataFrame,
        analysis_date: datetime = None,
        period_type: str = "月次"
    ) -> Tuple[pd.DataFrame, str]:
        """
        メトリクスデータをCSV形式で出力
        
        Args:
            summary_df: 粗利分析のサマリーデータ
            chart_df: 粗利分析のチャートデータ
            analysis_date: 分析基準日
            period_type: 期間タイプ
            
        Returns:
            Tuple[pd.DataFrame, str]: (メトリクスデータフレーム, ファイル名)
        """
        try:
            if analysis_date is None:
                analysis_date = datetime.now()
            
            # 期間情報設定
            period_info = self._calculate_period(analysis_date, period_type, chart_df)
            
            # メトリクス計算
            metrics_data = []
            
            # 1. 全体メトリクス
            overall_metrics = self._calculate_overall_metrics(
                summary_df, chart_df, period_info
            )
            metrics_data.extend(overall_metrics)
            
            # 2. 診療科別メトリクス
            dept_metrics = self._calculate_department_metrics(
                summary_df, period_info
            )
            metrics_data.extend(dept_metrics)
            
            # 3. 時系列メトリクス
            trend_metrics = self._calculate_trend_metrics(
                chart_df, period_info
            )
            metrics_data.extend(trend_metrics)
            
            # データフレーム作成
            metrics_df = pd.DataFrame(metrics_data)
            
            # ファイル名生成
            filename = self._generate_filename(analysis_date, period_type)
            
            return metrics_df, filename
            
        except Exception as e:
            logger.error(f"粗利メトリクス出力エラー: {e}")
            raise
    
    def _calculate_period(self, analysis_date: datetime, period_type: str, chart_df: pd.DataFrame) -> Dict:
        """期間情報を計算"""
        if not chart_df.empty and '月' in chart_df.columns:
            # チャートデータから実際の期間を取得
            chart_df['月'] = pd.to_datetime(chart_df['月'])
            min_date = chart_df['月'].min()
            max_date = chart_df['月'].max()
            
            return {
                "type": period_type,
                "start_date": min_date,
                "end_date": max_date,
                "label": f"{min_date.strftime('%Y年%m月')}〜{max_date.strftime('%Y年%m月')}",
                "latest_month": max_date
            }
        else:
            # フォールバック
            return {
                "type": period_type,
                "start_date": analysis_date,
                "end_date": analysis_date,
                "label": analysis_date.strftime('%Y年%m月'),
                "latest_month": analysis_date
            }
    
    def _calculate_overall_metrics(
        self, 
        summary_df: pd.DataFrame, 
        chart_df: pd.DataFrame,
        period_info: Dict
    ) -> List[Dict]:
        """全体メトリクス計算"""
        metrics = []
        
        if summary_df.empty:
            return metrics
        
        # 総診療科数
        total_depts = len(summary_df)
        metrics.append({
            "診療科名": "全体",
            "メトリクス名": "総診療科数",
            "値": total_depts,
            "単位": "科",
            "期間": period_info["label"],
            "期間タイプ": period_info["type"],
            "カテゴリ": "全体指標",
            "データ種別": "実績",
            "計算日時": datetime.now().isoformat(),
            "アプリ名": self.app_name
        })
        
        # 目標達成診療科数（100%以上）
        achieved_depts = len(summary_df[summary_df['直近月達成率'] >= 100])
        metrics.append({
            "診療科名": "全体",
            "メトリクス名": "目標達成診療科数",
            "値": achieved_depts,
            "単位": "科",
            "期間": period_info["label"],
            "期間タイプ": period_info["type"],
            "カテゴリ": "全体指標",
            "データ種別": "実績",
            "計算日時": datetime.now().isoformat(),
            "アプリ名": self.app_name
        })
        
        # 平均達成率
        avg_achievement = summary_df['直近月達成率'].mean()
        if not pd.isna(avg_achievement):
            metrics.append({
                "診療科名": "全体",
                "メトリクス名": "平均達成率",
                "値": round(avg_achievement, 1),
                "単位": "%",
                "期間": period_info["label"],
                "期間タイプ": period_info["type"],
                "カテゴリ": "全体指標",
                "データ種別": "実績",
                "計算日時": datetime.now().isoformat(),
                "アプリ名": self.app_name
            })
        
        # 達成率分布
        if not summary_df['直近月達成率'].empty:
            high_performers = len(summary_df[summary_df['直近月達成率'] >= 110])  # 110%以上
            good_performers = len(summary_df[(summary_df['直近月達成率'] >= 100) & (summary_df['直近月達成率'] < 110)])  # 100-110%
            under_performers = len(summary_df[summary_df['直近月達成率'] < 100])  # 100%未満
            
            metrics.extend([
                {
                    "診療科名": "全体",
                    "メトリクス名": "高達成診療科数_110%以上",
                    "値": high_performers,
                    "単位": "科",
                    "期間": period_info["label"],
                    "期間タイプ": period_info["type"],
                    "カテゴリ": "達成率分布",
                    "データ種別": "実績",
                    "計算日時": datetime.now().isoformat(),
                    "アプリ名": self.app_name
                },
                {
                    "診療科名": "全体",
                    "メトリクス名": "標準達成診療科数_100-110%",
                    "値": good_performers,
                    "単位": "科",
                    "期間": period_info["label"],
                    "期間タイプ": period_info["type"],
                    "カテゴリ": "達成率分布",
                    "データ種別": "実績",
                    "計算日時": datetime.now().isoformat(),
                    "アプリ名": self.app_name
                },
                {
                    "診療科名": "全体",
                    "メトリクス名": "未達成診療科数_100%未満",
                    "値": under_performers,
                    "単位": "科",
                    "期間": period_info["label"],
                    "期間タイプ": period_info["type"],
                    "カテゴリ": "達成率分布",
                    "データ種別": "実績",
                    "計算日時": datetime.now().isoformat(),
                    "アプリ名": self.app_name
                }
            ])
        
        return metrics
    
    def _calculate_department_metrics(
        self,
        summary_df: pd.DataFrame,
        period_info: Dict
    ) -> List[Dict]:
        """診療科別メトリクス計算"""
        metrics = []
        
        if summary_df.empty:
            return metrics
        
        for _, row in summary_df.iterrows():
            dept_name = row['診療科']
            
            # 直近月達成率
            if pd.notna(row['直近月達成率']):
                metrics.append({
                    "診療科名": dept_name,
                    "メトリクス名": "直近月達成率",
                    "値": round(row['直近月達成率'], 1),
                    "単位": "%",
                    "期間": period_info["label"],
                    "期間タイプ": period_info["type"],
                    "カテゴリ": "診療科別実績",
                    "データ種別": "実績",
                    "計算日時": datetime.now().isoformat(),
                    "アプリ名": self.app_name
                })
            
            # 今年度平均達成率
            if pd.notna(row['今年度平均達成率']):
                metrics.append({
                    "診療科名": dept_name,
                    "メトリクス名": "今年度平均達成率",
                    "値": round(row['今年度平均達成率'], 1),
                    "単位": "%",
                    "期間": period_info["label"],
                    "期間タイプ": period_info["type"],
                    "カテゴリ": "診療科別実績",
                    "データ種別": "実績",
                    "計算日時": datetime.now().isoformat(),
                    "アプリ名": self.app_name
                })
            
            # 過去6ヶ月平均達成率
            if pd.notna(row['過去6ヶ月平均達成率']):
                metrics.append({
                    "診療科名": dept_name,
                    "メトリクス名": "過去6ヶ月平均達成率",
                    "値": round(row['過去6ヶ月平均達成率'], 1),
                    "単位": "%",
                    "期間": period_info["label"],
                    "期間タイプ": period_info["type"],
                    "カテゴリ": "診療科別実績",
                    "データ種別": "実績",
                    "計算日時": datetime.now().isoformat(),
                    "アプリ名": self.app_name
                })
            
            # 全体比率
            if pd.notna(row['全体比率']):
                metrics.append({
                    "診療科名": dept_name,
                    "メトリクス名": "全体比率",
                    "値": round(row['全体比率'], 1),
                    "単位": "%",
                    "期間": period_info["label"],
                    "期間タイプ": period_info["type"],
                    "カテゴリ": "診療科別実績",
                    "データ種別": "実績",
                    "計算日時": datetime.now().isoformat(),
                    "アプリ名": self.app_name
                })
            
            # 昨年度同期比
            if pd.notna(row['昨年度同期比']):
                metrics.append({
                    "診療科名": dept_name,
                    "メトリクス名": "昨年度同期比",
                    "値": round(row['昨年度同期比'], 1),
                    "単位": "%",
                    "期間": period_info["label"],
                    "期間タイプ": period_info["type"],
                    "カテゴリ": "診療科別比較",
                    "データ種別": "実績",
                    "計算日時": datetime.now().isoformat(),
                    "アプリ名": self.app_name
                })
            
            # 評価コメント（カテゴリとして）
            evaluation = row['評価コメント']
            evaluation_score = self._convert_evaluation_to_score(evaluation)
            
            metrics.append({
                "診療科名": dept_name,
                "メトリクス名": "傾向評価スコア",
                "値": evaluation_score,
                "単位": "ポイント",
                "期間": period_info["label"],
                "期間タイプ": period_info["type"],
                "カテゴリ": "診療科別評価",
                "データ種別": "評価",
                "計算日時": datetime.now().isoformat(),
                "アプリ名": self.app_name,
                "備考": evaluation
            })
        
        return metrics
    
    def _calculate_trend_metrics(
        self,
        chart_df: pd.DataFrame,
        period_info: Dict
    ) -> List[Dict]:
        """時系列トレンドメトリクス計算"""
        metrics = []
        
        if chart_df.empty or '月' not in chart_df.columns:
            return metrics
        
        try:
            # 月列を日付型に変換
            chart_df['月'] = pd.to_datetime(chart_df['月'])
            
            # 最新3ヶ月の変動係数（安定性指標）
            latest_3_months = chart_df['月'].nlargest(3).min()
            recent_data = chart_df[chart_df['月'] >= latest_3_months]
            
            if not recent_data.empty:
                # 全体の変動係数
                achievement_rates = recent_data['達成率']
                if len(achievement_rates) > 1:
                    cv = (achievement_rates.std() / achievement_rates.mean()) * 100
                    
                    metrics.append({
                        "診療科名": "全体",
                        "メトリクス名": "直近3ヶ月変動係数",
                        "値": round(cv, 2),
                        "単位": "%",
                        "期間": f"{latest_3_months.strftime('%Y年%m月')}以降",
                        "期間タイプ": "3ヶ月",
                        "カテゴリ": "安定性分析",
                        "データ種別": "実績",
                        "計算日時": datetime.now().isoformat(),
                        "アプリ名": self.app_name
                    })
            
            # 診療科別トレンド分析
            for dept_name, group in chart_df.groupby('診療科'):
                group = group.sort_values('月')
                
                if len(group) >= 3:  # 最低3ヶ月のデータが必要
                    # 線形トレンド計算
                    x = range(len(group))
                    y = group['達成率'].values
                    
                    # 線形回帰でトレンド係数計算
                    if len(x) > 1:
                        slope = np.polyfit(x, y, 1)[0]  # 傾き
                        
                        metrics.append({
                            "診療科名": dept_name,
                            "メトリクス名": "月次トレンド係数",
                            "値": round(slope, 2),
                            "単位": "%/月",
                            "期間": period_info["label"],
                            "期間タイプ": period_info["type"],
                            "カテゴリ": "診療科別トレンド",
                            "データ種別": "分析",
                            "計算日時": datetime.now().isoformat(),
                            "アプリ名": self.app_name
                        })
                    
                    # 最新月 vs 平均との乖離
                    latest_rate = group.iloc[-1]['達成率']
                    avg_rate = group['達成率'].mean()
                    deviation = latest_rate - avg_rate
                    
                    metrics.append({
                        "診療科名": dept_name,
                        "メトリクス名": "最新月平均乖離",
                        "値": round(deviation, 1),
                        "単位": "%",
                        "期間": period_info["label"],
                        "期間タイプ": period_info["type"],
                        "カテゴリ": "診療科別トレンド",
                        "データ種別": "分析",
                        "計算日時": datetime.now().isoformat(),
                        "アプリ名": self.app_name
                    })
        
        except Exception as e:
            logger.warning(f"トレンドメトリクス計算エラー: {e}")
        
        return metrics
    
    def _convert_evaluation_to_score(self, evaluation: str) -> int:
        """評価コメントを数値スコアに変換"""
        if "改善傾向" in str(evaluation):
            return 3  # ポジティブ
        elif "悪化傾向" in str(evaluation):
            return 1  # ネガティブ
        else:
            return 2  # ニュートラル
    
    def _generate_filename(self, analysis_date: datetime, period_type: str) -> str:
        """ファイル名生成"""
        date_str = analysis_date.strftime("%Y%m%d")
        return f"{date_str}_{self.app_name}_メトリクス_{period_type}.csv"
    
    def create_downloadable_csv(self, metrics_df: pd.DataFrame) -> io.BytesIO:
        """ダウンロード可能なCSVファイルを作成"""
        output = io.BytesIO()
        metrics_df.to_csv(output, index=False, encoding='utf-8-sig')
        output.seek(0)
        return output


def create_gross_profit_metrics_export_interface():
    """粗利メトリクス出力インターフェース"""
    try:
        st.subheader("📊 粗利メトリクス出力")
        
        # セッションからデータ取得
        summary_df = st.session_state.get('gross_profit_summary_df', pd.DataFrame())
        chart_df = st.session_state.get('gross_profit_chart_df', pd.DataFrame())
        
        if summary_df.empty or chart_df.empty:
            st.info("📊 粗利データを処理してからメトリクス出力をご利用ください。")
            st.markdown("メインページで粗利目標ファイルと実績ファイルをアップロードし、データ処理を完了させてください。")
            return
        
        # データ概要表示
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("診療科数", len(summary_df))
        with col2:
            st.metric("分析期間", f"{len(chart_df)}ヶ月分")
        with col3:
            avg_rate = summary_df['直近月達成率'].mean()
            st.metric("平均達成率", f"{avg_rate:.1f}%" if not pd.isna(avg_rate) else "---")
        
        # エクスポート設定
        st.markdown("---")
        col1, col2 = st.columns(2)
        
        with col1:
            period_type = st.selectbox(
                "分析期間タイプ",
                ["月次", "四半期", "年次"],
                help="メトリクス分析の期間タイプを選択してください"
            )
        
        with col2:
            analysis_date = st.date_input(
                "基準日",
                value=datetime.now().date(),
                help="分析の基準となる日付を選択してください"
            )
        
        # プレビュー表示
        if st.button("📋 メトリクスプレビュー", type="secondary"):
            with st.spinner("メトリクス計算中..."):
                try:
                    exporter = GrossProfitMetricsExporter()
                    metrics_df, filename = exporter.export_metrics_csv(
                        summary_df, chart_df, datetime.combine(analysis_date, datetime.min.time()), period_type
                    )
                    
                    st.success(f"✅ メトリクス計算完了: {len(metrics_df)}件のメトリクス")
                    
                    # プレビュー表示
                    with st.expander("📄 メトリクスプレビュー", expanded=True):
                        st.dataframe(metrics_df.head(20), use_container_width=True)
                        if len(metrics_df) > 20:
                            st.caption(f"... 他 {len(metrics_df) - 20} 件のメトリクス")
                    
                    # カテゴリ別サマリー
                    category_summary = metrics_df['カテゴリ'].value_counts()
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        st.write("**カテゴリ別メトリクス数**")
                        for category, count in category_summary.items():
                            st.write(f"• {category}: {count}件")
                    
                    with col2:
                        st.write("**含まれる診療科**")
                        departments = metrics_df[metrics_df['診療科名'] != '全体']['診療科名'].unique()
                        for dept in departments[:10]:
                            st.write(f"• {dept}")
                        if len(departments) > 10:
                            st.write(f"... 他 {len(departments) - 10} 診療科")
                    
                    # セッションに保存
                    st.session_state['preview_gp_metrics_df'] = metrics_df
                    st.session_state['preview_gp_filename'] = filename
                    
                except Exception as e:
                    st.error(f"❌ メトリクス計算エラー: {e}")
                    logger.error(f"粗利メトリクス計算エラー: {e}")
        
        # CSV出力
        st.markdown("---")
        
        if st.button("📥 CSV出力", type="primary"):
            try:
                # プレビューデータがあればそれを使用、なければ新規計算
                if 'preview_gp_metrics_df' in st.session_state:
                    metrics_df = st.session_state['preview_gp_metrics_df']
                    filename = st.session_state['preview_gp_filename']
                else:
                    with st.spinner("メトリクス計算中..."):
                        exporter = GrossProfitMetricsExporter()
                        metrics_df, filename = exporter.export_metrics_csv(
                            summary_df, chart_df, datetime.combine(analysis_date, datetime.min.time()), period_type
                        )
                
                # CSV出力
                exporter = GrossProfitMetricsExporter()
                csv_data = exporter.create_downloadable_csv(metrics_df)
                
                st.download_button(
                    label="💾 粗利メトリクスCSVダウンロード",
                    data=csv_data,
                    file_name=filename,
                    mime="text/csv",
                    help=f"{filename} をダウンロードします。ポータル統合用の標準化された粗利メトリクスデータです。"
                )
                
                st.success(f"✅ CSV出力準備完了: {len(metrics_df)}件のメトリクス")
                
            except Exception as e:
                st.error(f"❌ CSV出力エラー: {e}")
                logger.error(f"粗利CSV出力エラー: {e}")
        
        # 使用方法説明
        with st.expander("ℹ️ 使用方法とデータ形式"):
            st.markdown("""
            ### 📋 出力される粗利メトリクス
            
            **全体指標**
            - 総診療科数
            - 目標達成診療科数
            - 平均達成率
            - 達成率分布（高達成・標準・未達成）
            
            **診療科別指標**
            - 直近月達成率
            - 今年度平均達成率
            - 過去6ヶ月平均達成率
            - 全体比率（粗利貢献度）
            - 昨年度同期比
            - 傾向評価スコア
            
            **トレンド分析**
            - 月次トレンド係数
            - 最新月平均乖離
            - 直近3ヶ月変動係数
            
            ### 🔧 ポータル統合について
            
            出力されるCSVファイルは以下の標準形式で統一されています：
            - 診療科名、メトリクス名、値、単位、期間などの共通フォーマット
            - 手術分析・入退院分析アプリと統合表示可能
            - ポータルwebページで自動読み込み・グラフ化対応
            
            ### 💡 活用例
            
            - **経営会議資料**: 診療科別業績サマリー
            - **部門評価**: 目標達成状況の客観的評価
            - **改善計画**: トレンド分析による課題特定
            - **予算策定**: 過去実績に基づく将来計画
            """)
    
    except Exception as e:
        st.error(f"❌ 粗利メトリクス出力インターフェースエラー: {e}")
        logger.error(f"粗利メトリクス出力インターフェースエラー: {e}")


if __name__ == "__main__":
    # テスト用
    st.title("粗利メトリクス出力テスト")
    create_gross_profit_metrics_export_interface()