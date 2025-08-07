# html_generator.py (Google Analytics対応版)
import pandas as pd
import json
from datetime import datetime
from typing import Optional

def format_rate(rate):
    """達成率をフォーマットする補助関数"""
    if pd.isna(rate):
        return "---"
    return f"{rate:.1f}%"

def get_performance_class(rate):
    """達成率に基づいてパフォーマンスクラスを返す"""
    if pd.isna(rate):
        return "info"
    elif rate >= 100:
        return "success"
    elif rate >= 90:
        return "warning"
    else:
        return "danger"

# ▼▼▼【修正箇所】引数に google_analytics_id を追加 ▼▼▼
def generate_html(summary_df, chart_df, google_analytics_id: Optional[str] = None):
    """
    サマリーとチャートデータからインタラクティブなHTMLレポートを生成する。
    Streamlit-OR-Dashboard風の統一デザインを適用。
    Google Analytics トラッキングコードの埋め込みに対応。
    """
    # 1. チャート用データをJavaScriptが扱いやすいJSON形式に変換
    chart_json_data = {}
    for dept_name, group in chart_df.groupby('診療科'):
        group['月'] = group['月'].dt.strftime('%Y-%m-%d')
        chart_json_data[dept_name] = group.to_dict('records')
    
    chart_json_string = json.dumps(chart_json_data, ensure_ascii=False)

    # 2. HTMLのカード部分を生成（新デザインシステム適用）
    cards_html = ""
    for index, row in summary_df.iterrows():
        perf_class = get_performance_class(row['直近月達成率'])
        rank_badge_class = "rank-badge-gold" if index == 0 else "rank-badge-silver" if index == 1 else "rank-badge-bronze" if index == 2 else ""
        progress_width = min(row['直近月達成率'], 100) if pd.notna(row['直近月達成率']) else 0
        trend_icon, trend_class = "", ""
        if row['評価コメント'] == "改善傾向 👍":
            trend_icon, trend_class = '↗', 'up'
        elif row['評価コメント'] == "悪化傾向 👎":
            trend_icon, trend_class = '↘', 'down'
        else:
            trend_icon, trend_class = '→', 'neutral'
    
        profit_share_val = row['全体比率'] if pd.notna(row['全体比率']) else 0
        yoy_comparison_val = row['昨年度同期比'] if pd.notna(row['昨年度同期比']) else 0
        
        cards_html += f"""
        <div class="metric-card {perf_class}" onclick="showChart('{row['診療科']}')">
            <div class="rank-badge {rank_badge_class}">{index + 1}</div>
            <div class="metric-header">
                <div class="metric-title">{row['診療科']}</div>
                <div class="trend-indicator {trend_class}"><span class="trend-icon">{trend_icon}</span></div>
            </div>
            <div class="metric-content">
                <div class="main-metric-row">
                    <div class="main-metric">
                        <div class="metric-value">{format_rate(row['直近月達成率'])}</div>
                        <div class="metric-label">直近月達成率</div>
                    </div>
                    <div class="sub-metric">
                        <div class="item-value">{format_rate(row['過去6ヶ月平均達成率'])}</div>
                        <div class="item-label">6ヶ月平均</div>
                    </div>
                </div>
                <div class="progress-bar"><div class="progress-fill" style="width: {progress_width}%"></div></div>
                <div class="metric-grid">
                    <div class="metric-item">
                        <div class="item-value">{format_rate(row['今年度平均達成率'])}</div>
                        <div class="item-label">今年度平均</div>
                    </div>
                    <div class="metric-item">
                        <div class="item-value">{format_rate(profit_share_val)}</div>
                        <div class="item-label">全体比率</div>
                    </div>
                    <div class="metric-item">
                        <div class="item-value">{format_rate(yoy_comparison_val)}</div>
                        <div class="item-label">昨年度同期比</div>
                    </div>
                    <div class="metric-item">
                        <div class="item-value evaluation">{row['評価コメント']}</div>
                        <div class="item-label">トレンド</div>
                    </div>
                </div>
            </div>
        </div>
        """

    total_depts = len(summary_df)
    achieved_depts = len(summary_df[summary_df['直近月達成率'] >= 100])
    avg_achievement = summary_df['直近月達成率'].mean() if not summary_df['直近月達成率'].empty else 0
    
    # ▼▼▼【修正箇所】Google Analyticsのトラッキングコードを生成 ▼▼▼
    ga_script_html = ""
    if google_analytics_id:
        # f-string内でJSの{}をエスケープするために{{}}を使用
        ga_script_html = f"""
    <script async src="https://www.googletagmanager.com/gtag/js?id={google_analytics_id}"></script>
    <script>
      window.dataLayer = window.dataLayer || [];
      function gtag(){{{{dataLayer.push(arguments);}}}}
      gtag('js', new Date());
      gtag('config', '{google_analytics_id}');
    </script>"""

    # 3. HTML全体のテンプレートを作成
    # ▼▼▼【修正箇所】テンプレートに {ga_script_html} を追加 ▼▼▼
    html_template = """
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>診療科別 入外粗利レポート</title>
    {ga_script_html}
    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
    <style>
        :root {{
            --primary: #3b82f6; --primary-dark: #2563eb; --success: #10b981; --warning: #f59e0b;
            --danger: #ef4444; --neutral: #6b7280; --background: #f8fafc; --surface: #ffffff;
            --text-primary: #1e293b; --text-secondary: #64748b; --text-muted: #94a3b8; --border: #e2e8f0;
            --shadow-sm: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
            --shadow-md: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
            --shadow-lg: 0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -2px rgba(0, 0, 0, 0.05);
            --radius-sm: 0.375rem; --radius-md: 0.5rem; --radius-lg: 0.75rem; --transition: all 0.2s ease;
        }}
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
            background-color: var(--background); color: var(--text-primary); line-height: 1.5;
            -webkit-font-smoothing: antialiased; -moz-osx-font-smoothing: grayscale;
        }}
        .container {{ max-width: 1400px; margin: 0 auto; padding: 2rem; }}
        .header {{
            background: var(--surface); border-radius: var(--radius-lg); padding: 2.5rem;
            margin-bottom: 2rem; box-shadow: var(--shadow-sm); text-align: center; position: relative;
        }}
        .portal-home-button {{
            position: absolute; top: 1.5rem; left: 1.5rem; background: var(--primary); color: white;
            padding: 0.5rem 1.25rem; border-radius: var(--radius-md); text-decoration: none;
            font-size: 0.875rem; font-weight: 500; display: inline-flex; align-items: center;
            gap: 0.5rem; transition: var(--transition);
        }}
        .portal-home-button:hover {{ background: var(--primary-dark); transform: translateY(-1px); box-shadow: var(--shadow-md); }}
        h1 {{ font-size: 2.25rem; font-weight: 700; color: var(--text-primary); margin-bottom: 0.5rem; }}
        .subtitle {{ font-size: 1.125rem; color: var(--text-secondary); margin-bottom: 1rem; }}
        .timestamp {{
            font-size: 0.875rem; color: var(--text-muted); background: var(--background);
            padding: 0.375rem 1rem; border-radius: 9999px; display: inline-block;
        }}
        .summary-stats {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 1rem; margin-bottom: 2rem; }}
        .stat-card {{ background: var(--surface); padding: 1.5rem; border-radius: var(--radius-lg); box-shadow: var(--shadow-sm); text-align: center; }}
        .stat-value {{ font-size: 2rem; font-weight: 700; color: var(--primary); line-height: 1; margin-bottom: 0.5rem; }}
        .stat-label {{ font-size: 0.875rem; color: var(--text-secondary); font-weight: 500; }}
        .guidance-box {{
            background: #eff6ff; border: 1px solid #dbeafe; border-radius: var(--radius-lg); padding: 1rem 1.5rem;
            margin-bottom: 2rem; display: flex; align-items: center; gap: 1rem;
        }}
        .guidance-icon {{ font-size: 1.5rem; color: var(--primary); }}
        .guidance-text {{ font-size: 0.875rem; color: var(--text-primary); line-height: 1.5; }}
        .info-panel {{ background: var(--surface); border-radius: var(--radius-lg); margin-bottom: 2rem; overflow: hidden; box-shadow: var(--shadow-sm); }}
        .info-panel-header {{
            background: var(--background); padding: 1rem 1.5rem; cursor: pointer; display: flex;
            align-items: center; justify-content: space-between; border-bottom: 1px solid var(--border); transition: var(--transition);
        }}
        .info-panel-header:hover {{ background: #f1f5f9; }}
        .info-panel-title {{ display: flex; align-items: center; gap: 0.75rem; font-weight: 600; color: var(--text-primary); }}
        .info-panel-toggle {{ color: var(--text-secondary); transition: transform 0.2s ease; }}
        .info-panel.collapsed .info-panel-toggle {{ transform: rotate(-90deg); }}
        .info-panel-content {{ padding: 1.5rem; display: block; }}
        .info-panel.collapsed .info-panel-content {{ display: none; }}
        .info-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 1rem; }}
        .info-section {{ background: var(--background); border-radius: var(--radius-md); padding: 1.25rem; border-left: 3px solid var(--primary); }}
        .info-section-title {{ font-weight: 600; color: var(--text-primary); margin-bottom: 0.75rem; font-size: 0.9375rem; }}
        .info-section-content {{ font-size: 0.875rem; color: var(--text-secondary); line-height: 1.6; }}
        #homepage {{ display: block; }}
        #chartpage {{ display: none; }}
        .cards-container {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(320px, 1fr)); gap: 1.25rem; }}
        .metric-card {{
            background: var(--surface); border-radius: var(--radius-lg); padding: 1.5rem; position: relative; cursor: pointer;
            transition: var(--transition); box-shadow: var(--shadow-sm); border: 1px solid var(--border); display: flex; flex-direction: column;
        }}
        .metric-card:hover {{ transform: translateY(-2px); box-shadow: var(--shadow-lg); }}
        .metric-card.success {{ border-top: 3px solid var(--success); }}
        .metric-card.warning {{ border-top: 3px solid var(--warning); }}
        .metric-card.danger {{ border-top: 3px solid var(--danger); }}
        .metric-card.info {{ border-top: 3px solid var(--neutral); }}
        .rank-badge {{
            position: absolute; top: 1rem; right: 1rem; width: 2rem; height: 2rem; border-radius: 50%; display: flex;
            align-items: center; justify-content: center; font-weight: 600; font-size: 0.875rem;
            background: var(--background); color: var(--text-secondary);
        }}
        .rank-badge-gold {{ background: #fef3c7; color: #d97706; }}
        .rank-badge-silver {{ background: #e5e7eb; color: #4b5563; }}
        .rank-badge-bronze {{ background: #fed7aa; color: #c2410c; }}
        .metric-header {{ display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 1.25rem; }}
        .metric-title {{ font-size: 1.125rem; font-weight: 600; color: var(--text-primary); flex: 1; padding-right: 3rem; }}
        .trend-indicator {{ display: flex; align-items: center; justify-content: center; width: 2rem; height: 2rem; border-radius: var(--radius-sm); background: var(--background); }}
        .trend-indicator.up {{ background: #d1fae5; color: var(--success); }}
        .trend-indicator.down {{ background: #fee2e2; color: var(--danger); }}
        .trend-indicator.neutral {{ background: #f3f4f6; color: var(--neutral); }}
        .trend-icon {{ font-size: 1.125rem; font-weight: 700; }}
        .metric-content {{ flex: 1; display: flex; flex-direction: column; }}
        .main-metric-row {{ display: flex; align-items: flex-start; gap: 1.5rem; margin-bottom: 1rem; }}
        .main-metric {{ flex: 1; text-align: left; }}
        .sub-metric {{ text-align: center; padding: 0.5rem; background: var(--background); border-radius: var(--radius-sm); min-width: 80px; }}
        .metric-value {{ font-size: 2.25rem; font-weight: 700; line-height: 1; color: var(--text-primary); margin-bottom: 0.25rem; }}
        .metric-card.success .metric-value {{ color: var(--success); }}
        .metric-card.warning .metric-value {{ color: var(--warning); }}
        .metric-card.danger .metric-value {{ color: var(--danger); }}
        .metric-label {{ font-size: 0.75rem; color: var(--text-muted); font-weight: 500; text-transform: uppercase; letter-spacing: 0.05em; }}
        .progress-bar {{ width: 100%; height: 0.375rem; background: var(--border); border-radius: 9999px; overflow: hidden; margin-bottom: 1.25rem; }}
        .progress-fill {{ height: 100%; background: var(--primary); border-radius: 9999px; transition: width 0.6s ease; }}
        .metric-card.success .progress-fill {{ background: var(--success); }}
        .metric-card.warning .progress-fill {{ background: var(--warning); }}
        .metric-card.danger .progress-fill {{ background: var(--danger); }}
        .metric-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 0.75rem; margin-top: auto; }}
        .metric-item {{ text-align: center; padding: 0.5rem; background: var(--background); border-radius: var(--radius-sm); display: flex; flex-direction: column; justify-content: center; }}
        .item-value {{ font-size: 1rem; font-weight: 600; color: var(--text-primary); margin-bottom: 0.125rem; }}
        .item-value.evaluation {{ font-size: 0.875rem; }}
        .item-label {{ font-size: 0.6875rem; color: var(--text-muted); font-weight: 500; text-transform: uppercase; letter-spacing: 0.025em; }}
        #backButton {{
            background: var(--primary); color: white; border: none; padding: 0.625rem 1.25rem; border-radius: var(--radius-md);
            cursor: pointer; font-size: 0.875rem; font-weight: 500; margin-bottom: 1.5rem; display: inline-flex; align-items: center; gap: 0.5rem; transition: var(--transition);
        }}
        #backButton:hover {{ background: var(--primary-dark); transform: translateY(-1px); box-shadow: var(--shadow-md); }}
        #chart-title {{ font-size: 1.5rem; font-weight: 700; color: var(--text-primary); margin-bottom: 1.5rem; }}
        #chart-container {{ background: var(--surface); padding: 2rem; border-radius: var(--radius-lg); box-shadow: var(--shadow-sm); min-height: 500px; }}
        @media (max-width: 768px) {{
            .container {{ padding: 1rem; }}
            .cards-container {{ grid-template-columns: 1fr; }}
            .summary-stats {{ grid-template-columns: 1fr; }}
            h1 {{ font-size: 1.75rem; }}
            .metric-value {{ font-size: 2rem; }}
            .header {{ padding-top: 4rem; }}
            .portal-home-button {{ top: 1rem; left: 1rem; padding: 0.375rem 1rem; font-size: 0.75rem; }}
            .main-metric-row {{ flex-direction: column; gap: 1rem; }}
            .sub-metric {{ width: 100%; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div id="homepage">
            <div class="header">
                <a href="../index.html" class="portal-home-button">🏠 ポータルTOPへ</a>
                <h1>診療科別 入外粗利レポート</h1>
                <p class="subtitle">2025年度診療科目標達成率の推移</p>
                <span class="timestamp">📅 {timestamp} 更新</span>
            </div>
            <div class="summary-stats">
                <div class="stat-card"><div class="stat-value">{total_depts}</div><div class="stat-label">診療科数</div></div>
                <div class="stat-card"><div class="stat-value">{achieved_depts}</div><div class="stat-label">目標達成</div></div>
                <div class="stat-card"><div class="stat-value">{avg_achievement:.1f}%</div><div class="stat-label">平均達成率</div></div>
            </div>
            <div class="guidance-box">
                <div class="guidance-icon">💡</div>
                <div class="guidance-text"><strong>ヒント:</strong> 各診療科のカードをクリックすると、月ごとの詳細な達成率推移グラフを表示できます。</div>
            </div>
            <div class="info-panel collapsed" id="infoPanel">
                <div class="info-panel-header" onclick="toggleInfoPanel()">
                    <div class="info-panel-title"><span class="info-panel-icon">📋</span><span>指標・評価の説明</span></div>
                    <span class="info-panel-toggle">▼</span>
                </div>
                <div class="info-panel-content">
                    <div class="info-grid">
                        <div class="info-section"><div class="info-section-title">🎯 直近月達成率</div><div class="info-section-content">最新月の実績値を目標値で割った達成率です。100%が目標ラインとなります。</div></div>
                        <div class="info-section"><div class="info-section-title">📈 パフォーマンス指標</div><div class="info-section-content">今年度平均と過去6ヶ月平均で中長期的な業績を評価します。</div></div>
                        <div class="info-section"><div class="info-section-title">📊 昨年度同期比</div><div class="info-section-content">今年度の実績合計を昨年度同期間と比較。成長率を示します。</div></div>
                        <div class="info-section"><div class="info-section-title">💰 全体比率</div><div class="info-section-content">各診療科の粗利が全体に占める割合。経営への貢献度を示します。</div></div>
                        <div class="info-section"><div class="info-section-title">✅ 評価コメント</div><div class="info-section-content">直近月と過去6ヶ月平均の比較による傾向分析。±5%を基準に判定します。</div></div>
                    </div>
                </div>
            </div>
            <div class="cards-container">{cards_html}</div>
        </div>
        <div id="chartpage">
            <button id="backButton" onclick="showHome()"><span>←</span><span>ダッシュボードに戻る</span></button>
            <h2 id="chart-title"></h2>
            <div id="chart-container"></div>
        </div>
    </div>
    <script>
        const chartData = {chart_json_string};
        const homePageDiv = document.getElementById('homepage');
        const chartPageDiv = document.getElementById('chartpage');
        const chartTitle = document.getElementById('chart-title');

        function showChart(deptName) {{
            homePageDiv.style.display = 'none';
            chartPageDiv.style.display = 'block';
            chartTitle.innerText = deptName + ' - 達成率推移';
            const data = chartData[deptName];
            if (!data) return;
            const dates = data.map(d => d['月']);
            const rates = data.map(d => d['達成率']);
            function calculateLinearRegression(x, y) {{
                const n = x.length; let sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;
                for (let i = 0; i < n; i++) {{ const xi = i; sumX += xi; sumY += y[i]; sumXY += xi * y[i]; sumX2 += xi * xi; }}
                const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
                const intercept = (sumY - slope * sumX) / n;
                return {{ slope, intercept }};
            }}
            const regression = calculateLinearRegression(dates, rates);
            const regressionY = dates.map((_, i) => regression.intercept + regression.slope * i);
            const trace = {{ x: dates, y: rates, mode: 'lines+markers', type: 'scatter', name: '達成率', line: {{ color: '#3b82f6', width: 3, shape: 'linear' }}, marker: {{ size: 8, color: '#3b82f6', line: {{ color: 'white', width: 2 }} }}, hovertemplate: '<b>%{{x|%Y年%m月}}</b><br>達成率: %{{y:.1f}}%<extra></extra>' }};
            const regressionTrace = {{ x: dates, y: regressionY, mode: 'lines', type: 'scatter', name: 'トレンド', line: {{ color: '#94a3b8', width: 2, dash: 'dot' }}, hovertemplate: '<b>%{{x|%Y年%m月}}</b><br>トレンド: %{{y:.1f}}%<extra></extra>' }};
            const minRate = Math.min(...rates); const maxRate = Math.max(...rates);
            const yPadding = (maxRate - minRate) * 0.1 || 10;
            const yMin = Math.max(0, minRate - yPadding); const yMax = maxRate + yPadding;
            const layout = {{
                title: {{ text: deptName + ' の達成率推移', font: {{ size: 18, family: 'sans-serif' }} }},
                xaxis: {{ title: '期間', showgrid: true, gridcolor: '#e2e8f0', tickformat: '%Y/%m', tickangle: -45 }},
                yaxis: {{ title: '達成率 (%)', tickformat: '.1f', showgrid: true, gridcolor: '#e2e8f0', range: [yMin, yMax] }},
                margin: {{ t: 60, l: 70, r: 40, b: 80 }}, plot_bgcolor: '#f8fafc', paper_bgcolor: 'white', hovermode: 'x unified',
                hoverlabel: {{ bgcolor: 'white', font: {{ size: 13 }}, bordercolor: '#e2e8f0' }},
                shapes: [{{ type: 'line', x0: dates[0], y0: 100, x1: dates[dates.length - 1], y1: 100, line: {{ color: '#ef4444', width: 2, dash: 'dash' }} }}],
                annotations: [{{ x: dates[dates.length - 1], y: 100, xref: 'x', yref: 'y', text: '目標ライン (100%)', showarrow: false, xanchor: 'right', yanchor: 'bottom', font: {{ size: 12, color: '#ef4444' }}, bgcolor: 'rgba(255, 255, 255, 0.9)', borderpad: 4 }}],
                legend: {{ orientation: 'h', yanchor: 'bottom', y: 1.02, xanchor: 'right', x: 1 }}
            }};
            const config = {{ responsive: true, displayModeBar: true, displaylogo: false, modeBarButtonsToRemove: ['pan2d', 'select2d', 'lasso2d', 'autoScale2d'] }};
            Plotly.newPlot('chart-container', [trace, regressionTrace], layout, config);
        }}
        function showHome() {{ homePageDiv.style.display = 'block'; chartPageDiv.style.display = 'none'; }}
        window.addEventListener('load', () => {{
            const cards = document.querySelectorAll('.metric-card');
            cards.forEach((card, index) => {{
                card.style.opacity = '0'; card.style.transform = 'translateY(20px)';
                setTimeout(() => {{ card.style.transition = 'opacity 0.4s ease, transform 0.4s ease'; card.style.opacity = '1'; card.style.transform = 'translateY(0)'; }}, index * 50);
            }});
        }});
        function toggleInfoPanel() {{ const panel = document.getElementById('infoPanel'); panel.classList.toggle('collapsed'); }}
    </script>
</body>
</html>
"""
    # .format()メソッドを使用してテンプレートに値を埋め込む
    # ▼▼▼【修正箇所】ga_script_html をフォーマット引数に追加 ▼▼▼
    return html_template.format(
        timestamp=datetime.now().strftime('%Y年%m月%d日 %H:%M'),
        total_depts=total_depts,
        achieved_depts=achieved_depts,
        avg_achievement=avg_achievement,
        cards_html=cards_html,
        chart_json_string=chart_json_string,
        ga_script_html=ga_script_html
    )