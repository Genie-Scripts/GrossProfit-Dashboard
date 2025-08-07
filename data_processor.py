# data_processor.py
import pandas as pd
import numpy as np
from datetime import datetime
from dateutil.relativedelta import relativedelta
import os

def load_data(file):
    """
    アップロードされたExcelまたはCSVファイルを読み込み、列名を日付オブジェクトに変換する。
    """
    if file is None:
        return None
    try:
        file_extension = os.path.splitext(file.name)[1].lower()
        if file_extension == '.csv':
            df = pd.read_csv(file)
        elif file_extension in ['.xlsx', '.xls']:
            df = pd.read_excel(file, header=0)
        else:
            return None # サポート外の形式

        print(f"読み込んだファイル: {file.name}")
        print(f"列名: {list(df.columns)}")
        print(f"データ形状: {df.shape}")
        
        # 列名を処理（日付列のみを変換）
        new_columns = {}
        for col in df.columns:
            col_str = str(col).lower()
            
            # 以下のキーワードを含む列は日付変換しない
            preserve_keywords = ['診療科', '名称', '目標', '粗利', '金額', 'target', 'goal', 'amount', '値']
            should_preserve = any(keyword in col_str for keyword in preserve_keywords)
            
            if should_preserve:
                new_columns[col] = col
            else:
                # 日付として解釈を試みる
                try:
                    # まず日付形式かチェック
                    parsed_date = pd.to_datetime(col, errors='coerce')
                    if pd.notna(parsed_date):
                        new_columns[col] = parsed_date
                    else:
                        # 日付でない場合は元のまま保持
                        new_columns[col] = col
                except (ValueError, TypeError):
                    new_columns[col] = col
        
        df = df.rename(columns=new_columns)
        
        # NaTの列名を持つ列を削除（ただし、意図的にNoneやNaTにした列は除く）
        valid_columns = [col for col in df.columns if pd.notna(col)]
        df = df[valid_columns]
        
        print(f"処理後の列名: {list(df.columns)}")
        
        return df
    except Exception as e:
        print(f"Error loading data: {e}")
        return None

def process_data(target_df, actual_df, today=datetime.now()):
    """
    目標と実績のデータフレームを処理し、サマリーと詳細チャート用のデータフレームを生成する。
    """
    if target_df is None or actual_df is None:
        return pd.DataFrame(), pd.DataFrame()

    print("\n=== 目標データの確認 ===")
    print(f"目標データの列: {list(target_df.columns)}")
    print(f"目標データの形状: {target_df.shape}")
    print(f"目標データの最初の5行:\n{target_df.head()}")
    
    print("\n=== 実績データの確認 ===")
    print(f"実績データの列名の型: {[type(col).__name__ for col in actual_df.columns]}")
    print(f"実績データの形状: {actual_df.shape}")

    # --- 1. データの準備 ---
    # 診療科列を特定（最初の列）
    dept_col_target = target_df.columns[0]
    dept_col_actual = actual_df.columns[0]
    
    print(f"\n診療科列（目標）: {dept_col_target}")
    print(f"診療科列（実績）: {dept_col_actual}")
    
    # 目標データの目標値列を特定
    target_value_col = None
    
    # 目標データの2列目以降から目標値列を探す
    if len(target_df.columns) > 1:
        # まず「目標」というキーワードを含む列を探す
        for col in target_df.columns[1:]:
            col_str = str(col).lower()
            if '目標' in col_str or 'target' in col_str or 'goal' in col_str:
                target_value_col = col
                print(f"目標値列を発見: {col}")
                break
        
        # 見つからない場合は2列目を使用
        if target_value_col is None and len(target_df.columns) > 1:
            target_value_col = target_df.columns[1]
            print(f"2列目を目標値列として使用: {target_value_col}")
    
    if target_value_col is None:
        print("Error: 目標値列が見つかりません。目標ファイルには診療科列と目標値列が必要です。")
        print(f"現在の列: {list(target_df.columns)}")
        return pd.DataFrame(), pd.DataFrame()
    
    # 目標値列のデータ型を確認し、必要なら数値に変換
    print(f"\n目標値列のデータ型: {target_df[target_value_col].dtype}")
    if not pd.api.types.is_numeric_dtype(target_df[target_value_col]):
        print("目標値列を数値に変換します...")
        target_df[target_value_col] = pd.to_numeric(
            target_df[target_value_col].astype(str).str.replace(',', '').str.replace('，', ''),
            errors='coerce'
        )
    
    print(f"目標値のサンプル:\n{target_df[[dept_col_target, target_value_col]].head()}")
    
    # 実績データから日付列を特定
    date_cols = sorted([col for col in actual_df.columns if isinstance(col, (datetime, pd.Timestamp))], 
                      key=lambda x: (x.year, x.month))
    
    if not date_cols:
        print("Error: 実績データに日付列が見つかりません")
        print(f"実績データの列: {list(actual_df.columns)}")
        return pd.DataFrame(), pd.DataFrame()
    
    print(f"\n日付列数: {len(date_cols)}")
    print(f"期間: {date_cols[0].strftime('%Y/%m')} 〜 {date_cols[-1].strftime('%Y/%m')}")
    
    # --- 2. 達成率の計算 ---
    chart_data = []
    processed_depts = 0
    
    # 各診療科ごとに処理
    for _, target_row in target_df.iterrows():
        dept_name = target_row[dept_col_target]
        target_value = target_row[target_value_col]
        
        # 実績データから同じ診療科のデータを取得
        actual_rows = actual_df[actual_df[dept_col_actual] == dept_name]
        
        if actual_rows.empty:
            print(f"Warning: {dept_name} の実績データが見つかりません")
            continue
            
        if pd.notna(target_value) and target_value > 0:
            actual_row = actual_rows.iloc[0]
            processed_depts += 1
            
            for month_col in date_cols:
                actual_value = actual_row.get(month_col)
                if pd.notna(actual_value):
                    # 実績値も数値に変換（必要な場合）
                    if isinstance(actual_value, str):
                        try:
                            actual_value = float(actual_value.replace(',', '').replace('，', ''))
                        except:
                            continue
                    
                    achievement_rate = (actual_value / target_value) * 100
                    
                    chart_data.append({
                        "診療科": dept_name,
                        "月": month_col,
                        "実績": actual_value,
                        "目標": target_value,
                        "達成率": achievement_rate
                    })
    
    print(f"\n処理した診療科数: {processed_depts}")
    
    chart_df = pd.DataFrame(chart_data)
    if chart_df.empty:
        print("Error: 達成率データが生成できませんでした")
        return pd.DataFrame(), pd.DataFrame()

    print(f"達成率データ生成完了: {len(chart_df)}レコード")

    # --- 3. 各種指標の計算 ---
    summary_data = []
    
    # 最新月を特定
    most_recent_month_date = chart_df['月'].max()
    print(f"\n最新月: {most_recent_month_date.strftime('%Y/%m')}")

    # ★★★ 新規追加：最新月の全診療科の粗利合計を計算 ★★★
    recent_month_df = chart_df[chart_df['月'] == most_recent_month_date]
    total_recent_profit = recent_month_df['実績'].sum()
    print(f"最新月の全体粗利合計: {total_recent_profit:,.0f}")
    
    # 各診療科の達成率サマリーを表示
    print("\n=== 直近月の達成率 ===")

    for dept_name, group in chart_df.groupby('診療科'):
        group = group.sort_values('月')
        
        # a. 直近月の達成率と実績値
        recent_rate_row = group[group['月'] == most_recent_month_date]
        recent_rate = recent_rate_row['達成率'].iloc[0] if not recent_rate_row.empty else np.nan
        recent_actual = recent_rate_row['実績'].iloc[0] if not recent_rate_row.empty else 0 # 実績値を取得
        
        # ★★★ 新規追加：全体比率を計算 ★★★
        profit_share = (recent_actual / total_recent_profit) * 100 if total_recent_profit > 0 else 0
        
        # デバッグ: 直近月の詳細
        if not recent_rate_row.empty:
            recent_target = recent_rate_row['目標'].iloc[0]
            print(f"{dept_name:15s}: 実績={recent_actual:12,.0f} 目標={recent_target:12,.0f} 達成率={recent_rate:6.1f}% 全体比率={profit_share:5.1f}%")

        # b. 今年度の平均達成率 (4月始まり)
        fy_start_year = today.year if today.month >= 4 else today.year - 1
        fy_start_date = pd.Timestamp(year=fy_start_year, month=4, day=1)
        fy_df = group[group['月'] >= fy_start_date]
        fy_avg_rate = fy_df['達成率'].mean() if not fy_df.empty else np.nan

        # c. 過去6ヵ月の平均達成率
        six_months_ago = most_recent_month_date - relativedelta(months=5)
        six_month_df = group[(group['月'] >= six_months_ago) & (group['月'] <= most_recent_month_date)]
        six_month_avg_rate = six_month_df['達成率'].mean() if not six_month_df.empty else np.nan

        # ★★★ 新規追加：昨年度同期比の計算 ★★★
        # 昨年度の同じ期間（今年度の開始月から最新月まで）を特定
        last_fy_start_date = pd.Timestamp(year=fy_start_year-1, month=4, day=1)
        last_fy_end_month = most_recent_month_date - relativedelta(years=1)
        
        # 今年度の実績合計
        current_fy_actual = fy_df['実績'].sum() if not fy_df.empty else 0
        
        # 昨年度同期間の実績合計
        last_fy_df = group[(group['月'] >= last_fy_start_date) & (group['月'] <= last_fy_end_month)]
        last_fy_actual = last_fy_df['実績'].sum() if not last_fy_df.empty else 0
        
        # 昨年度同期比率の計算
        yoy_comparison = ((current_fy_actual / last_fy_actual) * 100) if last_fy_actual > 0 else np.nan

        # d. 改善コメント
        comment = ""
        if pd.notna(recent_rate) and pd.notna(six_month_avg_rate):
            diff = recent_rate - six_month_avg_rate
            if diff > 5:
                comment = "改善傾向 👍"
            elif diff < -5:
                comment = "悪化傾向 👎"
            else:
                comment = "横ばい 😐"

        summary_data.append({
            "診療科": dept_name,
            "直近月達成率": recent_rate,
            "今年度平均達成率": fy_avg_rate,
            "過去6ヶ月平均達成率": six_month_avg_rate,
            "評価コメント": comment,
            "全体比率": profit_share,
            "昨年度同期比": yoy_comparison  # ★★★ 新しい指標を追加 ★★★
        })

    summary_df = pd.DataFrame(summary_data)
    
    # 直近月達成率でソート
    summary_df = summary_df.sort_values(by='直近月達成率', ascending=False, na_position='last').reset_index(drop=True)
    
    print(f"\nサマリーデータ生成完了: {len(summary_df)}診療科")
    
    # サマリーの統計情報
    print("\n=== 統計情報 ===")
    print(f"目標達成（100%以上）: {len(summary_df[summary_df['直近月達成率'] >= 100])}診療科")
    print(f"平均達成率: {summary_df['直近月達成率'].mean():.1f}%")
    print(f"最高達成率: {summary_df['直近月達成率'].max():.1f}%")
    print(f"最低達成率: {summary_df['直近月達成率'].min():.1f}%")

    return summary_df, chart_df