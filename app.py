import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import datetime
import calendar
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# =====================
# マニュアルの本文（日本語ヘッダー対応版）
# =====================
MANUAL_TEXT = """
### 1. 基本ルール
* **ファイル名**: `staff.xlsx`
* **シート構成**: 
    1. `staff`: スタッフ設定（必須）
    2. `ng_pair`: 禁止ペア設定（任意）

### 2. staffシートの入力内容
| 日本語ヘッダー | 説明 |
| :--- | :--- |
| **氏名** | スタッフの名前 |
| **雇用形態** | `常勤` `非常勤` `派遣` |
| **公休数** | 月に必要な公休（休み）の日数 |
| **早番可** 〜 **夜勤可** | `1`=可, `0`=不可 |
| **夜勤専従** | `1`=夜勤のみ割当 |
| **夜勤上限** | 月間の最大夜勤回数（空欄は制限なし） |
| **最大連勤** | 最大連勤数 (空欄=6連勤) |
| **前月夜勤** | 前月末が夜勤なら`1` |
| **前月連勤** | 前月末時点の連勤数 |
| **月** 〜 **日** | 曜日ごとの勤務可否 (`1`=可, `0`=不可) |
| **1** 〜 **31** | 希望シフト（休, 有, 早, 日, 遅, 夜） |

### 3. ng_pairシート（禁止ペア）
| staff1 | staff2 | 
| :--- | :--- |
| 山田 | 鈴木 |
※氏名を正確に入力してください。日勤帯（早・日・遅）での重複を避けます。
"""

# =====================
# ページ設定
# =====================
st.set_page_config(page_title="シフト自動作成ツール", layout="wide")

st.title("📅 シフト自動作成ツール (Ver.9)")
st.markdown("日本語ヘッダーのExcelファイルをアップロードして、作成ボタンを押してください。")

# =====================
# サイドバー
# =====================
with st.sidebar:
    st.header("設定")
    with st.expander("📖 使い方マニュアルを見る"):
        st.markdown(MANUAL_TEXT)
    st.markdown("---")
    
    default_year = datetime.date.today().year
    default_month = datetime.date.today().month + 1
    if default_month > 12:
        default_year += 1
        default_month = 1
        
    YEAR = st.number_input("作成する【年】", value=default_year, step=1)
    MONTH = st.number_input("作成する【月】", value=default_month, min_value=1, max_value=12, step=1)
    
    uploaded_file = st.file_uploader("Excelファイルをアップロード", type=["xlsx"])
    
    st.markdown("---")
    st.write("👥 **曜日別の必要人数設定**")
    default_req_data = {
        "月": [1, 1, 1, 2], "火": [1, 1, 1, 2], "水": [1, 1, 1, 2], "木": [1, 1, 1, 2],
        "金": [1, 1, 1, 2], "土": [1, 1, 1, 2], "日": [1, 1, 1, 2]
    }
    df_req_default = pd.DataFrame(default_req_data, index=["早", "日", "遅", "夜"])
    edited_req_df = st.data_editor(df_req_default, height=180)

    st.markdown("---")
    st.write("🔧 **詳細設定**")
    random_seed = st.number_input("再計算用乱数", value=0, step=1)

# =====================
# ヘルパー関数: モデル作成
# =====================
def create_shift_model(staff_df, ng_pairs, year, month, req_df, is_diagnostic=False):
    DAYS = calendar.monthrange(int(year), int(month))[1]
    SHIFT_TYPES = ["早", "日", "遅", "夜"]
    DAY_SHIFT_GROUP = ["早", "日", "遅"] 
    WEEKDAY_NAMES = ["月", "火", "水", "木", "金", "土", "日"]
    
    base_date = datetime.date(int(year), int(month), 1)
    weekdays_indices = [(base_date + datetime.timedelta(days=i)).weekday() for i in range(DAYS)]

    staff = staff_df["name"].tolist()
    staff_type = dict(zip(staff_df["name"], staff_df["type"]))
    off_days = dict(zip(staff_df["name"], staff_df["off_days"]))
    max_night_dict = dict(zip(staff_df["name"], staff_df["max_night"]))

    # 希望入力
    req_input = {s: {} for s in staff}
    for _, r in staff_df.iterrows():
        s = r["name"]
        for d in range(DAYS):
            val = r.get(f"req_shift_{d+1}", None)
            req_input[s][d] = str(val).strip() if pd.notna(val) and val != "" else None

    prev_night = dict(zip(staff_df["name"], staff_df.get("prev_night", [0]*len(staff))))
    prev_consecutive = dict(zip(staff_df["name"], staff_df.get("prev_consecutive", [0]*len(staff))))
    night_only = dict(zip(staff_df["name"], staff_df.get("night_only", [0]*len(staff))))

    can = {}
    weekday_can = {}
    for _, r in staff_df.iterrows():
        s = r["name"]
        can[s] = {"早": r["can_early"], "日": r["can_day"], "遅": r["can_late"], "夜": r["can_night"]}
        weekday_can[s] = {0: r["mon"], 1: r["tue"], 2: r["wed"], 3: r["thu"], 4: r["fri"], 5: r["sat"], 6: r["sun"]}

    limit_consecutive = dict(zip(staff_df["name"], staff_df.get("limit_consecutive", [6]*len(staff))))

    model = cp_model.CpModel()
    x = {}
    work_flag = {}

    for s in staff:
        for d in range(DAYS):
            for sh in SHIFT_TYPES:
                if can[s][sh] == 1:
                    x[s, d, sh] = model.NewBoolVar(f"x_{s}_{d}_{sh}")
                else:
                    x[s, d, sh] = None 

    for s in staff:
        for d in range(DAYS):
            model.Add(sum(x[s, d, sh] for sh in SHIFT_TYPES if x[s, d, sh] is not None) <= 1)

    # 人数制約
    shortage_vars = {} 
    for d in range(DAYS):
        wd_name = WEEKDAY_NAMES[weekdays_indices[d]]
        for sh in SHIFT_TYPES:
            active_staff = sum(x[s, d, sh] for s in staff if x[s, d, sh] is not None)
            req_num = int(req_df.at[sh, wd_name])
            if not is_diagnostic:
                model.Add(active_staff == req_num)
            else:
                shortage = model.NewIntVar(0, req_num, f"shortage_{d}_{sh}")
                shortage_vars[d, sh] = shortage
                model.Add(active_staff + shortage == req_num)

    # 勤務フラグ (連勤計算用)
    for s in staff:
        for d in range(DAYS):
            work_flag[s, d] = model.NewBoolVar(f"work_{s}_{d}")
            is_shifted = sum(x[s, d, sh] for sh in SHIFT_TYPES if x[s, d, sh] is not None)
            is_after_night = prev_night[s] if d == 0 else (x[s, d-1, "夜"] if x[s, d-1, "夜"] is not None else 0)
            model.Add(work_flag[s, d] == is_shifted + is_after_night)

    # 希望反映
    for s in staff:
        for d in range(DAYS):
            inp = req_input[s][d]
            if inp in ["休", "有"]:
                model.Add(work_flag[s, d] == 0)
            elif inp in SHIFT_TYPES and x[s, d, inp] is not None:
                model.Add(x[s, d, inp] == 1)

    # --- 新規制約 ---
    # 1. 遅番の翌日の早番禁止 (ハード)
    for s in staff:
        for d in range(DAYS - 1):
            if x[s, d, "遅"] is not None and x[s, d+1, "早"] is not None:
                model.Add(x[s, d, "遅"] + x[s, d+1, "早"] <= 1)

    # 2. 夜勤の連続抑制 (ハード/ソフト)
    consecutive_night_penalty = []
    for s in staff:
        for d in range(DAYS - 1):
            if x[s, d, "夜"] is not None and x[s, d+1, "夜"] is not None:
                # 3連勤以上の夜勤は一律禁止
                if d + 2 < DAYS and x[s, d+2, "夜"] is not None:
                     model.Add(x[s, d, "夜"] + x[s, d+1, "夜"] + x[s, d+2, "夜"] <= 2)
                
                # 2連夜勤にペナルティ
                is_double_night = model.NewBoolVar(f"double_night_{s}_{d}")
                model.AddBoolAnd([x[s, d, "夜"], x[s, d+1, "夜"]]).OnlyEnforceIf(is_double_night)
                model.AddBoolOr([x[s, d, "夜"].Not(), x[s, d+1, "夜"].Not()]).OnlyEnforceIf(is_double_night.Not())
                consecutive_night_penalty.append(is_double_night)

    # 3. 個人ごとの夜勤回数上限 (ハード)
    for s in staff:
        limit = max_night_dict.get(s)
        if pd.notna(limit) and limit >= 0:
            model.Add(sum(x[s, d, "夜"] for d in range(DAYS) if x[s, d, "夜"] is not None) <= int(limit))

    # --- 既存制約の維持 ---
    # 夜勤明けの翌日休み
    next_day_off_penalty = [] 
    for s in staff:
        for d in range(DAYS):
            if x[s, d, "夜"] is not None:
                if d + 1 < DAYS:
                    model.Add(sum(x[s, d+1, sh] for sh in SHIFT_TYPES if x[s, d+1, sh] is not None) == 0).OnlyEnforceIf(x[s, d, "夜"])
                if d + 2 < DAYS:
                    v = model.NewBoolVar(f"violation_{s}_{d}")
                    model.AddBoolAnd([x[s, d, "夜"], work_flag[s, d+2]]).OnlyEnforceIf(v)
                    next_day_off_penalty.append(v)

    # 連勤制限
    for s in staff:
        limit = int(limit_consecutive[s]) if pd.notna(limit_consecutive[s]) else 6
        for d in range(DAYS - limit):
            model.AddBoolOr([work_flag[s, d+k].Not() for k in range(limit + 1)])

    # 常勤確保
    full_time_staff = [s for s in staff if staff_type.get(s) == "常勤"]
    no_leader_penalty = []
    for d in range(DAYS):
        ft_vars = [x[s, d, sh] for s in full_time_staff for sh in DAY_SHIFT_GROUP if x[s, d, sh] is not None]
        if ft_vars:
            is_no_leader = model.NewBoolVar(f"no_leader_{d}")
            model.Add(sum(ft_vars) == 0).OnlyEnforceIf(is_no_leader)
            no_leader_penalty.append(is_no_leader)

    # NGペア
    ng_pair_penalty = []
    for (p1, p2) in ng_pairs:
        if p1 in staff and p2 in staff:
            for d in range(DAYS):
                p1_day = model.NewBoolVar(f"p1_d_{p1}_{d}")
                p2_day = model.NewBoolVar(f"p2_d_{p2}_{d}")
                model.Add(sum(x[p1, d, sh] for sh in DAY_SHIFT_GROUP if x[p1, d, sh] is not None) == 1).OnlyEnforceIf(p1_day)
                model.Add(sum(x[p2, d, sh] for sh in DAY_SHIFT_GROUP if x[p2, d, sh] is not None) == 1).OnlyEnforceIf(p2_day)
                clash = model.NewBoolVar(f"clash_{p1}_{p2}_{d}")
                model.AddBoolAnd([p1_day, p2_day]).OnlyEnforceIf(clash)
                ng_pair_penalty.append(clash)

    # 目的関数
    if not is_diagnostic:
        night_only_bonus = sum(sum(x[s, d, "夜"] for d in range(DAYS) if x[s, d, "夜"] is not None) for s in staff if night_only.get(s) == 1)
        model.Minimize(
            500 * sum(next_day_off_penalty) + 
            200 * sum(no_leader_penalty) + 
            150 * sum(consecutive_night_penalty) + # 夜勤連続の抑制
            100 * sum(ng_pair_penalty) - 
            1000 * night_only_bonus
        )
    else:
        model.Minimize(sum(shortage_vars.values()))

    return model, x, shortage_vars, req_input, prev_night, staff

# =====================
# メイン処理
# =====================
if st.button("シフトを作成する", type="primary"):
    if uploaded_file is None:
        st.error("Excelファイルをアップロードしてください。")
        st.stop()

    try:
        # Excel読み込み & カラム名変換
        df_raw = pd.read_excel(uploaded_file, sheet_name="staff")
        
        # 変換対応表
        rename_map = {
            "氏名": "name", "雇用形態": "type", "公休数": "off_days",
            "早番可": "can_early", "日勤可": "can_day", "遅番可": "can_late", "夜勤可": "can_night",
            "夜勤専従": "night_only", "夜勤上限": "max_night", "最大連勤": "limit_consecutive",
            "前月夜勤": "prev_night", "前月連勤": "prev_consecutive",
            "月": "mon", "火": "tue", "水": "wed", "木": "thu", "金": "fri", "土": "sat", "日": "sun"
        }
        # 1〜31の数字カラムを req_shift_n に変換
        for i in range(1, 32):
            rename_map[i] = f"req_shift_{i}"
            rename_map[str(i)] = f"req_shift_{i}"

        staff_df = df_raw.rename(columns=rename_map)

        # NGペア読み込み
        ng_pairs = []
        try:
            ng_df = pd.read_excel(uploaded_file, sheet_name="ng_pair")
            if "staff1" in ng_df.columns and "staff2" in ng_df.columns:
                ng_pairs = list(zip(ng_df["staff1"], ng_df["staff2"]))
        except: pass

        # 通常計算
        model, x, _, req_input, prev_night, staff = create_shift_model(staff_df, ng_pairs, YEAR, MONTH, edited_req_df)
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = 30.0
        solver.parameters.random_seed = int(random_seed)
        status = solver.Solve(model)

       if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            st.success("✅ シフト作成に成功しました！")
            
            # 結果集計
            DAYS = calendar.monthrange(int(YEAR), int(MONTH))[1]
            SHIFT_TYPES = ["早", "日", "遅", "夜"]
            res_rows = []
            
            for s in staff:
                row = []
                # ① カウント用変数の初期化
                c_night = 0  # 夜勤回数
                c_work = 0   # 勤務日数
                c_off = 0    # 休み（休＋明）
                c_paid = 0   # 有給
                
                for d in range(DAYS):
                    v = "休"
                    if req_input[s][d] == "有": 
                        v = "有"
                    else:
                        for sh in SHIFT_TYPES:
                            if x[s, d, sh] is not None and solver.Value(x[s, d, sh]) == 1:
                                v = sh; break
                        if v == "休":
                            if (d == 0 and prev_night[s] == 1) or (d > 0 and x[s, d-1, "夜"] is not None and solver.Value(x[s, d-1, "夜"]) == 1):
                                v = "明"
                    
                    row.append(v)
                    
                    # ② カウント加算処理
                    if v == "夜": c_night += 1
                    if v in ["早", "日", "遅", "夜"]: c_work += 1
                    if v in ["休", "明"]: c_off += 1
                    if v == "有": c_paid += 1
                
                # ループの最後で集計列を右端に追加
                row.extend([c_night, c_work, c_off, c_paid])
                res_rows.append(row)
            
            # データフレーム作成（集計用ヘッダー追加）
            date_cols = [f"{i+1}日" for i in range(DAYS)]
            count_cols = ["夜勤", "勤務日", "休み", "有給"]
            df_out = pd.DataFrame(res_rows, index=staff, columns=date_cols + count_cols)
            
            st.dataframe(df_out)

            # Excelダウンロードとスタイル適用
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_out.to_excel(writer, sheet_name='シフト結果')
                
                # ③ 「明」と「休」のセルを薄緑に着色
                wb = writer.book
                ws = writer.sheets['シフト結果']
                fill_green = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # シフト部分のみループ（集計列は着色しないため max_col を調整）
                # 1行目はヘッダー、A列は氏名なので、データは2行目・2列目(B列)から開始
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=2, max_col=DAYS+1):
                    for cell in row:
                        if cell.value in ["明", "休"]:
                            cell.fill = fill_green

            st.download_button("📥 Excelを保存", output.getvalue(), f"shift_{YEAR}_{MONTH}.xlsx")

        else:
            # 診断モード
            st.error("❌ 条件を満たすシフトが見つかりませんでした。不足箇所を特定します...")
            diag_model, _, shortage_vars, _, _, _ = create_shift_model(staff_df, [], YEAR, MONTH, edited_req_df, is_diagnostic=True)
            diag_solver = cp_model.CpSolver()
            if diag_solver.Solve(diag_model) in (cp_model.OPTIMAL, cp_model.FEASIBLE):
                error_report = []
                base_date = datetime.date(int(YEAR), int(MONTH), 1)
                WEEKDAY_NAMES = ["月", "火", "水", "木", "金", "土", "日"]
                
                for d in range(calendar.monthrange(int(YEAR), int(MONTH))[1]):
                    for sh in ["早", "日", "遅", "夜"]:
                        lack = diag_solver.Value(shortage_vars[d, sh])
                        if lack > 0:
                            wd = WEEKDAY_NAMES[(base_date + datetime.timedelta(days=d)).weekday()]
                            error_report.append({"日付": f"{d+1}日({wd})", "シフト": sh, "不足人数": f"{lack}名"})
                
                if error_report:
                    st.warning("⚠️ **以下の箇所で人数が不足しています：**")
                    st.table(pd.DataFrame(error_report))
                    st.info("💡 ヒント: 不足日の「希望休」を減らすか、サイドバーの「必要人数」を下げて再試行してください。")
                else:
                    st.warning("人数は足りていますが、遅早禁止や連勤制限などのルールが厳しすぎて作成できません。")

    except Exception as e:
        st.error(f"エラーが発生しました: {e}")
        st.info("Excelのカラム名がマニュアル通り（氏名、雇用形態...）になっているか確認してください。")

