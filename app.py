import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import datetime
import calendar
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# =====================
# マニュアルの本文（長いので変数に入れます）
# =====================
MANUAL_TEXT = """
### 1. 基本ルール
* **ファイル名**: `staff.xlsx`
* **シート名**: `staff` （変更不可）

### 2. 各列の入力内容
| 列名 | 説明 |
| :--- | :--- |
| **name** | スタッフの名前を記入（名前がダブらないように注意） |
| **type** | `常勤` または`非常勤` `派遣`（常勤は公休以外出勤） |
| **night_only** | `1`=夜勤専従 (優先的に夜勤を割当) |
| **off_days** | 月に必要な公休（休み）の日数 |
| **limit_consecutive** | 最大連勤数 (空欄=6連勤まで) |
| **can_〇〇** | (early=早番, day=日勤, late=遅番, night=夜勤)  `1`=可, `0`=不可 |
| **prev_night** | 前月末が夜勤なら`1` (1日が休みになります) |
| **prev_consecutive** | 前月末時点の連勤数を入力 |
| **mon**～**sun** | 曜日ごとの可否 (`1`=可, `0`=不可) |

### 3. 希望シフト (req_shift_1 ～ 31)
* **空欄**: おまかせ
* **休**: 希望休
* **有**: 有給 (公休とは別にカウント)
* **早/日/遅/夜**: シフト固定

### 4. 注意点
* 「夜勤不可」の人に「夜」を指定するとエラーになります。
* 公休数と希望休のバランスに注意してください。
"""

# =====================
# ページ設定
# =====================
st.set_page_config(page_title="シフト自動作成ツール", layout="wide")

st.title("📅 シフト自動作成ツール")
st.markdown("スタッフ設定ファイル(Excel)をアップロードして、作成ボタンを押してください。")

# =====================
# マニュアル表示機能 (ポップアップ)
# =====================
@st.dialog("📋 使い方マニュアル")
def show_manual():
    st.markdown(MANUAL_TEXT)

# =====================
# サイドバー：設定入力
# =====================
with st.sidebar:
    st.header("設定")
    
    # マニュアルボタン
    if st.button("📖 マニュアルを見る", use_container_width=True):
        show_manual()

    st.markdown("---")

    # 年月の入力
    default_year = datetime.date.today().year
    default_month = datetime.date.today().month + 1
    if default_month > 12:
        default_year += 1
        default_month = 1
        
    YEAR = st.number_input("作成する【年】", value=default_year, step=1)
    MONTH = st.number_input("作成する【月】", value=default_month, min_value=1, max_value=12, step=1)
    
    # ファイルアップロード
    uploaded_file = st.file_uploader("staff.xlsx をアップロード", type=["xlsx"])

    st.markdown("---")
    st.write("🔧 **詳細設定**")
    random_seed = st.number_input("再計算用乱数 (結果を変えたい時は数字を変更)", value=0, step=1)

# =====================
# メイン処理
# =====================
if st.button("シフトを作成する", type="primary"):
    if uploaded_file is None:
        st.error("エラー: Excelファイルをアップロードしてください。")
        st.stop()

    try:
        # 基本設定
        DAYS = calendar.monthrange(int(YEAR), int(MONTH))[1]
        st.info(f"{YEAR}年{MONTH}月 ({DAYS}日分) のシフトを計算中...")

        SHIFT_TYPES = ["早", "日", "遅", "夜"]
        REQUIRED = {"早": 1, "日": 1, "遅": 1, "夜": 2}

        base_date = datetime.date(int(YEAR), int(MONTH), 1)
        weekdays = [(base_date + datetime.timedelta(days=i)).weekday() for i in range(DAYS)]

        # Excel読み込み
        staff_df = pd.read_excel(uploaded_file, sheet_name="staff")
        staff = staff_df["name"].tolist()

        staff_type = dict(zip(staff_df["name"], staff_df["type"]))
        off_days = dict(zip(staff_df["name"], staff_df["off_days"]))

        # 希望入力
        req_input = {s: {} for s in staff}
        for _, r in staff_df.iterrows():
            s = r["name"]
            for d in range(DAYS):
                val = r.get(f"req_shift_{d+1}", None)
                if pd.isna(val) or val == "" or val == 0:
                    req_input[s][d] = None
                else:
                    req_input[s][d] = str(val).strip()

        prev_night = dict(zip(staff_df["name"], staff_df["prev_night"] if "prev_night" in staff_df.columns else [0]*len(staff)))
        prev_consecutive = dict(zip(staff_df["name"], staff_df["prev_consecutive"] if "prev_consecutive" in staff_df.columns else [0]*len(staff)))
        night_only = dict(zip(staff_df["name"], staff_df["night_only"] if "night_only" in staff_df.columns else [0]*len(staff)))

        can = {}
        weekday_can = {}
        for _, r in staff_df.iterrows():
            s = r["name"]
            can[s] = {"早": r["can_early"], "日": r["can_day"], "遅": r["can_late"], "夜": r["can_night"]}
            weekday_can[s] = {0: r["mon"], 1: r["tue"], 2: r["wed"], 3: r["thu"], 4: r["fri"], 5: r["sat"], 6: r["sun"]}

        limit_consecutive = dict(zip(staff_df["name"], staff_df["limit_consecutive"] if "limit_consecutive" in staff_df.columns else [6]*len(staff)))

        # モデル作成
        model = cp_model.CpModel()
        x = {}
        work_flag = {}

        for s in staff:
            for d in range(DAYS):
                for sh in SHIFT_TYPES:
                    x[s, d, sh] = model.NewBoolVar(f"x_{s}_{d}_{sh}") if can[s][sh] == 1 else None

        for s in staff:
            for d in range(DAYS):
                model.Add(sum(x[s, d, sh] for sh in SHIFT_TYPES if x[s, d, sh] is not None) <= 1)

        for d in range(DAYS):
            for sh in SHIFT_TYPES:
                model.Add(sum(x[s, d, sh] for s in staff if x[s, d, sh] is not None) == REQUIRED[sh])

        for s in staff:
            for d in range(DAYS):
                work_flag[s, d] = model.NewBoolVar(f"work_{s}_{d}")
                is_shifted = sum(x[s, d, sh] for sh in SHIFT_TYPES if x[s, d, sh] is not None)
                if d == 0:
                    is_after_night = prev_night[s]
                else:
                    is_after_night = x[s, d-1, "夜"] if x[s, d-1, "夜"] is not None else 0
                model.Add(work_flag[s, d] == is_shifted + is_after_night)

        for s in staff:
            for d in range(DAYS):
                inp = req_input[s][d]
                if inp is None: continue
                if inp == "休" or inp == "有":
                    model.Add(sum(x[s, d, sh] for sh in SHIFT_TYPES if x[s, d, sh] is not None) == 0)
                    model.Add(work_flag[s, d] == 0)
                elif inp in SHIFT_TYPES:
                    if x[s, d, inp] is not None:
                        model.Add(x[s, d, inp] == 1)
                        for other_sh in SHIFT_TYPES:
                            if other_sh != inp and x[s, d, other_sh] is not None:
                                model.Add(x[s, d, other_sh] == 0)

        for s in staff:
            for d in range(DAYS - 1):
                if x[s, d, "夜"] is not None:
                    model.Add(sum(x[s, d+1, sh] for sh in SHIFT_TYPES if x[s, d+1, sh] is not None) <= 1 - x[s, d, "夜"])

        for s in staff:
            limit = int(limit_consecutive[s]) if pd.notna(limit_consecutive[s]) else 6
            ng_days = limit + 1
            for d in range(DAYS - limit):
                window = [work_flag[s, d + k] for k in range(ng_days)]
                model.AddBoolOr([w.Not() for w in window])
            
            p_cons = prev_consecutive[s]
            if p_cons > 0:
                check_len = limit - p_cons + 1
                if check_len <= 0:
                    model.Add(work_flag[s, 0] == 0)
                elif check_len <= DAYS:
                    window = [work_flag[s, k] for k in range(check_len)]
                    model.AddBoolOr([w.Not() for w in window])

        for s in staff:
            for d in range(DAYS):
                if weekday_can[s][weekdays[d]] == 0:
                    model.Add(sum(x[s, d, sh] for sh in SHIFT_TYPES if x[s, d, sh] is not None) == 0)

        for s in staff:
            total_work = sum(work_flag[s, d] for d in range(DAYS))
            paid_leave = sum(1 for d in range(DAYS) if req_input[s][d] == "有")
            if staff_type[s] == "常勤":
                model.Add(total_work == DAYS - off_days[s] - paid_leave)
            else:
                model.Add(total_work <= DAYS - off_days[s] - paid_leave)

        # 目的関数
        night_count = {}
        for s in staff:
            night_count[s] = model.NewIntVar(0, DAYS, f"night_{s}")
            model.Add(night_count[s] == sum(x[s, d, "夜"] for d in range(DAYS) if x[s, d, "夜"] is not None))
        
        target_workers = [s for s in staff if can[s]["夜"] == 1 and night_only.get(s, 0) == 0]
        balance_penalty = []
        if target_workers:
            max_night = model.NewIntVar(0, DAYS, "max_night")
            min_night = model.NewIntVar(0, DAYS, "min_night")
            for s in target_workers:
                model.Add(night_count[s] <= max_night)
                model.Add(night_count[s] >= min_night)
            diff = model.NewIntVar(0, DAYS, "night_diff")
            model.Add(diff == max_night - min_night)
            balance_penalty.append(diff)

        night_penalty = []
        dispatch_penalty = []
        night_maximization_bonus = []
        night_only_staff = [s for s in staff if night_only.get(s, 0) == 1]

        for s in staff:
            for d in range(DAYS):
                if x[s, d, "夜"] is not None and night_only.get(s, 0) == 0:
                    night_penalty.append(x[s, d, "夜"])
                if staff_type[s] == "派遣":
                    for sh in SHIFT_TYPES:
                        if x[s, d, sh] is not None:
                            dispatch_penalty.append(x[s, d, sh])
        
        for s in night_only_staff:
            night_maximization_bonus.append(night_count[s])

        model.Minimize(
            10 * sum(night_penalty) +
            100 * sum(dispatch_penalty) +
            100 * sum(balance_penalty)
            - 1000 * sum(night_maximization_bonus)
        )

        # Solve
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = 30.0 
        solver.parameters.random_seed = int(random_seed)

        status = solver.Solve(model)

        if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            st.success(f"✅ 作成成功！ (Status: {solver.StatusName(status)})")

            result = []
            night_counts = [] 
            holiday_counts = [] 

            for s in staff:
                row = []
                for d in range(DAYS):
                    v = ""
                    for sh in SHIFT_TYPES:
                        if x[s, d, sh] is not None and solver.Value(x[s, d, sh]) == 1:
                            v = sh
                    row.append(v)
                result.append(row)
                
                n_count = solver.Value(night_count[s])
                w_days = sum(solver.Value(work_flag[s, d]) for d in range(DAYS))
                h_count = DAYS - w_days 
                night_counts.append(n_count)
                holiday_counts.append(h_count)
            
            df_out = pd.DataFrame(result, index=staff, columns=[f"{i+1}日" for i in range(DAYS)])
            df_out["夜勤回数"] = night_counts
            df_out["休日数"] = holiday_counts

            st.dataframe(df_out)

            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_out.to_excel(writer, sheet_name='Result')
            
            output.seek(0)
            wb = load_workbook(output)
            ws = wb.active
            green = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid")

            for i, s in enumerate(staff, start=2):
                off_need = off_days[s]
                off_cnt = 0
                for d in range(DAYS):
                    cell = ws.cell(row=i, column=d+2)
                    inp = req_input[s][d]
                    
                    if d > 0 and ws.cell(row=i, column=d+1).value == "夜":
                        cell.value = "明"
                        cell.fill = green
                        continue
                    if inp == "有":
                        cell.value = "有"
                        cell.fill = green
                        continue
                    if inp == "休":
                        cell.value = "休"
                        cell.fill = green
                        off_cnt += 1
                        continue
                    if (cell.value is None or cell.value == "") and off_cnt < off_need:
                        cell.value = "休"
                        cell.fill = green
                        off_cnt += 1
            
            final_output = BytesIO()
            wb.save(final_output)
            final_output.seek(0)

            st.download_button(
                label="📥 Excelファイルをダウンロード",
                data=final_output,
                file_name=f"shift_{YEAR}_{MONTH}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        else:
            st.error("❌ 解が見つかりませんでした。条件を緩和してください。")

    except Exception as e:
        st.error(f"エラーが発生しました: {e}")

