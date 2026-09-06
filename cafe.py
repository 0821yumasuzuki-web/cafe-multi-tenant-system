import streamlit as st
import pandas as pd
import os
import json
import time
from datetime import datetime, timezone, timedelta
import streamlit.components.v1 as components

# --- ページ基本設定 ---
st.set_page_config(
    page_title="マルチテナント型 カフェ総合管理システム", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 定数・ファイルパス ---
ACCOUNTS_FILE = "accounts.json"
JST = timezone(timedelta(hours=9))

def get_jst_now_str():
    return datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S")

def get_org_file_path(org_id):
    return f"data_{org_id}.xlsx"

# --- アカウント情報管理 ---
def load_accounts():
    if not os.path.exists(ACCOUNTS_FILE):
        return {}
    with open(ACCOUNTS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_accounts(accounts):
    with open(ACCOUNTS_FILE, "w", encoding="utf-8") as f:
        json.dump(accounts, f, ensure_ascii=False, indent=4)

# --- 初期データ定義 ---
DEFAULT_MENU = [
    {"メニューID": 1, "メニュー名": "スペシャルドリンクセット", "価格": 300, "イメージ色": "青", "カラーコード": "#1E90FF"},
    {"メニューID": 2, "メニュー名": "お茶会セット", "価格": 100, "イメージ色": "緑", "カラーコード": "#2E8B57"},
    {"メニューID": 3, "メニュー名": "おつまみセット", "価格": 100, "イメージ色": "赤", "カラーコード": "#CD5C5C"},
]

DEFAULT_TABLES = [
    {"テーブル番号": "1番卓", "定員": 2, "エリア": "メインフロア右", "ステータス": "空席"},
    {"テーブル番号": "2番卓", "定員": 2, "エリア": "メインフロア右", "ステータス": "空席"},
    {"テーブル番号": "3番卓", "定員": 4, "エリア": "メインフロア右", "ステータス": "空席"},
    {"テーブル番号": "4番卓", "定員": 4, "エリア": "メインフロア左", "ステータス": "空席"},
    {"テーブル番号": "5番卓", "定員": 2, "エリア": "メインフロア左", "ステータス": "空席"},
    {"テーブル番号": "6番卓", "定員": 2, "エリア": "メインフロア左", "ステータス": "空席"},
    {"テーブル番号": "7番卓", "定員": 2, "エリア": "メインフロア左", "ステータス": "空席"},
    {"テーブル番号": "8番卓", "定員": 4, "エリア": "受付・入口", "ステータス": "空席"},
]

# --- Excelデータ操作 ---
def init_org_excel(org_id):
    file_path = get_org_file_path(org_id)
    if not os.path.exists(file_path):
        df_menu = pd.DataFrame(DEFAULT_MENU)
        df_orders = pd.DataFrame(columns=["注文番号", "日時", "テーブル番号", "合計金額", "ステータス"])
        for _, m in df_menu.iterrows():
            df_orders[f"{m['メニュー名']}点数"] = 0
        df_tables = pd.DataFrame(DEFAULT_TABLES)
        save_org_sheets(org_id, df_menu, df_orders, df_tables)

def save_org_sheets(org_id, df_menu, df_orders, df_tables):
    file_path = get_org_file_path(org_id)
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        df_menu.to_excel(writer, sheet_name="メニュー情報", index=False)
        df_orders.to_excel(writer, sheet_name="注文履歴", index=False)
        df_tables.to_excel(writer, sheet_name="座席情報", index=False)

def load_org_data(org_id):
    file_path = get_org_file_path(org_id)
    if not os.path.exists(file_path):
        init_org_excel(org_id)
    df_menu = pd.read_excel(file_path, sheet_name="メニュー情報")
    df_orders = pd.read_excel(file_path, sheet_name="注文履歴")
    df_tables = pd.read_excel(file_path, sheet_name="座席情報")
    return df_menu, df_orders, df_tables

def update_table_status(org_id, table_name, status):
    df_menu, df_orders, df_tables = load_org_data(org_id)
    idx = df_tables[df_tables["テーブル番号"] == table_name].index
    if len(idx) > 0:
        df_tables.loc[idx, "ステータス"] = status
        save_org_sheets(org_id, df_menu, df_orders, df_tables)

def save_order(org_id, table_no, counts_dict, total_price):
    df_menu, df_orders, df_tables = load_org_data(org_id)
    next_id = 1 if len(df_orders) == 0 else int(df_orders["注文番号"].max()) + 1
    new_order = {
        "注文番号": next_id,
        "日時": get_jst_now_str(),
        "テーブル番号": table_no,
        "合計金額": total_price,
        "ステータス": "未提供"
    }
    for _, row in df_menu.iterrows():
        col_name = f"{row['メニュー名']}点数"
        new_order[col_name] = counts_dict.get(row['メニュー名'], 0)
    df_orders = pd.concat([df_orders, pd.DataFrame([new_order])], ignore_index=True)
    
    t_idx = df_tables[df_tables["テーブル番号"] == table_no].index
    if len(t_idx) > 0:
        df_tables.loc[t_idx, "ステータス"] = "利用中"
    save_org_sheets(org_id, df_menu, df_orders, df_tables)
    return next_id

def update_order_status(org_id, order_id, new_status):
    df_menu, df_orders, df_tables = load_org_data(org_id)
    df_orders.loc[df_orders["注文番号"] == order_id, "ステータス"] = new_status
    save_org_sheets(org_id, df_menu, df_orders, df_tables)

# --- SESSION STATE 初期化 ---
if "logged_org" not in st.session_state:
    st.session_state.logged_org = None
if "selected_mode" not in st.session_state:
    st.session_state.selected_mode = None
if "is_admin" not in st.session_state:
    st.session_state.is_admin = False
if "confirm_logout" not in st.session_state:
    st.session_state.confirm_logout = False
if "admin_view_page" not in st.session_state:
    st.session_state.admin_view_page = "⚙️ メニュー・レイアウト設定"

# --- ⑤ UI表示制御 (CSS) ---
# 管理者モード時のみサイドバーを表示し、それ以外は非表示
css_hide_elements = """
<style>
[data-testid="stHeader"] { display: none !important; }
footer { display: none !important; }
#MainMenu { display: none !important; }
.stDeployButton { display: none !important; }
</style>
"""

css_hide_sidebar = """
<style>
[data-testid="stSidebar"] { display: none !important; }
</style>
"""

if not st.session_state.is_admin:
    st.markdown(css_hide_elements + css_hide_sidebar, unsafe_allow_html=True)
else:
    st.markdown(css_hide_elements, unsafe_allow_html=True)

# 右クリック・ショートカット無効化
disable_devtools = """
<script>
const setupApp = function() {
    const targetDoc = window.parent.document || document;
    targetDoc.addEventListener('contextmenu', e => e.preventDefault(), true);
    targetDoc.addEventListener('keydown', e => {
        if (e.keyCode === 123 || ((e.ctrlKey || e.metaKey) && e.shiftKey && (e.keyCode === 73 || e.keyCode === 74)) || ((e.ctrlKey || e.metaKey) && e.keyCode === 85)) {
            e.preventDefault(); return false;
        }
    }, true);
};
if (document.readyState === "complete" || document.readyState === "interactive") setupApp();
else window.addEventListener("DOMContentLoaded", setupApp);
</script>
"""
components.html(disable_devtools, height=0, width=0)

# --- 共通ログアウト表示関数（① モード選択画面へ戻る） ---
def render_header_logout_button():
    col_a, col_b = st.columns([4, 1])
    with col_a:
        st.caption(f"🏫 ログイン中: **{st.session_state.logged_org}** ｜ モード: **{st.session_state.selected_mode}**")
    with col_b:
        if st.button("🚪 モード選択へ戻る", key="top_logout_btn", use_container_width=True):
            st.session_state.confirm_logout = True

    if st.session_state.confirm_logout:
        with st.container(border=True):
            st.warning("⚠️ **現在の画面を終了してモード選択画面へ戻りますか？**")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("はい（戻る）", type="primary", use_container_width=True):
                    st.session_state.selected_mode = None
                    st.session_state.is_admin = False
                    st.session_state.confirm_logout = False
                    st.rerun()
            with c2:
                if st.button("いいえ（キャンセル）", use_container_width=True):
                    st.session_state.confirm_logout = False
                    st.rerun()

# --- 🔐 ① 新規登録 ＆ ログイン画面 ---
if not st.session_state.logged_org:
    st.markdown("<br>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        with st.container(border=True):
            st.title("🏫 カフェ管理システム")
            
            active_tab = st.radio("メニュー選択", ["🔑 ログイン", "📝 新規団体登録"], horizontal=True, label_visibility="collapsed")
            accounts = load_accounts()

            if active_tab == "🔑 ログイン":
                org_input = st.text_input("団体ID", key="login_org_id")
                pass_input = st.text_input("ログインパスワード", type="password", key="login_pass")

                if st.button("ログイン", type="primary", use_container_width=True):
                    if org_input in accounts:
                        if pass_input == accounts[org_input]["password"]:
                            st.session_state.logged_org = org_input
                            init_org_excel(org_input)
                            st.success("ログインしました！")
                            time.sleep(0.3)
                            st.rerun()
                        else:
                            st.error("パスワードが正しくありません。")
                    else:
                        st.error("指定された団体IDが存在しません。")

            else:  # 新規団体登録
                new_org_id = st.text_input("希望の団体ID（例: class3a）", key="reg_org_id")
                new_pass = st.text_input("ログインパスワード設定", type="password", key="reg_pass")

                if st.button("新規団体を登録する", type="primary", use_container_width=True):
                    if not new_org_id or not new_pass:
                        st.error("団体IDとパスワードの両方を入力してください。")
                    elif new_org_id in accounts:
                        st.error("その団体IDは既に使われています。別のIDを指定してください。")
                    else:
                        accounts[new_org_id] = {
                            "password": new_pass,
                            "admin_password": None,
                            "is_first_login": True
                        }
                        save_accounts(accounts)
                        init_org_excel(new_org_id)
                        st.success("登録完了！ログインタブからログインしてください。")
                        time.sleep(1.2)
                        st.rerun()
    st.stop()

# アカウント情報取得
org_id = st.session_state.logged_org
accounts = load_accounts()
org_acc = accounts[org_id]

# --- ② 初回ログイン時 セットアップウィザード ---
if org_acc.get("is_first_login", False):
    st.title("🎉 ご登録ありがとうございます！")
    st.subheader("最初に店舗のメニューや卓数の初期設定を行いますか？")
    st.info("※設定は後から「管理者モード」でいつでも自由に変更できます。")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("⚙️ 今すぐ初期設定をする", type="primary", use_container_width=True):
            org_acc["is_first_login"] = False
            save_accounts(accounts)
            st.session_state.selected_mode = "👑 管理者モード"
            st.session_state.is_admin = True
            st.session_state.admin_view_page = "⚙️ メニュー・レイアウト設定"
            st.rerun()
    with col2:
        if st.button("⏩ 今回はスキップ（デフォルト設定で使用）", use_container_width=True):
            org_acc["is_first_login"] = False
            save_accounts(accounts)
            st.success("デフォルト設定で開始します！")
            time.sleep(0.5)
            st.rerun()
    st.stop()

# --- ③ ログイン後のモード選択メニュー ---
if not st.session_state.selected_mode:
    st.markdown("<br>", unsafe_allow_html=True)
    st.title(f"🏫 団体: {org_id} ｜ 起動モード選択")
    st.caption("利用する機能を選択してください。")
    st.markdown("---")

    c1, c2, c3 = st.columns(3)
    with c1:
        with st.container(border=True):
            st.header("🛒 レジ")
            st.write("注文入力と伝票発行")
            if st.button("「レジ」を起動", use_container_width=True, type="primary"):
                st.session_state.selected_mode = "🛒 レジ"
                st.rerun()
    with c2:
        with st.container(border=True):
            st.header("👨‍🍳 バックヤード")
            st.write("調理確認・提供チェック")
            if st.button("「バックヤード」を起動", use_container_width=True, type="primary"):
                st.session_state.selected_mode = "👨‍🍳 バックヤード"
                st.rerun()
    with c3:
        with st.container(border=True):
            st.header("🏢 店内退席確認")
            st.write("受付での空席・退席管理")
            if st.button("「退席確認」を起動", use_container_width=True, type="primary"):
                st.session_state.selected_mode = "🏢 店内退席確認"
                st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)
    c4, c5 = st.columns(2)
    with c4:
        with st.container(border=True):
            st.header("🗺️ 店内状況表示")
            st.write("リアルタイム座席マップ")
            if st.button("「店内状況」を起動", use_container_width=True):
                st.session_state.selected_mode = "🗺️ 店内状況表示"
                st.rerun()
    with c5:
        with st.container(border=True):
            st.header("👑 管理者")
            st.write("全画面閲覧・売上確認・各種設定")
            if st.button("「管理者モード」を起動", use_container_width=True):
                st.session_state.selected_mode = "👑 管理者モード"
                st.rerun()

    st.markdown("---")
    if st.button("🚪 団体ログイン画面へ（別団体でログイン）"):
        st.session_state.logged_org = None
        st.rerun()
    st.stop()

# --- 👑 管理者モード認証処理 ---
if st.session_state.selected_mode == "👑 管理者モード" and not st.session_state.is_admin:
    st.title("👑 管理者モード認証")
    admin_pw = org_acc.get("admin_password")

    # 初回管理者モード起動時（パスワード未設定の場合）
    if admin_pw is None:
        st.info("💡 管理者モード用のパスワードが未設定です。新しい管理者パスワードを設定してください。")
        new_admin_pw = st.text_input("新規 管理者パスワード", type="password", key="set_admin_pw")
        if st.button("管理者パスワードを保存して起動", type="primary"):
            if new_admin_pw:
                org_acc["admin_password"] = new_admin_pw
                save_accounts(accounts)
                st.session_state.is_admin = True
                st.success("パスワードを設定しました！")
                time.sleep(0.5)
                st.rerun()
            else:
                st.error("パスワードを入力してください。")
    else:
        input_admin_pw = st.text_input("管理者パスワードを入力", type="password", key="chk_admin_pw")
        c1, c2 = st.columns([1, 4])
        with c1:
            if st.button("認証", type="primary", use_container_width=True):
                if input_admin_pw == admin_pw:
                    st.session_state.is_admin = True
                    st.success("認証成功！")
                    time.sleep(0.3)
                    st.rerun()
                else:
                    st.error("管理者パスワードが違います。")
        with c2:
            if st.button("キャンセル（モード選択へ）"):
                st.session_state.selected_mode = None
                st.rerun()
    st.stop()

# 共通データの読込
df_menu, df_orders, df_tables = load_org_data(org_id)

# ----------------------------------------------------
# 画面描画ロジックの定義
# ----------------------------------------------------

def render_pos_page():
    st.title("🛒 注文入力（レジフロント）")
    for _, row in df_menu.iterrows():
        key = f"qty_{org_id}_{row['メニューID']}"
        if key not in st.session_state:
            st.session_state[key] = 0

    if "selected_table" not in st.session_state:
        st.session_state.selected_table = df_tables.iloc[0]["テーブル番号"] if len(df_tables) > 0 else "1番卓"

    st.markdown("### ☕ メニュー選択")
    cols = st.columns(min(len(df_menu), 4) if len(df_menu) > 0 else 1)
    counts = {}
    
    for i, row in df_menu.iterrows():
        m_id = row['メニューID']
        m_name = row['メニュー名']
        key = f"qty_{org_id}_{m_id}"
        
        target_col = cols[i % len(cols)]
        with target_col:
            st.markdown(
                f'<div style="border-left: 8px solid {row.get("カラーコード", "#1E90FF")}; padding-left: 10px; margin-bottom: 5px;">'
                f'<h4 style="margin-bottom:0px;">{m_name}</h4>'
                f'<p style="color: gray; margin-top:2px;">価格: {row["価格"]}円</p></div>', 
                unsafe_allow_html=True
            )
            
            b_minus, b_num, b_plus = st.columns([1, 1.5, 1])
            with b_minus:
                if st.button("➖", key=f"btn_m_{m_id}"):
                    if st.session_state[key] > 0:
                        st.session_state[key] -= 1
                        st.rerun()
            with b_num:
                st.markdown(f"<div style='text-align:center; font-size:1.5em; font-weight:bold;'>{st.session_state[key]}</div>", unsafe_allow_html=True)
            with b_plus:
                if st.button("➕", key=f"btn_p_{m_id}"):
                    st.session_state[key] += 1
                    st.rerun()
            counts[m_name] = st.session_state[key]

    st.markdown("---")
    st.markdown("### 🪑 割り当てるテーブル選択")
    t_cols = st.columns(min(len(df_tables), 8) if len(df_tables) > 0 else 1)
    
    for idx, (_, t_row) in enumerate(df_tables.iterrows()):
        t_name = t_row["テーブル番号"]
        is_vacant = (t_row["ステータス"] == "空席")
        btn_label = f"🟢 {t_name}" if is_vacant else f"🔴 {t_name}"
        btn_type = "primary" if st.session_state.selected_table == t_name else "secondary"
        
        with t_cols[idx % len(t_cols)]:
            if st.button(btn_label, key=f"tbl_btn_{t_name}", type=btn_type, use_container_width=True):
                st.session_state.selected_table = t_name
                st.rerun()

    selected_t_info = df_tables[df_tables["テーブル番号"] == st.session_state.selected_table].iloc[0]
    selected_is_vacant = (selected_t_info["ステータス"] == "空席")
    selected_cap = selected_t_info["定員"]
    
    total_items = sum(counts.values())
    total_price = sum(counts[row["メニュー名"]] * row["価格"] for _, row in df_menu.iterrows())
    
    st.markdown("---")
    st.markdown(f"## 💰 合計金額: <span style='color: #FF4B4B;'>{total_price} 円</span> （{total_items} 点） ｜ 選択: **{st.session_state.selected_table}** (定員:{selected_cap}名)", unsafe_allow_html=True)

    if not selected_is_vacant:
        st.warning(f"⚠️ **注意：{st.session_state.selected_table} は現在「利用中」です。**（追加注文・相席として登録可能）")

    if st.button("🚀 注文確定", type="primary", use_container_width=True, disabled=(total_items == 0)):
        order_id = save_order(org_id, st.session_state.selected_table, counts, total_price)
        st.success(f"🎉 注文確定！ 【注文番号: #{order_id} / テーブル: {st.session_state.selected_table}】")
        for _, row in df_menu.iterrows():
            st.session_state[f"qty_{org_id}_{row['メニューID']}"] = 0
        time.sleep(1.0)
        st.rerun()

def render_backyard_page():
    st.title("👨‍🍳 バックヤード（調理・提供）")
    active_orders = df_orders[df_orders["ステータス"] == "未提供"]
    
    c_left, c_right = st.columns([3, 1])
    with c_left:
        st.markdown(f"### 📋 未提供注文： {len(active_orders)} 件")
    with c_right:
        if st.button("🔄 最新情報に更新", use_container_width=True):
            st.rerun()

    for _, order in active_orders.iloc[::-1].iterrows():
        items_text = []
        for _, m_row in df_menu.iterrows():
            col_name = f"{m_row['メニュー名']}点数"
            if col_name in order and order[col_name] > 0:
                items_text.append(f"■ {m_row['メニュー名']}: {order[col_name]}点")
        
        c1, c2 = st.columns([3, 1])
        with c1:
            st.info(f"🪑 **{order['テーブル番号']}** ｜ 注文番号: #{order['注文番号']} ({order['日時']})\n\n" + " ｜ ".join(items_text))
        with c2:
            if st.button(f"✅ 提供完了 (#{order['注文番号']})", key=f"done_{order['注文番号']}", type="primary", use_container_width=True):
                update_order_status(org_id, order["注文番号"], "提供済み")
                st.rerun()

def render_reception_page():
    st.title("🏢 受付（退席確認）")
    cols = st.columns(2)
    for idx, (_, t_row) in enumerate(df_tables.iterrows()):
        t_name = t_row["テーブル番号"]
        status = t_row["ステータス"]
        cap = t_row["定員"]
        pending = df_orders[(df_orders["テーブル番号"] == t_name) & (df_orders["ステータス"] == "未提供")]
        
        with cols[idx % 2]:
            with st.container(border=True):
                c1, c2 = st.columns([2, 1])
                with c1:
                    st.subheader(f"🪑 {t_name} (定員{cap}名)")
                    if status == "利用中":
                        if len(pending) > 0:
                            st.warning(f"🔴 お客様在席中（⚠️ 未提供あり: {len(pending)}件）")
                        else:
                            st.error("🔴 お客様在席中（✅ 提供完了）")
                    else:
                        st.success("🟢 空席")
                with c2:
                    if status == "利用中":
                        if len(pending) > 0:
                            if st.button("⚠️ 強制退席", key=f"rec_out_{t_name}"):
                                update_table_status(org_id, t_name, "空席")
                                st.rerun()
                        else:
                            if st.button("📤 退席（空席）", key=f"rec_out_{t_name}", type="primary"):
                                update_table_status(org_id, t_name, "空席")
                                st.rerun()
                    else:
                        st.button("🟢 待機中", key=f"rec_idle_{t_name}", disabled=True)

def render_map_page():
    st.title("🗺️ リアルタイム店内座席マップ")
    areas = df_tables["エリア"].unique() if "エリア" in df_tables.columns else ["メインフロア"]
    
    for area in areas:
        st.subheader(f"📍 {area}")
        area_tables = df_tables[df_tables["エリア"] == area]
        cols = st.columns(min(len(area_tables), 4) if len(area_tables) > 0 else 1)
        
        for idx, (_, t_row) in enumerate(area_tables.iterrows()):
            t_name = t_row["テーブル番号"]
            status = t_row["ステータス"]
            cap = t_row["定員"]
            
            with cols[idx % len(cols)]:
                with st.container(border=True):
                    st.markdown(f"**🪑 {t_name}** ({cap}人席)")
                    if status == "利用中":
                        st.markdown("<span style='color:red; font-weight:bold;'>🔴 利用中</span>", unsafe_allow_html=True)
                    else:
                        st.markdown("<span style='color:green; font-weight:bold;'>🟢 空席</span>", unsafe_allow_html=True)

# ==========================================
# モードごとのメイン表示処理
# ==========================================

# 通常モード（非管理者）の場合
if st.session_state.selected_mode != "👑 管理者モード":
    render_header_logout_button()
    if st.session_state.selected_mode == "🛒 レジ":
        render_pos_page()
    elif st.session_state.selected_mode == "👨‍🍳 バックヤード":
        render_backyard_page()
    elif st.session_state.selected_mode == "🏢 店内退席確認":
        render_reception_page()
    elif st.session_state.selected_mode == "🗺️ 店内状況表示":
        render_map_page()

# 👑 管理者モード（全画面切替可能なサイドバー付き）
else:
    st.sidebar.title(f"👑 {org_id} 管理メニュー")
    
    page = st.sidebar.radio(
        "画面を選択してください",
        [
            "⚙️ メニュー・レイアウト設定",
            "📊 売上データ確認・取消",
            "🔐 管理者パスワード変更",
            "🧹 データ初期化",
            "--- 各機能の確認 ---",
            "🛒 レジ画面",
            "👨‍🍳 バックヤード画面",
            "🏢 店内退席確認画面",
            "🗺️ 店内状況表示画面"
        ],
        index=0
    )

    st.sidebar.markdown("---")
    if st.sidebar.button("🚪 モード選択へ戻る", use_container_width=True):
        st.session_state.selected_mode = None
        st.session_state.is_admin = False
        st.rerun()

    if st.sidebar.button("🚪 団体ログアウト", type="primary", use_container_width=True):
        st.session_state.logged_org = None
        st.session_state.selected_mode = None
        st.session_state.is_admin = False
        st.rerun()

    # 管理者用メインコンテンツ切り替え
    if page == "⚙️ メニュー・レイアウト設定":
        st.title("⚙️ 店舗初期設定・カスタマイズ")
        tab1, tab2 = st.tabs(["☕ メニュー設定", "🪑 テーブル設定"])
        with tab1:
            st.caption("※メニューの追加・編集・削除を行えます。変更後は下の「保存」を押してください。")
            edited_menu = st.data_editor(df_menu, num_rows="dynamic", use_container_width=True, key="menu_ed")
            if st.button("💾 メニュー設定を保存", type="primary"):
                save_org_sheets(org_id, edited_menu, df_orders, df_tables)
                st.success("保存しました！")
                time.sleep(0.5)
                st.rerun()
        with tab2:
            st.caption("※座席レイアウトや卓数の編集が行えます。")
            edited_tables = st.data_editor(df_tables, num_rows="dynamic", use_container_width=True, key="tbl_ed")
            if st.button("💾 テーブル設定を保存", type="primary"):
                save_org_sheets(org_id, df_menu, df_orders, edited_tables)
                st.success("保存しました！")
                time.sleep(0.5)
                st.rerun()

    elif page == "📊 売上データ確認・取消":
        st.title("📊 売上データ確認")
        valid_orders = df_orders[df_orders["ステータス"] != "取消済み"]
        st.metric("総売上金額", f"{valid_orders['合計金額'].sum():,} 円")
        st.markdown("---")
        st.subheader("🛠️ 注文履歴一覧")
        
        for _, order in df_orders.iloc[::-1].iterrows():
            c1, c2 = st.columns([4, 1])
            with c1:
                st.write(f"#{order['注文番号']} ｜ {order['テーブル番号']} ｜ {order['合計金額']}円 ｜ ステータス: {order['ステータス']}")
            with c2:
                if order["ステータス"] != "取消済み":
                    if st.button("❌ 取消", key=f"cncl_{order['注文番号']}"):
                        update_order_status(org_id, order["注文番号"], "取消済み")
                        st.rerun()

    elif page == "🔐 管理者パスワード変更":
        st.title("🔐 管理者パスワード変更")
        curr_admin_pass = st.text_input("新しい管理者パスワードを入力", type="password")
        if st.button("パスワードを変更する", type="primary"):
            if curr_admin_pass:
                org_acc["admin_password"] = curr_admin_pass
                save_accounts(accounts)
                st.success("管理者パスワードを更新しました！")
            else:
                st.error("パスワードを入力してください。")

    elif page == "🧹 データ初期化":
        st.title("🧹 データ初期化")
        st.warning("⚠️ データの初期化を行うと元に戻せません。")
        st.markdown("#### 1. 売上・注文履歴のみクリア（営業2日目など）")
        if st.button("🔥 注文・売上履歴をクリア"):
            df_orders_clean = pd.DataFrame(columns=["注文番号", "日時", "テーブル番号", "合計金額", "ステータス"])
            for _, m in df_menu.iterrows():
                df_orders_clean[f"{m['メニュー名']}点数"] = 0
            df_tables["ステータス"] = "空席"
            save_org_sheets(org_id, df_menu, df_orders_clean, df_tables)
            st.success("売上履歴をリセットしました！")
            time.sleep(0.5)
            st.rerun()

        st.markdown("---")
        st.markdown("#### 2. 全設定を初期化")
        if st.button("🚨 全データを初期化（デフォルトへ戻す）", type="primary"):
            init_org_excel(org_id)
            df_m = pd.DataFrame(DEFAULT_MENU)
            df_o = pd.DataFrame(columns=["注文番号", "日時", "テーブル番号", "合計金額", "ステータス"])
            df_t = pd.DataFrame(DEFAULT_TABLES)
            save_org_sheets(org_id, df_m, df_o, df_t)
            st.success("全データを初期化しました！")
            time.sleep(0.5)
            st.rerun()

    # 各専用画面のプレビュー・操作
    elif page == "🛒 レジ画面":
        render_pos_page()
    elif page == "👨‍🍳 バックヤード画面":
        render_backyard_page()
    elif page == "🏢 店内退席確認画面":
        render_reception_page()
    elif page == "🗺️ 店内状況表示画面":
        render_map_page()
