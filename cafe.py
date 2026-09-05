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

# --- 🛡️ UIカスタマイズ ＆ 保護設定 ---
hide_and_protect_style = """
    <style>
    [data-testid="stHeader"] > div:nth-child(2) { display: none !important; visibility: hidden !important; }
    .stActionButton, [data-testid="stActionButton"], [data-testid="stHeaderActionElements"] { display: none !important; visibility: hidden !important; }
    #MainMenu { visibility: hidden !important; display: none !important; }
    .stAppDeployButton { display: none !important; visibility: hidden !important; }
    footer { visibility: hidden !important; display: none !important; }
    [data-testid="stSidebar"] { display: block !important; visibility: visible !important; }
    body { user-select: none; -webkit-user-select: none; }
    input, textarea { user-select: text !important; -webkit-user-select: text !important; }
    </style>
"""
st.markdown(hide_and_protect_style, unsafe_allow_html=True)

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

# --- 定数・ファイルパス ---
ACCOUNTS_FILE = "accounts.json"
JST = timezone(timedelta(hours=9))

def get_jst_now_str():
    return datetime.now(JST).strftime("%Y-%m-%d %H:%M:%S")

def get_org_file_path(org_id):
    return f"data_{org_id}.xlsx"

# --- アカウント情報管理関数 ---
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

# --- 団体別Excelデータ操作 ---
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
if "is_admin" not in st.session_state:
    st.session_state.is_admin = False
if "authenticated_pages" not in st.session_state:
    st.session_state.authenticated_pages = set()

# --- 🔐 ① ログイン ＆ 団体登録画面 ---
if not st.session_state.logged_org:
    st.markdown("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.container(border=True):
            st.title("🏢 文化祭カフェ 総合ログイン")
            tab_login, tab_register = st.tabs(["🔑 ログイン", "📝 新規団体登録"])
            accounts = load_accounts()

            # --- ログイン タブ ---
            with tab_login:
                org_input = st.text_input("団体ID（ユーザーネーム）", key="login_org_id")
                login_role = st.radio("ログインモード", ["管理者モード", "一般スタッフモード"], horizontal=True)
                pass_input = st.text_input("パスワード", type="password", key="login_pass")

                if st.button("ログイン", type="primary", use_container_width=True):
                    if org_input in accounts:
                        org_acc = accounts[org_input]
                        if login_role == "管理者モード":
                            if pass_input == org_acc["passwords"]["admin"]:
                                st.session_state.logged_org = org_input
                                st.session_state.is_admin = True
                                st.success("管理者としてログインしました！")
                                init_org_excel(org_input)
                                time.sleep(0.5)
                                st.rerun()
                            else:
                                st.error("管理者パスワードが正しくありません。")
                        else:
                            st.session_state.logged_org = org_input
                            st.session_state.is_admin = False
                            st.success("スタッフモードでログインしました！")
                            init_org_excel(org_input)
                            time.sleep(0.5)
                            st.rerun()
                    else:
                        st.error("指定された団体IDが存在しません。")

            # --- 新規登録 タブ ---
            with tab_register:
                new_org_id = st.text_input("新規 団体ID（例: class3a）", key="reg_org_id")
                new_admin_pass = st.text_input("🔑 管理者用パスワード", type="password", key="reg_admin_pass")
                st.caption("※ 各画面の閲覧用個別パスワードを設定してください")
                
                reg_p_register = st.text_input("🛒 レジ画面用パスワード", value="1111", type="password")
                reg_p_backyard = st.text_input("👨‍🍳 バックヤード画面用パスワード", value="2222", type="password")
                reg_p_reception = st.text_input("🏢 受付画面用パスワード", value="3333", type="password")
                reg_p_map = st.text_input("🗺️ 座席マップ画面用パスワード", value="4444", type="password")
                reg_p_sales = st.text_input("📊 売上確認画面用パスワード", value="5555", type="password")

                if st.button("新規アカウントを作成", type="primary", use_container_width=True):
                    if not new_org_id or not new_admin_pass:
                        st.error("団体IDと管理者用パスワードは必須です。")
                    elif new_org_id in accounts:
                        st.error("その団体IDは既に使われています。別のIDを指定してください。")
                    else:
                        accounts[new_org_id] = {
                            "passwords": {
                                "admin": new_admin_pass,
                                "🛒 注文入力画面": reg_p_register,
                                "👨‍🍳 バックヤード画面": reg_p_backyard,
                                "🏢 受付（退席確認）": reg_p_reception,
                                "🗺️ 店内座席マップ": reg_p_map,
                                "📊 売計データ確認・取消": reg_p_sales
                            }
                        }
                        save_accounts(accounts)
                        init_org_excel(new_org_id)
                        st.success("登録が完了しました！ログインタブからログインしてください。")
    st.stop()

# --- 🚀 ログイン後のシステム管理領域 ---
org_id = st.session_state.logged_org
accounts = load_accounts()
org_passwords = accounts[org_id]["passwords"]

# サイドバー表示
st.sidebar.title(f"🏫 団体: {org_id}")
if st.session_state.is_admin:
    st.sidebar.success("👑 管理者モードで接続中")
else:
    st.sidebar.info("👤 スタッフモードで接続中")

# メニュー選択肢の構築
menu_options = [
    "🛒 注文入力画面", 
    "🏢 受付（退席確認）", 
    "👨‍🍳 バックヤード画面", 
    "🗺️ 店内座席マップ", 
    "📊 売計データ確認・取消"
]
if st.session_state.is_admin:
    menu_options.append("⚙️ 店舗・レイアウト設定")
    menu_options.append("🧹 データ初期化")

mode = st.sidebar.radio("画面を選択してください", menu_options)

st.sidebar.markdown("---")
if st.sidebar.button("🔄 画面情報更新", use_container_width=True):
    st.rerun()

if st.sidebar.button("🚪 ログアウト", use_container_width=True):
    st.session_state.logged_org = None
    st.session_state.is_admin = False
    st.session_state.authenticated_pages = set()
    st.rerun()

# --- 🔒 画面個別のアクセス権限チェック ---
def check_page_permission(page_name):
    if st.session_state.is_admin:
        return True
    if page_name in st.session_state.authenticated_pages:
        return True
    
    st.title(f"🔒 {page_name}")
    st.warning(f"このページにアクセスするには「{page_name}」専用のパスワードが必要です。")
    pw_input = st.text_input("パスワードを入力", type="password", key=f"auth_pw_{page_name}")
    if st.button("閲覧許可", type="primary"):
        target_pw = org_passwords.get(page_name, "")
        if pw_input == target_pw:
            st.session_state.authenticated_pages.add(page_name)
            st.success("認証成功！")
            time.sleep(0.3)
            st.rerun()
        else:
            st.error("パスワードが違います。")
    return False

# データ読込
df_menu, df_orders, df_tables = load_org_data(org_id)

# --- 各画面のレンダリング ---

# 1. 注文入力画面
if mode == "🛒 注文入力画面":
    if check_page_permission(mode):
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
            st.warning(f"⚠️ **注意：{st.session_state.selected_table} は現在「利用中」です。**（追加注文・相席としてそのまま確定可能です）")
        if total_items > selected_cap:
            st.warning(f"⚠️ **注意：定員を超えています。**（定員 {selected_cap} 名に対し {total_items} 点選択中）")

        if st.button("🚀 注文確定", type="primary", use_container_width=True, disabled=(total_items == 0)):
            order_id = save_order(org_id, st.session_state.selected_table, counts, total_price)
            st.success(f"🎉 注文確定！ 【注文番号: #{order_id} / テーブル: {st.session_state.selected_table}】")
            for _, row in df_menu.iterrows():
                st.session_state[f"qty_{org_id}_{row['メニューID']}"] = 0
            time.sleep(1.0)
            st.rerun()

# 2. 受付（退席確認）画面
elif mode == "🏢 受付（退席確認）":
    if check_page_permission(mode):
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

# 3. バックヤード画面
elif mode == "👨‍🍳 バックヤード画面":
    if check_page_permission(mode):
        st.title("👨‍🍳 バックヤード（調理・提供）")
        active_orders = df_orders[df_orders["ステータス"] == "未提供"]
        st.markdown(f"### 📋 未提供注文： {len(active_orders)} 件")
        
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
                if st.button(f"✅ 提供完了 (#{order['注文番号']})", key=f"done_{order['注文番号']}", type="primary"):
                    update_order_status(org_id, order["注文番号"], "提供済み")
                    st.rerun()

# 4. 店内座席マップ画面
elif mode == "🗺️ 店内座席マップ":
    if check_page_permission(mode):
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

# 5. 売計データ確認・取消画面
elif mode == "📊 売計データ確認・取消":
    if check_page_permission(mode):
        st.title("📊 売計データ確認 ＆ 注文取消")
        valid_orders = df_orders[df_orders["ステータス"] != "取消済み"]
        
        st.metric("総売上金額", f"{valid_orders['合計金額'].sum():,} 円")
        st.markdown("---")
        st.subheader("🛠️ 注文履歴")
        
        for _, order in df_orders.iloc[::-1].iterrows():
            c1, c2 = st.columns([4, 1])
            with c1:
                st.write(f"#{order['注文番号']} ｜ {order['テーブル番号']} ｜ {order['合計金額']}円 ｜ ステータス: {order['ステータス']}")
            with c2:
                if order["ステータス"] != "取消済み":
                    if st.button("❌ 取消", key=f"cncl_{order['注文番号']}"):
                        update_order_status(org_id, order["注文番号"], "取消済み")
                        st.rerun()

# 6. ⚙️ 店舗・レイアウト設定（管理者専用）
elif mode == "⚙️ 店舗・レイアウト設定":
    if st.session_state.is_admin:
        st.title("⚙️ 店舗・レイアウト設定（管理者専用）")
        
        tab1, tab2 = st.tabs(["☕ メニュー編集", "🪑 テーブル・レイアウト編集"])
        
        with tab1:
            st.subheader("現在のメニュー構成")
            edited_menu = st.data_editor(df_menu, num_rows="dynamic", use_container_width=True, key="menu_editor")
            if st.button("💾 メニュー設定を保存", type="primary"):
                save_org_sheets(org_id, edited_menu, df_orders, df_tables)
                st.success("メニュー情報を更新しました！")
                time.sleep(0.5)
                st.rerun()

        with tab2:
            st.subheader("現在のテーブル設定")
            edited_tables = st.data_editor(df_tables, num_rows="dynamic", use_container_width=True, key="table_editor")
            if st.button("💾 テーブル設定を保存", type="primary"):
                save_org_sheets(org_id, df_menu, df_orders, edited_tables)
                st.success("テーブル情報を更新しました！")
                time.sleep(0.5)
                st.rerun()

# 7. 🧹 データ初期化（管理者専用）
elif mode == "🧹 データ初期化":
    if st.session_state.is_admin:
        st.title("🧹 データ初期化（管理者専用）")
        st.warning("⚠️ データの初期化を行うと元に戻すことはできません。慎重に行ってください。")
        
        st.markdown("### 1. 注文・売上履歴のみリセット（営業2日目開始時等）")
        if st.button("🔥 注文・売上履歴を完全削除", type="secondary"):
            df_orders_clean = pd.DataFrame(columns=["注文番号", "日時", "テーブル番号", "合計金額", "ステータス"])
            for _, m in df_menu.iterrows():
                df_orders_clean[f"{m['メニュー名']}点数"] = 0
            df_tables["ステータス"] = "空席"
            save_org_sheets(org_id, df_menu, df_orders_clean, df_tables)
            st.success("注文履歴と卓のステータスをリセットしました！")
            time.sleep(0.5)
            st.rerun()
            
        st.markdown("---")
        st.markdown("### 2. 全データ・設定の初期化（デフォルトに戻す）")
        if st.button("🚨 メニュー・卓設定含め全データを初期化", type="primary"):
            init_org_excel(org_id)
            df_m = pd.DataFrame(DEFAULT_MENU)
            df_o = pd.DataFrame(columns=["注文番号", "日時", "テーブル番号", "合計金額", "ステータス"])
            df_t = pd.DataFrame(DEFAULT_TABLES)
            save_org_sheets(org_id, df_m, df_o, df_t)
            st.success("初期状態に全リセットしました！")
            time.sleep(0.5)
            st.rerun()
