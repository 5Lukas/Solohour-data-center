import streamlit as st
from seatable_api import Base
from seatable_api.constants import ColumnTypes
import pandas as pd
import numpy as np  
import io
import time  
import json
import os
import xlsxwriter

st.set_page_config(page_title="数据处理中台", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

# ==========================================
# 本地配置存储引擎
# ==========================================
CONFIG_FILE = "seatable_etl_config.json"

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {"templates": {}, "last_used": {}}
    return {"templates": {}, "last_used": {}}

def save_config(data):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

if 'db_config' not in st.session_state:
    st.session_state.db_config = load_config()

def get_current_state():
    state = {
        "read_token": st.session_state.get("read_token", ""),
        "write_token": st.session_state.get("write_token", ""),
        "target_table": st.session_state.get("target_table_name", "Table1"),
        "table_names": [st.session_state.get(f"tbl_input_{i}", "") for i in range(st.session_state.num_tables)],
        "mappings": [],
        "vlookups": [],
        "transforms": []
    }
    for m in range(st.session_state.num_mappings):
        m_dict = {"sources": {}, "target": st.session_state.get(f"target_{m}", "")}
        for idx in range(st.session_state.num_tables):
            m_dict["sources"][idx] = st.session_state.get(f"map_{m}_tbl_{idx}", "-- 不选 --")
        state["mappings"].append(m_dict)
        
    for v in range(st.session_state.num_vlookups):
        state["vlookups"].append({
            "token": st.session_state.get(f"vtok_{v}", ""),
            "table": st.session_state.get(f"vtbl_{v}", ""),
            "main_key": st.session_state.get(f"vmk_{v}", ""),
            "ref_key": st.session_state.get(f"vrk_{v}", ""),
            "pull_col": st.session_state.get(f"vpc_{v}", ""),
            "out_name": st.session_state.get(f"vout_{v}", "")
        })
        
    for t in range(st.session_state.num_transforms):
        state["transforms"].append({
            "target_col": st.session_state.get(f"tcol_{t}", ""),
            "mode": st.session_state.get(f"tmode_{t}", "📅 日期格式化 (YYYY/MM/DD)"),
            "rate_val": st.session_state.get(f"trate_{t}", "1.0"),
            "out_name": st.session_state.get(f"tout_{t}", "")
        })
    return state

def apply_state(state):
    st.session_state.num_tables = max(2, len(state.get("table_names", [])))
    st.session_state.num_mappings = max(1, len(state.get("mappings", [])))
    st.session_state.num_vlookups = max(0, len(state.get("vlookups", [])))
    st.session_state.num_transforms = max(0, len(state.get("transforms", [])))
    st.session_state.loaded_state = state 

# ==========================================
# 基础工具函数
# ==========================================
@st.cache_resource
def get_base_client(token, server_url='https://cloud.seatable.cn'):
    base = Base(token, server_url)
    base.auth()
    return base

def get_all_rows_safely(base_client, table_name):
    all_rows = []
    start = 0; limit = 1000; max_retries = 3 
    while True:
        for attempt in range(max_retries):
            try:
                rows = base_client.list_rows(table_name, start=start, limit=limit)
                break 
            except Exception as e:
                if "503" in str(e) or attempt < max_retries - 1: time.sleep(2) 
                else: raise e 
        if not rows: break
        all_rows.extend(rows)
        start += limit
        time.sleep(0.2) 
    return all_rows

# ==========================================
# 状态初始化
# ==========================================
if 'connected' not in st.session_state: st.session_state.connected = False
if 'table_cols' not in st.session_state: st.session_state.table_cols = {}
if 'num_tables' not in st.session_state: st.session_state.num_tables = 2
if 'num_mappings' not in st.session_state: st.session_state.num_mappings = 3
if 'num_vlookups' not in st.session_state: st.session_state.num_vlookups = 1
if 'num_transforms' not in st.session_state: st.session_state.num_transforms = 1

def add_table(): st.session_state.num_tables += 1
def remove_table(): st.session_state.num_tables = max(1, st.session_state.num_tables - 1)
def add_mapping(): st.session_state.num_mappings += 1
def remove_mapping(): st.session_state.num_mappings = max(1, st.session_state.num_mappings - 1)
def add_vlookup(): st.session_state.num_vlookups += 1
def remove_vlookup(): st.session_state.num_vlookups = max(0, st.session_state.num_vlookups - 1)
def add_transform(): st.session_state.num_transforms += 1
def remove_transform(): st.session_state.num_transforms = max(0, st.session_state.num_transforms - 1)

ls = st.session_state.get("loaded_state", st.session_state.db_config.get("last_used", {}))

# ==========================================
# UI 样式注入
# ==========================================
st.markdown(\"""
<style>
    .step-header {
        padding: 10px 15px;
        border-radius: 8px;
        font-weight: 600;
        font-size: 1.2rem;
        margin-bottom: 15px;
        color: #333;
    }
    .bg-merge { background-color: #E3F2FD; border-left: 5px solid #2196F3; }
    .bg-vlookup { background-color: #E8F5E9; border-left: 5px solid #4CAF50; }
    .bg-transform { background-color: #FFF3E0; border-left: 5px solid #FF9800; }
    .bg-execute { background-color: #F3E5F5; border-left: 5px solid #9C27B0; }
</style>
\""", unsafe_allow_html=True)

# ==========================================
# 侧边栏：配置与数据源
# ==========================================
st.sidebar.markdown("### 💾 配置存档中心")
template_names = ["-- 选择配置模板 --"] + list(st.session_state.db_config["templates"].keys())
selected_tpl = st.sidebar.selectbox("读取模板", template_names, label_visibility="collapsed")
if st.sidebar.button("📂 载入所选配置", use_container_width=True):
    if selected_tpl != "-- 选择配置模板 --":
        apply_state(st.session_state.db_config["templates"][selected_tpl])
        st.rerun()

new_tpl_name = st.sidebar.text_input("将当前配置保存为模板", placeholder="输入专属模板名称...")
if st.sidebar.button("💾 存档当前配置", use_container_width=True):
    if new_tpl_name.strip():
        st.session_state.db_config["templates"][new_tpl_name.strip()] = get_current_state()
        save_config(st.session_state.db_config)
        st.sidebar.success("🎉 配置存档成功！")
        time.sleep(0.5)
        st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("### 🔌 API 核心配置")

read_token = st.sidebar.text_input("🔑 读取 Token (源数据)", value=ls.get("read_token", ""), type="password", key="read_token")
write_token = st.sidebar.text_input("🔑 写入 Token (目标表)", value=ls.get("write_token", ""), type="password", key="write_token")
target_table_name = st.sidebar.text_input("🎯 目标写入表名", value=ls.get("target_table", "Table1"), key="target_table_name") 

st.sidebar.markdown("---")
st.sidebar.markdown("### 📁 数据源表配置")

table_names = []
loaded_tables = ls.get("table_names", ["TK-ID-01-ALLORDER", "SP-ID-01-ALLORDER"])
for i in range(st.session_state.num_tables):
    default_val = loaded_tables[i] if i < len(loaded_tables) else ""
    t_name = st.sidebar.text_input(f"源表 {i+1}", value=default_val, key=f"tbl_input_{i}")
    if t_name.strip(): table_names.append(t_name.strip())

col_btn1, col_btn2 = st.sidebar.columns(2)
col_btn1.button("➕ 源表", on_click=add_table, use_container_width=True)
col_btn2.button("➖ 源表", on_click=remove_table, use_container_width=True)

st.sidebar.markdown("<br>", unsafe_allow_html=True)
if st.sidebar.button("🔄 连接并获取表结构", type="primary", use_container_width=True):
    with st.spinner("正在扫描数据库列结构..."):
        try:
            base_read = get_base_client(read_token)
            st.session_state.table_cols = {}
            for t_name in table_names:
                cols_info = base_read.list_columns(t_name)
                st.session_state.table_cols[t_name] = [c['name'] for c in cols_info]
            st.session_state.connected = True
            st.sidebar.success("✅ 连接通道建立成功！")
            st.session_state.db_config["last_used"] = get_current_state()
            save_config(st.session_state.db_config)
        except Exception as e:
            st.sidebar.error(f"❌ 连接失败: {e}")
            st.session_state.connected = False

# ==========================================
# 主界面
# ==========================================
st.markdown("<h1 style='text-align: center; color: #333;'>📊 企业级数据流转中台</h1>", unsafe_allow_html=True)
st.write("")

if st.session_state.connected:
    
    # ----------------=====================================================
    # 步骤一：数据合并
    # ----------------=====================================================
    st.markdown("<div class='step-header bg-merge'>🧬 步骤 1：多源数据纵向合并</div>", unsafe_allow_html=True)
    
    num_cols = len(table_names) + 1
    ui_cols = st.columns(num_cols)
    for idx, t_name in enumerate(table_names): ui_cols[idx].write(f"📥 **来源: {t_name}**")
    ui_cols[-1].write("🎯 **合并至目标列**")
    
    mappings = []
    loaded_mappings = ls.get("mappings", [])
    
    for m in range(st.session_state.num_mappings):
        row_cols = st.columns(num_cols)
        mapping_dict = {"Sources": {}, "Target": ""}
        has_selection = False
        lm = loaded_mappings[m] if m < len(loaded_mappings) else {"sources": {}, "target": ""}
        
        for idx, t_name in enumerate(table_names):
            options = ["-- 不选 --"] + st.session_state.table_cols.get(t_name, [])
            default_sel = lm["sources"].get(str(idx), "-- 不选 --") if isinstance(lm["sources"], dict) else lm["sources"].get(idx, "-- 不选 --")
            sel_idx = options.index(default_sel) if default_sel in options else 0
            with row_cols[idx]:
                sel = st.selectbox(f"{t_name}_{m}", options, index=sel_idx, key=f"map_{m}_tbl_{idx}", label_visibility="collapsed")
                mapping_dict["Sources"][t_name] = sel
                if sel != "-- 不选 --": has_selection = True
                    
        with row_cols[-1]:
            target_col = st.text_input(f"tcol_{m}", value=lm.get("target", ""), key=f"target_{m}", placeholder="输入合并后的列名", label_visibility="collapsed")
            mapping_dict["Target"] = target_col.strip()
            
        if has_selection and mapping_dict["Target"]: mappings.append(mapping_dict)
            
    col_m1, col_m2, _ = st.columns([1, 1, 6])
    col_m1.button("➕ 添加合并列", on_click=add_mapping, use_container_width=True)
    col_m2.button("➖ 移除合并列", on_click=remove_mapping, use_container_width=True)

    # ----------------=====================================================
    # 步骤二：跨表匹配 (VLOOKUP)
    # ----------------=====================================================
    st.write("")
    st.markdown("<div class='step-header bg-vlookup'>🔗 步骤 2：跨表数据关联匹配 (VLOOKUP)</div>", unsafe_allow_html=True)

    vlookup_rules = []
    loaded_vlookups = ls.get("vlookups", [])

    if st.session_state.num_vlookups > 0:
        th_cols = st.columns([1.5, 1.5, 1.5, 1.5, 1.5, 1.5])
        th_cols[0].write("🔑 外部 Token (留空同主)")
        th_cols[1].write("📄 外部表名")
        th_cols[2].write("📌 主表匹配列")
        th_cols[3].write("📌 外部匹配列")
        th_cols[4].write("📥 提取目标列")
        th_cols[5].write("🎯 写入新列名")

    for v in range(st.session_state.num_vlookups):
        lv = loaded_vlookups[v] if v < len(loaded_vlookups) else {}
        r_cols = st.columns([1.5, 1.5, 1.5, 1.5, 1.5, 1.5])
        with r_cols[0]: v_token = st.text_input(f"vtok_{v}", value=lv.get("token", ""), type="password", key=f"vtok_{v}", label_visibility="collapsed")
        with r_cols[1]: v_table = st.text_input(f"vtbl_{v}", value=lv.get("table", ""), key=f"vtbl_{v}", placeholder="输入外部表名", label_visibility="collapsed")
        with r_cols[2]: v_main_key = st.text_input(f"vmk_{v}", value=lv.get("main_key", ""), key=f"vmk_{v}", placeholder="如: 订单号", label_visibility="collapsed")
        with r_cols[3]: v_ref_key = st.text_input(f"vrk_{v}", value=lv.get("ref_key", ""), key=f"vrk_{v}", placeholder="如: 单号", label_visibility="collapsed")
        with r_cols[4]: v_pull_col = st.text_input(f"vpc_{v}", value=lv.get("pull_col", ""), key=f"vpc_{v}", placeholder="如: 出库数量", label_visibility="collapsed")
        with r_cols[5]: v_out_name = st.text_input(f"vout_{v}", value=lv.get("out_name", ""), key=f"vout_{v}", placeholder="最终呈现列名", label_visibility="collapsed")
            
        if v_table.strip() and v_main_key.strip() and v_ref_key.strip() and v_pull_col.strip() and v_out_name.strip():
            vlookup_rules.append({
                "token": v_token.strip() if v_token.strip() else read_token,
                "table": v_table.strip(),
                "main_key": v_main_key.strip(),
                "ref_key": v_ref_key.strip(),
                "pull_col": v_pull_col.strip(),
                "out_name": v_out_name.strip()
            })

    col_v1, col_v2, _ = st.columns([1, 1, 6])
    col_v1.button("➕ 添加匹配规则", on_click=add_vlookup, use_container_width=True)
    col_v2.button("➖ 移除匹配规则", on_click=remove_vlookup, use_container_width=True)

    # ----------------=====================================================
    # 步骤三：本表列清洗换算
    # ----------------=====================================================
    st.write("")
    st.markdown("<div class='step-header bg-transform'>🛠️ 步骤 3：列数据清洗与汇率换算</div>", unsafe_allow_html=True)

    TRANSFORM_MODES = ["📅 日期格式化 (YYYY/MM/DD)", "💱 汇率除法换算"]
    transform_rules = []
    loaded_transforms = ls.get("transforms", [])

    if st.session_state.num_transforms > 0:
        th_cols = st.columns([2, 2, 2, 2])
        th_cols[0].write("📌 待处理主表列")
        th_cols[1].write("⚙️ 处理方式")
        th_cols[2].write("🔢 汇率数值 (仅汇率模式有效)")
        th_cols[3].write("🎯 写入目标列 (可填原列名)")

    for t in range(st.session_state.num_transforms):
        lt = loaded_transforms[t] if t < len(loaded_transforms) else {}
        r_cols = st.columns([2, 2, 2, 2])
        
        with r_cols[0]:
            t_col = st.text_input(f"tcol_{t}", value=lt.get("target_col", ""), key=f"tr_col_{t}", placeholder="如: 下单时间 / 金额", label_visibility="collapsed")
            
        with r_cols[1]:
            mode_val = lt.get("mode", TRANSFORM_MODES[0])
            mode_idx = TRANSFORM_MODES.index(mode_val) if mode_val in TRANSFORM_MODES else 0
            t_mode = st.selectbox(f"tmode_{t}", TRANSFORM_MODES, index=mode_idx, key=f"tr_mode_{t}", label_visibility="collapsed")
            
        with r_cols[2]:
            if "汇率" in t_mode:
                t_rate = st.text_input(f"trate_{t}", value=lt.get("rate_val", "1.0"), key=f"tr_rate_{t}", placeholder="填入数字, 如: 7.2", label_visibility="collapsed")
            else:
                t_rate = st.text_input(f"trate_{t}", value="-", disabled=True, key=f"tr_rate_{t}", label_visibility="collapsed")
                
        with r_cols[3]:
            t_out = st.text_input(f"tout_{t}", value=lt.get("out_name", ""), key=f"tr_out_{t}", placeholder="填入最终列名", label_visibility="collapsed")
            
        if t_col.strip() and t_out.strip():
            transform_rules.append({
                "target_col": t_col.strip(),
                "mode": t_mode,
                "rate_val": t_rate.strip() if t_rate.strip() != "-" else "1.0",
                "out_name": t_out.strip()
            })

    col_t1, col_t2, _ = st.columns([1, 1, 6])
    col_t1.button("➕ 添加清洗规则", on_click=add_transform, use_container_width=True)
    col_t2.button("➖ 移除清洗规则", on_click=remove_transform, use_container_width=True)

    # ----------------=====================================================
    # 步骤四：执行流转
    # ----------------=====================================================
    st.write("")
    st.markdown("<div class='step-header bg-execute'>🚀 步骤 4：全局执行中枢</div>", unsafe_allow_html=True)
    
    if st.button("▶️ 启动流转引擎并回写线上数据", type="primary", use_container_width=True):
        if not mappings:
            st.warning("⚠️ 请配置至少一条基础数据合并规则。")
        else:
            st.session_state.db_config["last_used"] = get_current_state()
            save_config(st.session_state.db_config)
            
            with st.container():
                try:
                    base_read = get_base_client(read_token)
                    base_write = get_base_client(write_token)
                    
                    st.info("⏳ [1/4] 正在并行合并底层大表...")
                    all_table_data = {}
                    for t_name in table_names:
                        all_table_data[t_name] = get_all_rows_safely(base_read, t_name)
                    
                    final_records = []
                    for t_name, rows in all_table_data.items():
                        for row in rows:
                            new_row = {}
                            for m in mappings:
                                s_col = m["Sources"].get(t_name, "-- 不选 --")
                                if s_col != "-- 不选 --": new_row[m["Target"]] = row.get(s_col, "")
                            final_records.append(new_row)
                            
                    df_final = pd.DataFrame(final_records)
                    
                    if vlookup_rules:
                        st.info("⏳ [2/4] 启动跨表关联，调取外部数据 (VLOOKUP)...")
                        for idx, rule in enumerate(vlookup_rules):
                            ref_base = get_base_client(rule["token"])
                            ref_rows = get_all_rows_safely(ref_base, rule["table"])
                            df_ref = pd.DataFrame(ref_rows)
                            if df_ref.empty: continue
                                
                            if rule["ref_key"] in df_ref.columns and rule["pull_col"] in df_ref.columns:
                                temp_ref = f"__tmp_ref_{idx}"
                                df_ref_clean = pd.DataFrame({
                                    temp_ref: df_ref[rule["ref_key"]].astype(str),
                                    rule["out_name"]: df_ref[rule["pull_col"]]
                                }).drop_duplicates(subset=[temp_ref], keep='first')
                                
                                df_final[rule["main_key"]] = df_final[rule["main_key"]].astype(str)
                                df_final = pd.merge(df_final, df_ref_clean, left_on=rule["main_key"], right_on=temp_ref, how="left")
                                df_final = df_final.drop(columns=[temp_ref], errors='ignore')
                                
                    if transform_rules:
                        st.info("⏳ [3/4] 驱动清洗模块：修正格式与汇率核算...")
                        for rule in transform_rules:
                            t_col = rule["target_col"]
                            out_col = rule["out_name"]
                            if t_col not in df_final.columns:
                                df_final[t_col] = ""
                                
                            if "日期" in rule["mode"]:
                                # 【核心修复】：使用 apply 逐行解析，解决多数据源格式混合导致 Pandas 猜测失败变空值的问题
                                def safe_parse_date(d):
                                    val = str(d).strip()
                                    if val in ('', 'nan', 'None', 'NaT'): return pd.NaT
                                    try: return pd.to_datetime(val)
                                    except: return pd.NaT
                                
                                parsed_date = df_final[t_col].apply(safe_parse_date)
                                df_final[out_col] = parsed_date.dt.strftime('%Y/%m/%d').fillna("")
                                
                            elif "汇率" in rule["mode"]:
                                try: rate_num = float(rule["rate_val"])
                                except: rate_num = 1.0
                                v_money = pd.to_numeric(df_final[t_col], errors='coerce').fillna(0)
                                df_final[out_col] = np.where(rate_num == 0, 0, v_money / rate_num)
                                df_final[out_col] = df_final[out_col].round(2)
                    
                    st.info("⏳ [4/4] 模型校准与数据回灌中枢...")
                    try:
                        target_cols_info = base_write.list_columns(target_table_name)
                        existing_cols = [c['name'] for c in target_cols_info]
                        needed_cols = list(df_final.columns)
                        for col_name in existing_cols[1:]:
                            if col_name not in needed_cols:
                                try: base_write.delete_column(target_table_name, col_name)
                                except: pass
                        target_cols_info_fresh = base_write.list_columns(target_table_name)
                        current_existing_cols = [c['name'] for c in target_cols_info_fresh]
                        used_needed_cols = [c for c in current_existing_cols if c in needed_cols]
                        remaining_needed = [c for c in needed_cols if c not in used_needed_cols]
                        one_counter = 1
                        for idx, col_name in enumerate(current_existing_cols):
                            if col_name not in needed_cols:
                                if remaining_needed: base_write.rename_column(target_table_name, col_name, remaining_needed.pop(0))
                                else:
                                    try:
                                        base_write.rename_column(target_table_name, col_name, "1" if one_counter == 1 else f"1_{one_counter}")
                                        one_counter += 1
                                    except: pass
                        target_cols_info_final = base_write.list_columns(target_table_name)
                        final_existing_cols = [c['name'] for c in target_cols_info_final]
                        for col in needed_cols:
                            if col not in final_existing_cols:
                                base_write.insert_column(target_table_name, col, ColumnTypes.TEXT)
                    except Exception as e: pass 
                    
                    existing_rows = get_all_rows_safely(base_write, target_table_name)
                    if existing_rows:
                        existing_ids = [r['_id'] for r in existing_rows]
                        for i in range(0, len(existing_ids), 1000):
                            for attempt in range(3):
                                try:
                                    base_write.batch_delete_rows(target_table_name, existing_ids[i:i+1000])
                                    break
                                except Exception: time.sleep(2)
                            time.sleep(0.5) 
                    
                    final_records_to_send = []
                    for _, row_series in df_final.iterrows():
                        row_dict = {}
                        for col in df_final.columns:
                            val = row_series[col]
                            if pd.notna(val) and str(val).strip() != "": row_dict[col] = str(val)
                        if row_dict: final_records_to_send.append(row_dict)
                            
                    write_progress = st.progress(0)
                    total_records = len(final_records_to_send)
                    for i in range(0, total_records, 1000):
                        chunk = final_records_to_send[i: i+1000]
                        for attempt in range(3):
                            try:
                                base_write.batch_append_rows(target_table_name, chunk)
                                break
                            except Exception: time.sleep(2)
                        write_progress.progress(min((i + 1000) / total_records, 1.0))
                        time.sleep(0.5) 
                        
                    st.success(f"🎉 任务流执行圆满成功！已输送并覆盖 {total_records} 行标准清洗数据。")
                    st.dataframe(df_final.head(10), use_container_width=True)
                    
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer: df_final.to_excel(writer, index=False)
                    st.download_button("📥 导出本次执行备份 (Excel)", data=buffer.getvalue(), file_name="处理结果_备份.xlsx", mime="application/vnd.ms-excel")
                    
                except Exception as e: st.error(f"❌ 流程中断: {e}")