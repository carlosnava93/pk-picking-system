import streamlit as st
import pandas as pd
import numpy as np
import os
import re
import warnings
from datetime import datetime
import io

# ==================== Configuración Básica / 网页基础配置 ====================
st.set_page_config(page_title="Sistema de Picking Inteligente", page_icon="📦", layout="wide")

st.title("📦 Sistema de Generación Automática de Lista de Picking")
st.subheader("智能拣货单（Picking）自动生成系统")
st.markdown("---")

# Instrucciones en la barra lateral / 侧边栏说明
st.sidebar.header("💡 Instrucciones / 使用说明")
st.sidebar.info(
    "**[ESPAÑOL]**\n"
    "1. El sistema identificará los archivos automáticamente según las columnas.\n"
    "2. El sistema extrae automáticamente la fecha del archivo.\n"
    "3. Después de la conversión, verifique si el [Total de Cajas] es correcto.\n\n"
    "--- \n\n"
    "**[中文]**\n"
    "1. 系统会自动根据最新特征列名识别并处理文件。\n"
    "2. 系统自动从出库单中智能精准搜索业务日期用于命名。\n"
    "3. 转换完成后，请在右侧核对【总箱数】是否账实相符。"
)

# ==================== Componentes de Carga / 文件上传组件 ====================
col1, col2 = st.columns(2)
with col1:
    file_a = st.file_uploader(
        "1. Cargue la [Hoja de Salida / Canal] (.xlsx, .xls) | 请上传【当日出库单表 / 渠道表】", 
        type=["xlsx", "xls"]
    )
with col2:
    file_b = st.file_uploader(
        "2. Cargue la [Tabla de Inventario / Ubicación] (.xlsx, .xls) | 请上传【整柜仓库库存表 / 库位表】", 
        type=["xlsx", "xls"]
    )

# ==================== Lógica Central / 核心业务逻辑 ====================
if file_a and file_b:
    with st.spinner("🚀 Procesando datos con precisión, por favor espere... | 正在智能识别并精准处理数据，请稍候..."):
        try:
            warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
            
            # 优先读取原始列名进行识别
            df_check_a = pd.read_excel(file_a, nrows=2, dtype=str)
            cols_a = [str(c).strip().lower() for c in df_check_a.columns]
            
            df_check_b = pd.read_excel(file_b, nrows=2, dtype=str)
            cols_b = [str(c).strip().lower() for c in df_check_b.columns]
            
            df_outbound_raw = None
            df_inventory_raw = None
            
            if 'shipping service' in cols_a or 'outbound/出库单号' in cols_a or any('outbound' in x for x in cols_a):
                df_outbound_raw = pd.read_excel(file_a, dtype=str)
            elif 'customize barcode' in cols_a or 'cellno' in cols_a or any('barcode' in x for x in cols_a):
                df_inventory_raw = pd.read_excel(file_a, dtype=str)
                
            if 'shipping service' in cols_b or 'outbound/出库单号' in cols_b or any('outbound' in x for x in cols_b):
                df_outbound_raw = pd.read_excel(file_b, dtype=str)
            elif 'customize barcode' in cols_b or 'cellno' in cols_b or any('barcode' in x for x in cols_b):
                df_inventory_raw = pd.read_excel(file_b, dtype=str)
                
            if df_outbound_raw is None or df_inventory_raw is None:
                st.error("❌ Error de identificación: Asegúrese de que los archivos contengan las columnas correctas. | 智能识别失败！请确认上传的文件中包含正确的特征列名。")
            else:
                # 📅 【全表智能动态日期检索】
                fecha_extract = None
                for col_name in df_outbound_raw.columns:
                    if any(k in str(col_name).lower() for k in ['time', 'date', '时间', '日期', 'creation']):
                        sample_data = df_outbound_raw[col_name].dropna().head(15)
                        for val in sample_data:
                            match_date = re.search(r'(\d{4}[-/]\d{2}[-/]\d{2})', str(val))
                            if match_date:
                                fecha_extract = match_date.group(1).replace('/', '-')
                                break
                    if fecha_extract: break
                
                if not fecha_extract:
                    for val in df_outbound_raw.values.flatten()[:200]:
                        match_date = re.search(r'(\d{4}[-/]\d{2}[-/]\d{2})', str(val))
                        if match_date:
                            fecha_extract = match_date.group(1).replace('/', '-')
                            break
                            
                if not fecha_extract:
                    fecha_extract = datetime.now().strftime("%Y-%m-%d")

                FINAL_OUTPUT_FILE = f'Picking#{fecha_extract}.xlsx'

                # 映射标准西语表头名称
                outbound_header_map = {
                    'A': 'Orden De Salida', 'E': 'Cliente', 'AN': 'Codigo Original',
                    'M': 'Etiqueta', 'K': 'Codigo de Barra', 'AQ': 'Cajas',
                    '位置': 'Ubicacion Piso', 'U': 'Destinatario', 'O': 'Horario De Entrega',
                    'AO': 'Cubicos', 'AP': 'Pesos/KG'
                }

                def rename_cols_to_letters(df):
                    new_cols = []
                    for i in range(len(df.columns)):
                        idx = i
                        letters = ""
                        while idx >= 0:
                            letters = chr(idx % 26 + ord('A')) + letters
                            idx = idx // 26 - 1
                        new_cols.append(letters)
                    df.columns = new_cols
                    return df

                df_outbound = rename_cols_to_letters(df_outbound_raw.copy())
                df_inventory = rename_cols_to_letters(df_inventory_raw.copy())

                # ==================== VLOOKUP ====================
                df_inv_clean = df_inventory[['B', 'G']].dropna(subset=['B']).drop_duplicates(subset=['B'])
                df_inv_clean.rename(columns={'G': '位置'}, inplace=True)
                
                df_step1 = pd.merge(df_outbound, df_inv_clean, left_on='AN', right_on='B', how='left')
                
                if 'B_y' in df_step1.columns: df_step1.drop(columns=['B_y'], inplace=True)
                if 'B' in df_step1.columns and 'B_x' in df_step1.columns:
                    df_step1.drop(columns=['B'], inplace=True)
                    df_step1.rename(columns={'B_x': 'B'}, inplace=True)

                # ==================== Filtro / 过滤条件 ====================
                m_series = df_step1['M'].astype(str).str.strip() 
                u_series = df_step1['U'].astype(str).str.strip().str.upper() 
                
                condition_m = m_series.str.contains('正常派送', na=False) | m_series.str.contains('换箱唛', na=False) | m_series.str.contains('换产品标', na=False)
                condition_u = u_series.str.contains('CPA', na=False) | u_series.str.contains('RC03', na=False) | u_series.str.contains('MXCD14', na=False)
                
                df_filtered = df_step1[condition_m & condition_u].copy()

                if df_filtered.empty:
                    df_filtered = df_step1.copy()

                # ==================== Algoritmo AN / 条码压缩算法 ====================
                def smart_compress_barcodes(series):
                    barcodes = sorted(list(set(series.dropna().astype(str).str.strip())))
                    if not barcodes: return ""
                    parsed_groups = {}
                    for code in barcodes:
                        if '/' in code:
                            parts = code.split('/')
                            prefix = parts[0] + '/'
                            suffix_str = parts[1]
                        else:
                            match = re.search(r'^(.*?)(\d+)$', code)
                            if match:
                                prefix = match.group(1)
                                suffix_str = match.group(2)
                            else:
                                prefix = code
                                suffix_str = ""
                        val = int(suffix_str) if suffix_str.isdigit() else -1
                        pad_len = len(suffix_str) if suffix_str.startswith('0') else 0
                        if prefix not in parsed_groups: parsed_groups[prefix] = []
                        parsed_groups[prefix].append((val, suffix_str, pad_len, code))

                    final_result_blocks = []
                    for prefix, items in parsed_groups.items():
                        items.sort(key=lambda x: x[0])
                        prefix_segments = []
                        i, n = 0, len(items)
                        while i < n:
                            if items[i][0] == -1:
                                prefix_segments.append(items[i][3]); i += 1; continue
                            start_idx = i
                            while i + 1 < n and items[i+1][0] == items[i][0] + 1: i += 1
                            end_idx = i
                            if end_idx - start_idx < 2:
                                for k in range(start_idx, end_idx + 1):
                                    prefix_segments.append(items[k][3] if not prefix_segments else items[k][1])
                            else:
                                prefix_segments.append(f"{items[start_idx][3]} a {items[end_idx][1]}" if not prefix_segments else f"{items[start_idx][1]} a {items[end_idx][1]}")
                            i += 1
                        
                        combined_prefix_str = ""
                        for idx, seg in enumerate(prefix_segments):
                            if idx == 0: combined_prefix_str += seg
                            else:
                                if '/' in prefix: combined_prefix_str += ", " + seg
                                else:
                                    last_letter = prefix[-1] if prefix else ""
                                    combined_prefix_str += f", {last_letter}{seg}" if (last_letter.isalpha() and not seg.startswith(last_letter)) else ", " + seg
                        final_result_blocks.append(combined_prefix_str)
                    return ", ".join(final_result_blocks)

                # ==================== Agrupación / 分组与聚合 ====================
                df_filtered['Box_8_Key'] = df_filtered['AN'].astype(str).str.strip().str[:8]
                df_filtered['Group_Key'] = df_filtered['A'].astype(str).str.strip() + "_" + df_filtered['Box_8_Key']

                # 🎯 【精准修正体积方数提取】完美适配最新的数据，防止因字符串拆分单位导致方数算偏
                def parse_cbm(val):
                    if pd.isna(val): return 0.0
                    val_str = str(val).lower()
                    nums = [float(n) for n in re.findall(r'\d+\.?\d*', val_str)]
                    if len(nums) >= 3:
                        # 自动适配长*宽*高*数量的连乘计算
                        prod = nums[0] * nums[1] * nums[2]
                        # 检查单位，如果是 cm 级别
