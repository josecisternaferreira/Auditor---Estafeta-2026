import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

# Configuración de página
st.set_page_config(
    page_title="Auditor Estafeta Abril 2026",
    page_icon="📊",
    layout="wide"
)

# ========== TARIFAS ABRIL 2026 (CON AUMENTO COMBUSTIBLE) ==========
def crear_tarifas_abril_2026():
    """Tarifas Estafeta Abril 2026 SIN descuento (aplicar 20% al calcular)"""
    tarifas = {}
    
    # SANTIAGO COMO ORIGEN
    tarifas['SCL'] = {
        'SCL': {'carta': 4376, '5kg': 6759, '10kg': 7519, '15kg': 8949, '20kg': 10518, '25kg': 12059, '30kg': 13628, '35kg': 15197, '40kg': 16764, '45kg': 18338, '50kg': 19905, 'hasta100kg': 335, 'hasta500kg': 309, 'mas500kg': 278},
        'ANF': {'carta': 10386, '5kg': 29267, '10kg': 32675, '15kg': 38441, '20kg': 42609, '25kg': 46694, '30kg': 50723, '35kg': 55052, '40kg': 59462, '45kg': 63007, '50kg': 67498, 'hasta100kg': 731, 'hasta500kg': 644, 'mas500kg': 615},
        'CPO': {'carta': 8356, '5kg': 17149, '10kg': 18719, '15kg': 22208, '20kg': 24780, '25kg': 27377, '30kg': 30027, '35kg': 32381, '40kg': 34974, '45kg': 37572, '50kg': 39983, 'hasta100kg': 615, 'hasta500kg': 542, 'mas500kg': 512},
        'IQQ': {'carta': 14170, '5kg': 33705, '10kg': 37603, '15kg': 44230, '20kg': 49046, '25kg': 53045, '30kg': 58271, '35kg': 63278, '40kg': 68095, '45kg': 72610, '50kg': 77698, 'hasta100kg': 1025, 'hasta500kg': 893, 'mas500kg': 805},
        'CJC': {'carta': 12038, '5kg': 32558, '10kg': 36920, '15kg': 42666, '20kg': 47611, '25kg': 52155, '30kg': 56880, '35kg': 61422, '40kg': 67956, '45kg': 70790, '50kg': 75733, 'hasta100kg': 834, 'hasta500kg': 791, 'mas500kg': 673},
        'COQ': {'carta': 6662, '5kg': 11661, '10kg': 13807, '15kg': 15176, '20kg': 17260, '25kg': 18922, '30kg': 20829, '35kg': 22197, '40kg': 23955, '45kg': 26006, '50kg': 27704, 'hasta100kg': 526, 'hasta500kg': 461, 'mas500kg': 432},
        'LSC': {'carta': 6662, '5kg': 11661, '10kg': 13807, '15kg': 15176, '20kg': 17260, '25kg': 18922, '30kg': 20829, '35kg': 22197, '40kg': 23955, '45kg': 26006, '50kg': 27704, 'hasta100kg': 526, 'hasta500kg': 461, 'mas500kg': 432},
        'CCP': {'carta': 6457, '5kg': 9635, '10kg': 10649, '15kg': 12556, '20kg': 14650, '25kg': 16706, '30kg': 18799, '35kg': 20891, '40kg': 22982, '45kg': 25081, '50kg': 27173, 'hasta100kg': 467, 'hasta500kg': 431, 'mas500kg': 391},
        'LSQ': {'carta': 6777, '5kg': 10117, '10kg': 11181, '15kg': 13185, '20kg': 15383, '25kg': 17543, '30kg': 19741, '35kg': 21939, '40kg': 24134, '45kg': 26338, '50kg': 28534, 'hasta100kg': 489, 'hasta500kg': 454, 'mas500kg': 410},
        'ZCO': {'carta': 7323, '5kg': 10927, '10kg': 12076, '15kg': 14239, '20kg': 16613, '25kg': 18944, '30kg': 21318, '35kg': 23691, '40kg': 26062, '45kg': 28442, '50kg': 30815, 'hasta100kg': 529, 'hasta500kg': 489, 'mas500kg': 442},
        'PMC': {'carta': 9567, '5kg': 14274, '10kg': 15775, '15kg': 18600, '20kg': 21703, '25kg': 24752, '30kg': 27856, '35kg': 30959, '40kg': 34060, '45kg': 37170, '50kg': 40272, 'hasta100kg': 691, 'hasta500kg': 640, 'mas500kg': 579},
        'ZAL': {'carta': 7908, '5kg': 11801, '10kg': 13042, '15kg': 15379, '20kg': 17941, '25kg': 20460, '30kg': 23023, '35kg': 25586, '40kg': 28148, '45kg': 30718, '50kg': 33280, 'hasta100kg': 571, 'hasta500kg': 529, 'mas500kg': 478},
        'VRR': {'carta': 7908, '5kg': 11801, '10kg': 13042, '15kg': 15379, '20kg': 17941, '25kg': 20460, '30kg': 23023, '35kg': 25586, '40kg': 28148, '45kg': 30718, '50kg': 33280, 'hasta100kg': 571, 'hasta500kg': 529, 'mas500kg': 478},
        'VAP': {'carta': 5324, '5kg': 7945, '10kg': 8781, '15kg': 10353, '20kg': 12079, '25kg': 13775, '30kg': 15501, '35kg': 17227, '40kg': 18952, '45kg': 20682, '50kg': 22406, 'hasta100kg': 385, 'hasta500kg': 356, 'mas500kg': 321},
        'KNA': {'carta': 5324, '5kg': 7945, '10kg': 8781, '15kg': 10353, '20kg': 12079, '25kg': 13775, '30kg': 15501, '35kg': 17227, '40kg': 18952, '45kg': 20682, '50kg': 22406, 'hasta100kg': 385, 'hasta500kg': 356, 'mas500kg': 321},
        'RCG': {'carta': 4695, '5kg': 7007, '10kg': 7743, '15kg': 9130, '20kg': 10652, '25kg': 12148, '30kg': 13668, '35kg': 15191, '40kg': 16712, '45kg': 18238, '50kg': 19758, 'hasta100kg': 339, 'hasta500kg': 314, 'mas500kg': 284},
    }
    
    # ANTOFAGASTA COMO ORIGEN
    tarifas['ANF'] = {
        'SCL': {'carta': 10386, '5kg': 29267, '10kg': 32675, '15kg': 38441, '20kg': 42609, '25kg': 46694, '30kg': 50723, '35kg': 55052, '40kg': 59462, '45kg': 63007, '50kg': 67498, 'hasta100kg': 731, 'hasta500kg': 644, 'mas500kg': 615},
        'ANF': {'carta': 4376, '5kg': 6759, '10kg': 7519, '15kg': 8949, '20kg': 10518, '25kg': 12059, '30kg': 13628, '35kg': 15197, '40kg': 16764, '45kg': 18338, '50kg': 19905, 'hasta100kg': 335, 'hasta500kg': 309, 'mas500kg': 278},
        'CPO': {'carta': 6346, '5kg': 9800, '10kg': 10902, '15kg': 12975, '20kg': 15250, '25kg': 17485, '30kg': 19760, '35kg': 22035, '40kg': 24308, '45kg': 26590, '50kg': 28862, 'hasta100kg': 486, 'hasta500kg': 448, 'mas500kg': 403},
        'IQQ': {'carta': 5164, '5kg': 7976, '10kg': 8872, '15kg': 10559, '20kg': 12411, '25kg': 14230, '30kg': 16081, '35kg': 17932, '40kg': 19781, '45kg': 21638, '50kg': 23489, 'hasta100kg': 395, 'hasta500kg': 364, 'mas500kg': 328},
        'CJC': {'carta': 4595, '5kg': 7097, '10kg': 7895, '15kg': 9396, '20kg': 11043, '25kg': 12662, '30kg': 14309, '35kg': 15957, '40kg': 17602, '45kg': 19255, '50kg': 20900, 'hasta100kg': 352, 'hasta500kg': 325, 'mas500kg': 292},
        'CCP': {'carta': 16843, '5kg': 38902, '10kg': 43324, '15kg': 50997, '20kg': 57259, '25kg': 63400, '30kg': 69522, '35kg': 75944, '40kg': 82445, '45kg': 88089, '50kg': 94670, 'hasta100kg': 1198, 'hasta500kg': 1075, 'mas500kg': 1006},
        'LSC': {'carta': 6652, '5kg': 10274, '10kg': 11428, '15kg': 13601, '20kg': 15987, '25kg': 18329, '30kg': 20714, '35kg': 23099, '40kg': 25481, '45kg': 27873, '50kg': 30256, 'hasta100kg': 510, 'hasta500kg': 469, 'mas500kg': 422},
    }

    # CALAMA COMO ORIGEN
    tarifas['CJC'] = {
        'SCL': {'carta': 12038, '5kg': 32558, '10kg': 36920, '15kg': 42666, '20kg': 47611, '25kg': 52155, '30kg': 56880, '35kg': 61422, '40kg': 67956, '45kg': 70790, '50kg': 75733, 'hasta100kg': 834, 'hasta500kg': 791, 'mas500kg': 673},
        'ANF': {'carta': 4595, '5kg': 7097, '10kg': 7895, '15kg': 9396, '20kg': 11043, '25kg': 12662, '30kg': 14309, '35kg': 15957, '40kg': 17602, '45kg': 19255, '50kg': 20900, 'hasta100kg': 352, 'hasta500kg': 325, 'mas500kg': 292},
        'CJC': {'carta': 4376, '5kg': 6759, '10kg': 7519, '15kg': 8949, '20kg': 10518, '25kg': 12059, '30kg': 13628, '35kg': 15197, '40kg': 16764, '45kg': 18338, '50kg': 19905, 'hasta100kg': 335, 'hasta500kg': 309, 'mas500kg': 278},
        'LSC': {'carta': 5339, '5kg': 8245, '10kg': 9173, '15kg': 10918, '20kg': 12832, '25kg': 14711, '30kg': 16625, '35kg': 18540, '40kg': 20452, '45kg': 22372, '50kg': 24284, 'hasta100kg': 409, 'hasta500kg': 376, 'mas500kg': 339},
    }
    
    # COPIAPO COMO ORIGEN
    tarifas['CPO'] = {
        'SCL': {'carta': 8356, '5kg': 17149, '10kg': 18719, '15kg': 22208, '20kg': 24780, '25kg': 27377, '30kg': 30027, '35kg': 32381, '40kg': 34974, '45kg': 37572, '50kg': 39983, 'hasta100kg': 615, 'hasta500kg': 542, 'mas500kg': 512},
        'ANF': {'carta': 6346, '5kg': 9800, '10kg': 10902, '15kg': 12975, '20kg': 15250, '25kg': 17485, '30kg': 19760, '35kg': 22035, '40kg': 24308, '45kg': 26590, '50kg': 28862, 'hasta100kg': 486, 'hasta500kg': 448, 'mas500kg': 403},
        'CPO': {'carta': 4376, '5kg': 6759, '10kg': 7519, '15kg': 8949, '20kg': 10518, '25kg': 12059, '30kg': 13628, '35kg': 15197, '40kg': 16764, '45kg': 18338, '50kg': 19905, 'hasta100kg': 335, 'hasta500kg': 309, 'mas500kg': 278},
    }
    
    # CONCEPCION COMO ORIGEN
    tarifas['CCP'] = {
        'SCL': {'carta': 6457, '5kg': 9635, '10kg': 10649, '15kg': 12556, '20kg': 14650, '25kg': 16706, '30kg': 18799, '35kg': 20891, '40kg': 22982, '45kg': 25081, '50kg': 27173, 'hasta100kg': 467, 'hasta500kg': 431, 'mas500kg': 391},
        'ANF': {'carta': 16843, '5kg': 38902, '10kg': 43324, '15kg': 50997, '20kg': 57259, '25kg': 63400, '30kg': 69522, '35kg': 75944, '40kg': 82445, '45kg': 88089, '50kg': 94670, 'hasta100kg': 1198, 'hasta500kg': 1075, 'mas500kg': 1006},
        'CCP': {'carta': 4376, '5kg': 6759, '10kg': 7519, '15kg': 8949, '20kg': 10518, '25kg': 12059, '30kg': 13628, '35kg': 15197, '40kg': 16764, '45kg': 18338, '50kg': 19905, 'hasta100kg': 335, 'hasta500kg': 309, 'mas500kg': 278},
    }
    
    # LA SERENA COMO ORIGEN
    tarifas['LSC'] = {
        'SCL': {'carta': 6662, '5kg': 11661, '10kg': 13807, '15kg': 15176, '20kg': 17260, '25kg': 18922, '30kg': 20829, '35kg': 22197, '40kg': 23955, '45kg': 26006, '50kg': 27704, 'hasta100kg': 526, 'hasta500kg': 461, 'mas500kg': 432},
        'ANF': {'carta': 6652, '5kg': 10274, '10kg': 11428, '15kg': 13601, '20kg': 15987, '25kg': 18329, '30kg': 20714, '35kg': 23099, '40kg': 25481, '45kg': 27873, '50kg': 30256, 'hasta100kg': 510, 'hasta500kg': 469, 'mas500kg': 422},
        'CJC': {'carta': 5339, '5kg': 8245, '10kg': 9173, '15kg': 10918, '20kg': 12832, '25kg': 14711, '30kg': 16625, '35kg': 18540, '40kg': 20452, '45kg': 22372, '50kg': 24284, 'hasta100kg': 409, 'hasta500kg': 376, 'mas500kg': 339},
    }
    
    # IQUIQUE COMO ORIGEN
    tarifas['IQQ'] = {
        'SCL': {'carta': 14170, '5kg': 33705, '10kg': 37603, '15kg': 44230, '20kg': 49046, '25kg': 53045, '30kg': 58271, '35kg': 63278, '40kg': 68095, '45kg': 72610, '50kg': 77698, 'hasta100kg': 1025, 'hasta500kg': 893, 'mas500kg': 805},
        'ANF': {'carta': 5164, '5kg': 7976, '10kg': 8872, '15kg': 10559, '20kg': 12411, '25kg': 14230, '30kg': 16081, '35kg': 17932, '40kg': 19781, '45kg': 21638, '50kg': 23489, 'hasta100kg': 395, 'hasta500kg': 364, 'mas500kg': 328},
    }
    
    return tarifas

# CÓDIGOS DE CIUDAD
CODIGOS_CIUDAD = {
    'SANTIAGO': 'SCL', 'IQUIQUE': 'IQQ', 'CALAMA': 'CJC',
    'ANTOFAGASTA': 'ANF', 'COPIAPO': 'CPO', 'COQUIMBO': 'COQ',
    'LA SERENA': 'LSC', 'SERENA': 'LSC', 'VIÑA DEL MAR': 'KNA',
    'RANCAGUA': 'RCG', 'TALCA': 'ZCA', 'CHILLAN': 'YAI',
    'CONCEPCION': 'CCP', 'LOS ANGELES': 'LSQ', 'LOS-ANGELES': 'LSQ',
    'TEMUCO': 'ZCO', 'PUERTO MONTT': 'PMC', 'PUERTO-MONTT': 'PMC',
    'VALPARAISO': 'VAP', 'OSORNO': 'ZOS', 'VALDIVIA': 'ZAL',
    'VILLARICA': 'VRR', 'HUASCO': 'HCO',
    'CONSTITUCION': 'CTT'
}

# RAMALES (cargo adicional)
RAMALES = {
    'MEJILLONES': {'origen': 'ANTOFAGASTA', 'cargo': 30550},
}

def calcular_tarifa_base(peso_facturable, tarifas_ruta):
    """Calcula tarifa base según peso"""
    if peso_facturable <= 1:
        return tarifas_ruta.get('carta', 0), 'Carta (≤1kg)'
    elif peso_facturable <= 5:
        return tarifas_ruta.get('5kg', 0), 'Hasta 5kg'
    elif peso_facturable <= 10:
        return tarifas_ruta.get('10kg', 0), 'Hasta 10kg'
    elif peso_facturable <= 15:
        return tarifas_ruta.get('15kg', 0), 'Hasta 15kg'
    elif peso_facturable <= 20:
        return tarifas_ruta.get('20kg', 0), 'Hasta 20kg'
    elif peso_facturable <= 25:
        return tarifas_ruta.get('25kg', 0), 'Hasta 25kg'
    elif peso_facturable <= 30:
        return tarifas_ruta.get('30kg', 0), 'Hasta 30kg'
    elif peso_facturable <= 35:
        return tarifas_ruta.get('35kg', 0), 'Hasta 35kg'
    elif peso_facturable <= 40:
        return tarifas_ruta.get('40kg', 0), 'Hasta 40kg'
    elif peso_facturable <= 45:
        return tarifas_ruta.get('45kg', 0), 'Hasta 45kg'
    elif peso_facturable <= 50:
        return tarifas_ruta.get('50kg', 0), 'Hasta 50kg'
    else:
        base_50kg = tarifas_ruta.get('50kg', 0)
        kg_adicionales = peso_facturable - 50
        
        # LÓGICA CORRECTA: comparar kg_adicionales (EXCESO) con 100, 500
        if kg_adicionales <= 100:
            tarifa_kg = tarifas_ruta.get('hasta100kg', 0)
        elif kg_adicionales <= 500:
            tarifa_kg = tarifas_ruta.get('hasta500kg', 0)
        else:
            tarifa_kg = tarifas_ruta.get('mas500kg', 0)
        
        total = base_50kg + (kg_adicionales * tarifa_kg)
        detalle = f'Base 50kg + {kg_adicionales:.1f}kg × ${tarifa_kg}'
        return total, detalle

def auditar_excel(df):
    """Audita el archivo de Estafeta"""
    tarifas = crear_tarifas_abril_2026()
    resultados = []
    
    for idx, row in df.iterrows():
        try:
            orden = int(row['OrdFle'])
            origen_raw = str(row['Origen']).upper().strip()
            destino_raw = str(row['Destino']).upper().strip()
            
            peso_real = float(row['Kg.Bul']) if pd.notna(row['Kg.Bul']) else 0
            peso_vol = float(row['Kg.Vol.']) if pd.notna(row['Kg.Vol.']) else 0
            peso_fact = max(peso_real, peso_vol)
            cobrado_neto = float(row['Neto Fac']) if pd.notna(row['Neto Fac']) else 0
            
            # Verificar si es ramal
            cargo_ramal = 0
            es_ramal = False
            if destino_raw in RAMALES:
                ramal_info = RAMALES[destino_raw]
                if origen_raw == ramal_info['origen']:
                    cargo_ramal = ramal_info['cargo']
                    destino_para_tarifa = origen_raw
                    es_ramal = True
                else:
                    destino_para_tarifa = destino_raw
            else:
                destino_para_tarifa = destino_raw
            
            # Mapear códigos
            origen = CODIGOS_CIUDAD.get(origen_raw, origen_raw)
            destino = CODIGOS_CIUDAD.get(destino_para_tarifa, destino_para_tarifa)
            
            estado = 'SIN_TARIFA'
            diferencia = None
            correcto_desc = None
            tramo_desc = ''
            
            if origen in tarifas and destino in tarifas[origen]:
                tarifas_ruta = tarifas[origen][destino]
                tarifa_base, tramo_desc = calcular_tarifa_base(peso_fact, tarifas_ruta)
                
                tarifa_total_sin_desc = tarifa_base + cargo_ramal
                correcto_desc = round(tarifa_total_sin_desc * 0.80, 2)
                
                diferencia = cobrado_neto - correcto_desc
                
                if abs(diferencia) < 10:
                    estado = 'CORRECTO'
                elif diferencia > 0:
                    estado = 'SOBRECOBRO'
                else:
                    estado = 'COBRO_MENOS'
            
            if es_ramal:
                tramo_desc = f"{tramo_desc} + Ramal {destino_raw}"
            
            resultados.append({
                'Orden': orden,
                'Fecha': row['Fecha'] if 'Fecha' in row else '',
                'Origen': origen_raw,
                'Destino': destino_raw,
                'Peso_Real_kg': peso_real,
                'Peso_Vol_kg': peso_vol,
                'Peso_Fact_kg': peso_fact,
                'Tramo': tramo_desc,
                'Tarifa_Correcta': correcto_desc,
                'Cobrado_Neto': cobrado_neto,
                'Diferencia': diferencia,
                'Estado': estado,
            })
        except Exception as e:
            st.warning(f"Error procesando fila {idx}: {str(e)}")
            continue
    
    return pd.DataFrame(resultados)

def generar_excel_auditoria(df_audit):
    """Genera Excel con formato"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_audit.to_excel(writer, index=False, sheet_name='Auditoría')
        
        workbook = writer.book
        worksheet = writer.sheets['Auditoría']
        
        # Colores
        color_header = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
        color_ok = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
        color_error = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
        
        font_header = Font(bold=True, color='FFFFFF')
        
        # Formatear encabezados
        for cell in worksheet[1]:
            cell.fill = color_header
            cell.font = font_header
            cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # Formatear filas según estado
        for idx, row in enumerate(worksheet.iter_rows(min_row=2, max_row=worksheet.max_row), start=0):
            if idx < len(df_audit):
                estado = df_audit.iloc[idx]['Estado']
                estado_col = 12
                
                if estado == 'CORRECTO':
                    worksheet.cell(row=idx+2, column=estado_col).fill = color_ok
                elif estado == 'SOBRECOBRO':
                    worksheet.cell(row=idx+2, column=estado_col).fill = color_error
        
        # Ajustar anchos
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    output.seek(0)
    return output

# ========== INTERFAZ STREAMLIT ==========
st.title("📊 Auditor Estafeta Abril 2026")
st.info("🔥 **Tarifas actualizadas por aumento de combustible - Vigentes desde Abril 2026**")
st.markdown("---")

# Información
with st.expander("ℹ️ Instrucciones de uso"):
    st.markdown("""
    ### Cómo usar el auditor:
    1. **Sube tu Estado de Pago** de Estafeta (archivo Excel)
    2. El sistema detectará automáticamente las columnas necesarias
    3. Haz clic en **AUDITAR**
    4. Descarga los resultados en Excel
    
    ### Qué hace el auditor:
    - ✅ Verifica cada envío contra la tarifa oficial Abril 2026 (con aumento combustible)
    - ✅ Aplica automáticamente el 20% de descuento corporativo
    - ✅ Detecta ramales (ej: Mejillones) y aplica cargos adicionales
    - ✅ Identifica sobrecobros y subcobros
    - ✅ Genera reporte descargable con % de desviación
    """)

# Upload
uploaded_file = st.file_uploader(
    "📁 Sube tu Estado de Pago (Excel)",
    type=['xlsx', 'xls'],
    help="Archivo Excel con las columnas: OrdFle, Origen, Destino, Kg.Bul, Kg.Vol., Neto Fac"
)

if uploaded_file:
    try:
        # Leer Excel
        df = pd.read_excel(uploaded_file, header=4)
        df = df[pd.notna(df['OrdFle'])].copy()
        df = df[df['Origen'].notna()].copy()
        
        st.success(f"✅ Archivo cargado: {len(df)} órdenes encontradas")
        
        # Vista previa
        with st.expander("👀 Vista previa de datos"):
            st.dataframe(df.head(10))
        
        # Botón auditar
        if st.button("🚀 AUDITAR", type="primary", use_container_width=True):
            with st.spinner("Auditando..."):
                df_audit = auditar_excel(df)
            
            st.success("✅ Auditoría completada")
            
            # Métricas
            col1, col2, col3, col4, col5 = st.columns(5)

            total = len(df_audit)
            correctos = len(df_audit[df_audit['Estado'] == 'CORRECTO'])
            sobrecobros = len(df_audit[df_audit['Estado'] == 'SOBRECOBRO'])
            cobros_menos = len(df_audit[df_audit['Estado'] == 'COBRO_MENOS'])
            sin_tarifa = len(df_audit[df_audit['Estado'] == 'SIN_TARIFA'])

            col1.metric("Total Órdenes", total)
            col2.metric("✅ Correctos", correctos)
            col3.metric("🔴 Sobrecobros", sobrecobros)
            col4.metric("🟢 Cobros Menos", cobros_menos)
            col5.metric("⚠️ Sin Tarifa", sin_tarifa)
            
            # Resumen financiero
            st.markdown("### 💰 Resumen Financiero")
            col1, col2, col3 = st.columns(3)
            
            total_cobrado = df_audit['Cobrado_Neto'].sum()
            total_correcto = df_audit[df_audit['Tarifa_Correcta'].notna()]['Tarifa_Correcta'].sum()
            diferencia_neta = df_audit[df_audit['Diferencia'].notna()]['Diferencia'].sum()
            
            # Calcular porcentajes de desviación
            if total_correcto > 0:
                porcentaje_diferencia_neta = (diferencia_neta / total_correcto) * 100
            else:
                porcentaje_diferencia_neta = 0
            
            col1.metric("Total Cobrado", f"${total_cobrado:,.0f}")
            col2.metric("Total Correcto", f"${total_correcto:,.0f}")
            
            if diferencia_neta > 0:
                col3.metric(
                    "Diferencia Neta", 
                    f"${diferencia_neta:,.0f}", 
                    delta=f"{porcentaje_diferencia_neta:+.2f}% Sobrecobro", 
                    delta_color="inverse"
                )
            elif diferencia_neta < 0:
                col3.metric(
                    "Diferencia Neta", 
                    f"${abs(diferencia_neta):,.0f}", 
                    delta=f"{abs(porcentaje_diferencia_neta):.2f}% A favor", 
                    delta_color="normal"
                )
            else:
                col3.metric("Diferencia Neta", "$0", delta="0.00% Perfecto")
            
            # Detalles adicionales de desviación
            if sobrecobros > 0 or cobros_menos > 0:
                st.markdown("### 📊 Detalle de Desviaciones")
                col1, col2, col3 = st.columns(3)
                
                # Sobrecobros
                sobrecobros_df = df_audit[df_audit['Estado'] == 'SOBRECOBRO']
                if len(sobrecobros_df) > 0:
                    total_sobrecobros = sobrecobros_df['Diferencia'].sum()
                    porcentaje_sobrecobros = (total_sobrecobros / total_correcto) * 100 if total_correcto > 0 else 0
                    col1.metric(
                        "🔴 Total Sobrecobros", 
                        f"${total_sobrecobros:,.0f}",
                        delta=f"+{porcentaje_sobrecobros:.2f}%"
                    )
                else:
                    col1.metric("🔴 Total Sobrecobros", "$0", delta="0.00%")
                
                # Cobros menos
                cobros_menos_df = df_audit[df_audit['Estado'] == 'COBRO_MENOS']
                if len(cobros_menos_df) > 0:
                    total_cobros_menos = abs(cobros_menos_df['Diferencia'].sum())
                    porcentaje_cobros_menos = (total_cobros_menos / total_correcto) * 100 if total_correcto > 0 else 0
                    col2.metric(
                        "🟢 Total Cobros Menos", 
                        f"${total_cobros_menos:,.0f}",
                        delta=f"-{porcentaje_cobros_menos:.2f}%"
                    )
                else:
                    col2.metric("🟢 Total Cobros Menos", "$0", delta="0.00%")
                
                # Neto
                col3.metric(
                    "⚖️ Balance Neto",
                    f"${abs(diferencia_neta):,.0f}",
                    delta=f"{porcentaje_diferencia_neta:+.2f}%"
                )
            
            # Detalle de órdenes sin tarifa (ruta no encontrada en la tabla)
            if sin_tarifa > 0:
                st.markdown("### ⚠️ Órdenes Sin Tarifa (no encontradas en la tabla)")
                st.warning(f"Se encontraron **{sin_tarifa}** órdenes cuya ruta Origen-Destino no está en la tabla de tarifas. Deben revisarse manualmente.")
                sin_tarifa_df = df_audit[df_audit['Estado'] == 'SIN_TARIFA'][['Orden', 'Origen', 'Destino', 'Peso_Fact_kg', 'Cobrado_Neto']]
                st.dataframe(sin_tarifa_df, use_container_width=True)

            # Detalles de sobrecobros
            if sobrecobros > 0:
                st.markdown("### 🔴 Sobrecobros Detectados")
                sobrecobros_df = df_audit[df_audit['Estado'] == 'SOBRECOBRO'][['Orden', 'Origen', 'Destino', 'Peso_Fact_kg', 'Cobrado_Neto', 'Tarifa_Correcta', 'Diferencia']]
                st.dataframe(sobrecobros_df, use_container_width=True)
                
                total_reclamable = sobrecobros_df['Diferencia'].sum()
                st.info(f"💵 **Total reclamable: ${total_reclamable:,.0f}**")
            
            # Tabla completa
            with st.expander("📋 Ver auditoría completa"):
                st.dataframe(df_audit, use_container_width=True)
            
            # Descarga
            excel_output = generar_excel_auditoria(df_audit)
            
            st.download_button(
                label="📥 Descargar Auditoría (Excel)",
                data=excel_output,
                file_name=f"AUDITORIA_ESTAFETA_ABRIL_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    
    except Exception as e:
        st.error(f"❌ Error al procesar archivo: {str(e)}")
        st.info("Asegúrate de que el archivo tenga las columnas: OrdFle, Origen, Destino, Kg.Bul, Kg.Vol., Neto Fac")

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666;'>
    <small>Auditor Estafeta Abril 2026 | Tarifas actualizadas con aumento combustible | 20% descuento corporativo</small>
</div>
""", unsafe_allow_html=True)
