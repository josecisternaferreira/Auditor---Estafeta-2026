import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

# Configuración de página
st.set_page_config(
    page_title="Auditor Estafeta 2026",
    page_icon="📊",
    layout="wide"
)

# ========== TARIFAS 2026 COMPLETAS ==========
def crear_tarifas_2026():
    """Tarifas Estafeta 2026 SIN descuento (aplicar 20% al calcular)"""
    tarifas = {}
    
    # SANTIAGO COMO ORIGEN
    tarifas['SCL'] = {
        'SCL': {'carta': 3907, '5kg': 6035, '10kg': 6713, '15kg': 7990, '20kg': 9391, '25kg': 10767, '30kg': 12168, '50kg': 17772, 'hasta100kg': 299, 'hasta500kg': 276, 'mas500kg': 248},
        'ANF': {'carta': 9273, '5kg': 26131, '10kg': 29174, '15kg': 34322, '20kg': 38044, '25kg': 41691, '30kg': 45288, '35kg': 49154, '40kg': 53091, '45kg': 56256, '50kg': 60266, 'hasta100kg': 653, 'hasta500kg': 575, 'mas500kg': 549},
        'CPO': {'carta': 7461, '5kg': 15312, '10kg': 16713, '15kg': 19829, '20kg': 22125, '25kg': 24444, '30kg': 26810, '50kg': 35699, 'hasta100kg': 549, 'hasta500kg': 484, 'mas500kg': 457},
        'IQQ': {'carta': 12652, '5kg': 30094, '10kg': 33574, '15kg': 39491, '20kg': 43791, '25kg': 47362, '30kg': 52028, '50kg': 69373, 'hasta100kg': 915, 'hasta500kg': 797, 'mas500kg': 719},
        'CJC': {'carta': 10748, '5kg': 29070, '10kg': 32964, '15kg': 38095, '20kg': 42510, '25kg': 46567, '30kg': 50786, '50kg': 67619, 'hasta100kg': 745, 'hasta500kg': 706, 'mas500kg': 601},
        'COQ': {'carta': 5948, '5kg': 10412, '10kg': 12328, '15kg': 13550, '20kg': 15411, '25kg': 16895, '30kg': 18597, '50kg': 24736, 'hasta100kg': 470, 'hasta500kg': 412, 'mas500kg': 386},
        'LSC': {'carta': 5948, '5kg': 10412, '10kg': 12328, '15kg': 13550, '20kg': 15411, '25kg': 16895, '30kg': 18597, '50kg': 24736, 'hasta100kg': 470, 'hasta500kg': 412, 'mas500kg': 386},
        'CCP': {'carta': 5765, '5kg': 8603, '10kg': 9508, '15kg': 11211, '20kg': 13080, '25kg': 14916, '30kg': 16785, '50kg': 24262, 'hasta100kg': 417, 'hasta500kg': 385, 'mas500kg': 349},
        'LSQ': {'carta': 6053, '5kg': 9033, '10kg': 9983, '15kg': 11772, '20kg': 13734, '25kg': 15661, '30kg': 17624, '50kg': 25475, 'hasta100kg': 437, 'hasta500kg': 405, 'mas500kg': 366},
        'ZCO': {'carta': 6538, '5kg': 9756, '10kg': 10782, '15kg': 12713, '20kg': 14833, '25kg': 16914, '30kg': 19034, '50kg': 27513, 'hasta100kg': 472, 'hasta500kg': 437, 'mas500kg': 395},
        'PMC': {'carta': 8541, '5kg': 12745, '10kg': 14085, '15kg': 16608, '20kg': 19377, '25kg': 22096, '30kg': 24865, '50kg': 35942, 'hasta100kg': 617, 'hasta500kg': 571, 'mas500kg': 517},
        'ZAL': {'carta': 7061, '5kg': 10537, '10kg': 11645, '15kg': 13731, '20kg': 16019, '25kg': 18268, '30kg': 20556, '50kg': 29714, 'hasta100kg': 510, 'hasta500kg': 472, 'mas500kg': 427},
        'VRR': {'carta': 7061, '5kg': 10537, '10kg': 11645, '15kg': 13731, '20kg': 16019, '25kg': 18268, '30kg': 20556, '50kg': 29714, 'hasta100kg': 510, 'hasta500kg': 472, 'mas500kg': 427},
        'VAP': {'carta': 4754, '5kg': 7094, '10kg': 7840, '15kg': 9244, '20kg': 10785, '25kg': 12299, '30kg': 13840, '50kg': 20005, 'hasta100kg': 344, 'hasta500kg': 318, 'mas500kg': 287},
        'KNA': {'carta': 4754, '5kg': 7094, '10kg': 7840, '15kg': 9244, '20kg': 10785, '25kg': 12299, '30kg': 13840, '50kg': 20005, 'hasta100kg': 344, 'hasta500kg': 318, 'mas500kg': 287},
        'RCG': {'carta': 4192, '5kg': 6256, '10kg': 6913, '15kg': 8152, '20kg': 9511, '25kg': 10846, '30kg': 12204, '50kg': 17641, 'hasta100kg': 303, 'hasta500kg': 280, 'mas500kg': 254},
    }
    
    # ANTOFAGASTA COMO ORIGEN
    tarifas['ANF'] = {
        'SCL': {'carta': 9273, '5kg': 26131, '10kg': 29174, '15kg': 34322, '20kg': 38044, '25kg': 41691, '30kg': 45288, '50kg': 60266, 'hasta100kg': 653, 'hasta500kg': 575, 'mas500kg': 549},
        'CPO': {'carta': 5666, '5kg': 8750, '10kg': 9734, '15kg': 11585, '20kg': 13616, '25kg': 15612, '30kg': 17643, '50kg': 25770, 'hasta100kg': 434, 'hasta500kg': 400, 'mas500kg': 360},
        'IQQ': {'carta': 4611, '5kg': 7121, '10kg': 7921, '15kg': 9428, '20kg': 11081, '25kg': 12705, '30kg': 14358, '50kg': 20972, 'hasta100kg': 353, 'hasta500kg': 325, 'mas500kg': 293},
        'CJC': {'carta': 4103, '5kg': 6337, '10kg': 7049, '15kg': 8389, '20kg': 9860, '25kg': 11305, '30kg': 12776, '50kg': 18661, 'hasta100kg': 314, 'hasta500kg': 290, 'mas500kg': 261},
        'ANF': {'carta': 3907, '5kg': 6035, '10kg': 6713, '15kg': 7990, '20kg': 9391, '25kg': 10767, '30kg': 12168, '50kg': 17772, 'hasta100kg': 299, 'hasta500kg': 276, 'mas500kg': 248},
        'CCP': {'carta': 15038, '5kg': 34734, '10kg': 38682, '15kg': 45533, '20kg': 51124, '25kg': 56607, '30kg': 62073, '50kg': 84527, 'hasta100kg': 1070, 'hasta500kg': 960, 'mas500kg': 898},
    }
    
    # COPIAPO COMO ORIGEN
    tarifas['CPO'] = {
        'SCL': {'carta': 7461, '5kg': 15312, '10kg': 16713, '15kg': 19829, '20kg': 22125, '25kg': 24444, '30kg': 26810, '50kg': 35699, 'hasta100kg': 549, 'hasta500kg': 484, 'mas500kg': 457},
        'ANF': {'carta': 5666, '5kg': 8750, '10kg': 9734, '15kg': 11585, '20kg': 13616, '25kg': 15612, '30kg': 17643, '50kg': 25770, 'hasta100kg': 434, 'hasta500kg': 400, 'mas500kg': 360},
    }
    
    # CONCEPCION COMO ORIGEN
    tarifas['CCP'] = {
        'SCL': {'carta': 5765, '5kg': 8603, '10kg': 9508, '15kg': 11211, '20kg': 13080, '25kg': 14916, '30kg': 16785, '50kg': 24262, 'hasta100kg': 417, 'hasta500kg': 385, 'mas500kg': 349},
        'ANF': {'carta': 15038, '5kg': 34734, '10kg': 38682, '15kg': 45533, '20kg': 51124, '25kg': 56607, '30kg': 62073, '50kg': 84527, 'hasta100kg': 1070, 'hasta500kg': 960, 'mas500kg': 898},
    }
    
    # LA SERENA COMO ORIGEN
    tarifas['LSC'] = {
        'SCL': {'carta': 5948, '5kg': 10412, '10kg': 12328, '15kg': 13550, '20kg': 15411, '25kg': 16895, '30kg': 18597, '50kg': 24736, 'hasta100kg': 470, 'hasta500kg': 412, 'mas500kg': 386},
        'ANF': {'carta': 5939, '5kg': 9173, '10kg': 10204, '15kg': 12144, '20kg': 14274, '25kg': 16365, '30kg': 18495, '50kg': 27014, 'hasta100kg': 455, 'hasta500kg': 419, 'mas500kg': 377},
    }
    
    # IQUIQUE COMO ORIGEN
    tarifas['IQQ'] = {
        'SCL': {'carta': 12652, '5kg': 30094, '10kg': 33574, '15kg': 39491, '20kg': 43791, '25kg': 47362, '30kg': 52028, '50kg': 69373, 'hasta100kg': 915, 'hasta500kg': 797, 'mas500kg': 719},
        'ANF': {'carta': 4611, '5kg': 7121, '10kg': 7921, '15kg': 9428, '20kg': 11081, '25kg': 12705, '30kg': 14358, '50kg': 20972, 'hasta100kg': 353, 'hasta500kg': 325, 'mas500kg': 293},
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
        # - exceso ≤ 100kg → usa hasta100kg
        # - exceso ≤ 500kg → usa hasta500kg
        # - exceso > 500kg → usa mas500kg
        
        if kg_adicionales <= 100:  # exceso 1-100kg
            tarifa_kg = tarifas_ruta.get('hasta100kg', 0)
        elif kg_adicionales <= 500:  # exceso 101-500kg
            tarifa_kg = tarifas_ruta.get('hasta500kg', 0)
        else:  # exceso >500kg
            tarifa_kg = tarifas_ruta.get('mas500kg', 0)
        
        total = base_50kg + (kg_adicionales * tarifa_kg)
        detalle = f'Base 50kg + {kg_adicionales:.1f}kg × ${tarifa_kg}'
        return total, detalle

def auditar_excel(df):
    """Audita el archivo de Estafeta"""
    tarifas = crear_tarifas_2026()
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
                    destino_para_tarifa = origen_raw  # ANF → ANF
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
                estado_col = 12  # Columna Estado
                
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
st.title("📊 Auditor Estafeta 2026")
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
    - ✅ Verifica cada envío contra la tarifa oficial 2026
    - ✅ Aplica automáticamente el 20% de descuento corporativo
    - ✅ Detecta ramales (ej: Mejillones) y aplica cargos adicionales
    - ✅ Identifica sobrecobros y subcobros
    - ✅ Genera reporte descargable
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
            col1, col2, col3, col4 = st.columns(4)
            
            total = len(df_audit)
            correctos = len(df_audit[df_audit['Estado'] == 'CORRECTO'])
            sobrecobros = len(df_audit[df_audit['Estado'] == 'SOBRECOBRO'])
            cobros_menos = len(df_audit[df_audit['Estado'] == 'COBRO_MENOS'])
            
            col1.metric("Total Órdenes", total)
            col2.metric("✅ Correctos", correctos)
            col3.metric("🔴 Sobrecobros", sobrecobros)
            col4.metric("🟢 Cobros Menos", cobros_menos)
            
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
                file_name=f"AUDITORIA_ESTAFETA_{datetime.now().strftime('%Y%m%d')}.xlsx",
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
    <small>Auditor Estafeta 2026 | Tarifa Básica con 20% descuento corporativo | Incluye ramales</small>
</div>
""", unsafe_allow_html=True)
