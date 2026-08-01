import streamlit as st
import pandas as pd

# ---------------------------------------------------------
# Configuración de la página
# ---------------------------------------------------------
st.set_page_config(
    page_title="Posición Mensual de IVA - ARCA / AFIP",
    page_icon="📊",
    layout="wide"
)

# Estilos CSS personalizados para la interfaz
st.markdown("""
    <style>
    .stMetric {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 8px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    .status-card {
        padding: 20px;
        border-radius: 8px;
        font-size: 1.2rem;
        font-weight: bold;
        text-align: center;
        margin-bottom: 20px;
    }
    .status-favor-contribuyente {
        background-color: #d4edda;
        color: #155724;
        border: 1px solid #c3e6cb;
    }
    .status-favor-arca {
        background-color: #f8d7da;
        color: #721c24;
        border: 1px solid #f5c6cb;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📊 Control y Posición Mensual de IVA")
st.subheader("Liquidación rápida basada en 'Mis Comprobantes' (Emitidos y Recibidos)")
st.markdown("---")

# ---------------------------------------------------------
# Barra Lateral (Inputs)
# ---------------------------------------------------------
st.sidebar.header("📁 1. Carga de Archivos")
file_emitidos = st.sidebar.file_uploader(
    "Cargar 'Mis Comprobantes Emitidos' (.xlsx)", 
    type=["xlsx", "xls"]
)
file_recibidos = st.sidebar.file_uploader(
    "Cargar 'Mis Comprobantes Recibidos' (.xlsx)", 
    type=["xlsx", "xls"]
)

st.sidebar.markdown("---")
st.sidebar.header("💵 2. Saldos del Período Anterior")
saldo_tecnico_anterior = st.sidebar.number_input(
    "Saldo Técnico a Favor del período anterior ($)",
    min_value=0.0,
    value=0.0,
    step=1000.0,
    format="%.2f"
)

saldo_libre_disp_anterior = st.sidebar.number_input(
    "Saldo de Libre Disponibilidad anterior ($)",
    min_value=0.0,
    value=0.0,
    step=1000.0,
    format="%.2f"
)

# ---------------------------------------------------------
# Función de procesamiento de comprobantes
# ---------------------------------------------------------
def procesar_comprobantes(file, es_emitidos=True):
    """
    Procesa el archivo Excel de AFIP adaptando la fila de encabezado,
    multiplicando por el tipo de cambio (si es moneda extranjera) y 
    asignando signos (+ / -) según la naturaleza del comprobante.
    """
    try:
        # Los encabezados de AFIP se encuentran en la fila 1 (índice 1 de pandas)
        df = pd.read_excel(file, header=1)
        
        # Eliminar filas vacías o sin tipo de comprobante
        df = df.dropna(subset=['Tipo']).copy()
        
        # Asegurar tipos numéricos y limpiar valores nulos
        cols_numericas = ['Tipo Cambio', 'Neto Gravado Total', 'Total IVA', 'Imp. Total']
        for col in cols_numericas:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            else:
                df[col] = 0.0

        # Normalizar Tipo de Cambio
        df['Tipo Cambio'] = df['Tipo Cambio'].apply(lambda x: 1.0 if x <= 0 else x)
        
        # Recalcular valores en Pesos si hay Moneda Extranjera
        df['Neto_ARS'] = df['Neto Gravado Total'] * df['Tipo Cambio']
        df['IVA_ARS'] = df['Total IVA'] * df['Tipo Cambio']
        df['Total_ARS'] = df['Imp. Total'] * df['Tipo Cambio']
        
        # Asignar signo según tipo de comprobante
        def determinar_signo(tipo_comp):
            tipo = str(tipo_comp).upper()
            if "CRÉDITO" in tipo or "CREDITO" in tipo:
                return -1.0
            return 1.0

        df['Signo'] = df['Tipo'].apply(determinar_signo)
        
        # Aplicar signo a los montos
        df['Neto_Final'] = df['Neto_ARS'] * df['Signo']
        df['IVA_Final'] = df['IVA_ARS'] * df['Signo']
        df['Total_Final'] = df['Total_ARS'] * df['Signo']
        
        return df

    except Exception as e:
        st.error(f"Error al procesar el archivo: {str(e)}")
        return None

# ---------------------------------------------------------
# Lógica Principal de la Aplicación
# ---------------------------------------------------------
if file_emitidos and file_recibidos:
    df_emitidos = procesar_comprobantes(file_emitidos, es_emitidos=True)
    df_recibidos = procesar_comprobantes(file_recibidos, es_emitidos=False)

    if df_emitidos is not None and df_recibidos is not None:
        
        # Totales principales
        total_debito_fiscal = df_emitidos['IVA_Final'].sum()
        total_credito_fiscal = df_recibidos['IVA_Final'].sum()

        # ---------------------------------------------------------
        # CÁLCULO DE LA POSICIÓN DE IVA (Borrador F.2002)
        # ---------------------------------------------------------
        
        # 1. Determinación del Saldo Técnico
        diferencia_tecnica = total_debito_fiscal - total_credito_fiscal - saldo_tecnico_anterior
        
        if diferencia_tecnica < 0:
            saldo_tecnico_favor_contribuyente = abs(diferencia_tecnica)
            saldo_tecnico_favor_arca = 0.0
        else:
            saldo_tecnico_favor_contribuyente = 0.0
            saldo_tecnico_favor_arca = diferencia_tecnica

        # 2. Determinación del Saldo Final (incorporando Libre Disponibilidad)
        if saldo_tecnico_favor_arca > 0:
            diferencia_final = saldo_tecnico_favor_arca - saldo_libre_disp_anterior
            if diferencia_final > 0:
                saldo_final_a_pagar_arca = diferencia_final
                saldo_final_libre_disp_remanente = 0.0
            else:
                saldo_final_a_pagar_arca = 0.0
                saldo_final_libre_disp_remanente = abs(diferencia_final)
        else:
            saldo_final_a_pagar_arca = 0.0
            saldo_final_libre_disp_remanente = saldo_libre_disp_anterior

        # ---------------------------------------------------------
        # VISUALIZACIÓN 1: POSICIÓN DE IVA / LIQUIDACIÓN
        # ---------------------------------------------------------
        st.header("📋 1. Borrador de Declaración Jurada de IVA")

        col1, col2, col3 = st.columns(3)
        col1.metric("Débito Fiscal (Ventas)", f"$ {total_debito_fiscal:,.2f}")
        col2.metric("Crédito Fiscal (Compras)", f"$ {total_credito_fiscal:,.2f}")
        
        if saldo_tecnico_favor_contribuyente > 0:
            col3.metric("Saldo Técnico Resultante", f"$ {saldo_tecnico_favor_contribuyente:,.2f}", delta="A Favor del Contribuyente", delta_color="normal")
        else:
            col3.metric("Saldo Técnico Resultante", f"$ {saldo_tecnico_favor_arca:,.2f}", delta="A Favor de ARCA", delta_color="inverse")

        st.markdown("<br>", unsafe_allow_html=True)

        # Tabla detallada de la posición
        liquidacion_data = [
            {"Concepto": "(+) IVA Débito Fiscal (Ventas)", "Monto ($)": total_debito_fiscal},
            {"Concepto": "(-) IVA Crédito Fiscal (Compras)", "Monto ($)": -total_credito_fiscal},
            {"Concepto": "(-) Saldo Técnico a Favor del Período Anterior", "Monto ($)": -saldo_tecnico_anterior},
            {"Concepto": "<b>= SALDO TÉCNICO A FAVOR DEL CONTRIBUYENTE</b>" if saldo_tecnico_favor_contribuyente > 0 else "<b>= SALDO TÉCNICO A FAVOR DE ARCA</b>", 
             "Monto ($)": saldo_tecnico_favor_contribuyente if saldo_tecnico_favor_contribuyente > 0 else saldo_tecnico_favor_arca},
            {"Concepto": "(-) Saldo de Libre Disponibilidad Período Anterior", "Monto ($)": -saldo_libre_disp_anterior},
            {"Concepto": "<b>= SALDO FINAL DEL PERÍODO (A FAVOR CONTRIBUYENTE - LIBRE DISP.)</b>" if saldo_final_a_pagar_arca == 0 else "<b>= SALDO FINAL A PAGAR A ARCA</b>",
             "Monto ($)": saldo_final_libre_disp_remanente if saldo_final_a_pagar_arca == 0 else saldo_final_a_pagar_arca}
        ]

        df_liq = pd.DataFrame(liquidacion_data)
        df_liq['Monto ($)'] = df_liq['Monto ($)'].apply(lambda x: f"$ {x:,.2f}")
        
        st.table(df_liq)

        # Cartel visual destacado
        if saldo_final_a_pagar_arca > 0:
            st.markdown(
                f"""<div class="status-card status-favor-arca">
                    🚨 SALDO A PAGAR A ARCA: $ {saldo_final_a_pagar_arca:,.2f}
                </div>""", 
                unsafe_allow_html=True
            )
        else:
            st.markdown(
                f"""<div class="status-card status-favor-contribuyente">
                    ✅ SALDO A FAVOR DEL CONTRIBUYENTE<br>
                    <small>Técnico: $ {saldo_tecnico_favor_contribuyente:,.2f} | Libre Disponibilidad: $ {saldo_final_libre_disp_remanente:,.2f}</small>
                </div>""", 
                unsafe_allow_html=True
            )

        st.markdown("---")

        # ---------------------------------------------------------
        # VISUALIZACIÓN 2: RESUMEN DE VENTAS POR COMPROBANTE
        # ---------------------------------------------------------
        st.header("📈 2. Resumen de Ventas (Débito Fiscal)")
        
        df_resumen_ventas = df_emitidos.groupby('Tipo', as_index=False).agg({
            'Neto_Final': 'sum',
            'IVA_Final': 'sum',
            'Total_Final': 'sum'
        }).rename(columns={
            'Tipo': 'Tipo de Comprobante',
            'Neto_Final': 'Total Neto Gravado',
            'IVA_Final': 'Total IVA',
            'Total_Final': 'Total Comprobante'
        })

        totales_ventas = pd.DataFrame([{
            'Tipo de Comprobante': 'TOTAL VENTAS',
            'Total Neto Gravado': df_resumen_ventas['Total Neto Gravado'].sum(),
            'Total IVA': df_resumen_ventas['Total IVA'].sum(),
            'Total Comprobante': df_resumen_ventas['Total Comprobante'].sum()
        }])
        
        df_resumen_ventas_full = pd.concat([df_resumen_ventas, totales_ventas], ignore_index=True)
        
        st.dataframe(
            df_resumen_ventas_full.style.format({
                'Total Neto Gravado': "$ {:,.2f}",
                'Total IVA': "$ {:,.2f}",
                'Total Comprobante': "$ {:,.2f}"
            }),
            use_container_width=True
        )

        st.markdown("---")

        # ---------------------------------------------------------
        # VISUALIZACIÓN 3: RESUMEN DE COMPRAS POR COMPROBANTE
        # ---------------------------------------------------------
        st.header("🛒 3. Resumen de Compras (Crédito Fiscal)")
        
        df_resumen_compras = df_recibidos.groupby('Tipo', as_index=False).agg({
            'Neto_Final': 'sum',
            'IVA_Final': 'sum',
            'Total_Final': 'sum'
        }).rename(columns={
            'Tipo': 'Tipo de Comprobante',
            'Neto_Final': 'Total Neto Gravado',
            'IVA_Final': 'Total IVA',
            'Total_Final': 'Total Comprobante'
        })

        totales_compras = pd.DataFrame([{
            'Tipo de Comprobante': 'TOTAL COMPRAS',
            'Total Neto Gravado': df_resumen_compras['Total Neto Gravado'].sum(),
            'Total IVA': df_resumen_compras['Total IVA'].sum(),
            'Total Comprobante': df_resumen_compras['Total Comprobante'].sum()
        }])
        
        df_resumen_compras_full = pd.concat([df_resumen_compras, totales_compras], ignore_index=True)

        st.dataframe(
            df_resumen_compras_full.style.format({
                'Total Neto Gravado': "$ {:,.2f}",
                'Total IVA': "$ {:,.2f}",
                'Total Comprobante': "$ {:,.2f}"
            }),
            use_container_width=True
        )

else:
    st.info("👋 Por favor, carga en el panel de la izquierda los archivos Excel de 'Mis Comprobantes Emitidos' y 'Mis Comprobantes Recibidos' para generar la posición de IVA.")