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
st.subheader("Liquidación rápida basada en 'Mis Comprobantes' o Ingreso Manual")
st.markdown("---")

# ---------------------------------------------------------
# Función de procesamiento de comprobantes
# ---------------------------------------------------------
def procesar_comprobantes(file, es_emitidos=True):
    """
    Procesa el archivo Excel de AFIP adaptando la fila de encabezado,
    multiplicando por el tipo de cambio (si es moneda extranjera) y 
    asignando signos (+ / -) según la naturaleza del comprobante.
    Si es comprobantes recibidos, excluye comprobantes tipo "B".
    """
    try:
        df = pd.read_excel(file, header=1)
        df = df.dropna(subset=['Tipo']).copy()
        
        # Filtrar comprobantes B en compras
        if not es_emitidos:
            df = df[~df['Tipo'].astype(str).str.contains(r'\bB\b', regex=True, case=False)].copy()

        cols_numericas = ['Tipo Cambio', 'Neto Gravado Total', 'Total IVA', 'Imp. Total']
        for col in cols_numericas:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            else:
                df[col] = 0.0

        df['Tipo Cambio'] = df['Tipo Cambio'].apply(lambda x: 1.0 if x <= 0 else x)
        
        df['Neto_ARS'] = df['Neto Gravado Total'] * df['Tipo Cambio']
        df['IVA_ARS'] = df['Total IVA'] * df['Tipo Cambio']
        df['Total_ARS'] = df['Imp. Total'] * df['Tipo Cambio']
        
        def determinar_signo(tipo_comp):
            tipo = str(tipo_comp).upper()
            if "CRÉDITO" in tipo or "CREDITO" in tipo:
                return -1.0
            return 1.0

        df['Signo'] = df['Tipo'].apply(determinar_signo)
        
        df['Neto_Final'] = df['Neto_ARS'] * df['Signo']
        df['IVA_Final'] = df['IVA_ARS'] * df['Signo']
        df['Total_Final'] = df['Total_ARS'] * df['Signo']
        
        return df

    except Exception as e:
        st.error(f"Error al procesar el archivo: {str(e)}")
        return None

# ---------------------------------------------------------
# Barra Lateral (Inputs)
# ---------------------------------------------------------
st.sidebar.header("📁 1. Carga de Ventas (Débito Fiscal)")
modo_ventas = st.sidebar.radio(
    "Seleccionar método de carga de ventas:",
    ("Subir Excel 'Emitidos'", "Ingreso Manual (Total Facturado)")
)

df_emitidos = None

if modo_ventas == "Subir Excel 'Emitidos'":
    file_emitidos = st.sidebar.file_uploader(
        "Cargar 'Mis Comprobantes Emitidos' (.xlsx)", 
        type=["xlsx", "xls"]
    )
    if file_emitidos:
        df_emitidos = procesar_comprobantes(file_emitidos, es_emitidos=True)
else:
    monto_total_facturado = st.sidebar.number_input(
        "Monto Total Facturado con IVA ($)",
        min_value=0.0,
        value=0.0,
        step=10000.0,
        format="%.2f"
    )
    alicuota_sel = st.sidebar.selectbox(
        "Alícuota de IVA aplicable:",
        ("21.0%", "10.5%")
    )
    
    tasa = 0.21 if alicuota_sel == "21.0%" else 0.105
    neto_manual = monto_total_facturado / (1 + tasa) if monto_total_facturado > 0 else 0.0
    iva_manual = monto_total_facturado - neto_manual
    
    df_emitidos = pd.DataFrame([{
        'Tipo': f'Ventas Totales Declaradas (Manual {alicuota_sel})',
        'Neto_Final': neto_manual,
        'IVA_Final': iva_manual,
        'Total_Final': monto_total_facturado
    }])

st.sidebar.markdown("---")
st.sidebar.header("🛒 2. Carga de Compras (Crédito Fiscal)")
file_recibidos = st.sidebar.file_uploader(
    "Cargar 'Mis Comprobantes Recibidos' (.xlsx)", 
    type=["xlsx", "xls"]
)

st.sidebar.markdown("---")
st.sidebar.header("🎁 3. Beneficio de Reducción")
tiene_reduccion = st.sidebar.selectbox(
    "¿Cuenta con beneficio de reducción?",
    ("No", "Sí")
)

porcentaje_reduccion = 0.0
if tiene_reduccion == "Sí":
    opcion_reduccion = st.sidebar.selectbox(
        "Seleccionar % de reducción:",
        ("50%", "30%", "10%")
    )
    if opcion_reduccion == "50%":
        porcentaje_reduccion = 0.50
    elif opcion_reduccion == "30%":
        porcentaje_reduccion = 0.30
    elif opcion_reduccion == "10%":
        porcentaje_reduccion = 0.10

st.sidebar.markdown("---")
st.sidebar.header("💵 4. Saldos del Período Anterior")
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
# Lógica Principal de la Aplicación
# ---------------------------------------------------------
if df_emitidos is not None and file_recibidos is not None:
    df_recibidos = procesar_comprobantes(file_recibidos, es_emitidos=False)

    if df_recibidos is not None:
        
        # Totales principales
        total_debito_fiscal = df_emitidos['IVA_Final'].sum()
        total_credito_fiscal = df_recibidos['IVA_Final'].sum()

        # ---------------------------------------------------------
        # CÁLCULO DE LA POSICIÓN DE IVA CON REDUCCIÓN
        # ---------------------------------------------------------
        diferencia_bruta = total_debito_fiscal - total_credito_fiscal

        monto_reduccion = 0.0
        # La reducción aplica solo si DF > CF
        if diferencia_bruta > 0 and porcentaje_reduccion > 0:
            monto_reduccion = diferencia_bruta * porcentaje_reduccion

        saldo_despues_reduccion = diferencia_bruta - monto_reduccion

        # Subtotal tras restar saldo técnico anterior
        diferencia_tecnica = saldo_despues_reduccion - saldo_tecnico_anterior
        
        if diferencia_tecnica < 0:
            saldo_tecnico_favor_contribuyente = abs(diferencia_tecnica)
            saldo_tecnico_favor_arca = 0.0
        else:
            saldo_tecnico_favor_contribuyente = 0.0
            saldo_tecnico_favor_arca = diferencia_tecnica

        # Aplicación de Saldo de Libre Disponibilidad
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

        # Construcción de la tabla detallada de la liquidación
        liquidacion_data = [
            {"Concepto": "(+) IVA Débito Fiscal (Ventas)", "Monto ($)": total_debito_fiscal},
            {"Concepto": "(-) IVA Crédito Fiscal (Compras)", "Monto ($)": -total_credito_fiscal},
            {"Concepto": "= Saldo a Favor de ARCA (Previo a Reducción)" if diferencia_bruta > 0 else "= Saldo Técnico a Favor Contribuyente", 
             "Monto ($)": diferencia_bruta}
        ]

        if monto_reduccion > 0:
            liquidacion_data.append({
                "Concepto": f"(-) Reducción del {int(porcentaje_reduccion * 100)}% (Beneficio)", 
                "Monto ($)": -monto_reduccion
            })
            liquidacion_data.append({
                "Concepto": "= Saldo después de Reducción", 
                "Monto ($)": saldo_despues_reduccion
            })

        liquidacion_data.extend([
            {"Concepto": "(-) Saldo Técnico a Favor Período Anterior", "Monto ($)": -saldo_tecnico_anterior},
            {"Concepto": "= SALDO TÉCNICO RESULTANTE (A FAVOR CONTRIBUYENTE)" if saldo_tecnico_favor_contribuyente > 0 else "= SALDO TÉCNICO RESULTANTE (A FAVOR DE ARCA)", 
             "Monto ($)": saldo_tecnico_favor_contribuyente if saldo_tecnico_favor_contribuyente > 0 else saldo_tecnico_favor_arca},
            {"Concepto": "(-) Saldo de Libre Disponibilidad Período Anterior", "Monto ($)": -saldo_libre_disp_anterior},
            {"Concepto": "= SALDO FINAL REANUDADO (LIBRE DISPONIBILIDAD)" if saldo_final_a_pagar_arca == 0 else "= SALDO FINAL A PAGAR A ARCA",
             "Monto ($)": saldo_final_libre_disp_remanente if saldo_final_a_pagar_arca == 0 else saldo_final_a_pagar_arca}
        ])

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
    st.info("👋 Por favor, ingresá o cargá los datos de Ventas y el archivo Excel de 'Mis Comprobantes Recibidos' en el panel lateral para calcular la posición de IVA.")