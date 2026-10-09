import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode
import plotly.express as px
import plotly.graph_objects as go
import re

st.set_page_config(page_title="Tablero de Pagos ObSBA (Jerarquía)", layout="wide")
st.title("📊 Tablero de Pagos por Categorías del Sistema - ObSBA")

# ==========================================
# FUNCIONES DE FORMATO Y LIMPIEZA
# ==========================================
def formatear_moneda(valor):
    try:
        return f"$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return "$ 0,00"

def clean_money(x):
    if pd.isna(x): return 0.0
    if isinstance(x, (int, float)): return float(x)
    x = str(x).upper().replace('$', '').replace(' ', '').replace('ARS', '').strip()
    if x == '': return 0.0
    if ',' in x and '.' in x:
        if x.rfind(',') > x.rfind('.'):
            x = x.replace('.', '').replace(',', '.')
        else:
            x = x.replace(',', '')
    elif ',' in x:
        x = x.replace(',', '.')
    try:
        return float(x)
    except:
        return 0.0

formato_pesos = JsCode(r"""
function(params) {
    var val = Number(params.value);
    if (isNaN(val)) { return "$ 0,00"; }
    var parts = val.toFixed(2).split(".");
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ".");
    return "$ " + parts.join(",");
}
""")

# ==========================================
# CARGA Y EXTRACCIÓN DE DATOS
# ==========================================
st.sidebar.header("Carga de Datos")
archivo_pagos = st.sidebar.file_uploader("Subí el Reporte de Pagos (.csv o .xlsx)", type=['csv', 'xlsx'])

if archivo_pagos is not None:
    try:
        # LECTURA DEL ARCHIVO
        if archivo_pagos.name.endswith('.csv'):
            df_pagos = pd.read_csv(archivo_pagos, skiprows=6, skipinitialspace=True, encoding='latin1', sep=';')
        else:
            df_pagos = pd.read_excel(archivo_pagos, skiprows=6)
            
        df_pagos.columns = df_pagos.columns.str.strip()
        df_pagos.dropna(how='all', inplace=True)
        df_pagos.reset_index(drop=True, inplace=True)
        
        # CORTAR LA TABLA ANTES DE LOS TOTALES (Filtro Anti-Basura)
        filtro_basura = df_pagos.astype(str).apply(lambda col: col.str.contains(r'^\s*(TOTAL|MEP:|Cheque:)', case=False, na=False)).any(axis=1)
        if filtro_basura.any():
            indice_corte = filtro_basura.idxmax()
            df_pagos = df_pagos.iloc[:indice_corte].copy()
            st.sidebar.success("✅ Resumen final de Totales excluido.")

        # IDENTIFICACIÓN DE LAS COLUMNAS DE CATEGORÍA SEGÚN EL ÍNDICE
        # Columna Y = Índice 24, Columna AG = Índice 32, Columna AI = Índice 34
        try:
            col_Y = df_pagos.columns[24]  # Tipo Benef.
            col_AG = df_pagos.columns[32] # Desc. Obra
            col_AI = df_pagos.columns[34] # Desc. Proy.
            
            # Rellenar vacíos para que no rompa la agrupación
            df_pagos['Cat_Madre'] = df_pagos[col_AI].fillna('Sin Categoría Madre').astype(str)
            df_pagos['Cat_Inferior'] = df_pagos[col_AG].fillna('Sin Categoría Inferior').astype(str)
            df_pagos['Cat_Final'] = df_pagos[col_Y].fillna('Sin Categoría Final').astype(str)
            
        except IndexError:
            st.error("⚠️ El archivo no tiene las 35 columnas necesarias para extraer las categorías (Y, AG, AI).")
            st.stop()
            
        # EXTRACCIÓN DE ENTE Y PRESTADOR
        if 'Benef.OP' in df_pagos.columns:
            df_pagos['Ente_Crudo'] = df_pagos['Benef.OP'].astype(str).str.strip()
            df_pagos['Ente'] = df_pagos['Ente_Crudo'].str.extract(r'(^\d+)', expand=False).fillna('')
            
            col_desc = 'Desc.' if 'Desc.' in df_pagos.columns else 'Descripción' if 'Descripción' in df_pagos.columns else None
            df_pagos['Nombre_Extraido'] = df_pagos['Ente_Crudo'].str.replace(r'^\d+\.?\d*\s*-?\s*', '', regex=True).str.strip()
            
            if col_desc and col_desc in df_pagos.columns:
                df_pagos['Prestador'] = df_pagos[col_desc].replace('', pd.NA).fillna(df_pagos['Nombre_Extraido'])
            else:
                df_pagos['Prestador'] = df_pagos['Nombre_Extraido']
                
            df_pagos['Prestador'] = df_pagos['Prestador'].fillna('Prestador Desconocido')
        else:
            st.error("No se encontró la columna 'Benef.OP'.")
            st.stop()
            
        # LIMPIEZA FINANCIERA
        columnas_dinero = ['Imp.OP', 'Imp.Neto', 'Imp.Ret.']
        col_anul = 'Imp.Anul.' if 'Imp.Anul.' in df_pagos.columns else ('Perc. por terceros' if 'Perc. por terceros' in df_pagos.columns else None)
        if col_anul:
            columnas_dinero.append(col_anul)
            
        for col in columnas_dinero:
            if col in df_pagos.columns:
                df_pagos[col] = df_pagos[col].apply(clean_money)
            else:
                df_pagos[col] = 0.0 
                
        # FECHAS
        if 'Fecha Pago' in df_pagos.columns:
            df_pagos['Fecha_Obj'] = pd.to_datetime(df_pagos['Fecha Pago'], dayfirst=True, errors='coerce')
        else:
            df_pagos['Fecha_Obj'] = pd.NaT

        df_final = df_pagos.copy()

        # ==========================================
        # INTERFAZ: TABS, MÉTRICAS Y GRÁFICOS
        # ==========================================
        if not df_final['Fecha_Obj'].dropna().empty:
            fecha_maxima = df_final['Fecha_Obj'].max()
            fecha_minima = df_final['Fecha_Obj'].min()
            fecha_max_str = fecha_maxima.strftime("%d/%m/%Y")
            fecha_min_str = fecha_minima.strftime("%d/%m/%Y")
            
            df_dia = df_final[df_final['Fecha_Obj'] == fecha_maxima].copy()
            titulo_tab_2 = f"📈 Acumulado Histórico (Desde {fecha_min_str} al {fecha_max_str})"
        else:
            df_dia = df_final.copy()
            fecha_max_str = "Desconocida"
            titulo_tab_2 = "📈 Acumulado Histórico"

        tab1, tab2 = st.tabs([f"📅 Pagos del Día ({fecha_max_str})", titulo_tab_2])
        
        def render_tablero(df_mostrar, es_acumulado=False):
            if df_mostrar.empty:
                st.info("No hay datos para mostrar.")
                return

            st.markdown("---")
            st.header("Análisis por Categorías del Sistema")
            
            # Selector de Categoría Madre (AI)
            lista_cat_madre = ['Todas'] + sorted(df_mostrar['Cat_Madre'].unique().tolist())
            clave_filtro = "filtro_madre_acum" if es_acumulado else "filtro_madre_dia"
            cat_sel = st.selectbox("Seleccioná la Categoría Principal para filtrar los datos:", lista_cat_madre, key=clave_filtro)
            
            if cat_sel != 'Todas':
                df_reactivo = df_mostrar[df_mostrar['Cat_Madre'] == cat_sel]
            else:
                df_reactivo = df_mostrar

            if df_reactivo.empty:
                st.warning("No hay pagos para la categoría seleccionada.")
                return

            # MÉTTRICAS REACTIVAS
            col_m1, col_m2, col_m3, col_m4 = st.columns(4)
            titulo_metrica = "Total Gastado" if cat_sel == 'Todas' else f"Gastado ({cat_sel})"
            col_m1.metric(f"{titulo_metrica} Bruto", formatear_moneda(df_reactivo['Imp.OP'].sum()))
            col_m2.metric(f"{titulo_metrica} Neto", formatear_moneda(df_reactivo['Imp.Neto'].sum()))
            col_m3.metric("Retenciones", formatear_moneda(df_reactivo['Imp.Ret.'].sum()))
            if col_anul:
                col_m4.metric("Anulaciones", formatear_moneda(df_reactivo[col_anul].sum()))
            
            # GRÁFICO EVOLUTIVO REACTIVO
            if es_acumulado and not df_reactivo['Fecha_Obj'].dropna().empty:
                st.markdown("---")
                st.subheader(f"Evolución Diaria de Pagos" + ("" if cat_sel == 'Todas' else f" - {cat_sel}"))
                df_evo = df_reactivo.groupby('Fecha_Obj')[columnas_dinero].sum().reset_index().sort_values('Fecha_Obj')
                
                fig_evo = go.Figure()
                nombres = {'Imp.OP': 'Total Bruto', 'Imp.Neto': 'Total Neto', 'Imp.Ret.': 'Retenciones', col_anul: 'Anulaciones'}
                colores = {'Imp.OP': '#1f77b4', 'Imp.Neto': '#2ca02c', 'Imp.Ret.': '#ff7f0e', col_anul: '#d62728'}
                
                for col in columnas_dinero:
                    fig_evo.add_trace(go.Scatter(
                        x=df_evo['Fecha_Obj'], y=df_evo[col],
                        mode='lines+markers', name=nombres.get(col, col),
                        line=dict(color=colores.get(col, '#333333'), width=2)
                    ))
                
                fig_evo.update_layout(
                    xaxis_title='Fecha de Pago', yaxis_title='Importe ($)',
                    hovermode='x unified', legend_title='Tipo de Importe',
                    xaxis=dict(tickformat="%d/%m/%Y"), yaxis=dict(tickformat="$,.0f"), separators=",."
                )
                st.plotly_chart(fig_evo, use_container_width=True)

            col_tabla, col_grafico = st.columns([2, 1])
            
            with col_tabla:
                st.subheader("Matriz Desplegable (Jerarquía)")
                # Agrupamos por las 3 categorías jerárquicas más el prestador
                df_resumen = df_reactivo.groupby(['Cat_Madre', 'Cat_Inferior', 'Cat_Final', 'Prestador'])[columnas_dinero].sum().reset_index()
                df_resumen = df_resumen[(df_resumen[columnas_dinero] != 0).any(axis=1)]
                
                gb = GridOptionsBuilder.from_dataframe(df_resumen)
                # Configuración multinivel para que se despliegue en cascada
                gb.configure_column('Cat_Madre', rowGroup=True, hide=False, header_name="Cat. Principal (AI)")
                gb.configure_column('Cat_Inferior', rowGroup=True, hide=False, header_name="Cat. Media (AG)")
                gb.configure_column('Cat_Final', rowGroup=True, hide=False, header_name="Cat. Final (Y)")
                
                for col in columnas_dinero:
                    gb.configure_column(col, type=["numericColumn"], valueFormatter=formato_pesos)
                
                # Configurar para que agrupe por defecto al inicio
                gb.configure_grid_options(groupDefaultExpanded=0) 
                
                grid_options = gb.build()
                AgGrid(df_resumen, gridOptions=grid_options, height=500, theme='streamlit', allow_unsafe_jscode=True)

            with col_grafico:
                st.subheader("Participación Global (Cat. Principal)")
                # Torta en base a la Categoría Madre del total general (df_mostrar)
                df_torta = df_mostrar.groupby('Cat_Madre')['Imp.OP'].sum().reset_index()
                df_torta = df_torta[df_torta['Imp.OP'] > 0]
                
                if not df_torta.empty:
                    df_torta['Monto_Format'] = df_torta['Imp.OP'].apply(formatear_moneda)
                    fig_torta = px.pie(df_torta, values='Imp.OP', names='Cat_Madre', hole=0.4,
                                       custom_data=['Monto_Format'])
                    fig_torta.update_traces(
                        textposition='inside', 
                        textinfo='percent+label', 
                        hovertemplate='<b>%{label}</b><br>Participación: %{percent}<br>Monto: %{customdata[0]}<extra></extra>'
                    )
                    fig_torta.update_layout(showlegend=False)
                    st.plotly_chart(fig_torta, use_container_width=True)
                else:
                    st.info("No hay pagos brutos para graficar.")

        with tab1:
            render_tablero(df_dia, es_acumulado=False)
            
        with tab2:
            render_tablero(df_final, es_acumulado=True)

        # ==========================================
        # BUSCADOR DETALLADO 
        # ==========================================
        st.markdown("---")
        st.header("🔍 Buscador Detallado de Prestador")
        busqueda_texto = st.text_input("Ingresá el nombre o código de ente de un prestador para ver su historial:")
        
        if busqueda_texto:
            df_busq = df_final[df_final['Prestador'].str.contains(busqueda_texto, case=False, na=False) | 
                               df_final['Ente'].astype(str).str.contains(busqueda_texto, case=False, na=False)]
            
            if not df_busq.empty:
                st.success(f"Se encontraron {len(df_busq)} registros para '{busqueda_texto}'.")
                col_b1, col_b2, col_b3, col_b4 = st.columns(4)
                col_b1.metric("Bruto", formatear_moneda(df_busq['Imp.OP'].sum()))
                col_b2.metric("Neto", formatear_moneda(df_busq['Imp.Neto'].sum()))
                col_b3.metric("Retenciones", formatear_moneda(df_busq['Imp.Ret.'].sum()))
                if col_anul:
                    col_b4.metric("Anulaciones", formatear_moneda(df_busq[col_anul].sum()))
                
                df_mostrar_b = df_busq[['Fecha_Obj', 'Ente', 'Prestador', 'Cat_Madre', 'Cat_Inferior', 'Cat_Final'] + columnas_dinero].copy()
                for col in columnas_dinero:
                    if col in df_mostrar_b.columns:
                        df_mostrar_b[col] = df_mostrar_b[col].apply(formatear_moneda)
                        
                df_mostrar_b['Fecha_Obj'] = df_mostrar_b['Fecha_Obj'].dt.strftime('%d/%m/%Y')
                df_mostrar_b.rename(columns={'Fecha_Obj': 'Fecha'}, inplace=True)
                st.dataframe(df_mostrar_b, use_container_width=True)
            else:
                st.warning("No se encontraron registros para tu búsqueda.")

    except Exception as e:
        st.error(f"Error procesando los datos: {e}")
else:
    st.info("Subí el reporte de pagos en el panel de la izquierda para comenzar.")
