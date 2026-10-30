#Creamos el archivo de la APP en el interprete principal (Python)
#####################################################
# DASHBOARD CLV | ABASTO, VENTAS Y ANOVA
# Presentación de resultados para el socio formador
# Fuente: datos complementarios (3 hojas)
#   - Existencias_totales (corte al 15/10/2026)
#   - Histórico_de_Entradas_Salidas (2019, 2021-2026)
#   - Histórico_Ventas (05/01/2026 al 15/10/2026)
#####################################################
#Importamos librerias
import os
import base64
import unicodedata
import streamlit as st
import plotly.express as px
import pandas as pd
import numpy as np
import statsmodels.api as sm
from statsmodels.formula.api import ols
from statsmodels.stats.multicomp import pairwise_tukeyhsd
from scipy import stats

#####################################################
#DATOS A LLENAR (solo modificar esta sección)
ARCHIVO = "Datos_CLV.xlsx"
HOJA_EXIST = "Existencias_totales"
HOJA_SALIDAS = "Histórico_de_Entradas_Salidas"
HOJA_VENTAS = "Histórico_Ventas"
LOGO = "clv_logo.jpg"
#Opcional: grupo de k-means de cada artículo (exportado del notebook de clustering)
ARCHIVO_CLUSTERS = "clusters_productos.csv"
COL_GRUPO = "grupo"
FECHA_CORTE = "15/10/2026"
#####################################################

#Configuración general de la página
st.set_page_config(page_title="CLV | Abasto y Ventas", layout="wide")

#####################################################
#DISEÑO: colores y estilos
#####################################################
#Color principal (cambiar aquí si se quiere otro tono)
COLOR = "#1f6fb2"
COLOR_OSCURO = "#0f3d63"
#Colores con significado para el estado del inventario (iguales en todas las vistas)
COLORES_ESTADO = {"Normal": COLOR,
                  "Cobertura baja": "#c0392b",
                  "Sobreinventario": "#d68910",
                  "Sin movimiento 12m": "#7f8c8d"}
#Tono de tarjeta para cada estado
TONO_ESTADO = {"Normal": "base", "Cobertura baja": "alerta",
               "Sobreinventario": "aviso", "Sin movimiento 12m": "neutro"}

#Estilos CSS: encabezado, tarjetas de KPIs y títulos de sección
#(usan colores semitransparentes para verse bien en modo claro y oscuro)
CSS = """
<style>
.block-container {padding-top: 3.2rem;}
.encabezado {background: linear-gradient(90deg, COLOR_P 0%, COLOR_O 100%);
             padding: 18px 26px; border-radius: 12px; margin-bottom: 20px;
             display: flex; justify-content: space-between; align-items: center;}
.encabezado h1 {color: white; font-size: 1.7rem; margin: 0; padding: 0;}
.encabezado p {color: rgba(255,255,255,0.85); margin: 4px 0 0 0; font-size: 0.9rem;}
.encabezado img {height: 52px; background: white; border-radius: 8px; padding: 4px;}
.tarjeta {background: rgba(31,111,178,0.06); border: 1px solid rgba(128,128,128,0.2);
          border-left: 5px solid COLOR_P; border-radius: 10px; padding: 14px 16px;
          min-height: 108px; margin-bottom: 12px;}
.tarjeta .titulo {font-size: 0.75rem; text-transform: uppercase; letter-spacing: 0.05em;
                  opacity: 0.7; margin-bottom: 6px; font-weight: 600;}
.tarjeta .valor {font-size: 1.6rem; font-weight: 700; line-height: 1.2;}
.tarjeta .detalle {font-size: 0.8rem; opacity: 0.75; margin-top: 4px;}
.tarjeta.alerta {border-left-color: #c0392b; background: rgba(192,57,43,0.07);}
.tarjeta.aviso {border-left-color: #d68910; background: rgba(214,137,16,0.08);}
.tarjeta.neutro {border-left-color: #7f8c8d; background: rgba(127,140,141,0.08);}
.tarjeta.ok {border-left-color: #1e8449; background: rgba(30,132,73,0.07);}
.seccion {font-size: 1.15rem; font-weight: 600; margin: 24px 0 10px 0; padding-bottom: 6px;
          border-bottom: 2px solid rgba(31,111,178,0.35);}
.etiqueta {font-size: 0.72rem; text-transform: uppercase; letter-spacing: 0.06em;
           opacity: 0.6; margin: 16px 0 2px 0; font-weight: 700;}
.lectura {background: rgba(31,111,178,0.08); border: 1px solid rgba(31,111,178,0.25);
          border-radius: 10px; padding: 14px 18px; margin: 18px 0 8px 0; line-height: 1.55;}
.lectura .titulo {font-size: 0.75rem; text-transform: uppercase; letter-spacing: 0.06em;
                  font-weight: 700; color: COLOR_P; margin-bottom: 6px;}
.conclusion {border-left: 4px solid COLOR_P; padding: 6px 0 6px 14px; margin: 10px 0;}
.conclusion b {display: block; margin-bottom: 2px;}
</style>
""".replace("COLOR_P", COLOR).replace("COLOR_O", COLOR_OSCURO)
st.markdown(CSS, unsafe_allow_html=True)

######################################################
#FUNCIONES DE DISEÑO
######################################################
#Logo en texto base64 para poder ponerlo dentro del encabezado
def logo_html():
    if not os.path.exists(LOGO):
        return ""
    with open(LOGO, "rb") as f:
        codigo = base64.b64encode(f.read()).decode()
    tipo = "png" if LOGO.lower().endswith(".png") else "jpeg"
    return f"<img src='data:image/{tipo};base64,{codigo}'>"

#Franja de encabezado de cada vista
def encabezado(titulo, subtitulo=""):
    st.markdown(f"<div class='encabezado'><div><h1>{titulo}</h1><p>{subtitulo}</p></div>{logo_html()}</div>",
                unsafe_allow_html=True)

#Tarjeta de KPI (tono: base, alerta, aviso, neutro, ok)
def tarjeta(titulo, valor, detalle="", tono="base"):
    return (f"<div class='tarjeta {tono}'><div class='titulo'>{titulo}</div>"
            f"<div class='valor'>{valor}</div><div class='detalle'>{detalle}</div></div>")

#Fila de tarjetas: recibe una lista de tuplas (titulo, valor, detalle, tono)
def fila_tarjetas(lista):
    columnas = st.columns(len(lista))
    for col, datos_tarjeta in zip(columnas, lista):
        col.markdown(tarjeta(*datos_tarjeta), unsafe_allow_html=True)

#Título de sección dentro de una vista
def seccion(texto):
    st.markdown(f"<div class='seccion'>{texto}</div>", unsafe_allow_html=True)

#Caja de lectura del equipo (texto fijo, no editable)
def lectura(texto, titulo="Lectura del equipo"):
    st.markdown(f"<div class='lectura'><div class='titulo'>{titulo}</div>{texto}</div>",
                unsafe_allow_html=True)

#Etiqueta de grupo en la sidebar
def etiqueta_sidebar(texto):
    st.sidebar.markdown(f"<div class='etiqueta'>{texto}</div>", unsafe_allow_html=True)

#Estilo común para todas las gráficas de Plotly
def mostrar(fig):
    fig.update_layout(margin=dict(l=10, r=10, t=30, b=10),
                      paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                      legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, title=None))
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(gridcolor="rgba(128,128,128,0.2)")
    st.plotly_chart(fig)

######################################################
#FUNCIONES DE APOYO
######################################################
#Quita acentos, espacios sobrantes y pasa a mayúsculas
def normalizar_texto(serie):
    serie = serie.astype(str).str.strip().str.upper()
    serie = serie.apply(lambda x: unicodedata.normalize("NFKD", x).encode("ascii", "ignore").decode("ascii"))
    serie = serie.str.replace(r"\s+", " ", regex=True)
    return serie

#Llave para cruzar hojas por nombre de producto (solo letras y números)
def llave_nombre(serie):
    return normalizar_texto(serie).str.replace(r"[^A-Z0-9]", "", regex=True)

#Unificamos las marcas abreviadas de salidas y ventas con el nombre de existencias
MAPA_MARCAS = {
    "AG": "AMANN GIRRBACH",
    "JENSE": "JENSEN",
    "PHROZ": "PHROZEN",
    "SCHEF": "SCHEFTNER",
    "DEKEM": "DEKEMA",
    "DENTI": "DENTINE",
    "SCHIC": "SCHICK",
    "DURRD": "DURR DENTAL",
}
def unificar_marca(serie):
    return normalizar_texto(serie).replace(MAPA_MARCAS)

#Formato de p-valor: evita mostrar 0.0000
def formato_p(p):
    return "< 0.001" if p < 0.001 else f"{p:.4f}"

#Versión para textos HTML (el símbolo "<" se escribe como &lt;)
def formato_p_html(p):
    return formato_p(p).replace("<", "&lt;")

#Formato de dinero
def dinero(x):
    return f"${x:,.0f}"

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

######################################################
#RESULTADOS DEL CLUSTERING DE PRODUCTOS (actividad de clustering, histórico de compras)
#473 artículos con 5 o más compras (2021-2026), k-means con k = 7
#Valores = mediana de cada grupo; % del gasto sobre el total de los 473 artículos
######################################################
PERFILES = pd.DataFrame({
    "grupo": [6, 4, 3, 5, 0, 1, 2],
    "nombre": ["Equipos de alto valor", "Alto valor de compra reciente", "Compras grandes sin reposición",
               "Consumibles estrella", "Consumibles de alto volumen", "Precio medio, compra regular",
               "Cola larga de bajo valor"],
    "articulos": [32, 19, 28, 26, 152, 70, 146],
    "compras": [7, 5, 5, 12, 7, 9, 6],
    "unidades_compra": [5.2, 4.4, 31.4, 151.2, 105.6, 15.7, 9.1],
    "precio_unitario": [36736.3, 20251.8, 2549.7, 578.2, 80.7, 780.2, 178.9],
    "recencia_dias": [223.5, 0, 653.5, 15, 142, 34, 142],
    "antiguedad_dias": [1391.5, 714, 488.5, 1856, 1807, 1777.5, 1681],
    "pct_gasto": [49.9, 13.4, 15.5, 9.3, 5.9, 5.1, 0.9],
    "estrategia": ["Comprar bajo pedido o con control estricto: concentran la mitad del gasto",
                   "Vigilar cobertura: son caros y se están comprando ahora",
                   "Revisar si se descontinuaron: riesgo de inventario detenido",
                   "Asegurar abasto con reorden automático: son los de mayor rotación",
                   "Consolidar compras por volumen para negociar precio",
                   "Reposición regular según cobertura",
                   "Evaluar si conviene mantenerlos en catálogo"],
})
PERFILES["pct_articulos"] = PERFILES["articulos"] / PERFILES["articulos"].sum() * 100
PERFILES["etiqueta"] = "G" + PERFILES["grupo"].astype(str) + " · " + PERFILES["nombre"]

#Comparación de métodos (silhouette y tamaño del grupo mayor)
METODOS = pd.DataFrame({
    "Método": ["k-means (k = 7)", "Jerárquico · single", "Jerárquico · complete", "Jerárquico · average", "Jerárquico · ward"],
    "Silhouette": [0.349, 0.394, 0.302, 0.317, 0.241],
    "Grupo mayor (% de artículos)": [32.1, 96.0, 43.3, 79.5, 30.2],
    "Lectura": ["Buena separación y grupos con perfiles distintos (elegido)",
                "Separación alta, pero junta casi todo en un grupo",
                "Grupos desbalanceados",
                "Junta 8 de cada 10 artículos en un grupo",
                "Grupos balanceados, pero con la menor separación"],
})

######################################################
#Definimos la instancia (cache_data porque guardamos DataFrames)
@st.cache_data
######################################################
#Creamos la función de carga de datos
def load_data():
    #Lectura de las tres hojas del archivo
    df_exist = pd.read_excel(ARCHIVO, sheet_name=HOJA_EXIST)
    df_salidas = pd.read_excel(ARCHIVO, sheet_name=HOJA_SALIDAS)
    df_ventas = pd.read_excel(ARCHIVO, sheet_name=HOJA_VENTAS)

    #Quitamos espacios sobrantes en los nombres de columna (ej. "marca_proveniente ")
    for df in [df_exist, df_salidas, df_ventas]:
        df.columns = df.columns.str.strip()

    #---------------- EXISTENCIAS ----------------
    df_exist = df_exist.rename(columns={
        "nombre_articulos": "nombre_articulo",
        "marca_proveniente": "marca",
        "existencias_id": "unidades_existencia",
        "valor_inventario_articulo": "valor_inventario"})
    df_exist["marca"] = unificar_marca(df_exist["marca"])
    df_exist["llave"] = llave_nombre(df_exist["nombre_articulo"])

    #---------------- SALIDAS ----------------
    df_salidas = df_salidas.rename(columns={
        "ANO_PROVENIENTE": "anio",
        "MES_ANO": "mes",
        "SKU_INVENTARIO": "sku",
        "NOMBRE_PRODUCTO": "nombre_articulo",
        "MARCA_PROVENIENTE": "marca",
        "SALIDAS_": "salidas"})
    df_salidas["marca"] = unificar_marca(df_salidas["marca"])
    df_salidas["llave"] = llave_nombre(df_salidas["nombre_articulo"])
    #Fecha del primer día de cada mes (para graficar la serie de tiempo)
    df_salidas["periodo"] = pd.to_datetime(dict(year=df_salidas["anio"], month=df_salidas["mes"], day=1))

    #---------------- VENTAS ----------------
    #Quitamos filas duplicadas completas
    df_ventas = df_ventas.drop_duplicates()
    df_ventas = df_ventas.rename(columns={
        "folio_abreviado": "serie_folio",
        "folio_COMPLETO": "folio",
        "fecha_documento_PROVENIENTE": "fecha",
        "articulo_id_busqueda": "articulo_id",
        "nombre_articulo_en_existencia": "nombre_articulo",
        "marca_": "marca",
        "precio_unitari_o": "precio_unitario",
        "total_unidades_de_producto": "unidades",
        "total_neto_deproducto": "total_neto",
        "nombre_cliente_": "cliente",
        "estado_provenientw": "estado",
        "ciudad_mx": "ciudad",
        "CP_LOCAL": "cp",
        "vendedor_EWS": "vendedor"})
    #La fecha viene como texto dd.mm.yyyy
    df_ventas["fecha"] = pd.to_datetime(df_ventas["fecha"], format="%d.%m.%Y", errors="coerce")
    df_ventas["mes"] = df_ventas["fecha"].dt.month
    df_ventas["marca"] = unificar_marca(df_ventas["marca"])
    #Estados: quitamos acentos/mayúsculas; "MEXICO" eran ciudades del Estado de México
    df_ventas["estado"] = normalizar_texto(df_ventas["estado"]).replace({"MEXICO": "ESTADO DE MEXICO"})
    df_ventas["llave"] = llave_nombre(df_ventas["nombre_articulo"])
    #Las partidas con precio 0 se separan (no son venta con costo)
    df_ventas["con_costo"] = df_ventas["precio_unitario"] > 0

    return df_ventas, df_salidas, df_exist

###############################################################################
#Función del cruce: rotación y cobertura por artículo en existencia
@st.cache_data
def construir_cruce(df_ventas, df_salidas, df_exist, cobertura_baja, sobreinventario):
    #Ventana de los últimos 12 meses con salidas (termina en el último mes con datos)
    indice = df_salidas["anio"] * 12 + df_salidas["mes"]
    salidas_12 = df_salidas[indice > indice.max() - 12]
    sal_art = salidas_12.groupby("llave")["salidas"].sum().rename("salidas_12m")

    #Ventas con costo por artículo (2026)
    ven = df_ventas[df_ventas["con_costo"]]
    ven_art = ven.groupby("llave").agg(unidades_vendidas=("unidades", "sum"),
                                       monto_vendido=("total_neto", "sum"))

    #Clasificación ABC por monto vendido (A = 80% del monto, B = siguiente 15%, C = resto)
    abc = ven_art.sort_values("monto_vendido", ascending=False).copy()
    abc["acumulado"] = abc["monto_vendido"].cumsum() / abc["monto_vendido"].sum()
    abc["clase_abc"] = np.where(abc["acumulado"] <= 0.80, "A", np.where(abc["acumulado"] <= 0.95, "B", "C"))

    #Unimos todo sobre los artículos en existencia
    cruce = df_exist.merge(sal_art, left_on="llave", right_index=True, how="left")
    cruce = cruce.merge(abc[["unidades_vendidas", "monto_vendido", "clase_abc"]],
                        left_on="llave", right_index=True, how="left")
    cruce[["salidas_12m", "unidades_vendidas", "monto_vendido"]] = cruce[["salidas_12m", "unidades_vendidas", "monto_vendido"]].fillna(0)
    cruce["clase_abc"] = cruce["clase_abc"].fillna("Sin venta 2026")

    #Indicadores
    cruce["salidas_prom_mes"] = cruce["salidas_12m"] / 12
    cruce["cobertura_meses"] = np.where(cruce["salidas_prom_mes"] > 0,
                                        cruce["unidades_existencia"] / cruce["salidas_prom_mes"], np.nan)
    cruce["rotacion_anual"] = cruce["salidas_12m"] / cruce["unidades_existencia"]

    #Estado del inventario (umbrales didácticos: validar con la empresa)
    condiciones = [cruce["salidas_12m"] == 0,
                   cruce["cobertura_meses"] < cobertura_baja,
                   cruce["cobertura_meses"] > sobreinventario]
    estados = ["Sin movimiento 12m", "Cobertura baja", "Sobreinventario"]
    cruce["estado_inventario"] = np.select(condiciones, estados, default="Normal")

    #Artículos vendidos en 2026 que no están en existencias al corte
    sin_exist = ven_art[~ven_art.index.isin(df_exist["llave"])].copy()
    nombres = ven.drop_duplicates("llave").set_index("llave")[["nombre_articulo", "marca"]]
    sin_exist = sin_exist.join(nombres).sort_values("monto_vendido", ascending=False)

    return cruce, sin_exist

###############################################################################
#Cargo los datos obtenidos de la función "load_data"
df_ventas, df_salidas, df_exist = load_data()
###############################################################################
#CREACIÓN DEL DASHBOARD
##############################################################################
#SIDEBAR - PARTE BÁSICA (siempre visible): logo, título y navegación
if os.path.exists(LOGO):
    st.sidebar.image(LOGO)
st.sidebar.title("CLV | Abasto y Ventas")
st.sidebar.caption(f"Existencias al corte: {FECHA_CORTE}")

#Widget 1: Navegación entre vistas
etiqueta_sidebar("Navegación")
View = st.sidebar.radio(label="Vista", label_visibility="collapsed",
                        options=["Resumen ejecutivo",
                                 "Histórico de ventas",
                                 "Salidas de almacén",
                                 "Existencias al corte",
                                 "Rotación y cobertura",
                                 "Segmentación de productos",
                                 "ANOVA",
                                 "Conclusiones"])
st.sidebar.markdown("---")

#SIDEBAR - FILTROS ESPECIALES: cada vista agrega solo los que necesita (ver cada vista)
#Los umbrales de inventario solo se muestran en "Rotación y cobertura";
#en las demás vistas se usa el último valor elegido (por defecto 1 y 12 meses)
cobertura_baja = st.session_state.get("umbral_baja", 1.0)
sobreinventario = st.session_state.get("umbral_sobre", 12)
if View == "Rotación y cobertura":
    etiqueta_sidebar("Umbrales de alerta")
    cobertura_baja = st.sidebar.slider("Cobertura baja (menos de N meses)", 0.5, 6.0, cobertura_baja, 0.5)
    sobreinventario = st.sidebar.slider("Sobreinventario (más de N meses)", 6, 36, sobreinventario, 1)
    st.sidebar.caption("Umbrales didácticos: se deben validar con la empresa.")
    st.session_state["umbral_baja"] = cobertura_baja
    st.session_state["umbral_sobre"] = sobreinventario

#Construimos el cruce con los umbrales vigentes
cruce, sin_exist = construir_cruce(df_ventas, df_salidas, df_exist, cobertura_baja, sobreinventario)

#Datos de apoyo para varias vistas
ventas_cc = df_ventas[df_ventas["con_costo"]]
anio_actual = int(df_salidas["anio"].max())
ultimo_mes = int(df_salidas.loc[df_salidas["anio"] == anio_actual, "mes"].max())
mes_txt = MESES[ultimo_mes - 1]
#Salidas acumuladas del año contra el mismo periodo del año anterior (comparación justa)
ytd_actual = df_salidas[(df_salidas["anio"] == anio_actual) & (df_salidas["mes"] <= ultimo_mes)]["salidas"].sum()
ytd_previo = df_salidas[(df_salidas["anio"] == anio_actual - 1) & (df_salidas["mes"] <= ultimo_mes)]["salidas"].sum()
var_ytd = (ytd_actual / ytd_previo - 1) * 100 if ytd_previo > 0 else np.nan

#Salidas totales por año y crecimiento entre años completos (2022 -> 2025)
anual_total = df_salidas.groupby("anio")["salidas"].sum()
crecimiento = ((anual_total[2025] / anual_total[2022] - 1) * 100
               if 2022 in anual_total.index and 2025 in anual_total.index else np.nan)

#Mes atípico: el mes con más salidas de toda la serie (2021 en adelante) comparado con la mediana
totales_mes = df_salidas[df_salidas["anio"] >= 2021].groupby("periodo")["salidas"].sum()
pico_periodo = totales_mes.idxmax()
pico_valor = totales_mes.max()
pico_veces = pico_valor / totales_mes.median()
pico_txt = f"{MESES[pico_periodo.month - 1]} de {pico_periodo.year}"
#Solo se considera atípico si es al menos 3 veces un mes normal
hay_pico = pico_veces >= 3

#ANOVA de referencia: salidas totales por mes, 2022 en adelante (el que se presenta)
ref_mes = df_salidas[df_salidas["anio"] >= 2022].groupby(["anio", "mes"])["salidas"].sum().reset_index()
p_estacional = stats.f_oneway(*[g["salidas"] for _, g in ref_mes.groupby("mes")]).pvalue
estacional_sig = p_estacional < 0.05
#Textos que cambian según los datos (para que siempre digan la verdad)
tendencia_txt = f"{'crecer' if crecimiento >= 0 else 'caer'} {abs(crecimiento):.0f}%"
estable = abs(var_ytd) < 5

#Pareto de ventas: cuántos artículos generan el 80% del monto
venta_art = ventas_cc.groupby("nombre_articulo")["total_neto"].sum().sort_values(ascending=False)
n_80_total = int((venta_art.cumsum() / venta_art.sum() <= 0.80).sum()) + 1

#Totales del cruce por estado (se usan en Resumen, Rotación y Conclusiones)
valor_total = cruce["valor_inventario"].sum()
sin_mov = cruce[cruce["estado_inventario"] == "Sin movimiento 12m"]
baja = cruce[cruce["estado_inventario"] == "Cobertura baja"]
sobre = cruce[cruce["estado_inventario"] == "Sobreinventario"]
baja_a = int((baja["clase_abc"] == "A").sum())
pct_sin_exist = sin_exist["monto_vendido"].sum() / ventas_cc["total_neto"].sum() * 100

##############################################################################
# CONTENIDO DE LA VISTA 1
if View == "Resumen ejecutivo":
    encabezado("Resumen ejecutivo",
               "Ventas 05/01–15/10/2026 · Salidas 2021–2026 · Existencias al 15/10/2026")

    #Fila 1: ventas y salidas
    fila_tarjetas([
        ("Ventas 2026", dinero(ventas_cc["total_neto"].sum()), "partidas con costo"),
        ("Clientes atendidos", f"{df_ventas['cliente'].nunique():,}", f"{df_ventas['folio'].nunique():,} folios"),
        (f"Salidas ene–{mes_txt[:3]} {anio_actual}", f"{ytd_actual:,.0f}",
         f"{var_ytd:+.1f}% vs mismo periodo {anio_actual - 1}" if not np.isnan(var_ytd) else "",
         "ok" if var_ytd >= 0 else "alerta"),
        ("Valor del inventario", dinero(cruce["valor_inventario"].sum()), f"{len(df_exist):,} artículos"),
    ])
    #Fila 2: alertas de inventario
    fila_tarjetas([
        ("Sin movimiento 12 meses", f"{len(sin_mov):,}",
         f"{dinero(sin_mov['valor_inventario'].sum())} inmovilizados", "neutro"),
        ("Cobertura baja", f"{len(baja):,}", f"{(baja['clase_abc'] == 'A').sum()} son clase A", "alerta"),
        ("Sobreinventario", f"{(cruce['estado_inventario'] == 'Sobreinventario').sum():,}",
         f"más de {sobreinventario} meses de cobertura", "aviso"),
        ("Vendidos sin existencia", f"{len(sin_exist):,}",
         f"{dinero(sin_exist['monto_vendido'].sum())} vendidos en 2026", "alerta"),
    ])
    st.caption(f"Umbrales vigentes: cobertura baja < {cobertura_baja} mes(es), sobreinventario > {sobreinventario} meses "
               "(se ajustan en Rotación y cobertura).")

    #Fila 3: gráficas principales
    Contenedor_A, Contenedor_B = st.columns(2)
    with Contenedor_A:
        seccion("Salidas por año")
        anual = df_salidas[df_salidas["anio"] >= 2021].groupby("anio").agg(
            salidas=("salidas", "sum"), meses=("mes", "nunique")).reset_index()
        anual["anio"] = anual["anio"].astype(str)
        fig = px.bar(anual, x="anio", y="salidas", text="meses", color_discrete_sequence=[COLOR])
        fig.update_traces(texttemplate="%{text} meses", textposition="outside")
        mostrar(fig)
        st.caption("2020 no tiene registros y 2019 (1 mes) se omite. 2021 y 2026 tienen 10 meses."
                   + (f" La barra de {pico_periodo.year} incluye {pico_txt}, un mes atípico de {pico_valor:,.0f} salidas."
                      if hay_pico else ""))
    with Contenedor_B:
        seccion("Valor del inventario por estado")
        resumen_estado = cruce.groupby("estado_inventario")["valor_inventario"].sum().reset_index()
        fig = px.pie(resumen_estado, names="estado_inventario", values="valor_inventario", hole=0.55,
                     color="estado_inventario", color_discrete_map=COLORES_ESTADO)
        mostrar(fig)

    lectura(f"{'La operación se estabilizó: l' if estable else 'L'}as salidas de {anio_actual} van {var_ytd:+.1f}% "
            f"contra el mismo periodo de {anio_actual - 1}, después de {tendencia_txt} entre 2022 y 2025. "
            f"El inventario vale {dinero(valor_total)}: el {sobre['valor_inventario'].sum() / valor_total * 100:.0f}% "
            f"de ese valor está en sobreinventario y el {sin_mov['valor_inventario'].sum() / valor_total * 100:.0f}% "
            f"en artículos sin movimiento. El riesgo más urgente son los <b>{baja_a} artículos clase A con cobertura baja</b>.")

##############################################################################
# CONTENIDO DE LA VISTA 2
if View == "Histórico de ventas":
    encabezado("Histórico de ventas", "Partidas de factura del 05/01/2026 al 15/10/2026")

    #Filtros especiales de esta vista
    etiqueta_sidebar("Filtros de ventas")
    marcas_sel = st.sidebar.multiselect("Marcas", sorted(df_ventas["marca"].unique()), placeholder="Todas")
    incluir_cero = st.sidebar.toggle("Incluir partidas con precio 0", value=False)
    df_v = df_ventas if incluir_cero else ventas_cc
    if marcas_sel:
        df_v = df_v[df_v["marca"].isin(marcas_sel)]

    fila_tarjetas([
        ("Monto vendido", dinero(df_v["total_neto"].sum()), ""),
        ("Unidades", f"{df_v['unidades'].sum():,.0f}", ""),
        ("Folios", f"{df_v['folio'].nunique():,}", f"{df_v['cliente'].nunique():,} clientes"),
        ("Ticket promedio", dinero(df_v.groupby("folio")["total_neto"].sum().mean()), "por folio"),
    ])

    #Fila 1: tendencia mensual y ranking
    Contenedor_A, Contenedor_B = st.columns(2)
    with Contenedor_A:
        seccion("Monto vendido por mes")
        mensual = df_v.groupby("mes")["total_neto"].sum().reset_index()
        fig = px.line(mensual, x="mes", y="total_neto", markers=True, color_discrete_sequence=[COLOR])
        fig.update_xaxes(dtick=1)
        mostrar(fig)
        st.caption("Octubre solo tiene datos hasta el día 15.")
    with Contenedor_B:
        seccion("Ranking")
        Col_1, Col_2 = st.columns(2)
        Agrupar = Col_1.selectbox("Agrupar por", ["marca", "estado", "vendedor", "serie_folio"])
        top_n = Col_2.slider("Número de categorías", 5, 25, 10)
        ranking = df_v.groupby(Agrupar)["total_neto"].sum().nlargest(top_n).reset_index()
        fig = px.bar(ranking, x="total_neto", y=Agrupar, orientation="h", color_discrete_sequence=[COLOR])
        fig.update_yaxes(autorange="reversed")
        mostrar(fig)

    #Fila 2: Pareto de artículos (ABC)
    seccion("Pareto de artículos (clasificación ABC)")
    pareto = df_v.groupby("nombre_articulo")["total_neto"].sum().sort_values(ascending=False).reset_index()
    pareto["porcentaje_acumulado"] = pareto["total_neto"].cumsum() / pareto["total_neto"].sum() * 100
    pareto["articulos"] = np.arange(1, len(pareto) + 1)
    n_80 = int((pareto["porcentaje_acumulado"] <= 80).sum()) + 1
    fig = px.line(pareto, x="articulos", y="porcentaje_acumulado", color_discrete_sequence=[COLOR])
    fig.add_hline(y=80, line_dash="dash")
    mostrar(fig)
    st.info(f"{n_80:,} de {len(pareto):,} artículos ({n_80 / len(pareto) * 100:.1f}%) generan el 80% del monto vendido.")

    marca_top = df_v.groupby("marca")["total_neto"].sum().sort_values(ascending=False)
    mes_top = mensual.loc[mensual["total_neto"].idxmax(), "mes"] if len(mensual) else 1
    lectura(f"La venta está concentrada: <b>{n_80:,} artículos ({n_80 / len(pareto) * 100:.0f}%) generan el 80% del monto</b>, "
            f"así que el abasto de esos artículos es el que más protege el ingreso. "
            f"{marca_top.index[0].title()} es la marca con más venta ({marca_top.iloc[0] / marca_top.sum() * 100:.0f}% del monto) "
            f"y {MESES[int(mes_top) - 1]} fue el mes más alto del año.")

##############################################################################
# CONTENIDO DE LA VISTA 3
if View == "Salidas de almacén":
    encabezado("Salidas de almacén", "Histórico mensual 2019 y 2021–2026 (sin registros en 2020)")

    #Filtros especiales de esta vista
    etiqueta_sidebar("Filtros de salidas")
    anios = sorted(df_salidas["anio"].unique())
    anios_sel = st.sidebar.multiselect("Años", anios, default=[a for a in anios if a >= 2021])
    marcas_sel = st.sidebar.multiselect("Marcas", sorted(df_salidas["marca"].unique()), placeholder="Todas")
    df_s = df_salidas[df_salidas["anio"].isin(anios_sel)]
    if marcas_sel:
        df_s = df_s[df_s["marca"].isin(marcas_sel)]

    mensual = df_s.groupby(["anio", "mes"])["salidas"].sum().reset_index()
    #La estacionalidad se calcula sin 2021 (igual que el ANOVA) para que el mes atípico no la distorsione
    mensual_sin_2021 = mensual[mensual["anio"] >= 2022]
    estacional = mensual_sin_2021.groupby("mes")["salidas"].mean().reset_index()
    fila_tarjetas([
        ("Salidas en el periodo", f"{df_s['salidas'].sum():,.0f}", f"{len(anios_sel)} años seleccionados"),
        ("Mes típico", f"{mensual_sin_2021['salidas'].median():,.0f}" if len(mensual_sin_2021) else "-",
         "salidas por mes (mediana 2022–2026)"),
        ("Mes atípico", pico_txt.capitalize() if hay_pico else "Ninguno",
         f"{pico_valor:,.0f} salidas ({pico_veces:.0f} veces lo normal)" if hay_pico else "sin picos fuera de lo normal",
         "aviso" if hay_pico else "base"),
        (f"Ene–{mes_txt[:3]} {anio_actual} vs {anio_actual - 1}",
         f"{var_ytd:+.1f}%" if not np.isnan(var_ytd) else "-", "todas las marcas",
         "ok" if var_ytd >= 0 else "alerta"),
    ])

    #Fila 1: serie de tiempo mensual
    seccion("Salidas mensuales")
    serie = df_s.groupby("periodo")["salidas"].sum().reset_index()
    fig = px.line(serie, x="periodo", y="salidas", markers=True, color_discrete_sequence=[COLOR])
    #Señalamos el mes atípico si está dentro de lo que se muestra
    if hay_pico and pico_periodo in set(serie["periodo"]):
        valor_visible = serie.loc[serie["periodo"] == pico_periodo, "salidas"].iloc[0]
        fig.add_annotation(x=pico_periodo, y=valor_visible, text=f"Mes atípico: {pico_txt}",
                           showarrow=True, arrowhead=2, ax=90, ay=10)
    mostrar(fig)
    st.caption("El hueco de 2020 corresponde a meses sin registros (no se imputaron).")

    #Fila 2: estacionalidad y marcas
    Contenedor_A, Contenedor_B = st.columns(2)
    with Contenedor_A:
        seccion("Estacionalidad: promedio por mes")
        fig = px.bar(estacional, x="mes", y="salidas", color_discrete_sequence=[COLOR])
        fig.update_xaxes(dtick=1)
        mostrar(fig)
        st.caption("Calculado con 2022 en adelante, igual que el ANOVA, para que el mes atípico no lo distorsione.")
    with Contenedor_B:
        seccion("Marcas con más salidas")
        top_n = st.slider("Número de marcas", 5, 20, 10)
        marcas = df_s.groupby("marca")["salidas"].sum().nlargest(top_n).reset_index()
        fig = px.bar(marcas, x="salidas", y="marca", orientation="h", color_discrete_sequence=[COLOR])
        fig.update_yaxes(autorange="reversed")
        mostrar(fig)

    marca_sal = df_s.groupby("marca")["salidas"].sum().sort_values(ascending=False)
    texto_marca = (f" {marca_sal.index[0].title()} es la marca que más se mueve "
                   f"({marca_sal.iloc[0] / marca_sal.sum() * 100:.0f}% de las salidas)." if len(marca_sal) else "")
    if estacional_sig:
        texto_est = (f"Hay meses con más demanda que otros: el ANOVA con 2022–2026 encuentra diferencias significativas "
                     f"entre meses (p {formato_p_html(p_estacional)}).")
    else:
        texto_est = (f"La demanda del almacén es estable a lo largo del año: el ANOVA con 2022–2026 no encuentra diferencias "
                     f"significativas entre meses (p = {p_estacional:.2f}).")
    texto_pico = (f" El pico de <b>{pico_txt}</b> ({pico_valor:,.0f} salidas, unas {pico_veces:.0f} veces un mes normal) "
                  f"es atípico, probablemente una carga inicial o un ajuste de inventario; por eso se excluye del análisis "
                  f"estadístico y conviene validarlo con CLV." if hay_pico else "")
    lectura(texto_est + texto_pico + texto_marca)

##############################################################################
# CONTENIDO DE LA VISTA 4
if View == "Existencias al corte":
    encabezado("Existencias al corte", f"Foto del inventario al {FECHA_CORTE} · Almacén General")

    #Filtros especiales de esta vista
    etiqueta_sidebar("Filtros de existencias")
    marcas_sel = st.sidebar.multiselect("Marcas", sorted(df_exist["marca"].unique()), placeholder="Todas")
    df_e = df_exist if not marcas_sel else df_exist[df_exist["marca"].isin(marcas_sel)]

    orden = df_e.sort_values("valor_inventario", ascending=False)
    top10 = orden.head(10)["valor_inventario"].sum() / orden["valor_inventario"].sum() * 100 if len(orden) else 0
    fila_tarjetas([
        ("Artículos", f"{len(df_e):,}", f"{df_e['marca'].nunique()} marcas"),
        ("Unidades", f"{df_e['unidades_existencia'].sum():,.0f}", "en existencia"),
        ("Valor del inventario", dinero(df_e["valor_inventario"].sum()), ""),
        ("Concentración", f"{top10:.1f}%", "del valor en 10 artículos", "aviso" if top10 > 30 else "base"),
        ("Valor en 0", f"{(df_e['valor_inventario'] == 0).sum():,}", "artículos por revisar", "neutro"),
    ])

    Contenedor_A, Contenedor_B = st.columns(2)
    with Contenedor_A:
        seccion("Valor del inventario por marca")
        por_marca = df_e.groupby("marca")["valor_inventario"].sum().sort_values(ascending=False).reset_index()
        fig = px.bar(por_marca, x="valor_inventario", y="marca", orientation="h", color_discrete_sequence=[COLOR])
        fig.update_yaxes(autorange="reversed")
        mostrar(fig)
    with Contenedor_B:
        seccion("10 artículos de mayor valor")
        st.dataframe(orden.head(10)[["nombre_articulo", "marca", "unidades_existencia", "valor_inventario"]],
                     hide_index=True)

    if len(por_marca):
        lectura(f"El valor del inventario está concentrado: los 10 artículos más caros suman el <b>{top10:.0f}%</b> "
                f"del valor y {por_marca.iloc[0]['marca'].title()} concentra el "
                f"{por_marca.iloc[0]['valor_inventario'] / por_marca['valor_inventario'].sum() * 100:.0f}%. "
                f"Hay {(df_e['valor_inventario'] == 0).sum():,} artículos con valor 0 que conviene revisar en el sistema "
                f"de CLV, porque tienen unidades pero no costo registrado.")

##############################################################################
# CONTENIDO DE LA VISTA 5
if View == "Rotación y cobertura":
    encabezado("Rotación y cobertura",
               "Existencias al corte vs salidas de los últimos 12 meses y ventas 2026")

    #Tarjetas por estado del inventario (siempre en el mismo orden y color)
    tarjetas_estado = []
    for estado in ["Normal", "Cobertura baja", "Sobreinventario", "Sin movimiento 12m"]:
        sub = cruce[cruce["estado_inventario"] == estado]
        tarjetas_estado.append((estado, f"{len(sub):,}",
                                f"{dinero(sub['valor_inventario'].sum())} en inventario", TONO_ESTADO[estado]))
    fila_tarjetas(tarjetas_estado)

    Contenedor_A, Contenedor_B = st.columns(2)
    with Contenedor_A:
        seccion("Existencias vs salidas (12 meses)")
        graf = cruce[cruce["salidas_12m"] > 0]
        fig = px.scatter(graf, x="salidas_12m", y="unidades_existencia", color="estado_inventario",
                         color_discrete_map=COLORES_ESTADO, hover_name="nombre_articulo",
                         log_x=True, log_y=True)
        mostrar(fig)
        st.caption("Escala logarítmica. Los artículos sin salidas no aparecen en esta gráfica.")
    with Contenedor_B:
        seccion("Valor por clase ABC y estado")
        tabla = cruce.pivot_table(index="clase_abc", columns="estado_inventario",
                                  values="valor_inventario", aggfunc="sum", fill_value=0)
        st.dataframe(tabla.style.format("${:,.0f}"))
        st.caption("Clase A: artículos que suman el 80% del monto vendido en 2026.")

    #Tabla de alertas con descarga
    seccion("Artículos en alerta")
    estado_sel = st.radio("Estado", ["Cobertura baja", "Sin movimiento 12m", "Sobreinventario"],
                          horizontal=True, label_visibility="collapsed")
    alertas = cruce[cruce["estado_inventario"] == estado_sel].sort_values("valor_inventario", ascending=False)
    columnas_tabla = ["articulo_id", "nombre_articulo", "marca", "clase_abc", "unidades_existencia",
                      "salidas_12m", "cobertura_meses", "valor_inventario"]
    st.dataframe(alertas[columnas_tabla].round(1), hide_index=True)
    st.download_button("Descargar CSV", alertas[columnas_tabla].to_csv(index=False).encode("utf-8"),
                       file_name=f"alertas_{estado_sel}.csv", mime="text/csv")

    with st.expander(f"Artículos vendidos en 2026 sin existencia al corte ({len(sin_exist):,})"):
        st.dataframe(sin_exist[["nombre_articulo", "marca", "unidades_vendidas", "monto_vendido"]].round(0))

    lectura(f"<b>Prioridad de compra:</b> {len(baja):,} artículos se pueden agotar en menos de {cobertura_baja:g} mes(es), "
            f"y {baja_a} son clase A (de los que generan el 80% de la venta). "
            f"<b>Dinero detenido:</b> {dinero(sobre['valor_inventario'].sum())} en sobreinventario y "
            f"{dinero(sin_mov['valor_inventario'].sum())} en artículos sin salidas en un año, que pueden liberarse "
            f"con promociones, liquidación o ajustando compras. <b>Por validar:</b> {len(sin_exist):,} artículos vendidos "
            f"en 2026 no aparecen en el inventario al corte; muchos pueden ser equipos bajo pedido, no faltantes.")

##############################################################################
# CONTENIDO DE LA VISTA: SEGMENTACIÓN DE PRODUCTOS (CLUSTERING)
if View == "Segmentación de productos":
    encabezado("Segmentación de productos",
               "Clustering del histórico de compras 2021–2026 · 473 artículos con 5 o más compras · k-means, 7 grupos")

    alto = PERFILES[PERFILES["grupo"].isin([6, 4, 3])]
    bajo = PERFILES[PERFILES["grupo"].isin([0, 2])]
    g6 = PERFILES[PERFILES["grupo"] == 6].iloc[0]
    fila_tarjetas([
        ("Artículos segmentados", "473", "7 grupos con k-means"),
        ("Calidad de los grupos", "0.349", "silhouette: mejor que el jerárquico balanceado"),
        ("Equipos de alto valor", f"{g6['pct_gasto']:.0f}%",
         f"del gasto en solo {g6['articulos']} artículos ({g6['pct_articulos']:.0f}%)", "aviso"),
        ("3 grupos de alto valor", f"{alto['pct_gasto'].sum():.0f}%",
         f"del gasto en {alto['articulos'].sum()} artículos ({alto['pct_articulos'].sum():.0f}%)", "aviso"),
    ])

    Contenedor_A, Contenedor_B = st.columns(2)
    with Contenedor_A:
        seccion("¿Dónde se va el dinero de las compras?")
        concentracion = PERFILES.melt(id_vars="etiqueta", value_vars=["pct_articulos", "pct_gasto"],
                                      var_name="medida", value_name="porcentaje")
        concentracion["medida"] = concentracion["medida"].map({"pct_articulos": "% de artículos",
                                                               "pct_gasto": "% del gasto"})
        fig = px.bar(concentracion, x="porcentaje", y="etiqueta", color="medida", orientation="h", barmode="group",
                     color_discrete_map={"% de artículos": "#9fb8d0", "% del gasto": COLOR},
                     labels={"porcentaje": "%", "etiqueta": ""})
        fig.update_yaxes(autorange="reversed")
        mostrar(fig)
    with Contenedor_B:
        seccion("Mapa de grupos: precio vs volumen")
        mapa = PERFILES.assign(clave="G" + PERFILES["grupo"].astype(str))
        fig = px.scatter(mapa, x="precio_unitario", y="unidades_compra", size="pct_gasto", text="clave",
                         log_x=True, log_y=True, size_max=60, hover_name="etiqueta",
                         color_discrete_sequence=[COLOR],
                         labels={"precio_unitario": "Precio unitario mediano (escala log)",
                                 "unidades_compra": "Unidades por compra (escala log)"})
        #G4 queda junto a G6: su etiqueta va abajo para que no se encimen
        fig.update_traces(textposition=["bottom center" if c == "G4" else "top center" for c in mapa["clave"]],
                          marker=dict(opacity=0.75))
        mostrar(fig)
        st.caption("Cada burbuja es un grupo; su tamaño es el % del gasto que concentra.")

    seccion("Perfil y estrategia por grupo")
    tabla_perfil = PERFILES[["grupo", "nombre", "articulos", "pct_gasto", "precio_unitario", "unidades_compra",
                             "compras", "recencia_dias"]].rename(columns={
        "grupo": "Grupo", "nombre": "Perfil", "articulos": "Artículos", "pct_gasto": "% del gasto",
        "precio_unitario": "Precio mediano", "unidades_compra": "Unidades por compra", "compras": "Compras",
        "recencia_dias": "Días desde la última compra"})
    st.dataframe(tabla_perfil.style.format({"% del gasto": "{:.1f}%", "Precio mediano": "${:,.0f}",
                                            "Unidades por compra": "{:.1f}", "Días desde la última compra": "{:.0f}"}),
                 hide_index=True)
    st.caption("Valores medianos de cada grupo.")
    st.dataframe(PERFILES[["grupo", "nombre", "estrategia"]].rename(columns={
        "grupo": "Grupo", "nombre": "Perfil", "estrategia": "Estrategia sugerida (propuesta del equipo)"}),
        hide_index=True)

    #Cruce con el inventario: solo si existe el archivo con el grupo de cada artículo
    if os.path.exists(ARCHIVO_CLUSTERS):
        seccion("¿En qué estado de inventario está cada grupo?")
        clusters = pd.read_csv(ARCHIVO_CLUSTERS)
        clusters.columns = clusters.columns.str.strip()
        clusters = clusters[["articulo_id", COL_GRUPO]].rename(columns={COL_GRUPO: "grupo"})
        cruce_g = cruce.merge(clusters, on="articulo_id", how="inner").merge(
            PERFILES[["grupo", "etiqueta"]], on="grupo", how="left")
        st.caption(f"{cruce_g['articulo_id'].nunique()} de 473 artículos segmentados tienen existencia al corte.")
        if len(cruce_g):
            Col_1, Col_2 = st.columns(2)
            with Col_1:
                conteo_g = cruce_g.groupby(["etiqueta", "estado_inventario"]).size().reset_index(name="articulos")
                fig = px.bar(conteo_g, x="articulos", y="etiqueta", color="estado_inventario", orientation="h",
                             color_discrete_map=COLORES_ESTADO, labels={"etiqueta": "", "articulos": "Artículos"})
                fig.update_yaxes(categoryorder="array", categoryarray=list(PERFILES["etiqueta"])[::-1])
                mostrar(fig)
            with Col_2:
                valor_g = cruce_g.pivot_table(index="etiqueta", columns="estado_inventario",
                                              values="valor_inventario", aggfunc="sum", fill_value=0)
                valor_g = valor_g.reindex([e for e in PERFILES["etiqueta"] if e in valor_g.index])
                st.dataframe(valor_g.style.format("${:,.0f}"))
                st.caption("Valor del inventario al corte por grupo y estado.")

    with st.expander("¿Por qué k-means y no el agrupamiento jerárquico?"):
        st.dataframe(METODOS.style.format({"Silhouette": "{:.3f}", "Grupo mayor (% de artículos)": "{:.1f}%"}),
                     hide_index=True)
        st.markdown("El silhouette mide qué tan separados están los grupos (más alto es mejor). El jerárquico con "
                    "*single* tiene el valor más alto, pero junta el 96% de los artículos en un grupo, así que no "
                    "distingue perfiles. K-means da buena separación con grupos útiles. Además, los grupos de alto "
                    "valor coinciden en ambos métodos (26 de 28, 17 de 19 y 28 de 32 artículos), lo que respalda el resultado.")

    lectura(f"Las compras de CLV están muy concentradas: <b>{alto['articulos'].sum()} artículos "
            f"({alto['pct_articulos'].sum():.0f}%) se llevan el {alto['pct_gasto'].sum():.0f}% del gasto</b>, y solo los "
            f"{g6['articulos']} equipos de alto valor concentran la mitad. En cambio, {bajo['articulos'].sum()} artículos "
            f"({bajo['pct_articulos'].sum():.0f}%) suman apenas el {bajo['pct_gasto'].sum():.0f}%. Por eso cada grupo pide "
            "una estrategia distinta: control estricto en los de alto valor, abasto asegurado en los consumibles estrella "
            "y revisión de los que dejaron de comprarse (grupo 3), que pueden terminar como inventario detenido.")

##############################################################################
# CONTENIDO DE LA VISTA 6
if View == "ANOVA":
    encabezado("ANOVA", "¿Las diferencias entre grupos son reales o son azar?")

    #Filtros especiales de esta vista: configuración básica
    etiqueta_sidebar("Configuración del ANOVA")
    Base = st.sidebar.selectbox("Base de análisis", ["Salidas totales por mes",
                                                     "Salidas por artículo-mes",
                                                     "Ventas por partida",
                                                     "Inventario por artículo"])
    #Para las bases de salidas: 2021 es incompleto (10 meses) y tiene un mes atípico,
    #por eso por defecto el ANOVA empieza en 2022
    if Base in ["Salidas totales por mes", "Salidas por artículo-mes"]:
        excluir_2021 = st.sidebar.toggle("Excluir 2021 (incompleto y con mes atípico)", value=True)
        anio_inicio = 2022 if excluir_2021 else 2021
    if Base == "Salidas totales por mes":
        #Cada fila es un mes del almacén
        df_a = df_salidas[df_salidas["anio"] >= anio_inicio].groupby(["anio", "mes"])["salidas"].sum().reset_index()
        Lista_num = ["salidas"]
        Lista_cat = ["mes", "anio"]
    elif Base == "Salidas por artículo-mes":
        df_a = df_salidas[df_salidas["anio"] >= anio_inicio]
        Lista_num = ["salidas"]
        Lista_cat = ["mes", "anio", "marca"]
    elif Base == "Ventas por partida":
        df_a = ventas_cc
        Lista_num = ["total_neto", "unidades", "precio_unitario"]
        Lista_cat = ["marca", "estado", "serie_folio", "vendedor", "mes"]
    else:
        df_a = cruce[cruce["salidas_12m"] > 0]
        Lista_num = ["rotacion_anual", "cobertura_meses", "valor_inventario"]
        Lista_cat = ["marca", "clase_abc", "estado_inventario"]

    Variable_y = st.sidebar.selectbox(label="Variable dependiente (y)", options=Lista_num)
    Variable_x = st.sidebar.selectbox(label="Variable independiente (x)", options=Lista_cat)

    #Opciones técnicas: escondidas en un desplegable
    with st.sidebar.expander("Opciones avanzadas"):
        usar_log = st.toggle("Transformar y con logaritmo", value=Base != "Salidas totales por mes")
        min_obs = st.slider("Mínimo de observaciones por grupo", 2, 100, 4 if Base == "Salidas totales por mes" else 30)
        max_grupos = st.slider("Máximo de grupos (los más frecuentes)", 2, 15, 12)

    #Preparamos los datos: quitamos nulos, grupos pequeños y dejamos los más frecuentes
    datos = df_a[[Variable_x, Variable_y]].dropna().copy()
    datos[Variable_x] = datos[Variable_x].astype(str)
    conteo = datos[Variable_x].value_counts()
    grupos_validos = conteo[conteo >= min_obs].head(max_grupos).index
    datos = datos[datos[Variable_x].isin(grupos_validos)]
    datos["y"] = np.log1p(datos[Variable_y]) if usar_log else datos[Variable_y]
    etiqueta_y = f"log(1 + {Variable_y})" if usar_log else Variable_y

    st.caption(f"Base: {Base} · {len(datos):,} observaciones · {datos[Variable_x].nunique()} grupos · y = {etiqueta_y}")

    if datos[Variable_x].nunique() < 2:
        st.warning("Se necesitan al menos 2 grupos. Baja el mínimo de observaciones en Opciones avanzadas.")
    else:
        grupos = [grupo["y"] for nombre, grupo in datos.groupby(Variable_x)]
        f_valor, p_valor = stats.f_oneway(*grupos)
        #Tamaño del efecto: eta cuadrada = SS entre grupos / SS total
        media_total = datos["y"].mean()
        ss_entre = sum(len(g) * (g.mean() - media_total) ** 2 for g in grupos)
        ss_total = ((datos["y"] - media_total) ** 2).sum()
        eta2 = ss_entre / ss_total
        #Supuesto de varianzas iguales y alternativa no paramétrica
        lev_p = stats.levene(*grupos).pvalue
        kw_h, kw_p = stats.kruskal(*grupos)
        tamano = "pequeño" if eta2 < 0.06 else ("mediano" if eta2 < 0.14 else "grande")

        #Fila 1: tarjetas de resultados
        fila_tarjetas([
            ("F (ANOVA)", f"{f_valor:.2f}", f"{len(grupos)} grupos"),
            ("p-valor", formato_p(p_valor), "significativo" if p_valor < 0.05 else "no significativo",
             "ok" if p_valor < 0.05 else "neutro"),
            ("η² (tamaño del efecto)", f"{eta2 * 100:.1f}%",
             f"efecto {tamano}" if p_valor < 0.05 else "no concluyente (p ≥ 0.05)"),
            ("Levene", formato_p(lev_p), "varianzas distintas" if lev_p < 0.05 else "varianzas similares",
             "aviso" if lev_p < 0.05 else "base"),
            ("Kruskal-Wallis", formato_p(kw_p), f"H = {kw_h:.2f}", "ok" if kw_p < 0.05 else "neutro"),
        ])
        if p_valor < 0.05:
            st.success(f"Hay diferencias significativas entre los grupos de **{Variable_x}**: "
                       f"el factor explica el {eta2 * 100:.1f}% de la variación (efecto {tamano}).")
        else:
            st.warning(f"No hay diferencias significativas entre los grupos de **{Variable_x}**.")
        if lev_p < 0.05:
            st.caption("Levene indica varianzas distintas: apóyate también en Kruskal-Wallis, que no requiere ese supuesto.")

        #Fila 2: gráficas
        Contenedor_A, Contenedor_B = st.columns(2)
        # Contenedor A: Gráfico de barras (media ± error estándar)
        with Contenedor_A:
            seccion("Media por grupo con error estándar")
            resumen = datos.groupby(Variable_x)["y"].agg(['mean', 'sem']).reset_index()
            #Mes y año se ordenan en el tiempo; las demás variables de mayor a menor
            if Variable_x in ["mes", "anio"]:
                resumen = resumen.sort_values(Variable_x, key=lambda s: s.astype(int))
            else:
                resumen = resumen.sort_values('mean', ascending=False)
            fig = px.bar(resumen, x=Variable_x, y="mean", error_y="sem", color_discrete_sequence=[COLOR],
                         labels={"mean": f"Media de {etiqueta_y}"})
            fig.update_xaxes(type="category")
            mostrar(fig)
        # Contenedor B: distribución por grupo
        with Contenedor_B:
            seccion("Distribución por grupo")
            orden_x = list(resumen[Variable_x])
            fig = px.box(datos, x=Variable_x, y="y", labels={"y": etiqueta_y}, color_discrete_sequence=[COLOR],
                         category_orders={Variable_x: orden_x})
            fig.update_xaxes(type="category")
            mostrar(fig)

        #Fila 3: prueba post-hoc
        seccion("Tukey: ¿qué grupos difieren entre sí?")
        tukey = pairwise_tukeyhsd(datos["y"], datos[Variable_x])
        tabla_tukey = pd.DataFrame(tukey.summary().data[1:], columns=tukey.summary().data[0])
        significativos = tabla_tukey[tabla_tukey["reject"] == True]
        st.write(f"{len(significativos)} de {len(tabla_tukey)} pares son significativamente distintos.")
        st.dataframe(significativos[["group1", "group2", "meandiff", "p-adj"]], hide_index=True)

        #Fila 4: ANOVA factorial (dos factores)
        with st.expander("ANOVA factorial (dos factores)"):
            opciones_2 = [c for c in Lista_cat if c != Variable_x]
            if len(opciones_2) == 0:
                st.write("Esta base solo tiene un factor disponible.")
            else:
                Variable_x2 = st.selectbox("Segundo factor", opciones_2)
                datos2 = df_a[[Variable_x, Variable_x2, Variable_y]].dropna().copy()
                datos2[Variable_x] = datos2[Variable_x].astype(str)
                datos2[Variable_x2] = datos2[Variable_x2].astype(str)
                datos2 = datos2[datos2[Variable_x].isin(grupos_validos)]
                #Dejamos solo las categorías más frecuentes del segundo factor
                top2 = datos2[Variable_x2].value_counts().head(max_grupos).index
                datos2 = datos2[datos2[Variable_x2].isin(top2)]
                datos2["y"] = np.log1p(datos2[Variable_y]) if usar_log else datos2[Variable_y]

                #Si alguna combinación tiene una sola observación, la interacción no se puede estimar
                #(ej. mes x año en salidas totales): usamos el modelo aditivo sin interacción
                con_interaccion = datos2.groupby([Variable_x, Variable_x2]).size().min() >= 2
                signo = "*" if con_interaccion else "+"
                modelo = ols(f"y ~ C(Q('{Variable_x}')) {signo} C(Q('{Variable_x2}'))", data=datos2).fit()
                tabla = sm.stats.anova_lm(modelo, typ=2)
                if con_interaccion:
                    tabla.index = [Variable_x, Variable_x2, "Interacción", "Residual"]
                else:
                    tabla.index = [Variable_x, Variable_x2, "Residual"]
                tabla["eta2_parcial"] = tabla["sum_sq"] / (tabla["sum_sq"] + tabla.loc["Residual", "sum_sq"])
                tabla.loc["Residual", "eta2_parcial"] = np.nan
                st.dataframe(tabla.round(4))
                if con_interaccion:
                    p_inter = tabla.loc["Interacción", "PR(>F)"]
                    if p_inter < 0.05:
                        st.success(f"La interacción es significativa (p {formato_p(p_inter)}): "
                                   f"el efecto de {Variable_x} depende de {Variable_x2}.")
                    else:
                        st.info(f"La interacción no es significativa (p {formato_p(p_inter)}).")
                else:
                    st.info("Hay combinaciones con una sola observación, así que se ajustó el modelo "
                            "sin interacción (efectos de cada factor por separado).")
                medias = datos2.groupby([Variable_x, Variable_x2])["y"].mean().reset_index()
                fig = px.line(medias, x=Variable_x, y="y", color=Variable_x2, markers=True,
                              labels={"y": f"Media de {etiqueta_y}"})
                mostrar(fig)

        #Lectura del resultado: cambia según la base y la variable elegidas
        if p_valor < 0.05:
            grupo_alto = resumen.sort_values("mean", ascending=False).iloc[0][Variable_x]
            texto = (f"<b>{Variable_x.capitalize()} sí genera diferencias reales</b> en {Variable_y} "
                     f"(p {formato_p_html(p_valor)}) y explica el {eta2 * 100:.0f}% de la variación (efecto {tamano}). "
                     f"El grupo más alto es <b>{grupo_alto}</b>; la tabla de Tukey muestra qué pares son distintos.")
        else:
            texto = (f"<b>{Variable_x.capitalize()} no genera diferencias significativas</b> en {Variable_y} "
                     f"(p = {p_valor:.2f}); las diferencias que se ven en la gráfica están dentro de la variación normal.")
            if Base == "Salidas totales por mes" and Variable_x == "mes":
                texto += (" Para CLV significa que <b>su demanda no es estacional</b>: no necesita compras por temporada, "
                          "sino gestionar cada artículo (ver Rotación y cobertura).")
            elif Variable_x == "anio":
                texto += (" Aunque visualmente hay una tendencia de crecimiento, el ANOVA trata los años como grupos "
                          "sin orden y la variación entre meses es alta; el crecimiento se reporta de forma descriptiva.")
        if Base in ["Salidas totales por mes", "Salidas por artículo-mes"] and excluir_2021:
            texto += " Se excluyó 2021 por estar incompleto" + (f" y por el mes atípico de {pico_txt}." if hay_pico else ".")
        lectura(texto)

##############################################################################
# CONTENIDO DE LA VISTA 7
if View == "Conclusiones":
    encabezado("Conclusiones y recomendaciones", "Lo que dicen juntos todos los análisis del tablero")

    pct_sobre = sobre["valor_inventario"].sum() / valor_total * 100
    pct_sin_mov = sin_mov["valor_inventario"].sum() / valor_total * 100

    #Tarjetas con los números clave
    fila_tarjetas([
        ("Salidas 2022 → 2025", f"{crecimiento:+.0f}%", f"{anio_actual}: {var_ytd:+.1f}% vs {anio_actual - 1}",
         "ok" if crecimiento >= 0 else "alerta"),
        ("Estacionalidad", "Significativa" if estacional_sig else "No significativa",
         f"ANOVA por mes, p {formato_p_html(p_estacional)}" if estacional_sig else f"ANOVA por mes, p = {p_estacional:.2f}"),
        ("Dinero detenido", f"{pct_sobre + pct_sin_mov:.0f}%",
         f"del inventario ({dinero(sobre['valor_inventario'].sum() + sin_mov['valor_inventario'].sum())})", "aviso"),
        ("Clase A en riesgo", f"{baja_a}", f"con cobertura menor a {cobertura_baja:g} mes(es)", "alerta"),
    ])

    #Conclusiones integradas: cada una junta varias secciones del tablero
    seccion("Conclusiones integradas")
    conclusiones = [
        (f"1. La demanda {'creció' if crecimiento >= 0 else 'cayó'}{' y hoy está estable' if estable else ''}"
         f"{', sin temporadas' if not estacional_sig else ', con meses de mayor demanda'}.",
         f"Las salidas {'crecieron' if crecimiento >= 0 else 'cayeron'} {abs(crecimiento):.0f}% entre 2022 y 2025 "
         f"y en {anio_actual} van {var_ytd:+.1f}% contra el año anterior. "
         + (f"El ANOVA confirma que ningún mes vende más que otro (p = {p_estacional:.2f}). " if not estacional_sig
            else f"El ANOVA muestra diferencias entre meses (p {formato_p_html(p_estacional)}). ")
         + "<i>Salidas de almacén · ANOVA</i>"),
        ("2. El problema no es cuánto inventario hay, sino de qué.",
         f"El {pct_sobre:.0f}% del valor está en sobreinventario y el {pct_sin_mov:.0f}% en artículos sin movimiento, "
         f"mientras {len(baja):,} artículos están por agotarse. <i>Existencias · Rotación y cobertura</i>"),
        ("3. Pocos artículos sostienen la venta, y algunos están en riesgo.",
         f"{n_80_total:,} artículos generan el 80% del monto vendido; {baja_a} de ellos tienen cobertura baja. "
         "Un faltante ahí pesa mucho más que en el resto. <i>Histórico de ventas · Rotación y cobertura</i>"),
        ("4. Hay ventas que no pasan por el inventario.",
         f"{len(sin_exist):,} artículos vendidos en 2026 ({pct_sin_exist:.0f}% del monto) no aparecen en las existencias "
         "al corte. Pueden ser equipos bajo pedido o faltantes: hay que distinguirlos. <i>Ventas · Existencias</i>"),
        ("5. Las compras repiten el mismo patrón de concentración.",
         f"En el clustering de compras, {PERFILES[PERFILES['grupo'].isin([6, 4, 3])]['articulos'].sum()} artículos "
         f"(17%) concentran el {PERFILES[PERFILES['grupo'].isin([6, 4, 3])]['pct_gasto'].sum():.0f}% del gasto, y 32 equipos "
         "de alto valor se llevan la mitad. Gestionar bien esos pocos artículos protege la mayor parte del dinero. "
         "<i>Segmentación de productos · Histórico de ventas</i>"),
        ("6. La calidad de los datos limita algunos análisis.",
         (f"El mes atípico de {pico_txt} ({pico_valor:,.0f} salidas) y l" if hay_pico else "L")
         + "a falta de un identificador común entre hojas obligaron a "
         + ("excluir datos y " if hay_pico else "") + "cruzar por nombre de producto. <i>Salidas · Rotación</i>"),
    ]
    for titulo, cuerpo in conclusiones:
        st.markdown(f"<div class='conclusion'><b>{titulo}</b>{cuerpo}</div>", unsafe_allow_html=True)

    seccion("Recomendaciones para CLV")
    st.markdown(f"""
1. **Reponer primero los {baja_a} artículos clase A con cobertura baja** (lista descargable en Rotación y cobertura).
2. **{"Planear compras por artículo y no por temporada" if not estacional_sig else "Planear compras por artículo y anticipar los meses de mayor demanda"}**, usando la cobertura en meses como referencia.
3. **Aplicar una estrategia por segmento**: control estricto y compra bajo pedido en equipos de alto valor, reorden automático en consumibles estrella y revisión de los artículos que dejaron de comprarse.
4. **Liberar dinero detenido**: revisar los artículos en sobreinventario y sin movimiento para promociones, liquidación o ajuste de compras.
5. **Clasificar los artículos vendidos sin existencia** en bajo pedido o faltantes reales.
6. **Agregar `articulo_id` a la hoja de salidas** para cruzar las bases sin depender del nombre.
7. **Validar con CLV** {"el mes atípico de " + pico_txt + " y " if hay_pico else ""}los umbrales de alerta del tablero.
""")

    seccion("Limitaciones")
    st.markdown(f"""
- El cruce entre hojas se hizo por **nombre de producto**: las salidas no traen `articulo_id`.
- No hay registros de **2020**; 2019 solo tiene un mes; 2021 y 2026 están incompletos.
{"- **" + pico_txt.capitalize() + "** tiene un volumen atípico y se excluye del análisis estadístico." if hay_pico else ""}
- La hoja de entradas y salidas **solo trae salidas**; no se pueden calcular las entradas.
- Las existencias son una foto al **{FECHA_CORTE}**; la cobertura supone que la demanda se mantiene.
- Los umbrales de alerta son didácticos y deben validarse con la empresa.
""")