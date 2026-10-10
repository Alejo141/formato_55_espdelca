import io

import numpy as np
import pandas as pd
import streamlit as st

COL_KEY = "COD_LOCALIDAD"
COLS_SALDO = ["SALDO_CAR_90A360", "SALDO_CAR_90"]
DECIMALES = 2

st.set_page_config(page_title="Distribución de saldos por localidad", layout="wide")
st.title("Distribución de SALDO_CAR por COD_LOCALIDAD")
st.caption(
    "Toma los saldos del Archivo 1 (un registro por localidad) y los reparte en partes iguales "
    "entre las filas repetidas de la misma localidad en el Archivo 2."
)


# ----------------------------------------------------------------- lectura
def elegir_hoja(archivo, etiqueta):
    """Devuelve la hoja que contiene COD_LOCALIDAD y los saldos (o la elegida por el usuario)."""
    xls = pd.ExcelFile(archivo)
    candidatas = []
    for h in xls.sheet_names:
        cols = xls.parse(h, nrows=0).columns
        if COL_KEY in cols:
            candidatas.append(h)
    if not candidatas:
        candidatas = xls.sheet_names
    if len(candidatas) == 1:
        return candidatas[0]
    return st.selectbox(f"Hoja de {etiqueta}", candidatas, key=etiqueta)


def normalizar_codigo(s: pd.Series) -> pd.Series:
    """Código como texto sin ceros a la izquierda ni '.0' (0523400000001 == 523400000001)."""
    s = s.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    return s.str.lstrip("0")


# ----------------------------------------------------------------- lógica
def repartir_exacto(valor: float, n: int, decimales: int = DECIMALES) -> np.ndarray:
    """Divide `valor` en n partes con `decimales` decimales cuya suma es EXACTAMENTE `valor`.
    Trabaja en centavos enteros; el residuo se reparte de a 1 centavo."""
    factor = 10 ** decimales
    total = int(round(valor * factor))
    base, resto = divmod(total, n)
    partes = np.full(n, base, dtype=np.int64)
    partes[:resto] += 1
    return partes / factor


def distribuir(df1: pd.DataFrame, df2: pd.DataFrame):
    df1 = df1.copy()
    df2 = df2.copy()
    df1["_k"] = normalizar_codigo(df1[COL_KEY])
    df2["_k"] = normalizar_codigo(df2[COL_KEY])

    if df1["_k"].duplicated().any():
        dup = df1.loc[df1["_k"].duplicated(keep=False), COL_KEY].unique()[:10]
        raise ValueError(
            f"El Archivo 1 debe tener un único registro por {COL_KEY}. Repetidos (ej.): {list(dup)}"
        )

    saldos = df1.set_index("_k")[COLS_SALDO].apply(pd.to_numeric, errors="coerce").fillna(0.0)
    conteo = df2.groupby("_k")["_k"].transform("size")
    df2["N_REGISTROS_LOCALIDAD"] = conteo

    resultado = df2.copy()
    for c in COLS_SALDO:
        resultado[c] = np.nan  # se sobrescribe lo que traiga el archivo 2

    for k, idx in df2.groupby("_k").groups.items():
        if k not in saldos.index:
            continue
        for c in COLS_SALDO:
            resultado.loc[idx, c] = repartir_exacto(float(saldos.at[k, c]), len(idx))

    # Códigos del archivo 1 sin correspondencia en el archivo 2
    sin_destino = df1[~df1["_k"].isin(df2["_k"])].drop(columns="_k")
    # Códigos del archivo 2 sin correspondencia en el archivo 1
    sin_origen = df2[~df2["_k"].isin(df1["_k"])].drop(columns=["_k", "N_REGISTROS_LOCALIDAD"])

    # Conciliación
    asignados = df1[df1["_k"].isin(df2["_k"])]
    filas = []
    for c in COLS_SALDO:
        t1_total = round(float(pd.to_numeric(df1[c]).sum()), 2)
        t1_asig = round(float(pd.to_numeric(asignados[c]).sum()), 2)
        t2 = round(float(resultado[c].sum()), 2)
        filas.append(
            {
                "Campo": c,
                "Suma Archivo 1 (total)": t1_total,
                "Suma Archivo 1 (códigos presentes en Archivo 2)": t1_asig,
                "Suma Resultado (Archivo 2)": t2,
                "Diferencia (códigos cruzados)": round(t1_asig - t2, 2),
                "Sin asignar (códigos solo en Archivo 1)": round(t1_total - t1_asig, 2),
            }
        )
    conciliacion = pd.DataFrame(filas)

    resultado = resultado.drop(columns="_k")
    return resultado, conciliacion, sin_destino, sin_origen


def a_excel(hojas: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
        for nombre, d in hojas.items():
            d.to_excel(w, sheet_name=nombre[:31], index=False)
    return buf.getvalue()


# ----------------------------------------------------------------- interfaz
c1, c2 = st.columns(2)
with c1:
    f1 = st.file_uploader("Archivo 1 (saldos por localidad, código único)", type=["xlsx", "xls"])
with c2:
    f2 = st.file_uploader("Archivo 2 (localidades repetidas a distribuir)", type=["xlsx", "xls"])

if f1 and f2:
    cc1, cc2 = st.columns(2)
    with cc1:
        h1 = elegir_hoja(f1, "Archivo 1")
    with cc2:
        h2 = elegir_hoja(f2, "Archivo 2")

    df1 = pd.read_excel(f1, sheet_name=h1)
    df2 = pd.read_excel(f2, sheet_name=h2)

    faltan = [
        f"Archivo {n}: {c}"
        for n, d in ((1, df1), (2, df2))
        for c in [COL_KEY] + (COLS_SALDO if n == 1 else [])
        if c not in d.columns
    ]
    if faltan:
        st.error("Faltan columnas -> " + "; ".join(faltan))
        st.stop()

    st.write(f"Archivo 1: **{len(df1):,}** filas · Archivo 2: **{len(df2):,}** filas")

    if st.button("Distribuir saldos", type="primary"):
        try:
            res, conc, sin_destino, sin_origen = distribuir(df1, df2)
        except ValueError as e:
            st.error(str(e))
            st.stop()

        st.session_state["salida"] = (res, conc, sin_destino, sin_origen)

if "salida" in st.session_state and f1 and f2:
    res, conc, sin_destino, sin_origen = st.session_state["salida"]

    st.subheader("Conciliación de sumatorias")
    st.dataframe(conc, use_container_width=True, hide_index=True)

    ok = (conc["Diferencia (códigos cruzados)"].abs() < 0.005).all()
    if ok and sin_destino.empty:
        st.success("Las sumatorias de ambos archivos coinciden exactamente.")
    elif ok:
        st.warning(
            f"Las sumatorias coinciden para los códigos que existen en ambos archivos, pero "
            f"{len(sin_destino)} código(s) del Archivo 1 no existen en el Archivo 2, por lo que su saldo "
            "no se pudo asignar (ver columna 'Sin asignar' y la hoja SIN_DESTINO del Excel)."
        )
    else:
        st.error("Las sumatorias NO coinciden. Revise los datos.")

    if not sin_origen.empty:
        st.info(f"{len(sin_origen)} fila(s) del Archivo 2 no tienen saldo en el Archivo 1 (quedan vacías).")

    t1, t2, t3, t4 = st.tabs(["Resultado", "Distribuidos (repetidos)", "Sin destino (Archivo 1)", "Sin saldo (Archivo 2)"])
    with t1:
        st.dataframe(res, use_container_width=True)
    with t2:
        st.dataframe(res[res["N_REGISTROS_LOCALIDAD"] > 1].sort_values(COL_KEY), use_container_width=True)
    with t3:
        st.dataframe(sin_destino, use_container_width=True)
    with t4:
        st.dataframe(sin_origen, use_container_width=True)

    st.download_button(
        "Descargar Excel resultado",
        data=a_excel(
            {
                "RESULTADO": res.drop(columns="N_REGISTROS_LOCALIDAD"),
                "CONCILIACION": conc,
                "SIN_DESTINO": sin_destino,
                "SIN_SALDO": sin_origen,
            }
        ),
        file_name="resultado_saldos_distribuidos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
