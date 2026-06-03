"""
actualizar_dashboard_v2.py
==========================
Lee unidades.xlsx + unidades (1).xlsx, filtra Departamentos,
recalcula todos los KPIs y reconstruye Dashboard_LAR_Final.html
incluyendo la seccion Por Liberar.
"""

import sys, re, json, os
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
import psycopg2
from datetime import datetime
from pathlib import Path

# ── CONFIG DB (env vars para CI, fallback a valores locales) ───────────────
DB = dict(
    host=    os.environ.get("DB_HOST",     "187.127.29.98"),
    port=int(os.environ.get("DB_PORT",     "5432")),
    dbname=  os.environ.get("DB_NAME",     "SQLLAR"),
    user=    os.environ.get("DB_USER",     "sqllar_user"),
    password=os.environ.get("DB_PASSWORD", "vyGaGaBkkMNNdGJyFPjTNcXxIYfIIkBm"),
    connect_timeout=10
)

# ── Modo CI (GitHub Actions) ────────────────────────────────────────────────
_CI    = os.environ.get("LAR_CI") == "1"
_HERE  = Path(__file__).parent  # directorio del script

# Mapping códigos POP Estate → nombres del dashboard
ESTADO_MAP  = {"100": "Disponible", "200": "Arrendado", "400": "No Disponible"}
SUBEST_MAP  = {               # (estado_raw, sub_estado_raw) → sub label
    ("100", "800"): "reservada",
    ("100", "600"): "por arrendar",
    ("100", "510"): "en obra",
    ("100", ""):    "por arrendar",   # disponible sin sub = libre para arrendar
    ("200", "700"): "por liberar",
    ("200", "800"): "por renovar",
    ("200", "200"): "activo",
    ("200", ""):    "activo",
    ("400", "510"): "en obra",
    ("400", ""):    "no disponible",
}

# ── RUTAS ──────────────────────────────────────────────────────────────────
if _CI:
    # GitHub Actions: script en src/, output en raíz del repo
    _ROOT      = _HERE.parent
    SRC_COLLECTIVE = _HERE / "unidades_collective.xlsx"
    HTML_SRC   = _HERE / "template.html"
    HTML_OUT   = _ROOT / "Dashboard_LAR_Final.html"
    HTML_BACKUP= _ROOT / "Dashboard_LAR_Final.html"
    HTML_REPO  = _ROOT / "Dashboard_LAR_Final.html"
else:
    # Local: rutas originales
    _DRIVE     = Path("C:/Users/mstipicevic/OneDrive - BNV/Escritorio/Claude/CODE/Ocupación")
    SRC_COLLECTIVE = _DRIVE / "unidades_collective.xlsx"
    HTML_SRC   = Path("C:/Users/mstipicevic/Downloads/MatiCode/src/template.html")
    HTML_OUT   = Path("C:/Users/mstipicevic/Downloads/Dashboard_LAR_Final_v2.html")
    HTML_BACKUP= _DRIVE / "Dashboard_LAR_Final.html"
    HTML_REPO  = Path("C:/Users/mstipicevic/Downloads/MatiCode/Dashboard_LAR_Final.html")
SRC_LAR = None  # no se usa (datos vienen de PostgreSQL)

DATE_STR = datetime.today().strftime("%d/%m/%Y")

# ── 1. LEER Y FILTRAR ──────────────────────────────────────────────────────
def load_data():
    """
    Fuente híbrida:
      - 12 proyectos LAR  → PostgreSQL (POP Estate, datos en vivo)
      - Collective Bustamante → Excel (OneDrive)
    """
    # ── PostgreSQL: 12 proyectos ──────────────────────────────────────────
    print("  Conectando a PostgreSQL (12 proyectos)...")
    conn = psycopg2.connect(**DB)
    sql = """
        SELECT
            p.nombre                              AS proyecto,
            u.nombre                              AS unidad,
            u.tipologia                           AS tipologia,
            COALESCE(u.raw->>'modelo','')         AS modelo,
            u.estado                              AS estado_raw,
            COALESCE(u.raw->>'sub_estado','')     AS sub_raw
        FROM public.unidades u
        JOIN public.propiedades p ON u.propiedad_id = p.id
        WHERE u.nombre LIKE '%-DEPA-%'
        ORDER BY p.nombre, u.nombre
    """
    df_db = pd.read_sql(sql, conn)
    conn.close()
    print(f"  DB: {len(df_db)} departamentos (12 proyectos)")

    df_db["_estado"] = df_db["estado_raw"].map(ESTADO_MAP).fillna("Desconocido")
    df_db["_sub"]    = df_db.apply(
        lambda r: SUBEST_MAP.get((str(r["estado_raw"]), str(r["sub_raw"])),
                                  str(r["sub_raw"]).lower()),
        axis=1
    )
    df_db["_tip"]  = df_db["tipologia"].astype(str).str.strip()
    df_db["_proj"] = df_db["proyecto"].astype(str).str.strip()
    df_db["Nombre"] = df_db["unidad"]
    df_db["Modelo"] = df_db["modelo"]

    # ── Excel: Collective Bustamante ──────────────────────────────────────
    print("  Excel → Collective Bustamante...")
    df_cb = pd.read_excel(SRC_COLLECTIVE, sheet_name="Unidades", header=1)
    df_cb = df_cb[df_cb["Tipo de unidad"] == "Departamento"].copy()
    tip_col = [c for c in df_cb.columns if "ipolog" in c and "Especial" not in c][0]
    df_cb["_tip"]    = df_cb[tip_col].astype(str).str.strip()
    df_cb["_estado"] = df_cb["Estado"].astype(str).str.strip()
    df_cb["_proj"]   = df_cb["Propiedad"].astype(str).str.strip()
    df_cb["Nombre"]  = df_cb["Nombre"].astype(str)
    df_cb["Modelo"]  = df_cb["Modelo"].astype(str) if "Modelo" in df_cb.columns else ""
    raw_sub = df_cb["Sub estado"].astype(str).str.strip().str.lower()
    def _map_sub_excel(row, sub):
        if sub in ("nan", "", "none", "nat"):
            return {"Arrendado": "activo", "Disponible": "por arrendar"}.get(row["_estado"], "no disponible")
        return sub
    df_cb["_sub"] = [_map_sub_excel(row, sub) for row, sub in
                     zip(df_cb.itertuples(), raw_sub)]
    print(f"  Excel: {len(df_cb)} departamentos (Collective Bustamante)")

    # ── Combinar ──────────────────────────────────────────────────────────
    cols = ["_proj","_tip","_estado","_sub","Nombre","Modelo"]
    df = pd.concat([df_db[cols], df_cb[cols]], ignore_index=True)
    return df


def load_data_excel_fallback():
    """Fallback: lee desde Excel si la DB no está disponible."""
    print("  [FALLBACK] Leyendo desde Excel...")
    df1 = pd.read_excel(SRC_LAR,        sheet_name="Unidades", header=1)
    df2 = pd.read_excel(SRC_COLLECTIVE, sheet_name="Unidades", header=1)
    df  = pd.concat([df1, df2], ignore_index=True)
    df  = df[df["Tipo de unidad"] == "Departamento"].copy()
    tip_col = [c for c in df.columns if "ipolog" in c and "Especial" not in c][0]
    df["_tip"]    = df[tip_col].astype(str).str.strip()
    df["_estado"] = df["Estado"].astype(str).str.strip()
    df["_sub"]    = df["Sub estado"].astype(str).str.strip().str.lower()
    df["_proj"]   = df["Propiedad"].astype(str).str.strip()
    return df

# ── 2. CALCULAR METRICAS ───────────────────────────────────────────────────
def compute(df):
    total       = len(df)
    # "por arrendar" (estado 100 / sub 600) se contabiliza dentro de Arrendadas
    _is_arr  = (df["_estado"] == "Arrendado") | (df["_sub"] == "por arrendar")
    _is_disp = (df["_estado"] == "Disponible") & (df["_sub"] != "por arrendar")
    arrendadas  = int(_is_arr.sum())
    disponibles = int(_is_disp.sum())
    no_disp     = int((df["_estado"] == "No Disponible").sum())
    en_obra     = int((df["_sub"] == "en obra").sum())
    reservadas  = int((df["_sub"] == "reservada").sum())
    por_arrendar= int((df["_sub"] == "por arrendar").sum())
    por_renovar = int((df["_sub"] == "por renovar").sum())
    por_liberar = int(((df["_estado"]=="Arrendado") & (df["_sub"]=="por liberar")).sum())
    ocup_global = round(arrendadas / total * 100, 4)

    # Por proyecto
    proj_list = []
    for proj, g in df.groupby("_proj"):
        tot = len(g)
        arr = int(((g["_estado"]=="Arrendado") | (g["_sub"]=="por arrendar")).sum())
        dis = int(((g["_estado"]=="Disponible") & (g["_sub"]!="por arrendar")).sum())
        nod = int((g["_estado"]=="No Disponible").sum())
        eob = int((g["_sub"]=="en obra").sum())
        res = int((g["_sub"]=="reservada").sum())
        paa = int((g["_sub"]=="por arrendar").sum())
        prv = int((g["_sub"]=="por renovar").sum())
        pli = int(((g["_estado"]=="Arrendado")&(g["_sub"]=="por liberar")).sum())
        pct = round(arr / tot * 100, 4) if tot else 0
        gap = round(pct - 95, 4)
        uds_needed = max(0, round(0.95 * tot) - arr)
        proj_list.append({
            "Propiedad": proj, "Total": float(tot), "Arrendados": float(arr),
            "Disponibles": float(dis), "No_Disp": float(nod), "En_Obra": float(eob),
            "Reservadas": float(res), "Por_Arrendar": float(paa),
            "Por_Liberar": float(pli), "Por_Renovar": float(prv),
            "Pct_Ocup": round(pct/100, 4), "Gap": round(gap/100, 4),
            "Uds_Needed": float(uds_needed)
        })

    proj_desc = sorted(proj_list, key=lambda x: -x["Pct_Ocup"])
    proj_asc  = sorted(proj_list, key=lambda x:  x["Pct_Ocup"])

    # Tipologias para heatmap
    tips = sorted(df["_tip"].unique())
    projs_hm = [p["Propiedad"] for p in proj_asc]
    hm_pct = []
    for proj in projs_hm:
        row = []
        for tip in tips:
            sub = df[(df["_proj"]==proj) & (df["_tip"]==tip)]
            t = len(sub); a = int(((sub["_estado"]=="Arrendado")|(sub["_sub"]=="por arrendar")).sum())
            row.append(round(a/t*100, 1) if t else None)
        hm_pct.append(row)

    # Pipeline por proyecto (ordenado por total pipeline desc)
    pipe = []
    for p in proj_list:
        total_pipe = int(p["Por_Arrendar"]+p["Por_Liberar"]+p["Por_Renovar"]+p["Reservadas"]+p["En_Obra"])
        if total_pipe > 0:
            pipe.append({**{k: int(v) if k != "Propiedad" else v
                           for k,v in p.items()
                           if k in ["Propiedad","Por_Arrendar","Por_Liberar","Por_Renovar","Reservadas","En_Obra"]},
                         "Total_Pipe": total_pipe})
    pipe = sorted(pipe, key=lambda x: -x["Total_Pipe"])

    # Risk (Por_Liberar / Arrendados)
    risk = []
    for p in proj_list:
        arr = p["Arrendados"]
        pli = p["Por_Liberar"]
        if arr > 0 and pli > 0:
            risk.append({**p, "Risk_Pct": round(pli/arr*100, 1)})
    risk = sorted(risk, key=lambda x: -x["Risk_Pct"])

    # Disponibles por proyecto (para tabla y JS) — excluye "por arrendar" (ya van en arrendadas)
    disp = df[(df["_estado"]=="Disponible") & (df["_sub"]!="por arrendar")].copy()
    disp_table = disp[["_proj","_tip","Nombre","Modelo"]].copy()
    disp_table.columns = ["Propiedad","Tipologia","Nombre","Modelo"]
    unidades_disp = []
    for _, row in disp.iterrows():
        unidades_disp.append({
            "Propiedad": str(row["_proj"]),
            "Tipologia": str(row["_tip"]),
            "Modelo":    str(row["Modelo"]) if pd.notna(row["Modelo"]) else "",
            "Nombre":    str(row["Nombre"]) if pd.notna(row["Nombre"]) else "",
            "SubEstado": str(row["_sub"])   if pd.notna(row["_sub"])   else "",
        })

    # tipo_data para chart de Composicion por Tipologia
    tips_all = sorted(df["_tip"].unique().tolist())
    tipo_data = {"tipologias": tips_all, "projects": projs_hm}
    for tip in tips_all:
        totals_t, arr_t, pct_t = [], [], []
        for proj in projs_hm:
            sub = df[(df["_proj"]==proj) & (df["_tip"]==tip)]
            t = len(sub); a = int(((sub["_estado"]=="Arrendado")|(sub["_sub"]=="por arrendar")).sum())
            totals_t.append(t); arr_t.append(a)
            pct_t.append(round(a/t*100, 1) if t else 0)
        tipo_data[tip] = {"total": totals_t, "arrendados": arr_t, "pct": pct_t}

    # Por Liberar detalle
    pol = df[(df["_estado"]=="Arrendado") & (df["_sub"]=="por liberar")]
    pol_by_proj = pol.groupby("_proj").size().sort_values(ascending=False).to_dict()
    pol_by_tipo = pol.groupby("_tip").size().sort_values(ascending=False).to_dict()
    pol_matrix  = pol.groupby(["_proj","_tip"]).size().unstack(fill_value=0)

    # Reservadas detalle
    res = df[df["_sub"]=="reservada"]
    res_by_proj = res.groupby("_proj").size().sort_values(ascending=False).to_dict()
    res_by_tipo = res.groupby("_tip").size().sort_values(ascending=False).to_dict()
    res_matrix  = res.groupby(["_proj","_tip"]).size().unstack(fill_value=0)
    # Detalle unit-level con Nombre y Modelo
    res_detalle = []
    for _, row in res.iterrows():
        res_detalle.append({
            "Propiedad": str(row["_proj"]),
            "Tipologia": str(row["_tip"]),
            "Modelo":    str(row["Modelo"]) if pd.notna(row["Modelo"]) else "",
            "Nombre":    str(row["Nombre"]) if pd.notna(row["Nombre"]) else "",
        })

    # Colectiva y problematicos para historico
    cb   = df[df["_proj"]=="Collective Bustamante"]
    cb_pct = round(int(((cb["_estado"]=="Arrendado")|(cb["_sub"]=="por arrendar")).sum())/len(cb)*100, 2) if len(cb) else 0
    imu  = df[df["_proj"]=="IMU San Cristóbal"]
    imu_pct = round(int(((imu["_estado"]=="Arrendado")|(imu["_sub"]=="por arrendar")).sum())/len(imu)*100, 2) if len(imu) else 0
    blend = df[df["_proj"]=="Blend Apoquindo"]
    blend_pct = round(int(((blend["_estado"]=="Arrendado")|(blend["_sub"]=="por arrendar")).sum())/len(blend)*100, 2) if len(blend) else 0

    return {
        "total": total, "arrendadas": arrendadas, "disponibles": disponibles,
        "no_disp": no_disp, "en_obra": en_obra, "reservadas": reservadas,
        "por_arrendar": por_arrendar, "por_renovar": por_renovar,
        "por_liberar": por_liberar,
        "ocup_global": ocup_global,
        "proj_desc": proj_desc, "proj_asc": proj_asc,
        "tips": tips, "projs_hm": projs_hm, "hm_pct": hm_pct,
        "pipe": pipe, "risk": risk,
        "disp_table": disp_table,
        "pol_by_proj": pol_by_proj, "pol_by_tipo": pol_by_tipo,
        "pol_matrix": pol_matrix,
        "res_by_proj": res_by_proj, "res_by_tipo": res_by_tipo,
        "res_matrix": res_matrix, "res_detalle": res_detalle,
        "cb_pct": cb_pct, "imu_pct": imu_pct, "blend_pct": blend_pct,
        "unidades_disp": unidades_disp,
        "tipo_data": tipo_data,
    }

# ── 2b. CARGAR CONTRATOS (vencimientos) ───────────────────────────────────
def load_precios_disponibles():
    """Min/Max precio arriendo (UF) + n_disp por (proyecto, tipologia, modelo)."""
    try:
        conn = psycopg2.connect(**DB)
        cur  = conn.cursor()
        cur.execute("""
            SELECT p.nombre,
                   u.tipologia,
                   COALESCE(u.raw->>'modelo','') AS modelo,
                   MIN(u.precio_monto)           AS precio_min,
                   MAX(u.precio_monto)           AS precio_max,
                   COUNT(*)                      AS n_disp,
                   u.precio_divisa
            FROM public.unidades u
            JOIN public.propiedades p ON u.propiedad_id = p.id
            WHERE u.nombre LIKE '%%-DEPA-%%'
              AND u.estado = '100'
              AND u.precio_monto IS NOT NULL
              AND u.precio_monto > 0
            GROUP BY p.nombre, u.tipologia, COALESCE(u.raw->>'modelo',''), u.precio_divisa
            ORDER BY p.nombre, u.tipologia, modelo
        """)
        rows = cur.fetchall()
        conn.close()
        precios = {}
        for proj, tip, modelo, min_p, max_p, n_d, divisa in rows:
            div = str(divisa or "UF")
            if proj not in precios:
                precios[proj] = {}
            if tip not in precios[proj]:
                precios[proj][tip] = {"min": float(min_p), "max": float(max_p),
                                      "n_disp": 0, "divisa": div, "modelos": []}
            entry = precios[proj][tip]
            entry["min"]     = min(entry["min"], float(min_p))
            entry["max"]     = max(entry["max"], float(max_p))
            entry["n_disp"] += int(n_d)
            entry["modelos"].append({"modelo": str(modelo), "min": float(min_p),
                                     "max": float(max_p), "n_disp": int(n_d)})
        n = sum(len(v) for v in precios.values())
        print(f"  Precios disponibles (DB): {n} combinaciones")
        _merge_collective_precios(precios)
        return precios
    except Exception as e:
        print(f"  [WARN] No se pudo cargar precios DB: {e}")
        precios = {}
        _merge_collective_precios(precios)
        return precios


def _merge_collective_precios(precios):
    """Agrega precios de Collective Bustamante desde Excel al dict precios."""
    if not SRC_COLLECTIVE.exists():
        print(f"  [WARN] No se encontró {SRC_COLLECTIVE}")
        return
    try:
        u  = pd.read_excel(SRC_COLLECTIVE, sheet_name="Unidades", header=1)
        pr = pd.read_excel(SRC_COLLECTIVE, sheet_name="Precios",  header=1)
        disp = u[(u["Tipo de unidad"] == "Departamento") & (u["Estado"] == "Disponible")].copy()
        p_lista = pr[(pr["Tipo"] == "Lista") & (pr["Concepto"] == "Arriendo") & (pr["Monto"] > 0)]
        merged = p_lista.merge(
            disp[["Nombre","Tipología","Modelo"]].rename(columns={"Nombre":"Unidad"}),
            on="Unidad", how="inner"
        )
        proj = "Collective Bustamante"
        precios.setdefault(proj, {})
        for (tip, modelo), g in merged.groupby(["Tipología","Modelo"]):
            tip, modelo = str(tip), str(modelo)
            mn, mx, nd = float(g["Monto"].min()), float(g["Monto"].max()), int(len(g))
            if tip not in precios[proj]:
                precios[proj][tip] = {"min": mn, "max": mx, "n_disp": 0, "divisa": "UF", "modelos": []}
            e = precios[proj][tip]
            e["min"] = min(e["min"], mn)
            e["max"] = max(e["max"], mx)
            e["n_disp"] += nd
            e["modelos"].append({"modelo": modelo, "min": mn, "max": mx, "n_disp": nd})
        n_tips = len(precios[proj])
        print(f"  Precios Collective Bustamante (Excel): {n_tips} tipologías, {sum(e['n_disp'] for e in precios[proj].values())} unidades disp.")
    except Exception as e:
        print(f"  [WARN] Precios Collective desde Excel: {e}")

def load_uf():
    """Obtiene el valor actual de la UF: primero mindicador.cl, luego BD."""
    import urllib.request as _urlreq, json as _json2
    try:
        req = _urlreq.Request(
            "https://mindicador.cl/api/uf",
            headers={"User-Agent": "Mozilla/5.0"}
        )
        with _urlreq.urlopen(req, timeout=6) as r:
            data = _json2.loads(r.read())
        uf = float(data["serie"][0]["valor"])
        fecha = data["serie"][0]["fecha"][:10]
        print(f"  UF ({fecha}): ${uf:,.2f}")
        return uf
    except Exception as e:
        print(f"  [WARN] mindicador.cl falló ({e}) — obteniendo UF desde BD...")
    # Fallback: valor más reciente de conversión en liquidacion_cargos
    try:
        import psycopg2 as _pg
        conn = _pg.connect(**DB)
        cur  = conn.cursor()
        cur.execute("""
            SELECT divisa_conversion_valor, divisa_conversion_fecha
            FROM liquidacion_cargos
            WHERE divisa='Unidad de fomento' AND divisa_conversion_valor IS NOT NULL
            ORDER BY divisa_conversion_fecha DESC LIMIT 1
        """)
        row = cur.fetchone()
        conn.close()
        if row:
            uf = float(row[0])
            print(f"  UF desde BD ({row[1]}): ${uf:,.2f}")
            return uf
    except Exception as e2:
        print(f"  [WARN] BD UF fallback falló: {e2}")
    return None


def load_tendencias_proyectos():
    """Ocupacion activa por proyecto: mes actual vs mes anterior (para flechas de tendencia)."""
    try:
        conn = psycopg2.connect(**DB)
        cur  = conn.cursor()
        cur.execute("""
            SELECT p.nombre,
                   COUNT(DISTINCT CASE
                       WHEN c.cargo_concepto = 'Arriendo'
                        AND c.fecha_inicio <= (date_trunc('month', NOW()) + INTERVAL '1 month'
                                               - INTERVAL '1 day')::date
                        AND c.fecha_fin    >= date_trunc('month', NOW())::date
                       THEN c.unidad_id END) AS now_activos,
                   COUNT(DISTINCT CASE
                       WHEN c.cargo_concepto = 'Arriendo'
                        AND c.fecha_inicio <= (date_trunc('month', NOW())
                                               - INTERVAL '1 day')::date
                        AND c.fecha_fin    >= (date_trunc('month', NOW())
                                               - INTERVAL '1 month')::date
                       THEN c.unidad_id END) AS prev_activos
            FROM public.contratos c
            JOIN public.unidades u ON c.unidad_id = u.id
            JOIN public.propiedades p ON u.propiedad_id = p.id
            WHERE u.nombre LIKE '%%-DEPA-%%'
            GROUP BY p.nombre
        """)
        rows = cur.fetchall()
        conn.close()
        result = {proj: {"now": int(now or 0), "prev": int(prev or 0)}
                  for proj, now, prev in rows}
        print(f"  Tendencias: {len(result)} proyectos con datos historicos")
        return result
    except Exception as e:
        print(f"  [WARN] No se pudo cargar tendencias: {e}")
        return {}


def build_hero_section(m, uf_valor=None, hist_data=None):
    """4 KPI cards at page top + global search bar."""
    total      = int(m.get("total", 0))
    arrendadas = int(m.get("arrendadas", 0))
    por_lib    = int(m.get("por_liberar", 0))
    proj_desc  = m.get("proj_desc", [])
    n_total    = len(proj_desc)
    n_below    = sum(1 for p in proj_desc if p.get("Pct_Ocup", 0) < 0.95)
    ocup_pct   = round(arrendadas / total * 100, 1) if total else 0
    ocup_clr   = "#16A34A" if ocup_pct >= 95 else "#D97706" if ocup_pct >= 85 else "#DC2626"
    # Trend vs prev month
    trend_html = ""
    if hist_data is not None and len(hist_data) >= 2:
        prev_pct = round(float(hist_data['pct'].iloc[-2]), 1)
        delta    = round(ocup_pct - prev_pct, 1)
        if delta > 0.05:
            trend_html = f'<span style="color:#16A34A;font-size:.68rem;font-weight:700">↑{delta}pp vs mes ant.</span>'
        elif delta < -0.05:
            trend_html = f'<span style="color:#DC2626;font-size:.68rem;font-weight:700">↓{abs(delta)}pp vs mes ant.</span>'
        else:
            trend_html = '<span style="color:#9CA3AF;font-size:.68rem">→ Sin cambio</span>'
    uf_str     = f"${uf_valor:,.2f}" if uf_valor else "—"
    import datetime
    today_str  = datetime.date.today().strftime("%d/%m/%Y")

    return f"""
<div id="hero-kpis" style="display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:18px">
  <div style="background:#fff;border-radius:14px;padding:18px 20px;border:1px solid #E2E8F0;border-left:4px solid {ocup_clr};box-shadow:0 1px 6px rgba(0,0,0,.05)">
    <div style="font-size:.58rem;color:#9CA3AF;font-weight:700;text-transform:uppercase;letter-spacing:.1em">Ocupaci&oacute;n Total</div>
    <div style="font-size:2rem;font-weight:800;color:{ocup_clr};line-height:1.2">{ocup_pct}%</div>
    <div style="font-size:.68rem;color:#6B7A8D;margin-top:2px">{arrendadas:,} / {total:,} unidades</div>
    <div style="margin-top:4px">{trend_html}</div>
  </div>
  <div style="background:#fff;border-radius:14px;padding:18px 20px;border:1px solid #E2E8F0;border-left:4px solid #D97706;box-shadow:0 1px 6px rgba(0,0,0,.05)">
    <div style="font-size:.58rem;color:#9CA3AF;font-weight:700;text-transform:uppercase;letter-spacing:.1em">Bajo Meta</div>
    <div style="font-size:2rem;font-weight:800;color:#D97706;line-height:1.2" id="hero-bajo-meta">{n_below}</div>
    <div style="font-size:.68rem;color:#6B7A8D;margin-top:2px">de {n_total} proyectos &mdash; meta <span id="hero-meta-pct">95</span>%</div>
  </div>
  <div style="background:#fff;border-radius:14px;padding:18px 20px;border:1px solid #E2E8F0;border-left:4px solid #DC2626;box-shadow:0 1px 6px rgba(0,0,0,.05)">
    <div style="font-size:.58rem;color:#9CA3AF;font-weight:700;text-transform:uppercase;letter-spacing:.1em">Por Liberar</div>
    <div style="font-size:2rem;font-weight:800;color:#DC2626;line-height:1.2">{por_lib}</div>
    <div style="font-size:.68rem;color:#6B7A8D;margin-top:2px">contratos vencidos activos</div>
  </div>
  <div style="background:#fff;border-radius:14px;padding:18px 20px;border:1px solid #E2E8F0;border-left:4px solid #0369A1;box-shadow:0 1px 6px rgba(0,0,0,.05)">
    <div style="font-size:.58rem;color:#9CA3AF;font-weight:700;text-transform:uppercase;letter-spacing:.1em">UF Hoy</div>
    <div style="font-size:2rem;font-weight:800;color:#0369A1;line-height:1.2">{uf_str}</div>
    <div style="font-size:.68rem;color:#6B7A8D;margin-top:2px">Valor al {today_str}</div>
  </div>
</div>
<div id="global-search-bar" style="margin-bottom:18px;display:flex;align-items:center;gap:10px">
  <div style="flex:1;position:relative;max-width:400px">
    <span style="position:absolute;left:11px;top:50%;transform:translateY(-50%);color:#9CA3AF;font-size:.8rem;pointer-events:none">&#128269;</span>
    <input id="global-search" type="text" placeholder="Buscar proyecto..."
           oninput="globalSearch(this.value)"
           style="width:100%;padding:8px 12px 8px 32px;border:1px solid #E2E8F0;border-radius:10px;font-size:.8rem;outline:none;box-sizing:border-box;background:#fff;transition:border-color .15s"
           onfocus="this.style.borderColor='#00A8B4'" onblur="this.style.borderColor='#E2E8F0'">
  </div>
  <span id="search-count" style="font-size:.71rem;color:#6B7A8D;white-space:nowrap;background:#F1F5F9;padding:3px 9px;border-radius:99px"></span>
</div>
"""


def load_contratos():
    """Carga contratos activos (estado=300) con fecha_fin para análisis de vencimientos."""
    import warnings; warnings.filterwarnings("ignore")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("""
        SELECT DISTINCT ON (c.contrato_id)
            c.contrato_id, c.folio,
            p.nombre   AS proyecto,
            u.nombre   AS unidad,
            u.tipologia,
            c.fecha_inicio::text,
            c.fecha_fin::text,
            c.ejecutivo
        FROM public.contratos c
        JOIN public.propiedades p ON c.propiedad_id = p.id
        JOIN public.unidades u    ON c.unidad_id    = u.id
        WHERE c.cargo_concepto = 'Arriendo'
          AND u.nombre LIKE '%%-DEPA-%%'
          AND c.estado_id = '300'
        ORDER BY c.contrato_id, c.fecha_actualizacion DESC
    """)
    rows = cur.fetchall()
    cols = [d[0] for d in cur.description]
    conn.close()
    df = pd.DataFrame(rows, columns=cols)
    df["fecha_fin"] = pd.to_datetime(df["fecha_fin"])
    today = pd.to_datetime(DATE_STR, dayfirst=True)
    df["dias"] = (df["fecha_fin"] - today).dt.days
    df["mes_venc"] = df["fecha_fin"].dt.to_period("M").astype(str)
    return df


def load_renovaciones():
    """Cuenta unidades con contrato nuevo (300) tras termino de contrato anterior, por período."""
    conn = psycopg2.connect(**DB)
    cur  = conn.cursor()
    cur.execute("""
        SELECT
            COUNT(DISTINCT c_n.unidad_id) FILTER (WHERE c_n.fecha_inicio >= CURRENT_DATE - 30)  AS r30,
            COUNT(DISTINCT c_n.unidad_id) FILTER (WHERE c_n.fecha_inicio >= CURRENT_DATE - 60)  AS r60,
            COUNT(DISTINCT c_n.unidad_id) FILTER (WHERE c_n.fecha_inicio >= CURRENT_DATE - 90)  AS r90,
            COUNT(DISTINCT c_n.propiedad_id) FILTER (WHERE c_n.fecha_inicio >= CURRENT_DATE - 30) AS p30
        FROM contratos c_n
        JOIN contratos c_a ON c_a.unidad_id = c_n.unidad_id
            AND c_a.estado_id IN ('400','410')
            AND c_a.fecha_fin >= c_n.fecha_inicio - INTERVAL '15 days'
            AND c_a.folio != c_n.folio
        WHERE c_n.estado_id = '300'
          AND c_n.cargo_concepto = 'Arriendo'
    """)
    r = cur.fetchone()
    conn.close()
    return {"r30": int(r[0] or 0), "r60": int(r[1] or 0),
            "r90": int(r[2] or 0), "p30": int(r[3] or 0)}


def load_historico(n_db):
    """Reconstruye % ocupación mensual real (últimos 12 meses) contando contratos activos por mes."""
    import warnings; warnings.filterwarnings("ignore")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("""
        SELECT
            to_char(m, 'YYYY-MM') AS mes,
            (
                SELECT COUNT(DISTINCT c.unidad_id)
                FROM public.contratos c
                JOIN public.unidades u ON c.unidad_id = u.id
                WHERE c.cargo_concepto = 'Arriendo'
                  AND u.nombre LIKE '%%-DEPA-%%'
                  AND c.fecha_inicio <= (m + INTERVAL '1 month' - INTERVAL '1 day')::date
                  AND c.fecha_fin    >= m::date
            ) AS activos
        FROM generate_series(
            date_trunc('month', NOW() - INTERVAL '11 months'),
            date_trunc('month', NOW()),
            '1 month'::interval
        ) AS m
        ORDER BY m
    """)
    rows = cur.fetchall()
    conn.close()
    df = pd.DataFrame(rows, columns=['mes', 'activos'])
    df['pct'] = (df['activos'] / n_db * 100).round(1).clip(upper=100)
    return df



def build_vencimientos_section(df, renov=None):
    """Construye la seccion HTML + JS de Vencimientos de Contratos."""
    today = pd.to_datetime(DATE_STR, dayfirst=True)

    urg  = df[df.dias.between(0, 30)]
    prox = df[df.dias.between(31, 60)]
    sig  = df[df.dias.between(61, 90)]
    trim = df[df.dias.between(0, 90)]
    n_urg, n_prox, n_sig, n_trim = len(urg), len(prox), len(sig), len(trim)

    proj30     = urg.groupby("proyecto").size().sort_values(ascending=False)
    proj30_js  = json.dumps(proj30.index.tolist(), ensure_ascii=False)
    proj30_val = json.dumps(proj30.values.tolist())

    meses      = pd.period_range(today.to_period("M"), periods=12, freq="M")
    mes_counts = [int((df["mes_venc"] == str(m)).sum()) for m in meses]
    mes_labels = json.dumps([str(m) for m in meses])
    mes_vals   = json.dumps(mes_counts)

    # Ejecutivo breakdown
    ejec_urg  = urg.groupby("ejecutivo").size().rename("urg")
    ejec_prox = prox.groupby("ejecutivo").size().rename("prox")
    ejec_df   = pd.concat([ejec_urg, ejec_prox], axis=1).fillna(0).astype(int)
    ejec_df["total"] = ejec_df["urg"] + ejec_df["prox"]
    ejec_df   = ejec_df.sort_values("total", ascending=False)
    ejec_names     = json.dumps(ejec_df.index.tolist(), ensure_ascii=False)
    ejec_urg_vals  = json.dumps(ejec_df["urg"].tolist())
    ejec_prox_vals = json.dumps(ejec_df["prox"].tolist())

    # Tabla detalle 60 dias
    det = df[df.dias.between(0, 60)].sort_values("dias")[
        ["proyecto","unidad","tipologia","fecha_fin","dias","ejecutivo"]
    ].copy()
    det["fecha_fin"] = det["fecha_fin"].dt.strftime("%d/%m/%Y")
    n_det = len(det)

    proyectos_opts = "".join(
        f'<option value="{p}">{p}</option>' for p in sorted(det["proyecto"].unique()))
    tipos_opts = "".join(
        f'<option value="{t}">{t}</option>' for t in sorted(det["tipologia"].dropna().unique()))
    ejec_opts = "".join(
        f'<option value="{e}">{e}</option>' for e in sorted(det["ejecutivo"].dropna().unique()))

    table_rows = ""
    for _, r in det.iterrows():
        color = "#DC2626" if r["dias"] <= 30 else "#D97706"
        badge = (f'<span style="background:{color}22;color:{color};padding:2px 8px;'
                 f'border-radius:99px;font-size:.72rem;font-weight:600">{int(r["dias"])} dias</span>')
        pa = str(r["proyecto"]).replace('"', "&quot;")
        ea = str(r["ejecutivo"]).replace('"', "&quot;")
        table_rows += (
            f'<tr data-proj="{pa}" data-tipo="{r["tipologia"]}" '
            f'data-ejec="{ea}" data-dias="{int(r["dias"])}">'
            f'<td onclick="showProjModal(this.parentNode.dataset.proj)" '
            f'style="cursor:pointer;color:#00A8B4;text-decoration:underline dotted">{r["proyecto"]}</td>'
            f'<td>{r["unidad"]}</td>'
            f'<td>{r["tipologia"]}</td><td>{r["fecha_fin"]}</td>'
            f'<td>{badge}</td><td style="font-size:.75rem">{r["ejecutivo"]}</td></tr>\n'
        )

    section = f"""
<div id="sec-vencimientos" class="sec">Vencimientos de Contratos</div>
<div class="sec-sub">Contratos vigentes (estado activo) &mdash; Datos al {DATE_STR}</div>

<div class="kg" style="grid-template-columns:repeat(5,1fr);max-width:1050px;margin-bottom:20px">
  <div class="kc" style="border-left:4px solid #DC2626">
    <div class="kl">Vencen en 30 d&iacute;as</div><div class="kv re">{n_urg}</div>
    <div class="ks">Renovaci&oacute;n urgente</div>
  </div>
  <div class="kc" style="border-left:4px solid #D97706">
    <div class="kl">Vencen 31&ndash;60 d&iacute;as</div><div class="kv or">{n_prox}</div>
    <div class="ks">Gestionar esta semana</div>
  </div>
  <div class="kc" style="border-left:4px solid #00A8B4">
    <div class="kl">Vencen 61&ndash;90 d&iacute;as</div><div class="kv ac">{n_sig}</div>
    <div class="ks">Planificar contacto</div>
  </div>
  <div class="kc">
    <div class="kl">Total pr&oacute;x. 30 d&iacute;as</div><div class="kv re">{n_urg}</div>
    <div class="ks">En {proj30.shape[0]} proyectos</div>
  </div>
  <div class="kc" style="border-left:4px solid #16A34A">
    <div class="kl">Renovados &uacute;lt. 30 d&iacute;as</div>
    <div class="kv" style="color:#16A34A">{(renov or {{}}).get('r30', '—')}</div>
    <div class="ks">{(renov or {{}}).get('p30', '—')} proyectos &middot; 90d: {(renov or {{}}).get('r90', '—')}</div>
  </div>
</div>

<div class="cr" style="grid-template-columns:1fr 1fr">
  <div class="cc"><div id="venc_bar_proj" style="height:340px"></div></div>
  <div class="cc"><div id="venc_timeline" style="height:340px"></div></div>
</div>

<div class="cf" style="margin-top:16px">
  <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;flex-wrap:wrap;gap:8px;">
    <b style="font-size:.82rem">Detalle &mdash; contratos que vencen en los pr&oacute;ximos 60 d&iacute;as
       (<span id="vf-count-label">{n_det}</span> unidades)</b>
    <button onclick="exportVencCSV()" style="padding:5px 14px;background:#00A8B4;color:#fff;border:none;border-radius:6px;font-size:.73rem;cursor:pointer;">Exportar CSV</button>
  </div>
  <div style="display:flex;flex-wrap:wrap;gap:8px;margin-bottom:12px;align-items:center;
              padding:10px 12px;background:#F8FAFC;border-radius:8px;border:1px solid #E2E8F0">
    <input id="vf-search" type="text" placeholder="Buscar unidad o ejecutivo..."
           oninput="filtrarVenc()"
           style="padding:5px 10px;border:1px solid #E2E8F0;border-radius:6px;font-size:.75rem;
                  min-width:200px;background:#fff;color:#1E2A38;outline:none;flex:1">
    <select id="vf-proj" onchange="filtrarVenc()"
            style="padding:5px 10px;border:1px solid #E2E8F0;border-radius:6px;font-size:.75rem;background:#fff;color:#1E2A38;outline:none">
      <option value="">Todos los proyectos</option>
      {proyectos_opts}
    </select>
    <select id="vf-tipo" onchange="filtrarVenc()"
            style="padding:5px 10px;border:1px solid #E2E8F0;border-radius:6px;font-size:.75rem;background:#fff;color:#1E2A38;outline:none">
      <option value="">Todas las tipolog&iacute;as</option>
      {tipos_opts}
    </select>
    <select id="vf-ejec" onchange="filtrarVenc()"
            style="padding:5px 10px;border:1px solid #E2E8F0;border-radius:6px;font-size:.75rem;background:#fff;color:#1E2A38;outline:none">
      <option value="">Todos los ejecutivos</option>
      {ejec_opts}
    </select>
    <select id="vf-urgencia" onchange="filtrarVenc()"
            style="padding:5px 10px;border:1px solid #E2E8F0;border-radius:6px;font-size:.75rem;background:#fff;color:#1E2A38;outline:none">
      <option value="">Todos los plazos</option>
      <option value="30">Urgente (&le;30 d&iacute;as)</option>
      <option value="60">Pr&oacute;ximo (31&ndash;60 d&iacute;as)</option>
    </select>
    <button onclick="limpiarFiltrosVenc()"
            style="padding:5px 10px;border:1px solid #E2E8F0;border-radius:6px;font-size:.72rem;background:#fff;color:#6B7A8D;cursor:pointer">&#10005; Limpiar</button>
    <span id="vf-count" style="font-size:.72rem;color:#6B7A8D"></span>
  </div>
  <div style="max-height:420px;overflow-y:auto;border-radius:8px;">
    <table id="venc-table">
      <thead><tr><th>Proyecto</th><th>Unidad</th><th>Tipolog&iacute;a</th><th>Vence</th><th>D&iacute;as rest.</th><th>Ejecutivo</th></tr></thead>
      <tbody>{table_rows}</tbody>
    </table>
  </div>
</div>
"""

    js = f"""
// ── Vencimientos charts ──────────────────────────────────────────────────────
var venc_projs      = {proj30_js};
var venc_vals90     = {proj30_val};
var venc_meses      = {mes_labels};
var venc_mes_v      = {mes_vals};
var venc_ejec_names = {ejec_names};
var venc_ejec_urg   = {ejec_urg_vals};
var venc_ejec_prox  = {ejec_prox_vals};

Plotly.newPlot("venc_bar_proj",[
  {{type:"bar",orientation:"h",
   y:venc_projs.slice().reverse(),x:venc_vals90.slice().reverse(),
   marker:{{color:venc_vals90.slice().reverse().map(function(v){{return v>=15?"#DC2626":v>=8?"#D97706":"#00A8B4";}})}},
   text:venc_vals90.slice().reverse().map(String),textposition:"outside",
   hovertemplate:"%{{y}}: %{{x}} contratos<extra></extra>"}}
],{{...base,
  title:{{text:"Vencimientos 30 días por Proyecto",font:{{size:13}}}},
  xaxis:{{gridcolor:GRID,color:AXIS,title:"Contratos"}},yaxis:{{gridcolor:GRID,color:AXIS}},
  margin:{{t:50,b:60,l:200,r:80}},height:340
}});

Plotly.newPlot("venc_timeline",[
  {{type:"bar",x:venc_meses,y:venc_mes_v,
   marker:{{color:venc_mes_v.map(function(v,i){{return i<1?"#DC2626":i<2?"#D97706":"#00A8B4";}})}},
   hovertemplate:"Mes %{{x}}: %{{y}} contratos<extra></extra>"}}
],{{...base,
  title:{{text:"Contratos que vencen por Mes (proximos 12 meses)",font:{{size:13}}}},
  xaxis:{{gridcolor:GRID,color:AXIS,tickangle:-40}},yaxis:{{gridcolor:GRID,color:AXIS,title:"Contratos"}},
  margin:{{t:50,b:120,l:60,r:20}},height:340
}});
</script>
"""
    return section, js








# ── Tipología grouping ────────────────────────────────────────────────────────
TIP_GROUPS = ['Studio', '1 Dorm', '2 Dorm', '3 Dorm']

def _tip_group(tip):
    """Clasificación por tipología (sin modelo)."""
    t = str(tip).upper().strip()
    if re.match(r'^(ST\b|STUD|0D)', t): return 'Studio'
    if re.match(r'^1D', t): return '1 Dorm'
    if re.match(r'^2D', t): return '2 Dorm'
    if re.match(r'^3D', t): return '3 Dorm'
    return '1 Dorm'   # fallback conservador


def _tip_group_full(tip, modelo):
    """
    Clasificación considerando también el campo modelo.
    En la BD los Studios tienen tipologia=1D1B pero modelo='Studio'/'Estudio A'/etc.
    """
    m = str(modelo or '').lower().strip()
    if 'studio' in m or 'estudio' in m:
        return 'Studio'
    return _tip_group(tip)


def load_vacancia():
    """
    Días de Vacancia por unidad Departamento disponible (estado=100).
    Vacancia = días transcurridos desde que terminó el último contrato pasado.
    Unidades sin historial de contratos se marcan con dias_vacancia=None.
    """
    sql = """
    WITH ultimo_contrato_pasado AS (
        -- Usa fecha_termino_real cuando existe (salida anticipada del inquilino),
        -- si no, usa fecha_fin (vencimiento contractual normal).
        SELECT DISTINCT ON (unidad_id)
            unidad_id,
            COALESCE(fecha_termino_real, fecha_fin) AS salida_real,
            fecha_fin                               AS fecha_fin_contrato
        FROM public.contratos
        WHERE COALESCE(fecha_termino_real, fecha_fin) < CURRENT_DATE
        ORDER BY unidad_id, COALESCE(fecha_termino_real, fecha_fin) DESC
    )
    SELECT
        p.nombre                                    AS proyecto,
        u.tipologia,
        u.raw->>'modelo'                            AS modelo,
        u.nombre                                    AS unidad,
        u.raw->>'sub_estado'                        AS sub_estado,
        ucp.salida_real                             AS ultima_salida,
        ucp.fecha_fin_contrato,
        (CURRENT_DATE - ucp.salida_real)::int       AS dias_vacancia
    FROM public.unidades u
    JOIN public.propiedades p ON p.id = u.propiedad_id
    LEFT JOIN ultimo_contrato_pasado ucp ON ucp.unidad_id = u.id
    WHERE u.nombre ILIKE '%%DEPA%%'
      AND u.raw->>'estado' = '100'
    ORDER BY dias_vacancia DESC NULLS LAST
    """
    try:
        conn = psycopg2.connect(**DB)
        with conn.cursor() as cur:
            cur.execute(sql)
            cols = [d[0] for d in cur.description]
            rows = [dict(zip(cols, r)) for r in cur.fetchall()]
        conn.close()
        print(f"  Vacancia: {len(rows)} unidades disponibles consultadas")
        return rows
    except Exception as e:
        print(f"  vacancia: omitido ({e})")
        return []


def build_vacancia_section(vacancia_rows):
    """
    Sección de Días de Vacancia:
      1. Tabla resumen proyecto × tipología (con colores semáforo)
      2. Selector de proyecto → tabla detalle de cada unidad con drill-down
    Solo Departamentos en BBDD.
    """
    if not vacancia_rows:
        return ""

    GRUPOS = ['Studio', '1 Dorm', '2 Dorm', '3 Dorm']

    def _tip(tip, modelo):
        return _tip_group_full(tip, modelo or '')

    def _color(d):
        if d is None: return ('#8896A6', '#F4F6FA')
        if d < 30:    return ('#16A34A', '#F0FDF4')
        if d < 60:    return ('#D97706', '#FFFBEB')
        if d < 90:    return ('#EA580C', '#FFF7ED')
        return            ('#DC2626', '#FEF2F2')

    # ── Procesar filas ─────────────────────────────────────────────────────
    from collections import defaultdict
    proj_grp  = defaultdict(lambda: defaultdict(list))   # solo unidades libres
    no_hist   = defaultdict(lambda: defaultdict(int))
    unit_data = defaultdict(list)   # para JS drill-down (todas)

    proyectos = []
    for r in vacancia_rows:
        proj     = r['proyecto']
        grp      = _tip(r['tipologia'], r['modelo'])
        dias     = r['dias_vacancia']
        mod      = (r['modelo'] or '').strip()
        unidad   = r['unidad']
        sub      = (r['sub_estado'] or '').strip()
        reservada = (sub == '800')
        ult      = str(r['ultima_salida']) if r['ultima_salida'] else None
        fin_c    = str(r['fecha_fin_contrato']) if r.get('fecha_fin_contrato') else None

        if proj not in proyectos:
            proyectos.append(proj)

        entry = {'u': unidad, 'g': grp, 'm': mod,
                 'd': int(dias) if dias is not None else None,
                 'f': ult, 'fc': fin_c, 'r': reservada}
        unit_data[proj].append(entry)

        if dias is None:
            no_hist[proj][grp] += 1
        elif not reservada:
            # Reservadas se excluyen de promedios — ya hay gestión activa
            proj_grp[proj][grp].append(int(dias))

    proyectos = sorted(proyectos)

    # ── TABLA RESUMEN (cada celda cliqueable: proyecto + tipología) ───────
    _td_style = ('padding:10px 16px;border-bottom:1px solid #F4F6FA;'
                 'cursor:pointer;transition:opacity .15s;')

    thead_mat = ('<thead><tr>'
                 '<th style="text-align:left;min-width:170px">Proyecto</th>')
    for g in GRUPOS:
        thead_mat += f'<th style="text-align:center">{g}</th>'
    thead_mat += '<th style="text-align:center">Prom.</th></tr></thead>'

    tbody_mat = '<tbody>'
    for proj in proyectos:
        row_days = []
        cells    = ''
        proj_js  = proj.replace("'", "\\'")
        for g in GRUPOS:
            vals = proj_grp[proj].get(g, [])
            nh   = no_hist[proj].get(g, 0)
            g_js = g.replace("'", "\\'")
            if vals:
                avg = round(sum(vals) / len(vals))
                row_days.append(avg)
                fg, bg = _color(avg)
                n_lbl = f'<div style="font-size:.59rem;opacity:.8;margin-top:2px">{len(vals)} ud.</div>'
                cells += (
                    f'<td style="{_td_style}text-align:center;background:{bg};color:{fg};font-weight:700" '
                    f'onclick="vacSetProj(\'{proj_js}\',\'{g_js}\')" '
                    f'title="Ver {g} — {proj}">'
                    f'{avg}d{n_lbl}</td>'
                )
            elif nh:
                cells += (
                    f'<td style="{_td_style}text-align:center;color:#8896A6;font-size:.78rem" '
                    f'onclick="vacSetProj(\'{proj_js}\',\'{g_js}\')" '
                    f'title="Ver {g} sin historial — {proj}">'
                    f'{nh} s/h</td>'
                )
            else:
                cells += f'<td style="text-align:center;color:#CBD5E1;padding:10px 16px">—</td>'

        proj_avg = round(sum(row_days) / len(row_days)) if row_days else None
        fg_p, bg_p = _color(proj_avg)
        avg_cell = (
            f'<td style="{_td_style}text-align:center;font-weight:800;color:{fg_p};background:{bg_p}" '
            f'onclick="vacSetProj(\'{proj_js}\')" title="Ver todo {proj}">{proj_avg}d</td>'
            if proj_avg is not None
            else '<td style="text-align:center;color:#CBD5E1;padding:10px 16px">—</td>'
        )
        # La celda del nombre del proyecto filtra el proyecto completo (sin grp)
        proj_cell = (
            f'<td style="{_td_style}font-weight:600" '
            f'onclick="vacSetProj(\'{proj_js}\')" title="Ver todo {proj}">{proj}</td>'
        )
        tbody_mat += f'<tr onmouseenter="vacRowHover(this,true)" onmouseleave="vacRowHover(this,false)">{proj_cell}{cells}{avg_cell}</tr>\n'
    tbody_mat += '</tbody>'

    leyenda = ''.join([
        f'<span style="display:inline-flex;align-items:center;gap:5px;margin-right:12px;font-size:.71rem">'
        f'<span style="width:9px;height:9px;border-radius:2px;background:{bg};display:inline-block"></span>'
        f'<span style="color:#4B5A6A">{lbl}</span></span>'
        for bg, _, lbl in [
            ('#F0FDF4','','< 30d'),('#FFFBEB','','30–59d'),
            ('#FFF7ED','','60–89d'),('#FEF2F2','','≥ 90d'),
        ]
    ])

    total_no_hist = sum(v for pr in no_hist.values() for v in pr.values())

    # ── JSON para drill-down JS ───────────────────────────────────────────
    # Orden: libres primero (peores arriba), luego reservadas, luego sin historial
    def _sort_key(x):
        if x['d'] is None: return (2, 0)
        if x['r']:         return (1, -(x['d']))
        return                    (0, -(x['d']))

    vac_json = json.dumps(
        {proj: sorted(units, key=_sort_key)
         for proj, units in unit_data.items()},
        ensure_ascii=False
    )
    proj_opts = ''.join(f'<option value="{p}">{p}</option>' for p in proyectos)

    section = f"""
<div id="sec-vacancia" class="sec">D&iacute;as de Vacancia &mdash; Departamentos</div>
<div class="sec-sub">
  D&iacute;as transcurridos desde el fin del &uacute;ltimo contrato por unidad disponible &mdash;
  solo proyectos en BBDD &mdash; excluye Collective Bustamante
</div>

<!-- ── Resumen matriz ── -->
<div class="cf" style="overflow-x:auto;margin-bottom:20px">
  <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:14px;flex-wrap:wrap;gap:8px">
    <div style="font-size:.78rem;font-weight:700;color:#4B5A6A">
      Resumen por proyecto &amp; tipolog&iacute;a
      <span style="font-weight:400;color:#8896A6;margin-left:8px">
        &#128204; Clic en fila para ver detalle de unidades
      </span>
    </div>
    <div>{leyenda}</div>
  </div>
  <table id="vac-table" style="min-width:640px">
    {thead_mat}
    {tbody_mat}
  </table>
  <div style="font-size:.72rem;color:#8896A6;margin-top:10px">
    <b style="color:#4B5A6A">{total_no_hist}</b>
    unidades sin historial de contratos &mdash; marcadas <em>s/h</em>, no incluidas en promedios.
  </div>
</div>

<!-- ── Drill-down por proyecto ── -->
<div class="cf" id="vac-detail-card" style="margin-bottom:20px">
  <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px;margin-bottom:16px">
    <div>
      <div style="font-size:.68rem;font-weight:700;color:#8896A6;text-transform:uppercase;letter-spacing:.07em;margin-bottom:4px">Detalle por proyecto</div>
      <div id="vac-proj-title" style="font-size:1.1rem;font-weight:800;color:#0F172A">
        Todos los proyectos
      </div>
    </div>
    <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">
      <select id="vac-proj-sel" onchange="vacSetProj(this.value||null,null)"
        style="background:#F8FAFC;border:1.5px solid #E2E8F0;border-radius:10px;
               padding:8px 12px;font-size:.82rem;font-family:inherit;color:#1A2332;
               cursor:pointer;min-width:200px">
        <option value="">— Todos los proyectos —</option>
        {proj_opts}
      </select>
      <span id="vac-grp-chip"
        style="display:none;align-items:center;gap:6px;background:#E0F7FA;color:#00838F;
               border-radius:8px;padding:5px 10px 5px 12px;font-size:.74rem;font-weight:700">
        <span></span>
        <button onclick="vacSetProj(_vacSelProj,null)"
          style="background:none;border:none;cursor:pointer;color:#00838F;font-size:.85rem;
                 line-height:1;padding:0 2px" title="Quitar filtro tipología">&#x2715;</button>
      </span>
      <span id="vac-count-badge"
        style="background:#F4F6FA;color:#4B5A6A;border-radius:8px;
               padding:6px 12px;font-size:.75rem;font-weight:700;white-space:nowrap">
      </span>
    </div>
  </div>
  <div style="overflow-x:auto">
    <table id="vac-unit-table">
      <thead>
        <tr>
          <th style="text-align:left">Unidad</th>
          <th style="text-align:left">Tipolog&iacute;a</th>
          <th style="text-align:left">Modelo</th>
          <th style="text-align:left">Proyecto</th>
          <th style="text-align:center">&Uacute;ltimo contrato</th>
          <th style="text-align:center;min-width:100px">D&iacute;as vacante</th>
        </tr>
      </thead>
      <tbody id="vac-unit-body"></tbody>
    </table>
  </div>
  <div id="vac-no-hist-note"
    style="display:none;font-size:.72rem;color:#8896A6;margin-top:10px;padding-top:10px;
           border-top:1px solid #F4F6FA">
  </div>
</div>

<script>
var _vacData = {vac_json};
var _vacSelProj = null;
var _vacSelGrp  = null;

function _vacColor(d) {{
  if(d===null||d===undefined) return ['#8896A6','#F4F6FA'];
  if(d<30)  return ['#16A34A','#F0FDF4'];
  if(d<60)  return ['#D97706','#FFFBEB'];
  if(d<90)  return ['#EA580C','#FFF7ED'];
  return ['#DC2626','#FEF2F2'];
}}

function vacRowHover(row, on) {{
  row.querySelectorAll('td[onclick]').forEach(function(td){{
    td.style.opacity = on ? '.75' : '1';
  }});
}}

function vacSetProj(proj, grp) {{
  _vacSelProj = proj || null;
  _vacSelGrp  = grp  || null;

  // Sync selector
  var sel = document.getElementById('vac-proj-sel');
  if(sel) sel.value = proj || '';

  // Breadcrumb title
  var title = document.getElementById('vac-proj-title');
  if(title) {{
    title.innerHTML = proj
      ? (grp
          ? '<span style="color:#8896A6;font-weight:400">' + proj + '</span>'
            + ' <span style="color:#CBD5E1;margin:0 6px">›</span> '
            + '<span style="color:#00A8B4">' + grp + '</span>'
          : proj)
      : 'Todos los proyectos';
  }}

  // Chip de filtro activo
  var chip = document.getElementById('vac-grp-chip');
  if(chip) {{
    if(grp) {{
      chip.style.display = 'inline-flex';
      chip.querySelector('span').textContent = grp;
    }} else {{
      chip.style.display = 'none';
    }}
  }}

  // Resaltar celda/fila activa en matriz
  document.querySelectorAll('#vac-table tbody tr').forEach(function(r) {{
    var firstTd = r.querySelector('td');
    var isProj  = proj && firstTd && firstTd.textContent.trim() === proj;
    r.style.background = isProj ? 'rgba(0,168,180,.05)' : '';
  }});

  // Construir lista de unidades
  var units = [];
  if(proj) {{
    units = (_vacData[proj] || []).map(function(u){{ return Object.assign({{proj:proj}},u); }});
  }} else {{
    Object.keys(_vacData).sort().forEach(function(p) {{
      (_vacData[p]||[]).forEach(function(u) {{ units.push(Object.assign({{proj:p}},u)); }});
    }});
  }}

  // Filtrar por tipología si se seleccionó
  if(grp) units = units.filter(function(u){{ return u.g === grp; }});

  var withHist    = units.filter(function(u){{return u.d!==null;}});
  var withoutHist = units.filter(function(u){{return u.d===null;}});

  var libres = withHist.filter(function(u){{return !u.r;}});
  var reserv = withHist.filter(function(u){{return u.r;}});
  var badge = document.getElementById('vac-count-badge');
  if(badge) badge.textContent =
    libres.length + ' libre'+(libres.length!==1?'s':'')+
    (reserv.length ? ' · '+reserv.length+' reservada'+(reserv.length!==1?'s':'') : '')+
    (withoutHist.length ? ' · '+withoutHist.length+' s/h' : '');

  var tbody = document.getElementById('vac-unit-body');
  if(!tbody) return;
  tbody.innerHTML = '';

  withHist.forEach(function(u) {{
    var clrs = _vacColor(u.r ? null : u.d);   // reservadas → gris
    var reservBadge = u.r
      ? '<span style="font-size:.66rem;font-weight:700;color:#2563EB;background:#EFF6FF;'
        +'border-radius:5px;padding:2px 7px;margin-left:5px">Reservada</span>'
      : '';
    // Mostrar fecha real de salida; si difiere del fin contractual, indicar ambas
    var fechaLabel = u.f || '—';
    if(u.fc && u.f && u.fc !== u.f)
      fechaLabel = u.f
        +'<div style="font-size:.65rem;color:#CBD5E1">Fin contrato: '+u.fc+'</div>';
    var tr = document.createElement('tr');
    tr.innerHTML =
      '<td style="font-weight:600;font-family:monospace;font-size:.82rem">'
        +u.u+reservBadge+'</td>'+
      '<td>'+u.g+'</td>'+
      '<td style="color:#8896A6;font-size:.8rem">'+(u.m||'—')+'</td>'+
      (proj?'':'<td>'+(u.proj||'')+'</td>')+
      '<td style="text-align:center;color:#8896A6;font-size:.8rem">'+fechaLabel+'</td>'+
      '<td style="text-align:center">'
        +(u.r
          ? '<span style="font-size:.8rem;color:#8896A6">'+u.d+'d</span>'
          : '<span style="display:inline-block;padding:3px 10px;border-radius:7px;'
            +'font-weight:700;font-size:.85rem;color:'+clrs[0]+';background:'+clrs[1]+'">'
            +u.d+'d</span>')
        +'</td>';
    tbody.appendChild(tr);
  }});

  // Sin historial al final
  var note = document.getElementById('vac-no-hist-note');
  if(withoutHist.length) {{
    withoutHist.forEach(function(u) {{
      var tr = document.createElement('tr');
      tr.style.opacity = '.6';
      tr.innerHTML =
        '<td style="font-weight:600;font-family:monospace;font-size:.82rem">'+u.u+'</td>'+
        '<td>'+u.g+'</td>'+
        '<td style="color:#8896A6;font-size:.8rem">'+(u.m||'—')+'</td>'+
        (proj?'':'<td>'+(u.proj||'')+'</td>')+
        '<td style="text-align:center;color:#8896A6;font-size:.8rem">Sin historial</td>'+
        '<td style="text-align:center"><span style="color:#8896A6;font-size:.8rem">s/h</span></td>';
      tbody.appendChild(tr);
    }});
    if(note) {{
      note.style.display='block';
      note.textContent = withoutHist.length+' unidad'+(withoutHist.length!==1?'es':'')+
        ' nunca arrendadas — sin contrato previo en el sistema.';
    }}
  }} else {{
    if(note) note.style.display='none';
  }}

  // Columna Proyecto: visible solo cuando se ven todos los proyectos
  var thRow = document.querySelector('#vac-unit-table thead tr');
  if(thRow) {{
    var ths = thRow.querySelectorAll('th');
    var hasProjCol = ths.length >= 6;
    if(proj && hasProjCol)      thRow.removeChild(ths[3]);
    else if(!proj && !hasProjCol) {{
      var th = document.createElement('th');
      th.textContent = 'Proyecto'; th.style.textAlign = 'left';
      thRow.insertBefore(th, ths[3]);
    }}
  }}

  // Scroll automático al panel de detalle
  var card = document.getElementById('vac-detail-card');
  if(card && (proj || grp)) {{
    setTimeout(function(){{
      card.scrollIntoView({{behavior:'smooth', block:'start'}});
    }}, 60);
  }}
}}

// Inicializar con todos los proyectos
vacSetProj(null, null);
</script>
"""
    return section


def build_disponibilidad_table(df, m):
    """
    Tabla de Disponibilidad y Ocupación por proyecto y grupo tipológico
    (Studio / 1 Dorm / 2 Dorm / 3 Dorm).
    Base Arrendable = Total − No Disponible.
    % Ocupación     = Arrendados / Base Arrendable.
    Exportable a Google Sheets vía CSV con BOM UTF-8.
    """
    date_fn = DATE_STR.replace('/', '-')

    # ── Compute per-project ────────────────────────────────────────────────
    rows_data = []
    for proj, g in df.groupby('_proj'):
        total   = len(g)
        no_disp = int((g['_estado'] == 'No Disponible').sum())
        base    = total - no_disp
        arr     = int(((g['_estado'] == 'Arrendado') | (g['_sub'] == 'por arrendar')).sum())
        pct     = round(arr / total * 100, 1) if total else 0.0

        disp_g       = g[(g['_estado'] == 'Disponible') & (g['_sub'] != 'por arrendar')].copy()
        disp_g['_grp'] = disp_g.apply(
            lambda r: _tip_group_full(r['_tip'], r.get('Modelo', '')), axis=1)
        grp_counts   = disp_g.groupby('_grp').size()

        rows_data.append({
            'proj':       proj,
            'Studio':     int(grp_counts.get('Studio', 0)),
            '1 Dorm':     int(grp_counts.get('1 Dorm', 0)),
            '2 Dorm':     int(grp_counts.get('2 Dorm', 0)),
            '3 Dorm':     int(grp_counts.get('3 Dorm', 0)),
            'total_disp': int(((g['_estado'] == 'Disponible') & (g['_sub'] != 'por arrendar')).sum()),
            'arr':        arr,
            'base':       base,
            'total':      total,
            'pct':        pct,
        })

    # Sort: lowest occupancy first (most critical at top)
    rows_data.sort(key=lambda x: x['pct'])

    # ── Max values for color intensity ─────────────────────────────────────
    max_vals = {grp: max((r[grp] for r in rows_data), default=1) or 1 for grp in TIP_GROUPS}
    max_disp = max((r['total_disp'] for r in rows_data), default=1) or 1

    def _red_bg(val, mx):
        """White → light red gradient based on val/max ratio."""
        if not val:
            return ''
        alpha = min(val / mx, 1.0)
        r_ = int(255 - alpha * (255 - 254))
        g_ = int(255 - alpha * (255 - 202))
        b_ = int(255 - alpha * (255 - 202))
        return f'background:rgb({r_},{g_},{b_})'

    def _pct_clr(pct):
        return '#16A34A' if pct >= 95 else '#D97706' if pct >= 85 else '#DC2626'

    def _pct_bg(pct):
        return '#F0FDF4' if pct >= 95 else '#FFFBEB' if pct >= 85 else '#FEF2F2'

    # ── Totals ─────────────────────────────────────────────────────────────
    t = {
        'Studio':     sum(r['Studio'] for r in rows_data),
        '1 Dorm':     sum(r['1 Dorm'] for r in rows_data),
        '2 Dorm':     sum(r['2 Dorm'] for r in rows_data),
        '3 Dorm':     sum(r['3 Dorm'] for r in rows_data),
        'total_disp': sum(r['total_disp'] for r in rows_data),
        'arr':        sum(r['arr'] for r in rows_data),
        'base':       sum(r['base'] for r in rows_data),
    }
    t['total'] = sum(r['total'] for r in rows_data)
    t['pct']   = round(t['arr'] / t['total'] * 100, 1) if t['total'] else 0.0

    # ── Build table rows HTML ──────────────────────────────────────────────
    table_rows_html = ''
    for r in rows_data:
        clr = _pct_clr(r['pct']); bg = _pct_bg(r['pct'])
        cells = ''
        for grp in TIP_GROUPS:
            v  = r[grp]
            rb = _red_bg(v, max_vals[grp])
            cells += f'<td style="text-align:center;{rb}">{v if v else "&#x2014;"}</td>'
        # Total disp
        rb_td = _red_bg(r['total_disp'], max_disp)
        cells += f'<td style="text-align:center;font-weight:600;{rb_td}">{r["total_disp"] if r["total_disp"] else "&#x2014;"}</td>'
        cells += f'<td style="text-align:center">{r["arr"]:,}</td>'
        cells += f'<td style="text-align:center">{r["total"]:,}</td>'
        cells += (f'<td style="text-align:center;font-weight:700;color:{clr};background:{bg}">'
                  f'{r["pct"]}%</td>')
        table_rows_html += f'<tr><td style="font-weight:500">{r["proj"]}</td>{cells}</tr>\n'

    # Totals row
    tc = ''; tp_clr = _pct_clr(t['pct']); tp_bg = _pct_bg(t['pct'])
    for grp in TIP_GROUPS:
        tc += f'<td style="text-align:center;font-weight:700;color:#00A8B4">{t[grp]}</td>'
    tc += f'<td style="text-align:center;font-weight:700;color:#00A8B4">{t["total_disp"]}</td>'
    tc += f'<td style="text-align:center;font-weight:700">{t["arr"]:,}</td>'
    tc += f'<td style="text-align:center;font-weight:700">{t["total"]:,}</td>'
    tc += (f'<td style="text-align:center;font-weight:700;color:{tp_clr};background:{tp_bg}">'
           f'{t["pct"]}%</td>')

    rows_js = json.dumps(rows_data, ensure_ascii=False)

    section = f"""
<div id="sec-disponibilidad" class="sec">Disponibilidad y Ocupaci&oacute;n por Proyecto</div>
<div class="sec-sub">Unidades disponibles agrupadas por tipolog&iacute;a &mdash; Datos al {DATE_STR}</div>

<div style="margin-bottom:14px;display:flex;justify-content:flex-end">
  <button onclick="exportDispGSheets()"
          style="display:flex;align-items:center;gap:6px;padding:7px 16px;background:#0F9D58;
                 color:#fff;border:none;border-radius:8px;cursor:pointer;font-size:.78rem;
                 font-weight:700;box-shadow:0 2px 8px rgba(15,157,88,.25);transition:opacity .15s"
          onmouseover="this.style.opacity='.85'" onmouseout="this.style.opacity='1'">
    &#9315; Exportar Google Sheets
  </button>
</div>

<div class="cf" style="overflow-x:auto;margin-bottom:6px">
  <table id="disp-table" style="min-width:720px">
    <thead>
      <tr style="background:#00A8B4;color:#fff">
        <th style="text-align:left;min-width:180px">Proyecto</th>
        <th style="text-align:center">Studio</th>
        <th style="text-align:center">1 Dorm</th>
        <th style="text-align:center">2 Dorm</th>
        <th style="text-align:center">3 Dorm</th>
        <th style="text-align:center">Total Disp.</th>
        <th style="text-align:center">Arrendados</th>
        <th style="text-align:center">Total</th>
        <th style="text-align:center">% Ocup.</th>
      </tr>
    </thead>
    <tbody>
      {table_rows_html}
      <tr style="border-top:2px solid #00A8B4;background:#F4FAFB">
        <td><b>TOTAL</b></td>
        {tc}
      </tr>
    </tbody>
  </table>
</div>
<div style="font-size:.69rem;color:#9CA3AF;margin-bottom:20px">
  % Ocup. = Arrendados / Total &nbsp;&middot;&nbsp;
  Intensidad roja &rarr; mayor disponibilidad relativa
</div>
"""

    js = f"""
// ── Disponibilidad table ─────────────────────────────────────────────────────
var disp_rows = {rows_js};

function exportDispGSheets() {{
  var hdr = ['Proyecto','Studio','1 Dorm','2 Dorm','3 Dorm',
             'Disponibles Total','Arrendados','Total','% Ocupacion'];
  var csvRows = [hdr];
  disp_rows.forEach(function(r) {{
    csvRows.push([r.proj, r.Studio, r['1 Dorm'], r['2 Dorm'], r['3 Dorm'],
                  r.total_disp, r.arr, r.total, r.pct + '%']);
  }});
  var ts=0,t1=0,t2=0,t3=0,td=0,ta=0,tb=0;
  disp_rows.forEach(function(r){{
    ts+=r.Studio; t1+=r['1 Dorm']; t2+=r['2 Dorm']; t3+=r['3 Dorm'];
    td+=r.total_disp; ta+=r.arr; tb+=r.total;
  }});
  var tp = tb>0 ? (ta/tb*100).toFixed(1)+'%' : '0%';
  csvRows.push(['TOTAL',ts,t1,t2,t3,td,ta,tb,tp]);

  var csv = csvRows.map(function(row) {{
    return row.map(function(c) {{
      var s = String(c).replace(/"/g,'""');
      return (s.indexOf(',')>=0||s.indexOf('"')>=0||s.indexOf('\\n')>=0) ? '"'+s+'"' : s;
    }}).join(',');
  }}).join('\\n');

  var blob = new Blob(['\\uFEFF'+csv], {{type:'text/csv;charset=utf-8'}});
  var url  = URL.createObjectURL(blob);
  var a    = document.createElement('a');
  a.href   = url;
  a.download = 'LAR_Disponibilidad_{date_fn}.csv';
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}}
"""
    return section, js


PROJ_PALETTE = [
    "#00A8B4","#0369A1","#7C3AED","#DB2777","#059669",
    "#D97706","#DC2626","#0891B2","#4F46E5","#BE123C",
    "#15803D","#B45309","#1D4ED8","#9333EA"
]

LOGO_URLS = {
    "spot nueva kennedy":  "https://largroup.cl/wp-content/uploads/2025/03/logo-spot.svg",
    "spot residence":      "https://largroup.cl/wp-content/uploads/2024/12/logo-spot-residence-1.png",
    "nomad bellet":        "https://largroup.cl/wp-content/uploads/2025/03/logo-nomad-bellet.svg",
    "nomad holley":        "https://largroup.cl/wp-content/uploads/2025/03/logo-nomad-holley.svg",
    "collective":          "https://largroup.cl/wp-content/uploads/2025/12/LOGO-CB-WEB.svg",
    "brooklyn":            "https://largroup.cl/wp-content/uploads/2025/03/logo-brooklyn-200x200-1.jpg",
    "nativo":              "https://largroup.cl/wp-content/uploads/2025/03/logo-nativo.png",
    "blend":               "https://largroup.cl/wp-content/uploads/2025/06/LOGO-BLEND-100X100.svg",
    "boldo":               "https://largroup.cl/wp-content/uploads/2025/03/logo-boldo.svg",
    "imu":                 "https://largroup.cl/wp-content/uploads/2025/12/IMU-Positivo-.png",
    "soho":                "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAxMDAgMTAwIiB3aWR0aD0iMzAiIGhlaWdodD0iMzAiPgogIDxyZWN0IHdpZHRoPSIxMDAiIGhlaWdodD0iMTAwIiBmaWxsPSIjMDAwIi8+CiAgPHRleHQgeD0iNTAiIHk9IjUyIiB0ZXh0LWFuY2hvcj0ibWlkZGxlIiBmb250LWZhbWlseT0iQXJpYWwgQmxhY2ssQXJpYWwsc2Fucy1zZXJpZiIKICAgICAgICBmb250LXNpemU9IjMwIiBmb250LXdlaWdodD0iOTAwIiBmaWxsPSJ3aGl0ZSI+U09ITzwvdGV4dD4KICA8dGV4dCB4PSI1MCIgeT0iNjciIHRleHQtYW5jaG9yPSJtaWRkbGUiIGZvbnQtZmFtaWx5PSJBcmlhbCxzYW5zLXNlcmlmIgogICAgICAgIGZvbnQtc2l6ZT0iNy41IiBmb250LXdlaWdodD0iNDAwIiBmaWxsPSJ3aGl0ZSIgbGV0dGVyLXNwYWNpbmc9IjIuNSI+QkFSUklPIElUQUxJQTwvdGV4dD4KPC9zdmc+",
    "park":                "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAxMDAgMTAwIiB3aWR0aD0iMzAiIGhlaWdodD0iMzAiPgogIDxyZWN0IHdpZHRoPSIxMDAiIGhlaWdodD0iMTAwIiBmaWxsPSIjMDAwIi8+CiAgPHRleHQgeD0iNTAiIHk9IjUyIiB0ZXh0LWFuY2hvcj0ibWlkZGxlIiBmb250LWZhbWlseT0iQXJpYWwgQmxhY2ssQXJpYWwsc2Fucy1zZXJpZiIKICAgICAgICBmb250LXNpemU9IjMyIiBmb250LXdlaWdodD0iOTAwIiBmaWxsPSJ3aGl0ZSI+UEFSSzwvdGV4dD4KICA8dGV4dCB4PSI1MCIgeT0iNjciIHRleHQtYW5jaG9yPSJtaWRkbGUiIGZvbnQtZmFtaWx5PSJBcmlhbCxzYW5zLXNlcmlmIgogICAgICAgIGZvbnQtc2l6ZT0iNyIgZm9udC13ZWlnaHQ9IjYwMCIgZmlsbD0id2hpdGUiIGxldHRlci1zcGFjaW5nPSIyIj5TQU5USUFHTyBDRU5UUk88L3RleHQ+Cjwvc3ZnPg==",
    "central":             "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAxMDAgMTAwIiB3aWR0aD0iMzAiIGhlaWdodD0iMzAiPgogIDxyZWN0IHdpZHRoPSIxMDAiIGhlaWdodD0iMTAwIiBmaWxsPSIjMDAwIi8+CiAgPHRleHQgeD0iNTAiIHk9IjUwIiB0ZXh0LWFuY2hvcj0ibWlkZGxlIiBmb250LWZhbWlseT0iQXJpYWwgQmxhY2ssQXJpYWwsc2Fucy1zZXJpZiIKICAgICAgICBmb250LXNpemU9IjIxIiBmb250LXdlaWdodD0iOTAwIiBmaWxsPSJ3aGl0ZSI+Q0VOVFJBTDwvdGV4dD4KICA8dGV4dCB4PSI1MCIgeT0iNjYiIHRleHQtYW5jaG9yPSJtaWRkbGUiIGZvbnQtZmFtaWx5PSJBcmlhbCxzYW5zLXNlcmlmIgogICAgICAgIGZvbnQtc2l6ZT0iNyIgZm9udC13ZWlnaHQ9IjYwMCIgZmlsbD0id2hpdGUiIGxldHRlci1zcGFjaW5nPSIyIj5TQU5USUFHTyBDRU5UUk88L3RleHQ+Cjwvc3ZnPg==",
}





def tipo_data_to_js(td):
    tips = td["tipologias"]
    projs = td["projects"]
    lines = ["var tipo_data = {",
             f'  tipologias: {json.dumps(tips, ensure_ascii=False)},',
             f'  projects: {json.dumps(projs, ensure_ascii=False)},']
    for tip in tips:
        d = td[tip]
        label = tip if re.match(r'^[A-Za-z_]\w*$', tip) else json.dumps(tip)
        lines.append(f'  {label}: {{total:{d["total"]},arrendados:{d["arrendados"]},pct:{d["pct"]}}},')
    lines.append("};")
    return "\n".join(lines)

# ── 4. CONSTRUIR SECCION POR LIBERAR ──────────────────────────────────────

def js_replace(html, var_name, new_value):
    pattern = rf'(var\s+{re.escape(var_name)}\s*=\s*)([^;]+)(;)'
    if not re.search(pattern, html):
        print(f"  SKIP var {var_name} (no encontrada)")
        return html
    return re.sub(pattern, rf'\g<1>{new_value}\g<3>', html)


def update_html(html, m):
    # KPI cards superiores (6 cards cliqueables)
    pct = round(m["ocup_global"], 1)
    old_kg = re.search(r'<div class="kg"[^>]*>.*?</div>\s*</div>', html, re.DOTALL)
    if old_kg:
        new_kg = f"""<div class="kg" style="grid-template-columns:repeat(6,1fr)">
  <div class="kc kc-link" onclick="irA('sec-alertas')" title="Ver Alertas"><div class="kl">Total Unidades</div><div class="kv">{m['total']:,}</div><div class="ks">Departamentos en cartera</div></div>
  <div class="kc kc-link" onclick="irA('sec-alertas')" title="Ver Alertas"><div class="kl">Arrendadas</div><div class="kv gr">{m['arrendadas']:,}</div><div class="ks">Estado: Arrendado</div></div>
  <div class="kc kc-link" onclick="irA('sec-disponibles')" title="Ver Disponibles"><div class="kl">Disponibles</div><div class="kv or">{m['disponibles']:,}</div><div class="ks">Libres para arrendar</div></div>
  <div class="kc kc-link" onclick="irA('sec-pipeline')" title="Ver Pipeline"><div class="kl">No Disponibles</div><div class="kv re">{m['no_disp']:,}</div><div class="ks">Incl. En Obra ({m['en_obra']})</div></div>
  <div class="kc kc-link" onclick="irA('sec-ocupacion')" title="Ver Ocupaci&oacute;n"><div class="kl">% Ocupaci&oacute;n Global</div><div class="kv ac">{pct}%</div><div class="ks">Target: 95% | Gap: {round(pct-95,1)}pp</div></div>
  <div class="kc kc-link" onclick="irA('sec-por-liberar')" title="Ver Por Liberar"><div class="kl">Por Liberar</div><div class="kv re">{m['por_liberar']}</div><div class="ks">{len(m['pol_by_proj'])} proyectos afectados</div></div>
</div>"""
        html = html[:old_kg.start()] + new_kg + html[old_kg.end():]

    # Variables JS numericas
    html = js_replace(html, "gpct", str(round(m["ocup_global"]/100, 4)))
    html = js_replace(html, "cb_pct", str(m["cb_pct"]))
    # imu_pct / blend_pct may not be standalone vars — patch via hv array
    html = re.sub(
        r'var hv\s*=\s*\[[^\]]+\]',
        f'var hv=[{round(m["ocup_global"],2)},{m["cb_pct"]},{m["imu_pct"]},{m["blend_pct"]}]',
        html
    )
    # Ensure imu_pct / blend_pct are declared standalone (after cb_pct line)
    if not re.search(r'var\s+imu_pct\s*=', html):
        html = re.sub(
            r'(var\s+cb_pct\s*=[^;]+;)',
            rf'\1\nvar imu_pct={m["imu_pct"]};\nvar blend_pct={m["blend_pct"]};',
            html, count=1
        )
    else:
        html = js_replace(html, "imu_pct",   str(m["imu_pct"]))
        html = js_replace(html, "blend_pct", str(m["blend_pct"]))

    # Arrays JS principales
    html = js_replace(html, "P", json.dumps(m["proj_desc"], ensure_ascii=False))
    html = js_replace(html, "A", json.dumps(m["proj_asc"],  ensure_ascii=False))
    html = js_replace(html, "proj_desc", json.dumps(m["proj_desc"], ensure_ascii=False))
    html = js_replace(html, "proj_asc",  json.dumps(m["proj_asc"],  ensure_ascii=False))
    html = js_replace(html, "hm_pct",    json.dumps(m["hm_pct"]))
    html = js_replace(html, "hm_proj",   json.dumps(m["projs_hm"],  ensure_ascii=False))
    html = js_replace(html, "hm_tip",    json.dumps(m["tips"],       ensure_ascii=False))
    html = js_replace(html, "pipe_data", json.dumps(m["pipe"],       ensure_ascii=False))
    html = js_replace(html, "risk_data", json.dumps(m["risk"],       ensure_ascii=False))

    # tipo_data (Composición por Tipología chart)
    new_tipo_js = tipo_data_to_js(m["tipo_data"])
    html = re.sub(r'var tipo_data\s*=\s*\{.*?\};', new_tipo_js, html, flags=re.DOTALL)

    # unidades_disp (tabla Disponibles)
    new_disp_js = f'var unidades_disp = {json.dumps(m["unidades_disp"], ensure_ascii=False)}'
    html = re.sub(r'var unidades_disp\s*=\s*\[.*?\]', new_disp_js, html, flags=re.DOTALL)

    # Pipeline KPI cards — actualiza valores y hace Reservadas clickeable
    html = re.sub(r'(<div class="kl">Por Liberar</div><div class="kv re">)\d+(</div><div class="ks">Riesgo)',
                  rf'\g<1>{m["por_liberar"]}\g<2>', html)
    html = re.sub(r'(<div class="kl">Por Renovar</div><div class="kv ac">)\d+(</div>)',
                  rf'\g<1>{m["por_renovar"]}\g<2>', html)
    html = re.sub(r'(<div class="kl">Por Arrendar</div><div class="kv or">)\d+(</div>)',
                  rf'\g<1>{m["por_arrendar"]}\g<2>', html)
    # Reservadas → clickeable
    html = re.sub(
        r'<div class="kc">(<div class="kl">Reservadas</div><div class="kv gr">)\d+(</div>[^<]*<div class="ks">En proceso de firma</div>)</div>',
        f'<div class="kc kc-link" onclick="irA(\'sec-reservadas\')" title="Ver Reservadas">\\g<1>{m["reservadas"]}\\g<2></div>',
        html
    )

    # Alerta de Por Liberar
    projs_pol = list(m["pol_by_proj"].keys())
    mayor_proj = projs_pol[0] if projs_pol else "N/A"
    mayor_n    = list(m["pol_by_proj"].values())[0] if projs_pol else 0
    html = re.sub(
        r'<b>\d+ unidades Por Liberar</b> en \d+ proyectos\. Mayor riesgo: <b>[^<]+</b> con <b>\d+ unidades</b>\.',
        f'<b>{m["por_liberar"]} unidades Por Liberar</b> en {len(projs_pol)} proyectos. Mayor riesgo: <b>{mayor_proj}</b> con <b>{mayor_n} unidades</b>.',
        html
    )

    # Fecha
    html = re.sub(r'Datos al \d{2}/\d{2}/\d{4}', f'Datos al {DATE_STR}', html)

    # IDs para secciones navegables
    html = html.replace('<div class="sec">Ocupación Global</div>',
                        '<div id="sec-ocupacion" class="sec">Ocupación Global</div>')
    html = html.replace('<div class="sec">Alertas por Proyecto</div>',
                        '<div id="sec-alertas" class="sec">Alertas por Proyecto</div>')
    html = html.replace('<div class="sec">Unidades Disponibles por Proyecto</div>',
                        '<div id="sec-disponibles" class="sec">Unidades Disponibles por Proyecto</div>')
    for tag in ['FASE 3 &mdash; Pipeline de Acción', 'FASE 3 &mdash; Pipeline de Acción']:
        html = html.replace(f'<div class="sec">{tag}</div>',
                            f'<div id="sec-pipeline" class="sec">{tag}</div>')

    # CSS hover cards + brand color + collapsible sections
    css = """.kc-link{cursor:pointer;transition:border-color .2s,transform .15s,box-shadow .2s}
.kc-link:hover{border-color:#00A8B4!important;transform:translateY(-3px);box-shadow:0 8px 24px rgba(0,168,180,.22);background:#F0FDFE!important}
.kc-link:hover .kl{color:#00A8B4}
/* Collapsible sections */
.sec{cursor:pointer;user-select:none;}
.sec-chevron{display:inline-block;font-size:.58rem;color:#9CA3AF;margin-left:7px;transition:transform .22s;vertical-align:middle;}
.sec.collapsed .sec-chevron{transform:rotate(180deg);}
@keyframes spin{to{transform:rotate(360deg);}}"""
    html = html.replace('</style>', css + '\n</style>', 1)

    # NOTE: Comparador Semanal id is added in main() AFTER all section insertions
    # so the marker '<div class="sec">Comparador Semanal</div>' stays intact for those insertions.

    return html

# ── 6. MAIN ────────────────────────────────────────────────────────────────

def build_res_section(m):
    res_by_proj = m["res_by_proj"]
    res_by_tipo = m["res_by_tipo"]
    res_matrix  = m["res_matrix"]
    total_res   = m["reservadas"]
    n_proj      = len(res_by_proj)
    tipo_dom    = list(res_by_tipo.keys())[0]  if res_by_tipo  else "N/A"
    tipo_dom_n  = list(res_by_tipo.values())[0] if res_by_tipo  else 0
    tipo_dom_pct= round(tipo_dom_n / total_res * 100) if total_res else 0
    mayor_proj  = list(res_by_proj.keys())[0]  if res_by_proj  else "N/A"
    mayor_n     = list(res_by_proj.values())[0] if res_by_proj  else 0

    tipos_cols = list(res_matrix.columns) if not res_matrix.empty else []

    # Tabla rows
    table_rows = ""
    for proj, total_p in sorted(res_by_proj.items(), key=lambda x: -x[1]):
        cells = ""
        for t in tipos_cols:
            v = int(res_matrix.loc[proj, t]) if proj in res_matrix.index and t in res_matrix.columns else 0
            cells += f'<td style="text-align:center">{v if v else "-"}</td>'
        table_rows += (
            f'<tr><td>{proj}</td>{cells}'
            f'<td style="text-align:center"><b style="color:#48BB78">{total_p}</b></td></tr>'
        )

    tipo_headers = "".join(f'<th style="text-align:center">{t}</th>' for t in tipos_cols)
    total_cells  = "".join(
        f'<td style="text-align:center;color:#00A8B4"><b>{int(res_by_tipo.get(t,0))}</b></td>'
        for t in tipos_cols)

    # JS data
    projs_js  = json.dumps(list(res_by_proj.keys()), ensure_ascii=False)
    totals_js = json.dumps(list(res_by_proj.values()))
    tipo_lbls = json.dumps(list(res_by_tipo.keys()),   ensure_ascii=False)
    tipo_vals = json.dumps(list(res_by_tipo.values()))

    stacked_traces = ""
    colors = ["#48BB78","#68D391","#9AE6B4","#00A8B4","#F6AD55"]
    for i, t in enumerate(tipos_cols):
        vals = [int(res_matrix.loc[p, t]) if p in res_matrix.index and t in res_matrix.columns else 0
                for p in res_by_proj.keys()]
        c = colors[i % len(colors)]
        stacked_traces += f'  {{type:"bar",name:{json.dumps(t)},x:res_projs,y:{vals},marker:{{color:"{c}"}}}},\n'

    section = f"""
<div id="sec-reservadas" class="sec">Reservadas &mdash; Detalle por Proyecto y Tipolog&iacute;a</div>
<div class="sec-sub">Unidades con subestado Reservada en POP Estate &mdash; Datos al {DATE_STR}</div>

<div class="kg" style="grid-template-columns:repeat(1,1fr);max-width:220px">
  <div class="kc"><div class="kl">Total Reservadas</div><div class="kv gr">{total_res}</div><div class="ks">{n_proj} proyectos</div></div>
</div>

<div class="ab" style="border-left-color:#48BB78">&#x1F4CB; <span><b>{mayor_proj}</b> lidera con <b>{mayor_n} unidades</b> reservadas. El <b>{tipo_dom_pct}% son tipolog&iacute;a {tipo_dom}</b>.</span></div>

<div class="cr">
  <div class="cc"><div id="res_bar_proj" style="height:420px"></div></div>
  <div class="cc"><div id="res_donut_tipo" style="height:420px"></div></div>
</div>
<div class="cf"><div id="res_stacked" style="height:380px"></div></div>

<div class="cf" style="margin-top:16px">
  <table>
    <thead>
      <tr><th>Proyecto</th>{tipo_headers}<th style="text-align:center;color:#16A34A"><b>Total</b></th></tr>
    </thead>
    <tbody>
      {table_rows}
      <tr style="border-top:2px solid #E2E8F0">
        <td><b>TOTAL</b></td>{total_cells}
        <td style="text-align:center"><b style="color:#48BB78">{total_res}</b></td>
      </tr>
    </tbody>
  </table>
</div>

"""
    js = f"""
// Reservadas charts
var res_projs  = {projs_js};
var res_totals = {totals_js};
Plotly.newPlot("res_bar_proj",[
  {{type:"bar",orientation:"h",
   y:res_projs.slice().reverse(),
   x:res_totals.slice().reverse(),
   marker:{{color:res_totals.slice().reverse().map(v=>v>=15?"#48BB78":v>=8?"#68D391":"#9AE6B4")}},
   text:res_totals.slice().reverse().map(String),textposition:"outside",
   hovertemplate:"%{{y}}: %{{x}} unidades<extra></extra>"}}
],{{...base,
  title:{{text:"Reservadas por Proyecto",font:{{size:13}}}},
  xaxis:{{gridcolor:GRID,color:AXIS,title:"Unidades",dtick:5}},
  yaxis:{{gridcolor:GRID,color:AXIS}},
  margin:{{t:50,b:60,l:200,r:80}},height:420
}});
Plotly.newPlot("res_donut_tipo",[
  {{type:"pie",hole:0.6,
   values:{tipo_vals},labels:{tipo_lbls},
   marker:{{colors:["#48BB78","#68D391","#9AE6B4","#00A8B4","#F6AD55"]}},
   textinfo:"label+percent",
   hovertemplate:"%{{label}}: %{{value}} uds (%{{percent}})<extra></extra>"}}
],{{...base,
  title:{{text:"Distribuci\\u00f3n por Tipolog\\u00eda",font:{{size:13}}}},
  annotations:[{{x:0.5,y:0.5,showarrow:false,align:"center",
    text:"<b><span style='font-size:24px;color:#16A34A'>{total_res}</span></b><br><span style='font-size:10px;color:#64748B'>total</span>"}}],
  showlegend:true,legend:{{orientation:"h",y:-0.1}},
  margin:{{t:50,b:60,l:20,r:20}},height:420
}});
Plotly.newPlot("res_stacked",[
{stacked_traces}],{{...base,
  title:{{text:"Reservadas: Desglose por Proyecto y Tipolog\\u00eda",font:{{size:13}}}},
  barmode:"stack",
  xaxis:{{gridcolor:GRID,color:AXIS,tickangle:-40}},
  yaxis:{{gridcolor:GRID,color:AXIS,title:"Unidades"}},
  legend:{{orientation:"h",y:-0.3}},
  margin:{{t:50,b:140,l:60,r:20}},height:380
}});
"""
    return section, js


# ── 5. ACTUALIZAR HTML ─────────────────────────────────────────────────────

def build_pol_section(m):
    pol_by_proj = m["pol_by_proj"]
    pol_by_tipo = m["pol_by_tipo"]
    pol_matrix  = m["pol_matrix"]
    total_pol   = m["por_liberar"]
    n_proj      = len(pol_by_proj)
    tipo_dom    = list(pol_by_tipo.keys())[0] if pol_by_tipo else "N/A"
    tipo_dom_n  = list(pol_by_tipo.values())[0] if pol_by_tipo else 0
    tipo_dom_pct= round(tipo_dom_n/total_pol*100) if total_pol else 0
    mayor_proj  = list(pol_by_proj.keys())[0] if pol_by_proj else "N/A"
    mayor_n     = list(pol_by_proj.values())[0] if pol_by_proj else 0

    # Tabla rows
    tipos_cols = list(pol_matrix.columns) if not pol_matrix.empty else []
    table_rows = ""
    for proj, row in sorted(pol_by_proj.items(), key=lambda x: -x[1]):
        cells = ""
        for t in tipos_cols:
            v = int(pol_matrix.loc[proj, t]) if proj in pol_matrix.index and t in pol_matrix.columns else 0
            cells += f'<td style="text-align:center">{v if v else "-"}</td>'
        si_v = pol_by_proj[proj] - sum(
            int(pol_matrix.loc[proj, t]) if proj in pol_matrix.index and t in pol_matrix.columns else 0
            for t in tipos_cols)
        table_rows += f'<tr><td>{proj}</td>{cells}<td style="text-align:center">{"" if si_v==0 else si_v}</td><td style="text-align:center"><b style="color:#FC8181">{pol_by_proj[proj]}</b></td></tr>'

    tipo_headers = "".join(f'<th style="text-align:center">{t}</th>' for t in tipos_cols)
    total_cells  = "".join(
        f'<td style="text-align:center;color:#00A8B4"><b>{int(pol_by_tipo.get(t,0))}</b></td>'
        for t in tipos_cols)
    si_total = total_pol - sum(pol_by_tipo.values())

    # JS data
    projs_js   = json.dumps(list(pol_by_proj.keys()), ensure_ascii=False)
    totals_js  = json.dumps(list(pol_by_proj.values()))
    tipo_lbls  = json.dumps(list(pol_by_tipo.keys()), ensure_ascii=False)
    tipo_vals  = json.dumps(list(pol_by_tipo.values()))

    # Stacked traces per tipologia
    stacked_traces = ""
    colors = ["#FC8181","#ED8936","#F6AD55","#68D391","#00A8B4"]
    for i, t in enumerate(tipos_cols):
        vals = [int(pol_matrix.loc[p, t]) if p in pol_matrix.index and t in pol_matrix.columns else 0
                for p in pol_by_proj.keys()]
        c = colors[i % len(colors)]
        stacked_traces += f'  {{type:"bar",name:{json.dumps(t)},x:pol_projs,y:{vals},marker:{{color:"{c}"}}}},\n'

    section = f"""
<div id="sec-por-liberar" class="sec">Por Liberar &mdash; Detalle por Proyecto y Tipolog&iacute;a</div>
<div class="sec-sub">Unidades arrendadas con subestado Por Liberar en POP Estate &mdash; Datos al {DATE_STR}</div>
<div class="sec-sub" style="color:#747678;font-size:.75rem;margin-top:4px">* Collective Bustamante, Blend Apoquindo y Boldo Club de Campo no registran unidades Por Liberar en POP Estate.</div>

<div class="kg" style="grid-template-columns:repeat(1,1fr);max-width:220px">
  <div class="kc"><div class="kl">Total Por Liberar</div><div class="kv re">{total_pol}</div><div class="ks">{n_proj} de 13 proyectos</div></div>
</div>

<div class="ab">&#x26A0;&#xFE0F; <span><b>{mayor_proj}</b> concentra el mayor riesgo con <b>{mayor_n} unidades</b>. El <b>{tipo_dom_pct}% son tipolog&iacute;a {tipo_dom}</b> &mdash; requieren campa&ntilde;a de reemplazo anticipada.</span></div>

<div class="cr">
  <div class="cc"><div id="pol_bar_proj" style="height:420px"></div></div>
  <div class="cc"><div id="pol_donut_tipo" style="height:420px"></div></div>
</div>
<div class="cf"><div id="pol_stacked" style="height:380px"></div></div>

<div class="cf" style="margin-top:16px">
  <table>
    <thead>
      <tr><th>Proyecto</th>{tipo_headers}<th style="text-align:center">S/I</th><th style="text-align:center;color:#DC2626"><b>Total</b></th></tr>
    </thead>
    <tbody>
      {table_rows}
      <tr style="border-top:2px solid #E2E8F0">
        <td><b>TOTAL</b></td>{total_cells}
        <td style="text-align:center;color:#747678"><b>{"" if si_total==0 else si_total}</b></td>
        <td style="text-align:center"><b style="color:#FC8181">{total_pol}</b></td>
      </tr>
    </tbody>
  </table>
</div>

"""
    js = f"""
// Por Liberar charts
var pol_projs  = {projs_js};
var pol_totals = {totals_js};
Plotly.newPlot("pol_bar_proj",[
  {{type:"bar",orientation:"h",
   y:pol_projs.slice().reverse(),
   x:pol_totals.slice().reverse(),
   marker:{{color:pol_totals.slice().reverse().map(v=>v>=15?"#FC8181":v>=8?"#ED8936":"#F6AD55")}},
   text:pol_totals.slice().reverse().map(String),textposition:"outside",
   hovertemplate:"%{{y}}: %{{x}} unidades<extra></extra>"}}
],{{...base,
  title:{{text:"Por Liberar por Proyecto",font:{{size:13}}}},
  xaxis:{{gridcolor:GRID,color:AXIS,title:"Unidades",dtick:5}},
  yaxis:{{gridcolor:GRID,color:AXIS}},
  margin:{{t:50,b:60,l:200,r:80}},height:420
}});
Plotly.newPlot("pol_donut_tipo",[
  {{type:"pie",hole:0.6,
   values:{tipo_vals},labels:{tipo_lbls},
   marker:{{colors:["#FC8181","#ED8936","#F6AD55","#747678","#00A8B4"]}},
   textinfo:"label+percent",
   hovertemplate:"%{{label}}: %{{value}} uds (%{{percent}})<extra></extra>"}}
],{{...base,
  title:{{text:"Distribuci\\u00f3n por Tipolog\\u00eda",font:{{size:13}}}},
  annotations:[{{x:0.5,y:0.5,showarrow:false,align:"center",
    text:"<b><span style='font-size:24px;color:#DC2626'>{total_pol}</span></b><br><span style='font-size:10px;color:#64748B'>total</span>"}}],
  showlegend:true,legend:{{orientation:"h",y:-0.1}},
  margin:{{t:50,b:60,l:20,r:20}},height:420
}});
Plotly.newPlot("pol_stacked",[
{stacked_traces}],{{...base,
  title:{{text:"Por Liberar: Desglose por Proyecto y Tipolog\\u00eda",font:{{size:13}}}},
  barmode:"stack",
  xaxis:{{gridcolor:GRID,color:AXIS,tickangle:-40}},
  yaxis:{{gridcolor:GRID,color:AXIS,title:"Unidades"}},
  legend:{{orientation:"h",y:-0.3}},
  margin:{{t:50,b:140,l:60,r:20}},height:380
}});
"""
    return section, js



def replace_full_div(html, start_idx, open_end_idx, new_html):
    """Replace a <div> block from start_idx to its matching </div>."""
    depth = 1
    i = open_end_idx
    while i < len(html) and depth > 0:
        if html[i:i+5] == "<div ": depth += 1
        elif html[i:i+6] == "<div>":  depth += 1
        elif html[i:i+6] == "</div>": depth -= 1
        i += 1
    return html[:start_idx] + new_html + html[i:]

def _get_logo_url(proj_name):
    name = proj_name.lower()
    for key in sorted(LOGO_URLS, key=len, reverse=True):
        if key in name:
            return LOGO_URLS[key]
    return None


def _proj_initials(name):
    words = [w for w in name.replace("-", " ").split() if w]
    return (words[0][0] + words[1][0]).upper() if len(words) >= 2 else name[:2].upper()


def _logo_svg(initials, color, size=48):
    cx, cy, fs = size // 2, int(size * 0.64), int(size * 0.33)
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{size}" height="{size}" viewBox="0 0 {size} {size}">'
        f'<rect width="{size}" height="{size}" rx="9" fill="{color}"/>'
        f'<text x="{cx}" y="{cy}" text-anchor="middle" font-family="Arial,sans-serif" '
        f'font-size="{fs}" font-weight="800" fill="white">{initials}</text></svg>'
    )


def _logo_box(init, color, url, size):
    """Img with SVG fallback via data attrs."""
    base = (f'class="proj-logo-box" data-init="{init}" data-color="{color}" '
            f'style="width:{size}px;height:{size}px;flex-shrink:0;border-radius:{max(6,size//5)}px;'
            f'background:#F8FAFC;overflow:hidden;display:flex;align-items:center;justify-content:center"')
    inner_size = size - 4
    if url:
        return (f'<div {base}>'
                f'<img src="{url}" width="{inner_size}" height="{inner_size}" '
                f'style="object-fit:contain" onerror="projLogoFallback(this)"></div>')
    return f'<div style="flex-shrink:0">{_logo_svg(init, color, size)}</div>'


def _fmt_clp(uf_val, uf_valor):
    if not uf_valor:
        return None
    clp = uf_val * uf_valor
    return f"${clp/1_000_000:.2f}M" if clp >= 1_000_000 else f"${clp/1_000:.0f}k"


def build_projects_section(m, vencs=None, precios=None, uf_valor=None, tendencias=None):
    """Vista por Proyecto v3 — sidebar de proyectos + panel de detalle inline."""

    sorted_projs   = sorted(m["proj_desc"], key=lambda x: x["Propiedad"])
    proj_color_map = {p["Propiedad"]: PROJ_PALETTE[i % len(PROJ_PALETTE)]
                      for i, p in enumerate(sorted_projs)}

    # ── Vencimientos 60d + mensual ────────────────────────────────────────────
    venc_by_proj  = {}
    venc_monthly  = {}
    if vencs is not None:
        det60 = vencs[vencs["dias"].between(0, 60)]
        venc_by_proj = det60.groupby("proyecto").size().to_dict()
        ref = pd.Timestamp.now().normalize()
        for p in sorted_projs:
            pname = p["Propiedad"]
            pdata = det60[det60["proyecto"] == pname]
            buckets = {}
            for _, row in pdata.iterrows():
                d = int(row["dias"]) if pd.notna(row["dias"]) else 999
                if d < 0:
                    continue
                mo = (ref + pd.Timedelta(days=int(d))).strftime("%b %Y")
                buckets[mo] = buckets.get(mo, 0) + 1
            if buckets:
                try:
                    items = sorted(buckets.items(),
                                   key=lambda x: pd.to_datetime(x[0], format="%b %Y"))
                except Exception:
                    items = list(buckets.items())
                venc_monthly[pname] = {"labels": [x[0] for x in items],
                                       "values": [x[1] for x in items]}
            else:
                venc_monthly[pname] = {"labels": [], "values": []}

    # ── Tendencias (delta pp mes anterior) ───────────────────────────────────
    proj_trends = {}
    for p in sorted_projs:
        name  = p["Propiedad"]
        t     = (tendencias or {}).get(name, {})
        total = int(p["Total"]) or 1
        if t and t.get("now") is not None:
            delta = round((t["now"] / total - t["prev"] / total) * 100, 1)
        else:
            delta = None
        proj_trends[name] = delta

    n_below_meta = sum(1 for p in sorted_projs if p["Pct_Ocup"] < 0.95)
    n_projs      = len(sorted_projs)

    # ── Sidebar items HTML ────────────────────────────────────────────────────
    sidebar_html = ""
    proj_logos   = {}

    for p in sorted_projs:
        name    = p["Propiedad"]
        color   = proj_color_map[name]
        init    = _proj_initials(name)
        pct     = round(p["Pct_Ocup"] * 100, 1)
        venc_n  = venc_by_proj.get(name, 0)

        logo_url = _get_logo_url(name)
        proj_logos[name] = logo_url or ""
        logo_html_sb = _logo_box(init, color, logo_url, 34)

        # Status dot
        is_crit  = p["Pct_Ocup"] < 0.85
        is_below = p["Pct_Ocup"] < 0.95
        dot_clr  = "#DC2626" if is_crit else "#D97706" if is_below else "#16A34A"
        dot_cls  = ' class="proj-dot-pulse"' if is_crit else ""

        # Trend arrow
        delta = proj_trends.get(name)
        if delta is not None and abs(delta) >= 0.1:
            if delta > 0:
                trend_html = f'<span style="color:#16A34A;font-weight:700">&#8593;{delta:.1f}pp</span>'
            else:
                trend_html = f'<span style="color:#DC2626;font-weight:700">&#8595;{abs(delta):.1f}pp</span>'
        else:
            trend_html = '<span style="color:#CBD5E1">&#8594;</span>'

        # Price pills (compact: tipología UF · CLP)
        proj_prices = (precios or {}).get(name, {})
        pills_html  = ""
        for tip, pinfo in sorted(proj_prices.items()):
            uf_val  = pinfo["min"]
            div     = pinfo["divisa"]
            clp_int = round(uf_val * uf_valor) if uf_valor and div == "UF" else 0
            clp_fmt = _fmt_clp(uf_val, uf_valor) or ""
            pills_html += (
                f'<span class="proj-price-pill" '
                f'data-uf="{uf_val}" data-clp="{clp_int}" data-div="{div}" data-tip="{tip}">'
                f'<b>{tip}</b> {uf_val:.1f}{div}'
                + (f' · {clp_fmt}' if clp_fmt else '')
                + '</span>'
            )

        name_js = name.replace("'", "\\'")
        name_id = "".join(c if c.isalnum() else "-" for c in name).strip("-").lower()

        sidebar_html += f"""
<div class="proj-list-item" id="pli-{name_id}" onclick="selectProj('{name_js}')"
     data-name="{name}" data-pct="{pct}" data-venc="{venc_n}"
     data-arr="{int(p['Arrendados'])}" data-disp="{int(p['Disponibles'])}"
     data-total="{int(p['Total'])}" data-pol="{int(p['Por_Liberar'])}">
  <div{dot_cls} style="margin-top:4px;width:9px;height:9px;flex-shrink:0;border-radius:99px;background:{dot_clr}"></div>
  {logo_html_sb}
  <div style="flex:1;min-width:0;overflow:hidden">
    <div style="font-size:.76rem;font-weight:700;color:#1E2A38;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;line-height:1.3">{name}</div>
    <div class="proj-sb-pct" style="font-size:.68rem;color:#6B7A8D;margin:1px 0 3px">{pct}% {trend_html}</div>
    <div class="proj-prices-row">{pills_html}</div>
  </div>
</div>"""

    # ── Cards HTML (overview panel) ───────────────────────────────────────────
    def _fc(v):
        """Format CLP value: 1.23M or 456k. Returns '' if zero."""
        if not v: return ""
        return f'${v/1e6:.2f}M' if v >= 1_000_000 else f'${round(v/1000)}k'

    cards_html = ""
    for p in sorted_projs:
        name    = p["Propiedad"]
        color   = proj_color_map[name]
        init    = _proj_initials(name)
        pct     = round(p["Pct_Ocup"] * 100, 1)
        tclr    = "#16A34A" if p["Pct_Ocup"] >= 0.95 else "#D97706" if p["Pct_Ocup"] >= 0.85 else "#DC2626"
        venc_n  = venc_by_proj.get(name, 0)
        pol_n   = int(p["Por_Liberar"])
        bar_w   = min(pct, 100)
        name_js = name.replace("'", "\\'")
        logo_url = proj_logos.get(name, "")
        logo_html_card = _logo_box(init, color, logo_url or None, 42)

        # Trend arrow for card
        delta = proj_trends.get(name)
        if delta is not None and abs(delta) >= 0.1:
            card_trend = (f'<span style="font-size:.66rem;color:#16A34A;font-weight:700">&#8593;{delta:.1f}pp</span>' if delta > 0
                          else f'<span style="font-size:.66rem;color:#DC2626;font-weight:700">&#8595;{abs(delta):.1f}pp</span>')
        else:
            card_trend = ""

        # Price table for card: Tipo | Disp | Desde | Hasta | F.Renta
        proj_prices = (precios or {}).get(name, {})
        card_rows   = ""
        for tip, pinfo in sorted(proj_prices.items()):
            mn  = pinfo["min"];  mx  = pinfo["max"]
            div = pinfo["divisa"];  nd = pinfo["n_disp"]
            mn_clp = round(mn * uf_valor) if uf_valor and div == "UF" else 0
            mx_clp = round(mx * uf_valor) if uf_valor and div == "UF" else 0
            fr_clp_s = (f'{_fc(mn_clp*3)}–{_fc(mx_clp*3)}') if mn_clp else ''
            card_rows += (
                f'<tr class="pct-row" '
                f'data-mn-uf="{mn:.2f}" data-mx-uf="{mx:.2f}" '
                f'data-mn-clp="{mn_clp}" data-mx-clp="{mx_clp}" '
                f'data-fr-mn-uf="{mn*3:.2f}" data-fr-mx-uf="{mx*3:.2f}" '
                f'data-fr-mn-clp="{mn_clp*3}" data-fr-mx-clp="{mx_clp*3}" '
                f'data-div="{div}">'
                f'<td class="pct-tip">{tip}</td>'
                f'<td class="pct-nd">{nd}</td>'
                f'<td class="pct-desde">{mn:.1f}&nbsp;{div}'
                f'<span class="pct-sub">{_fc(mn_clp)}</span></td>'
                f'<td class="pct-hasta">{mx:.1f}&nbsp;{div}'
                f'<span class="pct-sub">{_fc(mx_clp)}</span></td>'
                f'<td class="pct-fr">{mn*3:.1f}&ndash;{mx*3:.1f}&nbsp;{div}'
                f'<span class="pct-sub">{fr_clp_s}</span></td>'
                f'</tr>'
            )
        price_section = (
            f'<div style="margin-top:9px;padding-top:8px;border-top:1px solid #F1F5F9">'
            f'<table class="pct" cellspacing="0">'
            f'<thead><tr>'
            f'<th>Tipo</th><th>Disp</th>'
            f'<th class="pct-h-desde">Desde</th>'
            f'<th class="pct-h-hasta">Hasta</th>'
            f'<th>F.Renta</th>'
            f'</tr></thead><tbody>{card_rows}</tbody></table></div>'
        ) if card_rows else ""

        venc_b = (f'<div class="proj-stat"><span style="color:#DC2626;font-weight:700">{venc_n}</span> venc.</div>') if venc_n else ""
        pol_b  = (f'<div class="proj-stat"><span style="color:#DC2626;font-weight:700">{pol_n}</span> p.lib.</div>') if pol_n else ""

        cards_html += f"""
<div class="proj-card" onclick="selectProj('{name_js}')"
     data-name="{name}" data-pct="{pct}" data-venc="{venc_n}"
     data-arr="{int(p['Arrendados'])}" data-disp="{int(p['Disponibles'])}"
     data-total="{int(p['Total'])}" data-pol="{int(p['Por_Liberar'])}"
     style="border-top:4px solid {color}">
  <div style="display:flex;align-items:center;gap:9px;margin-bottom:10px">
    {logo_html_card}
    <div style="flex:1;min-width:0">
      <div style="font-size:.78rem;font-weight:700;color:#1E2A38;line-height:1.25;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">{name}</div>
      <div style="font-size:.66rem;color:#6B7A8D">{int(p['Total'])} deptos {card_trend}</div>
    </div>
  </div>
  <div style="margin-bottom:9px">
    <div style="display:flex;justify-content:space-between;align-items:baseline;margin-bottom:3px">
      <span style="font-size:.67rem;color:#6B7A8D">Ocupaci&oacute;n</span>
      <span class="proj-pct-val" style="font-size:.86rem;font-weight:800;color:{tclr}">{pct}%</span>
    </div>
    <div style="height:6px;background:#E2E8F0;border-radius:99px;overflow:hidden">
      <div class="proj-bar-fill" style="height:100%;width:{bar_w}%;background:{color};border-radius:99px"></div>
    </div>
  </div>
  <div class="proj-stats-row" style="display:flex;gap:5px;flex-wrap:wrap">
    <div class="proj-stat"><span class="proj-arr-val" style="color:#16A34A;font-weight:700">{int(p['Arrendados'])}</span> arr.</div>
    <div class="proj-stat"><span class="proj-disp-val" style="color:#D97706;font-weight:700">{int(p['Disponibles'])}</span> disp.</div>
    {pol_b}{venc_b}
  </div>
  {price_section}
</div>"""

    # ── JSON vars ─────────────────────────────────────────────────────────────
    colors_js      = json.dumps(proj_color_map, ensure_ascii=False)
    logos_js       = json.dumps(proj_logos,      ensure_ascii=False)
    venc_mo_js     = json.dumps(venc_monthly,    ensure_ascii=False)
    precios_js     = json.dumps(precios or {},   ensure_ascii=False)
    trends_js      = json.dumps(proj_trends,     ensure_ascii=False)
    uf_js          = str(round(uf_valor, 2)) if uf_valor else "null"

    uf_badge = (
        f'<span style="background:#EEF6FB;border:1px solid #BAE6FD;border-radius:99px;'
        f'padding:2px 10px;font-size:.7rem;color:#0369A1;font-weight:700;margin-left:8px">'
        f'UF ${uf_valor:,.2f}</span>'
    ) if uf_valor else ""

    section = f"""
<style>
.proj-dot-pulse{{animation:proj-pulse 1.8s ease-in-out infinite;}}
@keyframes proj-pulse{{0%,100%{{box-shadow:0 0 0 0 rgba(220,38,38,.5);}}50%{{box-shadow:0 0 0 6px rgba(220,38,38,0);}}}}
.proj-list-item{{display:flex;align-items:flex-start;gap:9px;padding:10px 12px;cursor:pointer;
  border-bottom:1px solid #F1F5F9;transition:background .13s;border-left:3px solid transparent;}}
.proj-list-item:hover{{background:#F0FAFB;}}
.proj-list-item.selected{{background:#E6F7F9;border-left-color:#00A8B4;}}
.proj-price-pill{{display:inline-flex;align-items:center;gap:2px;font-size:.62rem;background:#F0FDF4;
  border:1px solid #BBF7D0;border-radius:5px;padding:1px 6px;color:#15803D;margin:1px 3px 1px 0;}}
.proj-card{{background:#fff;border:1px solid #E2E8F0;border-radius:12px;padding:14px;
  cursor:pointer;transition:all .18s ease;}}
.proj-card:hover{{border-color:#00A8B4;box-shadow:0 6px 20px rgba(0,142,159,.1);transform:translateY(-2px);}}
.proj-stat{{font-size:.67rem;color:#6B7A8D;background:#F8FAFC;padding:2px 7px;border-radius:99px;}}
.proj-logo-box img{{display:block;}}
.pct{{width:100%;border-collapse:collapse;font-size:.59rem;table-layout:fixed;}}
.pct thead tr{{background:#F8FAFC;}}
.pct th{{padding:3px 4px;text-align:right;color:#9CA3AF;font-weight:700;text-transform:uppercase;letter-spacing:.03em;border-bottom:1px solid #F1F5F9;white-space:nowrap;overflow:hidden;}}
.pct th:first-child{{text-align:left;width:22%;}}
.pct th:nth-child(2){{width:11%;}}
.pct th:nth-child(3),.pct th:nth-child(4){{width:18%;}}
.pct th:nth-child(5){{width:31%;}}
.pct td{{padding:3px 4px;text-align:right;color:#374151;border-bottom:1px solid #F8FAFC;overflow:hidden;}}
.pct td:first-child{{text-align:left;font-weight:700;color:#1E2A38;}}
.pct-nd{{color:#D97706!important;font-weight:700!important;}}
.pct-desde{{color:#15803D!important;}}
.pct-hasta{{color:#0369A1!important;}}
.pct-fr{{color:#7C3AED!important;font-size:.57rem!important;white-space:normal!important;line-height:1.3;}}
.pct-sub{{display:block;font-size:.54rem;color:#9CA3AF;line-height:1.3;margin-top:1px;}}
@media(max-width:840px){{
  #proj-workspace{{flex-direction:column!important;}}
  #proj-sidebar-panel{{width:100%!important;max-height:260px;border-right:none!important;border-bottom:1px solid #E2E8F0;}}
}}
</style>

<div id="sec-proyectos" class="sec"
     style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;margin:28px 0 4px">
  <span>Vista por Proyecto&nbsp;<span class="sec-chevron">&#9650;</span></span>
  <button onclick="exportToExcel();event.stopPropagation();" title="Exportar tabla a Excel"
    style="display:flex;align-items:center;gap:6px;padding:7px 14px;background:#16A34A;color:#fff;
           border:none;border-radius:8px;cursor:pointer;font-size:.78rem;font-weight:700;
           box-shadow:0 2px 8px rgba(22,163,74,.25);transition:opacity .15s"
    onmouseover="this.style.opacity='.85'" onmouseout="this.style.opacity='1'">
    &#128229; Exportar Excel
  </button>
</div>
<div class="sec-sub" style="margin-bottom:12px">Selecciona un proyecto para ver su an&aacute;lisis completo &mdash; Datos al {DATE_STR}{uf_badge}</div>

<div id="proj-workspace"
     style="display:flex;border:1px solid #E2E8F0;border-radius:16px;overflow:hidden;
            min-height:580px;margin-bottom:32px;background:#fff">

  <!-- ── Sidebar izquierdo ───────────────────────────────────────────── -->
  <div id="proj-sidebar-panel"
       style="width:268px;flex-shrink:0;border-right:1px solid #E2E8F0;
              background:#F8FAFC;display:flex;flex-direction:column">

    <!-- Controles -->
    <div style="padding:12px 13px;border-bottom:1px solid #E2E8F0;background:#fff">
      <input id="proj-sb-search" type="text" placeholder="&#128269; Buscar proyecto..."
             oninput="filtrarSidebar()"
             style="width:100%;box-sizing:border-box;padding:6px 10px;border:1px solid #E2E8F0;
                    border-radius:8px;font-size:.75rem;background:#F8FAFC;color:#1E2A38;outline:none">
      <div style="display:flex;gap:6px;margin-top:8px">
        <select id="proj-sb-sort" onchange="sortSidebar(this.value)"
                style="flex:1;padding:5px 8px;border:1px solid #E2E8F0;border-radius:7px;
                       font-size:.72rem;background:#F8FAFC;color:#374151;cursor:pointer;outline:none">
          <option value="alpha">A &#8594; Z</option>
          <option value="ocup_desc">&#8595; Ocupaci&oacute;n</option>
          <option value="ocup_asc">&#8593; Ocupaci&oacute;n</option>
          <option value="venc">&#8595; Vencimientos</option>
        </select>
        <button id="proj-uf-btn" onclick="toggleUFMode()"
                title="Cambiar visualizaci&oacute;n de precios"
                style="padding:5px 9px;border:1px solid #BAE6FD;border-radius:7px;font-size:.72rem;
                       background:#EEF6FB;color:#0369A1;cursor:pointer;font-weight:700;white-space:nowrap">
          $ CLP
        </button>
      </div>
      <button id="proj-pol-btn" onclick="togglePorLiberar()"
              title="Proyecci&oacute;n contabilizando unidades por liberar como disponibles"
              style="width:100%;margin-top:6px;padding:5px 9px;border:1px solid #E2E8F0;
                     border-radius:7px;font-size:.72rem;background:#F8FAFC;color:#374151;
                     cursor:pointer;font-weight:600;display:flex;align-items:center;
                     justify-content:center;gap:5px;transition:all .15s">
        &#9654; Proyecci&oacute;n Por Liberar
      </button>
    </div>

    <!-- Lista proyectos -->
    <div id="proj-list-items" style="flex:1;overflow-y:auto">
      {sidebar_html}
    </div>

    <!-- Footer -->
    <div style="padding:8px 13px;border-top:1px solid #E2E8F0;background:#fff;
                font-size:.65rem;color:#9CA3AF;display:flex;justify-content:space-between">
      <span>{n_projs} proyectos</span>
      <span style="color:#DC2626;font-weight:600">{n_below_meta} bajo meta</span>
    </div>
  </div>

  <!-- ── Panel derecho ───────────────────────────────────────────────── -->
  <div id="proj-right-panel" style="flex:1;overflow-y:auto;background:#fff;min-width:0">

    <!-- Overview (estado inicial) -->
    <div id="proj-overview-panel" style="padding:18px 20px">
      <p style="font-size:.75rem;color:#9CA3AF;margin:0 0 14px;display:flex;align-items:center;gap:6px">
        <span style="font-size:1rem">&#128204;</span>
        Selecciona un proyecto en la lista para ver su an&aacute;lisis completo.
      </p>
      <div id="proj-cards-grid"
           style="display:grid;grid-template-columns:repeat(auto-fill,minmax(290px,1fr));gap:12px">
        {cards_html}
      </div>
    </div>

    <!-- Detalle proyecto (oculto por defecto) -->
    <div id="proj-detail-panel" style="display:none;padding:26px 30px 32px">
      <!-- Rellenado por JS: buildDetailHTML(projName) -->
    </div>
  </div>
</div>
"""

    js = f"""
// ── Vista por Proyecto v3 ────────────────────────────────────────────────────
var PROJ_COLORS  = {colors_js};
var PROJ_LOGOS   = {logos_js};
var PROJ_VENC_MO = {venc_mo_js};
var PROJ_PRECIOS = {precios_js};
var PROJ_TRENDS  = {trends_js};
var UF_VALOR     = {uf_js};
var _projUFMode  = true;   // true=UF primero, false=CLP primero
var _selProj     = null;

// ── Helpers de formato ───────────────────────────────────────────────────────
function _fmtClp(clp) {{
  if(!clp) return "";
  return clp >= 1000000 ? "$"+(clp/1000000).toFixed(2)+"M" : "$"+Math.round(clp/1000)+"k";
}}
function _dkpi(label, val, color) {{
  return '<div style="background:#F8FAFC;border-radius:8px;padding:9px 10px;border-bottom:3px solid '+color+'">'
    +'<div style="font-size:.6rem;color:#6B7A8D;font-weight:600;margin-bottom:2px">'+label+'</div>'
    +'<div style="font-size:1.18rem;font-weight:700;color:'+color+'">'+val+'</div></div>';
}}

// ── Logo fallback ─────────────────────────────────────────────────────────────
function projLogoFallback(img) {{
  var box=img.parentElement, init=box.dataset.init||"??", color=box.dataset.color||"#00A8B4";
  var sz=box.offsetWidth||34, cx=Math.round(sz/2), cy=Math.round(sz*.64), fs=Math.round(sz*.33);
  box.innerHTML='<svg xmlns="http://www.w3.org/2000/svg" width="'+sz+'" height="'+sz+'"'
    +' viewBox="0 0 '+sz+' '+sz+'"><rect width="'+sz+'" height="'+sz+'" rx="8" fill="'+color+'"/>'
    +'<text x="'+cx+'" y="'+cy+'" text-anchor="middle" font-family="Arial,sans-serif"'
    +' font-size="'+fs+'" font-weight="800" fill="white">'+init+'</text></svg>';
}}

// ── Sidebar: busqueda y sort ──────────────────────────────────────────────────
function filtrarSidebar() {{
  var q=((document.getElementById("proj-sb-search")||{{}}).value||"").toLowerCase();
  document.querySelectorAll(".proj-list-item").forEach(function(item) {{
    item.style.display=(item.dataset.name||"").toLowerCase().includes(q)?"":"none";
  }});
}}

function sortSidebar(mode) {{
  localStorage.setItem("projSort",mode);
  var c=document.getElementById("proj-list-items"); if(!c) return;
  var items=Array.from(c.querySelectorAll(".proj-list-item"));
  items.sort(function(a,b) {{
    if(mode==="alpha")     return (a.dataset.name||"").localeCompare(b.dataset.name||"","es");
    if(mode==="ocup_desc") return parseFloat(b.dataset.pct||0)-parseFloat(a.dataset.pct||0);
    if(mode==="ocup_asc")  return parseFloat(a.dataset.pct||0)-parseFloat(b.dataset.pct||0);
    if(mode==="venc")      return parseInt(b.dataset.venc||0)-parseInt(a.dataset.venc||0);
    return 0;
  }});
  items.forEach(function(item){{c.appendChild(item);}});
}}

// ── Toggle UF / CLP ──────────────────────────────────────────────────────────
function toggleUFMode() {{
  _projUFMode=!_projUFMode;
  localStorage.setItem("projUFMode",_projUFMode?"true":"false");
  var btn=document.getElementById("proj-uf-btn");
  if(btn) btn.innerHTML=_projUFMode?"$ CLP":"UF";
  _refreshPricePills();
  // Actualizar tabla de precios en detalle si hay proyecto abierto
  if(_selProj) _renderPriceTable(_selProj);
}}

function _refreshPricePills() {{
  document.querySelectorAll(".proj-price-pill").forEach(function(pill) {{
    var uf=parseFloat(pill.dataset.uf||0), clp=parseInt(pill.dataset.clp||0);
    var div=pill.dataset.div||"UF", tip=pill.dataset.tip||"";
    var ufStr=uf.toFixed(1)+" "+div, clpStr=clp?_fmtClp(clp):"";
    if(_projUFMode) {{
      pill.innerHTML="<b>"+tip+"</b> "+ufStr+(clpStr?" &middot; "+clpStr:"");
    }} else {{
      pill.innerHTML="<b>"+tip+"</b> "+(clpStr||ufStr)+(clpStr?" &middot; "+ufStr:"");
    }}
  }});
}}

// ── Inicializar CLP en cards (fallback cuando Python no pudo obtener UF) ─────
function _initCardCLP() {{
  if(!UF_VALOR) return;
  document.querySelectorAll(".pct-row").forEach(function(row) {{
    var mn=parseFloat(row.dataset.mnUf), mx=parseFloat(row.dataset.mxUf);
    if(!mn||row.dataset.div!=="UF") return;
    var mnC=Math.round(mn*UF_VALOR), mxC=Math.round(mx*UF_VALOR);
    var subs=row.querySelectorAll(".pct-sub");
    if(subs[0]&&!subs[0].textContent) subs[0].textContent=_fmtClp(mnC);
    if(subs[1]&&!subs[1].textContent) subs[1].textContent=_fmtClp(mxC);
    if(subs[2]&&!subs[2].textContent) subs[2].textContent=_fmtClp(mnC*3)+"–"+_fmtClp(mxC*3);
  }});
}}
window.addEventListener('load', _initCardCLP);

// ── Proyección Por Liberar ────────────────────────────────────────────────────
var _polMode = false;
function togglePorLiberar() {{
  _polMode = !_polMode;
  var btn = document.getElementById("proj-pol-btn");
  if(_polMode) {{
    btn.style.background="#FFF7ED"; btn.style.borderColor="#FED7AA";
    btn.style.color="#C2410C"; btn.innerHTML="&#9646;&#9646; Proyecci&oacute;n Activa";
  }} else {{
    btn.style.background="#F8FAFC"; btn.style.borderColor="#E2E8F0";
    btn.style.color="#374151"; btn.innerHTML="&#9654; Proyecci&oacute;n Por Liberar";
  }}

  // Actualizar cards
  document.querySelectorAll(".proj-card").forEach(function(card) {{
    var arr   = parseFloat(card.dataset.arr   || 0);
    var disp  = parseFloat(card.dataset.disp  || 0);
    var total = parseFloat(card.dataset.total || 1);
    var pol   = parseFloat(card.dataset.pol   || 0);
    var arrE  = _polMode ? arr - pol  : arr;
    var dispE = _polMode ? disp + pol : disp;
    var pctE  = total > 0 ? arrE / total * 100 : 0;
    var clr   = pctE >= 95 ? "#16A34A" : pctE >= 85 ? "#D97706" : "#DC2626";

    var pctEl  = card.querySelector(".proj-pct-val");
    var barEl  = card.querySelector(".proj-bar-fill");
    var arrEl  = card.querySelector(".proj-arr-val");
    var dispEl = card.querySelector(".proj-disp-val");
    if(pctEl)  {{ pctEl.textContent = pctE.toFixed(1)+"%"; pctEl.style.color = clr; }}
    if(barEl)  {{ barEl.style.width = Math.min(pctE,100)+"%"; }}
    if(arrEl)  {{ arrEl.textContent = Math.round(arrE); }}
    if(dispEl) {{ dispEl.textContent = Math.round(dispE); }}

    // Badge "por liberar"
    var badge = card.querySelector(".pol-proj-badge");
    if(_polMode && pol > 0) {{
      if(!badge) {{
        badge = document.createElement("div");
        badge.className = "proj-stat pol-proj-badge";
        badge.style.cssText = "color:#C2410C;background:#FFF7ED;border:1px solid #FED7AA;font-weight:700";
        badge.innerHTML = '+'+pol+' p.lib.';
        var sr = card.querySelector(".proj-stats-row");
        if(sr) sr.appendChild(badge);
      }}
    }} else if(badge) {{ badge.remove(); }}
  }});

  // Actualizar sidebar
  document.querySelectorAll(".proj-list-item").forEach(function(item) {{
    var arr   = parseFloat(item.dataset.arr   || 0);
    var total = parseFloat(item.dataset.total || 1);
    var pol   = parseFloat(item.dataset.pol   || 0);
    var arrE  = _polMode ? arr - pol : arr;
    var pctE  = total > 0 ? arrE / total * 100 : 0;
    var sbPct = item.querySelector(".proj-sb-pct");
    if(sbPct) {{
      var trend = sbPct.dataset.trendHtml || "";
      sbPct.innerHTML = pctE.toFixed(1)+"% "+(_polMode&&pol>0?'<span style="color:#C2410C;font-size:.62rem">(+'+pol+' p.l.)</span>':"");
    }}
  }});
}}

// ── Seleccionar proyecto ──────────────────────────────────────────────────────
function selectProj(projName) {{
  _selProj=projName;
  localStorage.setItem("lastSelProj",projName);
  if(history.pushState) history.pushState(null,null,"#proj-"+projName.replace(/[^a-zA-Z0-9]/g,"_"));

  // Highlight en sidebar
  document.querySelectorAll(".proj-list-item").forEach(function(item) {{
    item.classList.toggle("selected",item.dataset.name===projName);
  }});
  var selItem=document.querySelector(".proj-list-item.selected");
  if(selItem) selItem.scrollIntoView({{block:"nearest",behavior:"smooth"}});

  // Mostrar panel detalle, ocultar overview
  document.getElementById("proj-overview-panel").style.display="none";
  var dp=document.getElementById("proj-detail-panel");
  dp.style.display="block";
  dp.innerHTML=_buildDetailHTML(projName);

  // Charts Plotly con delay para animacion
  setTimeout(function(){{_renderDetailCharts(projName);}},380);
}}

function backToOverview() {{
  _selProj=null;
  localStorage.removeItem("lastSelProj");
  if(history.pushState) history.pushState(null,null,window.location.pathname+window.location.search);
  document.querySelectorAll(".proj-list-item").forEach(function(item){{item.classList.remove("selected");}});
  document.getElementById("proj-overview-panel").style.display="";
  document.getElementById("proj-detail-panel").style.display="none";
}}

// ── Construir HTML de detalle ─────────────────────────────────────────────────
function _buildDetailHTML(projName) {{
  if(typeof proj_desc==="undefined") return "<p>Sin datos</p>";
  var p=null;
  for(var i=0;i<proj_desc.length;i++){{if(proj_desc[i].Propiedad===projName){{p=proj_desc[i];break;}}}}
  if(!p) return "<p>Proyecto no encontrado</p>";

  var color=PROJ_COLORS[projName]||"#00A8B4";
  var pct=(p.Pct_Ocup*100).toFixed(1);
  var mt=typeof META_TARGET!=="undefined"?META_TARGET:0.95;
  var tgt=(mt*100).toFixed(0);
  var gap=((p.Pct_Ocup-mt)*100).toFixed(1);
  var stClr=p.Pct_Ocup>=mt?"#16A34A":p.Pct_Ocup>=mt-.1?"#D97706":"#DC2626";
  var badge=p.Pct_Ocup>=mt?"&#x2705; Cumple meta":p.Pct_Ocup>=mt-.1?"&#x1F7E1; En seguimiento":"&#x1F534; Cr&iacute;tico";

  // Trend arrow
  var delta=PROJ_TRENDS[projName];
  var trendBadge="";
  if(delta!==null&&delta!==undefined&&Math.abs(delta)>=.1) {{
    trendBadge='<span style="font-size:.72rem;font-weight:700;padding:1px 8px;border-radius:99px;margin-left:8px;background:'
      +(delta>0?"#F0FDF4;color:#16A34A":"#FEF2F2;color:#DC2626")+'">'
      +(delta>0?"&#8593;":"&#8595;")+Math.abs(delta).toFixed(1)+"pp vs mes ant.</span>";
  }}

  // Logo
  var logoUrl=PROJ_LOGOS[projName]||"";
  var words=projName.replace(/-/g," ").split(" ").filter(function(w){{return w;}});
  var initials=words.length>=2?(words[0][0]+words[1][0]).toUpperCase():projName.slice(0,2).toUpperCase();
  var svgLogo='<svg xmlns="http://www.w3.org/2000/svg" width="56" height="56" viewBox="0 0 56 56">'
    +'<rect width="56" height="56" rx="13" fill="'+color+'"/>'
    +'<text x="28" y="37" text-anchor="middle" font-family="Arial,sans-serif" font-size="19" font-weight="800" fill="white">'+initials+'</text></svg>';
  var logoHtml=logoUrl
    ?'<div class="proj-logo-box" data-init="'+initials+'" data-color="'+color+'"'
      +' style="width:56px;height:56px;flex-shrink:0;border-radius:12px;background:#F8FAFC;overflow:hidden;display:flex;align-items:center;justify-content:center">'
      +'<img src="'+logoUrl+'" width="50" height="50" style="object-fit:contain" onerror="projLogoFallback(this)"></div>'
    :svgLogo;

  // Vencimientos desde tabla
  var vRows=document.querySelectorAll("#venc-table tbody tr");
  var vList=[];
  vRows.forEach(function(r){{
    if(r.dataset.proj===projName) vList.push({{
      dias:parseInt(r.dataset.dias)||0,
      unidad:r.cells[1]?r.cells[1].textContent:"",
      tipo:r.cells[2]?r.cells[2].textContent:"",
      fecha:r.cells[3]?r.cells[3].textContent:"",
      ejec:r.cells[5]?r.cells[5].textContent:""
    }});
  }});
  vList.sort(function(a,b){{return a.dias-b.dias;}});

  var h="";

  // Barra de navegacion del detalle
  h+='<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:22px">';
  h+='<button onclick="backToOverview()" style="display:flex;align-items:center;gap:5px;background:none;border:1px solid #E2E8F0;border-radius:8px;padding:6px 12px;font-size:.75rem;color:#374151;cursor:pointer;font-weight:600">'
    +'&#8592; Volver</button>';
  // Prev / Next navigation
  var _allP=Array.from(document.querySelectorAll('.proj-list-item')).map(function(el){{return el.dataset.name;}});
  var _ci=_allP.indexOf(projName);
  _PROJ_PREV=_ci>0?_allP[_ci-1]:null;
  _PROJ_NEXT=_ci<_allP.length-1?_allP[_ci+1]:null;
  var pe=projName.replace(/'/g,"\\'");
  h+='<div style="display:flex;align-items:center;gap:5px">';
  h+=(_PROJ_PREV
    ?'<button onclick="prevProj()" title="'+_PROJ_PREV+'" style="padding:5px 10px;background:none;border:1px solid #E2E8F0;border-radius:7px;font-size:.85rem;color:#374151;cursor:pointer">&#8592;</button>'
    :'<button disabled style="padding:5px 10px;background:none;border:1px solid #F1F5F9;border-radius:7px;font-size:.85rem;color:#CBD5E1;cursor:default">&#8592;</button>');
  h+=(_PROJ_NEXT
    ?'<button onclick="nextProj()" title="'+_PROJ_NEXT+'" style="padding:5px 10px;background:none;border:1px solid #E2E8F0;border-radius:7px;font-size:.85rem;color:#374151;cursor:pointer">&#8594;</button>'
    :'<button disabled style="padding:5px 10px;background:none;border:1px solid #F1F5F9;border-radius:7px;font-size:.85rem;color:#CBD5E1;cursor:default">&#8594;</button>');
  h+='<button onclick="setProjFilter(&apos;'+pe+'&apos;)" '
    +'style="padding:5px 10px;background:#F8FAFC;border:1px solid #E2E8F0;border-radius:7px;font-size:.75rem;color:#374151;cursor:pointer">&#128204; Filtrar</button>';
  h+='<button onclick="exportProjData(&apos;'+pe+'&apos;)" '
    +'style="padding:5px 10px;background:#0369A1;color:#fff;border:none;border-radius:7px;font-size:.75rem;cursor:pointer;font-weight:600">&#8595; Exportar</button>';
  h+='</div></div>';

  // Header proyecto
  h+='<div style="display:flex;align-items:center;gap:14px;margin-bottom:18px">';
  h+=logoHtml;
  h+='<div>';
  h+='<div style="font-size:.58rem;text-transform:uppercase;letter-spacing:.12em;color:#9CA3AF;font-weight:600">PROYECTO</div>';
  h+='<h2 style="margin:2px 0;font-size:1.2rem;font-weight:800;color:#1E2A38">'+projName+'</h2>';
  h+='<div style="display:flex;align-items:center;gap:6px;font-size:.73rem;color:'+stClr+';font-weight:600">'+badge+trendBadge+'</div>';
  h+='</div></div>';

  // Barra de ocupacion
  h+='<div style="background:#F8FAFC;border-radius:10px;padding:14px;margin-bottom:16px;border:1px solid #E2E8F0">';
  h+='<div style="display:flex;justify-content:space-between;align-items:baseline;margin-bottom:8px">';
  h+='<span style="font-size:.72rem;color:#6B7A8D;font-weight:600;text-transform:uppercase;letter-spacing:.06em">Ocupaci&oacute;n</span>';
  h+='<span style="font-size:2rem;font-weight:800;color:'+stClr+'">'+pct+'%</span></div>';
  h+='<div style="height:10px;background:#E2E8F0;border-radius:99px;overflow:hidden;margin-bottom:7px">';
  h+='<div style="height:100%;width:'+Math.min(parseFloat(pct),100)+'%;background:'+color+';border-radius:99px"></div></div>';
  h+='<div style="display:flex;justify-content:space-between;font-size:.69rem;color:#9CA3AF">';
  h+='<span>Meta: <b>'+tgt+'%</b></span>';
  h+='<span>Gap: <b style="color:'+stClr+'">'+(parseFloat(gap)>=0?"+":"")+gap+'pp</b></span>';
  h+='<span>Necesita: <b style="color:'+stClr+'">'+(p.Uds_Needed>0?Math.round(p.Uds_Needed)+' uds':'&#x2713; OK')+'</b></span>';
  h+='</div></div>';

  // KPIs 3x2
  h+='<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:18px">';
  h+=_dkpi("Total",        Math.round(p.Total),       "#6B7A8D");
  h+=_dkpi("Arrendadas",   Math.round(p.Arrendados),  "#16A34A");
  h+=_dkpi("Disponibles",  Math.round(p.Disponibles), "#D97706");
  h+=_dkpi("No Disp.",     Math.round(p.No_Disp),     "#DC2626");
  h+=_dkpi("Por Liberar",  Math.round(p.Por_Liberar), "#DC2626");
  h+=_dkpi("Reservadas",   Math.round(p.Reservadas),  "#0369A1");
  h+='</div>';

  // Charts (side by side)
  h+='<div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:18px">';
  h+='<div><div style="font-size:.68rem;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#1E2A38;margin-bottom:6px">Tipolog&iacute;a</div>'
    +'<div id="dtipo-chart" style="height:190px"></div></div>';
  var vmo=PROJ_VENC_MO[projName];
  var hasVmo=vmo&&vmo.labels&&vmo.labels.length>0;
  h+='<div><div style="font-size:.68rem;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#1E2A38;margin-bottom:6px">Vencimientos por mes</div>'
    +(hasVmo?'<div id="dvenc-chart" style="height:190px"></div>':'<div style="height:190px;display:flex;align-items:center;justify-content:center;color:#CBD5E1;font-size:.75rem">Sin vencimientos en 60d</div>')+'</div>';
  h+='</div>';

  // Tabla de precios (UF y CLP, igual relevancia)
  var projPrices=PROJ_PRECIOS[projName]||{{}};
  var tips=Object.keys(projPrices).sort();
  if(tips.length>0) {{
    h+='<div style="margin-bottom:18px">';
    h+='<div style="font-size:.68rem;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#1E2A38;margin-bottom:8px">Precios desde (unidades disponibles)</div>';
    h+='<div id="detail-price-table"></div>';
    h+='</div>';
  }}

  // Contratos que vencen
  if(vList.length>0) {{
    h+='<div style="margin-bottom:18px">';
    h+='<div style="font-size:.68rem;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#1E2A38;margin-bottom:8px">Contratos que vencen pronto ('+vList.length+')</div>';
    h+='<div style="max-height:220px;overflow-y:auto;border:1px solid #F1F5F9;border-radius:10px">';
    vList.slice(0,20).forEach(function(v) {{
      var vc=v.dias<=30?"#DC2626":"#D97706";
      h+='<div style="display:flex;justify-content:space-between;align-items:center;padding:7px 11px;border-bottom:1px solid #F8FAFC;font-size:.73rem">';
      h+='<div><b>'+v.unidad+'</b> <span style="color:#9CA3AF;font-size:.67rem">('+v.tipo+')</span>';
      if(v.ejec) h+='<div style="font-size:.63rem;color:#9CA3AF">'+v.ejec+'</div>';
      h+='</div><div style="text-align:right;flex-shrink:0;margin-left:10px">';
      h+='<div style="font-size:.67rem;color:#9CA3AF;margin-bottom:2px">'+v.fecha+'</div>';
      h+='<span style="background:'+vc+'18;color:'+vc+';padding:1px 8px;border-radius:99px;font-size:.67rem;font-weight:700">'+v.dias+'d</span>';
      h+='</div></div>';
    }});
    if(vList.length>20) h+='<div style="padding:7px 11px;font-size:.67rem;color:#9CA3AF">+ '+(vList.length-20)+' m&aacute;s...</div>';
    h+='</div></div>';
  }}

  return h;
}}

// ── Tabla de precios (llamada tras render) ────────────────────────────────────
function _renderPriceTable(projName) {{
  var el=document.getElementById("detail-price-table"); if(!el) return;
  var projPrices=PROJ_PRECIOS[projName]||{{}};
  var tips=Object.keys(projPrices).sort();
  if(tips.length===0) {{ el.innerHTML=""; return; }}
  var tdB='padding:7px 10px;border-bottom:1px solid #F1F5F9;font-size:.74rem;vertical-align:middle',
      tdR='padding:7px 10px;border-bottom:1px solid #F1F5F9;font-size:.74rem;text-align:right;vertical-align:middle',
      tdG='padding:7px 10px;border-bottom:1px solid #F1F5F9;font-size:.74rem;text-align:right;color:#15803D;font-weight:700;vertical-align:middle',
      tdP='padding:7px 10px;border-bottom:1px solid #F1F5F9;font-size:.68rem;text-align:right;color:#7C3AED;vertical-align:middle',
      tdM='padding:5px 10px 5px 22px;border-bottom:1px solid #F8FAFC;font-size:.69rem;color:#6B7A8D;vertical-align:middle',
      tdMr='padding:5px 10px;border-bottom:1px solid #F8FAFC;font-size:.69rem;text-align:right;vertical-align:middle';
  function fmtUF(v,div){{return '<b>'+v.toFixed(1)+'</b> <span style="color:#9CA3AF;font-size:.65rem">'+div+'</span>';}}
  var rows="";
  tips.forEach(function(tip) {{
    var pr=projPrices[tip]; if(!pr) return;
    var div=pr.divisa||"UF";
    var mn=pr.min, mx=pr.max||mn;
    var mnC=UF_VALOR&&div==="UF"?Math.round(mn*UF_VALOR):0;
    var mxC=UF_VALOR&&div==="UF"?Math.round(mx*UF_VALOR):0;
    var nd=pr.n_disp||0, ndClr=nd>0?"#D97706":"#CBD5E1";
    function sub(v){{return v?'<span style="display:block;font-size:.65rem;color:#9CA3AF;font-weight:400;margin-top:1px">'+_fmtClp(v)+'</span>':''}}
    var desdeVal=fmtUF(mn,div)+sub(mnC);
    var hastaVal=fmtUF(mx,div)+sub(mxC);
    var frVal='<b>'+(mn*3).toFixed(1)+'</b>&ndash;<b>'+(mx*3).toFixed(1)+'</b> <span style="color:#9CA3AF;font-size:.65rem">'+div+'</span>'+
      (mnC?'<span style="display:block;font-size:.65rem;color:#9CA3AF;font-weight:400;margin-top:1px">'+_fmtClp(mnC*3)+'&ndash;'+_fmtClp(mxC*3)+'</span>':'');
    rows+='<tr>'+
      '<td style="'+tdB+';font-weight:700;color:#1E2A38">'+tip+'</td>'+
      '<td style="'+tdR+';font-weight:700;color:'+ndClr+'">'+nd+'</td>'+
      '<td style="'+tdG+'">'+desdeVal+'</td>'+
      '<td style="'+tdG+'">'+hastaVal+'</td>'+
      '<td style="'+tdP+'">'+frVal+'</td>'+
      '</tr>';
    var mods=(pr.modelos||[]).filter(function(m){{return m.modelo;}});
    if(mods.length>1) {{
      mods.forEach(function(m) {{
        var mmn=m.min, mmx=m.max||mmn, mnd=m.n_disp;
        var mmnC=UF_VALOR&&div==="UF"?Math.round(mmn*UF_VALOR):0;
        var mmxC=UF_VALOR&&div==="UF"?Math.round(mmx*UF_VALOR):0;
        var mDesde=mmn.toFixed(1)+" "+div+(mmnC?'<span style="display:block;font-size:.62rem;color:#9CA3AF;margin-top:1px">'+_fmtClp(mmnC)+'</span>':'');
        var mHasta=mmx.toFixed(1)+" "+div+(mmxC?'<span style="display:block;font-size:.62rem;color:#9CA3AF;margin-top:1px">'+_fmtClp(mmxC)+'</span>':'');
        var mFr=(mmn*3).toFixed(1)+"&ndash;"+(mmx*3).toFixed(1)+" "+div+
          (mmnC?'<span style="display:block;font-size:.62rem;color:#9CA3AF;margin-top:1px">'+_fmtClp(mmnC*3)+'&ndash;'+_fmtClp(mmxC*3)+'</span>':'');
        rows+='<tr>'+
          '<td style="'+tdM+'">&#x2514; '+m.modelo+'</td>'+
          '<td style="'+tdMr+';color:#D97706">'+mnd+'</td>'+
          '<td style="'+tdMr+';color:#6B7A8D;vertical-align:top">'+mDesde+'</td>'+
          '<td style="'+tdMr+';color:#6B7A8D;vertical-align:top">'+mHasta+'</td>'+
          '<td style="'+tdMr+';color:#9061F9;font-size:.66rem;vertical-align:top">'+mFr+'</td>'+
          '</tr>';
      }});
    }}
  }});
  function _th(txt,align,clr){{
    return '<th style="padding:8px 10px;text-align:'+align+';font-size:.64rem;color:'+clr+';font-weight:700;text-transform:uppercase;letter-spacing:.05em;border-bottom:1px solid #E2E8F0">'+txt+'</th>';
  }}
  el.innerHTML='<table style="width:100%;border-collapse:collapse;border:1px solid #E2E8F0;border-radius:10px;overflow:hidden">'+
    '<thead><tr style="background:#F8FAFC">'+
    _th("Tipología","left","#6B7A8D")+_th("Disp","right","#D97706")+
    _th("Desde","right","#15803D")+_th("Hasta","right","#15803D")+
    _th("Factor Renta (3×)","right","#7C3AED")+
    '</tr></thead><tbody>'+rows+'</tbody></table>';
}}


// ── Render Plotly en detalle ──────────────────────────────────────────────────
function _renderDetailCharts(projName) {{
  _renderPriceTable(projName);
  var color=PROJ_COLORS[projName]||"#00A8B4";
  var base={{paper_bgcolor:"#FFFFFF",plot_bgcolor:"#FFFFFF",
    font:{{family:"Arial,sans-serif",size:11,color:"#1A202C"}},
    margin:{{t:10,b:55,l:40,r:10}},
    xaxis:{{gridcolor:"#E2E8F0",linecolor:"#E2E8F0",tickfont:{{size:10}}}},
    yaxis:{{gridcolor:"#E2E8F0",linecolor:"#E2E8F0",tickfont:{{size:10}}}},
    hoverlabel:{{bgcolor:"#1E2A38",font:{{color:"#fff",size:11}}}}
  }};

  // Tipologia chart
  if(typeof tipo_data!=="undefined"&&tipo_data.projects) {{
    var idx=tipo_data.projects.indexOf(projName);
    if(idx>=0) {{
      var tips=[],arrs=[],disps=[];
      tipo_data.tipologias.forEach(function(tip) {{
        var d=tipo_data[tip]; if(!d) return;
        var tot=d.total[idx]||0; if(!tot) return;
        var arr=d.arrendados[idx]||0;
        tips.push(tip); arrs.push(arr); disps.push(tot-arr);
      }});
      if(tips.length) try {{
        Plotly.newPlot("dtipo-chart",[
          {{type:"bar",name:"Arrendadas",x:tips,y:arrs,marker:{{color:color}},hovertemplate:"%{{y}} uds<extra>Arrendadas</extra>"}},
          {{type:"bar",name:"Disponibles",x:tips,y:disps,marker:{{color:"#CBD5E1"}},hovertemplate:"%{{y}} uds<extra>Disponibles</extra>"}}
        ],Object.assign({{}},base,{{barmode:"stack",height:190,showlegend:true,
          legend:{{orientation:"h",x:0,y:1.12,font:{{size:10}}}},
          margin:{{t:30,b:55,l:35,r:10}}}}),{{responsive:true,displayModeBar:false}});
      }} catch(e) {{}}
    }}
  }}

  // Vencimientos mensual chart
  var vmo=PROJ_VENC_MO[projName];
  if(vmo&&vmo.labels&&vmo.labels.length) try {{
    Plotly.newPlot("dvenc-chart",[
      {{type:"bar",x:vmo.labels,y:vmo.values,marker:{{color:"#DC2626",opacity:.85}},
        hovertemplate:"%{{y}} contratos<extra></extra>"}}
    ],Object.assign({{}},base,{{height:190,showlegend:false,
      yaxis:Object.assign({{}},base.yaxis,{{dtick:1}}),
      margin:{{t:10,b:55,l:30,r:10}}}}),{{responsive:true,displayModeBar:false}});
  }} catch(e) {{}}
}}

// ── Export datos ──────────────────────────────────────────────────────────────
function exportProjData(projName) {{
  if(typeof proj_desc==="undefined") return;
  var p=null;
  for(var i=0;i<proj_desc.length;i++) {{if(proj_desc[i].Propiedad===projName){{p=proj_desc[i];break;}}}}
  if(!p) return;
  var rows=[];
  rows.push(["=== RESUMEN: "+projName+" ==="]);
  rows.push(["Indicador","Valor"]);
  rows.push(["Total unidades",Math.round(p.Total)]);
  rows.push(["Arrendadas",Math.round(p.Arrendados)]);
  rows.push(["Disponibles",Math.round(p.Disponibles)]);
  rows.push(["No disponibles",Math.round(p.No_Disp)]);
  rows.push(["Por liberar",Math.round(p.Por_Liberar)]);
  rows.push(["Reservadas",Math.round(p.Reservadas)]);
  rows.push(["Ocupacion %",(p.Pct_Ocup*100).toFixed(1)+"%"]);
  var delta=PROJ_TRENDS[projName];
  if(delta!==null&&delta!==undefined) rows.push(["Tendencia vs mes ant.",delta.toFixed(1)+"pp"]);
  rows.push([""]);
  // Precios
  var pr=PROJ_PRECIOS[projName]||{{}};
  var tips=Object.keys(pr).sort();
  if(tips.length) {{
    rows.push(["=== PRECIOS DESDE (UNIDADES DISPONIBLES) ==="]);
    rows.push(["Tipologia","Precio desde (UF)","Precio desde (CLP)"]);
    tips.forEach(function(tip) {{
      var info=pr[tip]; if(!info) return;
      var uf=info.min, div=info.divisa||"UF";
      var clp=UF_VALOR&&div==="UF"?("$"+Math.round(uf*UF_VALOR).toLocaleString("es-CL")):"";
      rows.push([tip,uf.toFixed(2)+" "+div,clp]);
    }});
    rows.push([""]);
  }}
  // Tipologia
  if(typeof tipo_data!=="undefined"&&tipo_data.projects) {{
    var idx=tipo_data.projects.indexOf(projName);
    if(idx>=0) {{
      rows.push(["=== TIPOLOGIA ==="]);
      rows.push(["Tipologia","Total","Arrendadas","Disponibles","Ocupacion %"]);
      tipo_data.tipologias.forEach(function(tip) {{
        var d=tipo_data[tip]; if(!d) return;
        var tot=d.total[idx]||0; if(!tot) return;
        var arr=d.arrendados[idx]||0;
        rows.push([tip,tot,arr,tot-arr,tot?(arr/tot*100).toFixed(1)+"%":"0%"]);
      }});
      rows.push([""]);
    }}
  }}
  // Vencimientos
  var vRows=document.querySelectorAll("#venc-table tbody tr");
  var vData=[];
  vRows.forEach(function(r){{
    if(r.dataset.proj===projName) vData.push([
      r.cells[0]?r.cells[0].textContent:"",r.cells[1]?r.cells[1].textContent:"",
      r.cells[2]?r.cells[2].textContent:"",r.cells[3]?r.cells[3].textContent:"",
      r.cells[4]?r.cells[4].textContent:"",r.cells[5]?r.cells[5].textContent:""
    ]);
  }});
  if(vData.length) {{
    rows.push(["=== CONTRATOS QUE VENCEN EN 60 DIAS ==="]);
    rows.push(["Proyecto","Unidad","Tipologia","Vencimiento","Dias","Ejecutivo"]);
    vData.forEach(function(r){{rows.push(r);}});
  }}
  var csv=rows.map(function(r){{
    return r.map(function(c){{var s=String(c).replace(/"/g,'""');return s.indexOf(",")>=0||s.indexOf('"')>=0?'"'+s+'"':s;}}).join(",");
  }}).join("\\n");
  var blob=new Blob(["\\uFEFF"+csv],{{type:"text/csv;charset=utf-8"}});
  var url=URL.createObjectURL(blob);
  var a=document.createElement("a"),t=new Date();
  a.href=url;
  a.download="Reporte_"+projName.replace(/[^a-zA-Z0-9]/g,"_")
    +"_"+String(t.getDate()).padStart(2,"0")+String(t.getMonth()+1).padStart(2,"0")+t.getFullYear()+".csv";
  document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
}}

// ── Badge en nav sidebar (proyectos bajo meta) ────────────────────────────────
(function() {{
  if(typeof proj_desc==="undefined") return;
  var nBelow=proj_desc.filter(function(p){{return p.Pct_Ocup<0.95;}}).length;
  if(!nBelow) return;
  document.querySelectorAll(".sbl").forEach(function(item) {{
    if(item.textContent.indexOf("Vista Proyectos")>=0) {{
      var badge=document.createElement("span");
      badge.textContent=nBelow;
      badge.style.cssText="background:#DC2626;color:#fff;border-radius:99px;padding:0 5px;"+
        "font-size:.58rem;margin-left:4px;font-weight:700;vertical-align:middle;line-height:1.6;display:inline-block;";
      item.appendChild(badge);
    }}
  }});
}})();

// ── Refresh sidebar dots cuando cambia meta ──────────────────────────────────
var _PROJ_PREV=null,_PROJ_NEXT=null;
function prevProj(){{if(_PROJ_PREV)selectProj(_PROJ_PREV);}}
function nextProj(){{if(_PROJ_NEXT)selectProj(_PROJ_NEXT);}}

function refreshSidebarStatus() {{
  var mt = typeof META_TARGET !== "undefined" ? META_TARGET : 0.95;
  var nBelow = 0;
  document.querySelectorAll(".proj-list-item").forEach(function(item) {{
    var pct = parseFloat(item.dataset.pct || 0) / 100;
    var isCrit  = pct < 0.85;
    var isBelow = pct < mt;
    var color   = isCrit ? "#DC2626" : isBelow ? "#D97706" : "#16A34A";
    var dot = item.firstElementChild;
    if(dot) {{
      dot.style.background = color;
      dot.className = isCrit ? "proj-dot-pulse" : "";
    }}
    if(isBelow) nBelow++;
  }});
  var footer = document.querySelector("#proj-sidebar-panel div:last-child span:last-child");
  if(footer) footer.textContent = nBelow + " bajo meta";
}}


// ── Excel export ─────────────────────────────────────────────────────────────
function exportToExcel() {{
  if(typeof XLSX==="undefined") {{
    var s=document.createElement("script");
    s.src="https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.mini.min.js";
    s.onload=function(){{_doExport();}};
    document.head.appendChild(s);
  }} else {{ _doExport(); }}
}}
function _doExport() {{
  var rows=[["Proyecto","Tipología","Modelo","Disponibles",
             "Desde UF","Hasta UF","Desde CLP","Hasta CLP",
             "F.Renta Desde UF","F.Renta Hasta UF","F.Renta Desde CLP","F.Renta Hasta CLP"]];
  Object.keys(PROJ_PRECIOS).sort().forEach(function(proj) {{
    var tips=PROJ_PRECIOS[proj]; if(!tips) return;
    Object.keys(tips).sort().forEach(function(tip) {{
      var pr=tips[tip]; if(!pr) return;
      var div=pr.divisa||"UF";
      var mods=(pr.modelos||[]).filter(function(m){{return m.modelo;}});
      var list=mods.length>1?mods:[{{modelo:"",min:pr.min,max:pr.max||pr.min,n_disp:pr.n_disp||0}}];
      list.forEach(function(m) {{
        var mn=m.min, mx=m.max||mn, nd=m.n_disp;
        var mnC=UF_VALOR&&div==="UF"?Math.round(mn*UF_VALOR):0;
        var mxC=UF_VALOR&&div==="UF"?Math.round(mx*UF_VALOR):0;
        rows.push([proj,tip,m.modelo,nd,
          parseFloat(mn.toFixed(2)),parseFloat(mx.toFixed(2)),mnC,mxC,
          parseFloat((mn*3).toFixed(2)),parseFloat((mx*3).toFixed(2)),mnC*3,mxC*3]);
      }});
    }});
  }});
  var ws=XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"]=[{{wch:28}},{{wch:10}},{{wch:18}},{{wch:8}},
               {{wch:10}},{{wch:10}},{{wch:12}},{{wch:12}},
               {{wch:14}},{{wch:14}},{{wch:14}},{{wch:14}}];
  var wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,"Precios y Factor Renta");
  XLSX.writeFile(wb,"LAR_Precios_"+new Date().toISOString().slice(0,10)+".xlsx");
}}
// ── Restaurar estado desde localStorage ──────────────────────────────────────
(function() {{
  var saved=localStorage.getItem("projSort");
  if(saved) {{ var el=document.getElementById("proj-sb-sort"); if(el){{ el.value=saved; sortSidebar(saved); }} }}
  var ufSaved=localStorage.getItem("projUFMode");
  if(ufSaved==="false") {{ _projUFMode=false; _refreshPricePills(); var btn=document.getElementById("proj-uf-btn"); if(btn) btn.innerHTML="UF"; }}
  var lastProj=localStorage.getItem("lastSelProj");
  if(lastProj) setTimeout(function(){{ selectProj(lastProj); }},250);
}})();
"""
    return section, js



def add_extra_features(html, m, hist_data, uf_valor=None):
    """
    Agrega al HTML:
    1. Meta de ocupacion configurable (slider)
    2. Grafico historico real desde contratos (12 meses)
    3. Filtro cruzado: click en proyecto filtra toda la pagina
    4. Modal de detalle por proyecto
    """

    # ── 0. BRAND COLOR — replace residual #008E9F from base template ─────────
    html = html.replace('#008E9F', '#00A8B4')

    # ── 0a. CSS MODERNO — override de estilos para elementos dinámicos ────────
    modern_css = """<style>
/* — Collapsible sections — */
.sec-body{padding:0 0 8px}
.sec-toggle{cursor:pointer;user-select:none}
.sec-toggle:hover .sec{opacity:.85}

/* — Tables (override to modern) — */
#disp-table th,#venc-table th,#res-table th{
  background:transparent;color:#8896A6;font-weight:700;
  font-size:.64rem;text-transform:uppercase;letter-spacing:.07em;
  padding:8px 14px;border-bottom:1.5px solid #EEF2F7;
}
#disp-table td,#venc-table td,#res-table td{
  padding:10px 14px;border-bottom:1px solid #F4F6FA;color:#1A2332;
}
#disp-table tbody tr:hover td,#venc-table tbody tr:hover td,#res-table tbody tr:hover td{
  background:rgba(0,168,180,.03);
}
#disp-table tbody tr:last-child td,#venc-table tbody tr:last-child td{border-bottom:none}

/* — KPI chips dentro de cards — */
.proj-stat{
  background:#F4F6FA;border-radius:8px;
  padding:3px 9px;font-size:.69rem;font-weight:600;color:#4B5A6A;
  border:none;
}

/* — Proyectos: proj-card — */
.proj-card{
  background:#fff;border:none!important;
  border-radius:18px!important;
  padding:16px!important;
  box-shadow:0 1px 3px rgba(0,0,0,.04),0 8px 28px rgba(0,0,0,.06)!important;
  transition:transform .22s ease,box-shadow .22s ease!important;
}
.proj-card:hover{
  transform:translateY(-4px)!important;
  box-shadow:0 4px 8px rgba(0,0,0,.05),0 20px 48px rgba(0,0,0,.1)!important;
}

/* — Barra de progreso de ocupación — */
.proj-bar-track{background:#EEF2F7!important;border-radius:99px!important}
.proj-bar-fill{border-radius:99px!important}

/* — Vista por Proyecto sidebar — */
#proj-sidebar{background:#fff!important;border-radius:16px!important;border:none!important;box-shadow:0 1px 3px rgba(0,0,0,.04),0 8px 28px rgba(0,0,0,.06)!important}
.proj-list-item{border-bottom:1px solid #F4F6FA!important;transition:background .14s!important}
.proj-list-item:hover{background:#F8FAFC!important}
.proj-list-item.selected{background:rgba(0,168,180,.07)!important;border-left:3px solid #00A8B4!important}

/* — Meta bar slider — */
#meta-bar{
  background:#fff;border-radius:14px;padding:14px 20px;
  border:none;box-shadow:0 1px 3px rgba(0,0,0,.04),0 4px 16px rgba(0,0,0,.05);
  margin-bottom:20px;
}

/* — Inputs y selects globales — */
select,input[type=text],input[type=number]{
  background:#F8FAFC;border:1.5px solid #E2E8F0;border-radius:10px;
  padding:7px 11px;font-family:'Aileron',Arial,sans-serif;color:#1A2332;
  font-size:.82rem;transition:border-color .15s;
}
select:focus,input[type=text]:focus,input[type=number]:focus{
  outline:none;border-color:#00A8B4;background:#fff;
}

/* — Botones primarios — */
button[style*="background:#00A8B4"],button[style*="background: #00A8B4"]{
  border-radius:10px!important;font-weight:700!important;
  transition:background .15s,transform .15s,box-shadow .15s!important;
  box-shadow:0 2px 10px rgba(0,168,180,.3)!important;
}

/* — Scrollbar minimalista — */
::-webkit-scrollbar{width:5px;height:5px}
::-webkit-scrollbar-track{background:transparent}
::-webkit-scrollbar-thumb{background:#CBD5E1;border-radius:99px}
::-webkit-scrollbar-thumb:hover{background:#94A3B8}
</style>"""
    html = html.replace('</head>', modern_css + '\n</head>', 1)

    # ── 0b. GATE DE ACCESO ────────────────────────────────────────────────────
    _PWD_HASH = '10ca7afef5a927a199f952212a07aff3a1a33aa98cf3b51afba65fd701c8f0d4'

    # Script en <head>: oculta el body inmediatamente si no hay sesión válida
    head_script = f"""<script>
(function(){{
  if(sessionStorage.getItem('lar_pwd_ok')!=='1'){{
    document.documentElement.style.visibility='hidden';
  }}
}})();
</script>"""
    html = html.replace('</head>', head_script + '\n</head>', 1)

    # Gate como primer hijo del <body>, z-index máximo
    gate_html = f"""<div id="lar-gate" style="
  position:fixed;top:0;left:0;right:0;bottom:0;
  width:100%;height:100%;
  z-index:2147483647;
  background:linear-gradient(135deg,#00A8B4 0%,#007A84 100%);
  display:flex;align-items:center;justify-content:center;
  font-family:'Aileron','Arial',sans-serif;
  visibility:visible">
  <div style="background:#fff;border-radius:20px;padding:40px 44px 36px;
              width:340px;box-shadow:0 24px 64px rgba(0,0,0,.3);text-align:center">
    <img src="static/LAR-logo.png" onerror="this.style.display='none'"
         style="height:48px;margin-bottom:20px;object-fit:contain">
    <div style="font-size:.68rem;font-weight:700;letter-spacing:.12em;
                text-transform:uppercase;color:#9CA3AF;margin-bottom:6px">
      Panel de Ocupaci&oacute;n
    </div>
    <h2 style="font-size:1.15rem;font-weight:800;color:#1E2A38;margin:0 0 24px">
      Acceso restringido
    </h2>
    <input id="gate-pwd" type="password" placeholder="Contrase&ntilde;a"
      onkeydown="if(event.key==='Enter')gateCheck()"
      style="width:100%;box-sizing:border-box;padding:11px 14px;
             border:1.5px solid #E2E8F0;border-radius:10px;font-size:.95rem;
             outline:none;color:#1E2A38;transition:border-color .15s;margin-bottom:10px"
      onfocus="this.style.borderColor='#00A8B4'"
      onblur="this.style.borderColor='#E2E8F0'">
    <div id="gate-err" style="font-size:.75rem;color:#DC2626;min-height:18px;margin-bottom:10px"></div>
    <button onclick="gateCheck()"
      style="width:100%;padding:13px;background:#00A8B4;color:#fff;border:none;
             border-radius:10px;font-size:.95rem;font-weight:700;cursor:pointer"
      onmouseover="this.style.background='#007A84'"
      onmouseout="this.style.background='#00A8B4'">
      Ingresar
    </button>
    <div style="font-size:.67rem;color:#CBD5E1;margin-top:16px">LAR Group &mdash; Uso interno</div>
  </div>
</div>

<script>
async function gateCheck() {{
  var pwd = (document.getElementById('gate-pwd').value || '').trim();
  var err = document.getElementById('gate-err');
  if (!pwd) {{ err.textContent = 'Ingresa la contraseña'; return; }}
  var enc = new TextEncoder();
  var buf = await crypto.subtle.digest('SHA-256', enc.encode(pwd));
  var hex = Array.from(new Uint8Array(buf)).map(function(b){{return b.toString(16).padStart(2,'0');}}).join('');
  if (hex === '{_PWD_HASH}') {{
    sessionStorage.setItem('lar_pwd_ok', '1');
    document.documentElement.style.visibility = '';
    var g = document.getElementById('lar-gate');
    g.style.opacity = '0'; g.style.transition = 'opacity .3s';
    setTimeout(function(){{ g.style.display = 'none'; }}, 320);
  }} else {{
    err.textContent = 'Contraseña incorrecta';
    document.getElementById('gate-pwd').value = '';
    document.getElementById('gate-pwd').focus();
  }}
}}
// Si ya autenticado: revelar página y ocultar gate
(function(){{
  if (sessionStorage.getItem('lar_pwd_ok') === '1') {{
    document.documentElement.style.visibility = '';
    var g = document.getElementById('lar-gate');
    if (g) g.style.display = 'none';
  }}
}})();
</script>
"""
    # Insertar gate como PRIMER hijo del body
    html = html.replace('<body>', '<body>\n' + gate_html, 1)

    # ── 1. META SLIDER — antes del primer KPI grid ───────────────────────────
    meta_html = """
<div id="meta-bar" style="display:flex;align-items:center;gap:12px;margin:16px 0 8px;
     flex-wrap:wrap;padding:10px 14px;background:#F0FDFE;border-radius:8px;
     border:1px solid #BAE6ED">
  <span style="font-size:.75rem;color:#6B7A8D;font-weight:600">&#127919; Meta de ocupaci&oacute;n:</span>
  <input type="range" id="meta-slider" min="80" max="100" step="1" value="95"
         oninput="actualizarMeta(this.value)"
         style="width:140px;accent-color:#00A8B4;cursor:pointer">
  <span id="meta-val" style="font-size:.9rem;font-weight:700;color:#00A8B4">95%</span>
  <span id="meta-feedback" style="font-size:.68rem;color:#0369A1;background:#EFF6FF;border-radius:99px;padding:2px 9px;font-weight:600;display:none"></span>
  <span style="font-size:.71rem;color:#9CA3AF">| La tabla de alertas se recalcula en tiempo real</span>
</div>
"""
    hero_html = build_hero_section(m, uf_valor=uf_valor, hist_data=hist_data)
    # Actualizar texto de fuente de datos
    html = html.replace("Fuente: unidades.xlsx", "Fuente: API PostgreSQL")
    html = html.replace("__DATE__", DATE_STR)
    html = html.replace("Target global: 95%", "Target global: <span id=\"ft-meta-pct\">95</span>%")
    kg_match = re.search(r'<div class="kg"', html)
    if kg_match:
        html = html[:kg_match.start()] + hero_html + meta_html + html[kg_match.start():]

    # ── 2. GRAFICO HISTORICO REAL — antes de sec-alertas ────────────────────
    if hist_data is not None and len(hist_data) > 0:
        hist_meses_js = json.dumps(hist_data['mes'].tolist())
        hist_pcts_js  = json.dumps(hist_data['pct'].tolist())
        hist_section  = f"""
<div style="font-size:.78rem;font-weight:700;text-transform:uppercase;letter-spacing:.06em;
     color:#1E2A38;border-bottom:2px solid #00A8B4;padding-bottom:4px;margin:24px 0 12px">
  Tendencia Hist&oacute;rica de Ocupaci&oacute;n &mdash; 12 meses reales
</div>
<div class="cf" style="margin-bottom:20px">
  <div id="hist-ocup-chart" style="height:290px"></div>
</div>
"""
        html = html.replace('<div id="sec-alertas"',
                            hist_section + '<div id="sec-alertas"', 1)
    else:
        hist_meses_js = "[]"
        hist_pcts_js  = "[]"

    # ── 3. MODAL HTML — antes de </body> ────────────────────────────────────
    modal_html = """
<div id="proj-modal" style="display:none;position:fixed;inset:0;z-index:2000;
     background:rgba(0,0,0,.45);justify-content:center;align-items:center;"
     onclick="if(event.target===this)closeModal()">
  <div style="background:#fff;border-radius:14px;padding:28px 28px 22px;max-width:640px;
              width:92%;max-height:84vh;overflow-y:auto;position:relative;
              box-shadow:0 20px 60px rgba(0,0,0,.28)">
    <button onclick="closeModal()" style="position:absolute;top:14px;right:16px;background:none;
            border:none;font-size:1.1rem;cursor:pointer;color:#6B7A8D;line-height:1">&#x2715;</button>
    <div id="proj-modal-content"></div>
  </div>
</div>
<div id="proj-filter-bar" style="display:none;position:sticky;top:0;z-index:900;
     background:#EFF6FF;border-bottom:2px solid #93C5FD;padding:7px 16px;
     display:none;align-items:center;gap:10px;flex-wrap:wrap;font-size:.76rem">
  <span>&#128204; Filtrando dashboard por proyecto:</span>
  <b id="proj-filter-name" style="color:#1D4ED8"></b>
  <button onclick="setProjFilter(null)"
          style="padding:2px 10px;background:#fff;border:1px solid #93C5FD;
                 border-radius:99px;font-size:.71rem;cursor:pointer;color:#1D4ED8">
    &#10005; Limpiar filtro
  </button>
</div>
"""
    html = html.replace('</body>', modal_html + '\n</body>', 1)

    # ── 4. JS interactivo ────────────────────────────────────────────────────
    interactive_js = f"""
<script>
// ── Historico real ───────────────────────────────────────────────────────────
var hist_meses = {hist_meses_js};
var hist_pcts  = {hist_pcts_js};
if(hist_meses.length && document.getElementById("hist-ocup-chart")) {{
  var metaLine = hist_meses.map(function(){{return 95;}});
  Plotly.newPlot("hist-ocup-chart",[
    {{type:"scatter",mode:"lines+markers",name:"Ocupacion %",
     x:hist_meses,y:hist_pcts,
     line:{{color:"#00A8B4",width:2.5}},
     marker:{{color:hist_pcts.map(function(v){{return v>=95?"#16A34A":v>=85?"#D97706":"#DC2626";}}),size:7}},
     fill:"tozeroy",fillcolor:"rgba(0,142,159,.08)",
     hovertemplate:"Mes %{{x}}: <b>%{{y:.1f}}%</b><extra></extra>"
    }},
    {{type:"scatter",mode:"lines",name:"Meta 95%",
     x:hist_meses,y:metaLine,
     line:{{color:"#DC2626",dash:"dot",width:1.5}},
     hoverinfo:"skip"
    }}
  ],{{...base,
    title:{{text:"Ocupacion mensual real — contratos activos / unidades DB (12 proyectos LAR)",font:{{size:13}}}},
    xaxis:{{gridcolor:GRID,color:AXIS,tickangle:-35}},
    yaxis:{{gridcolor:GRID,color:AXIS,title:"%",range:[65,100]}},
    showlegend:true,legend:{{orientation:"h",y:-0.28}},
    margin:{{t:50,b:110,l:55,r:20}},height:290
  }});
}}

// ── Meta configurable ────────────────────────────────────────────────────────
var META_TARGET = 0.95;
function actualizarMeta(v) {{
  META_TARGET = v / 100;
  document.getElementById("meta-val").textContent = v + "%";
  // Live feedback
  if(typeof proj_desc !== "undefined") {{
    var nCumple = proj_desc.filter(function(p){{return p.Pct_Ocup >= META_TARGET;}}).length;
    var fb = document.getElementById("meta-feedback");
    if(fb) {{ fb.style.display=""; fb.textContent = nCumple+"/"+proj_desc.length+" proyectos cumplen"; }}
    var hBM = document.getElementById("hero-bajo-meta");
    var hMP = document.getElementById("hero-meta-pct");
    if(hBM) hBM.textContent = proj_desc.length - nCumple;
    if(hMP) hMP.textContent = Math.round(v);
  }}
  recalcAlertas();
  if(typeof refreshSidebarStatus!=="undefined") refreshSidebarStatus();
  // Actualizar linea de meta en historico
  var el = document.getElementById("hist-ocup-chart");
  if(el && el.data && el.data.length > 1) {{
    Plotly.restyle("hist-ocup-chart", {{y: [hist_meses.map(function(){{return parseFloat(v);}})] }}, [1]);
  }}
  // Actualizar Target en bar_proj
  var bEl = document.getElementById("bar_proj");
  if(bEl && bEl.data && bEl.data.length > 1) {{
    var tv = parseFloat(v);
    var tName = "Target " + Math.round(v) + "%";
    var tHover = "Target: " + Math.round(v) + "%<extra></extra>";
    Plotly.restyle("bar_proj", {{
      y: [bEl.data[0].x.map(function(){{ return tv; }})],
      name: [tName],
      "hovertemplate": [tHover]
    }}, [1]);
  }}
  // Actualizar footer meta
  var ftmp = document.getElementById("ft-meta-pct");
  if(ftmp) ftmp.textContent = Math.round(v);
  // Actualizar Brecha Meta en KPI grid
  var bGap = document.getElementById("kpi-brecha-meta");
  var bTgt = document.getElementById("kpi-brecha-target");
  if(bGap) {{
    var gap = (gpct * 100 - parseFloat(v));
    bGap.textContent = gap.toFixed(1) + "pp";
    bGap.style.color = gap >= 0 ? "#16A34A" : "#DC2626";
  }}
  if(bTgt) bTgt.textContent = Math.round(v);
}}

function globalSearch(q) {{
  q = (q||"").toLowerCase().trim();
  var cnt = 0;
  // Sidebar items
  document.querySelectorAll(".proj-list-item").forEach(function(el) {{
    var show = !q || (el.dataset.name||"").toLowerCase().includes(q);
    el.style.display = show ? "" : "none";
    if(show) cnt++;
  }});
  // Overview cards
  document.querySelectorAll("[data-proj-card]").forEach(function(el) {{
    el.style.display = (!q || (el.dataset.projCard||"").toLowerCase().includes(q)) ? "" : "none";
  }});
  // All table rows (first td)
  document.querySelectorAll("table tbody tr").forEach(function(r) {{
    var td = r.querySelector("td"); var text = td ? td.textContent.toLowerCase() : "";
    r.style.display = (!q || text.includes(q)) ? "" : "none";
  }});
  var sc = document.getElementById("search-count");
  if(sc) sc.textContent = q ? cnt+" proyecto"+(cnt!==1?"s":"") : "";
}}

function recalcAlertas() {{
  if(typeof proj_desc === "undefined") return;
  var tbody = document.getElementById("alerts-tbody");
  if(!tbody) return;
  var sorted = proj_desc.slice().sort(function(a,b){{ return a.Pct_Ocup - b.Pct_Ocup; }});
  var rows = "";
  sorted.forEach(function(p) {{
    // Aplicar filtro de proyecto si esta activo
    if(PROJ_FILTER && p.Propiedad !== PROJ_FILTER) return;
    var pct  = (p.Pct_Ocup * 100).toFixed(1);
    var gap  = ((p.Pct_Ocup - META_TARGET) * 100);
    var needed = Math.max(0, Math.round(META_TARGET * p.Total) - p.Arrendados);
    var badge, gapStr, neededStr;
    if(p.Pct_Ocup >= META_TARGET) {{
      badge = "<span class='badge badge-green'>&#x2705; Cumple</span>";
      gapStr = "&mdash;"; neededStr = "&mdash;";
    }} else if(p.Pct_Ocup >= META_TARGET - 0.10) {{
      badge = "<span class='badge badge-orange'>&#x1F7E1; En seguimiento</span>";
      gapStr = "+" + Math.abs(gap).toFixed(1) + "pp";
      neededStr = needed + " uds.";
    }} else {{
      badge = "<span class='badge badge-red'>&#x1F534; Cr&iacute;tico</span>";
      gapStr = "+" + Math.abs(gap).toFixed(1) + "pp";
      neededStr = needed + " uds.";
    }}
    var dotCls = p.Pct_Ocup >= META_TARGET ? "sem-g" : p.Pct_Ocup >= (META_TARGET-0.1) ? "sem-y" : "sem-r";
    var projEsc = p.Propiedad.replace(/'/g, "\\\\'");
    rows += "<tr style='cursor:pointer' onclick='showProjModal(\\\"" + p.Propiedad.replace(/"/g,"&quot;") + "\\\")'>" +
      "<td style='font-weight:500'><span class='sem " + dotCls + "' title='" + pct + "%'></span>" + p.Propiedad + "</td>" +
      "<td><b>" + pct + "%</b></td><td>" + gapStr + "</td>" +
      "<td>" + neededStr + "</td><td>" + badge + "</td></tr>";
  }});
  tbody.innerHTML = rows;
}}

// ── Filtro cruzado por proyecto ──────────────────────────────────────────────
var PROJ_FILTER = null;
function setProjFilter(proj) {{
  PROJ_FILTER = proj;
  var bar  = document.getElementById("proj-filter-bar");
  var name = document.getElementById("proj-filter-name");
  if(bar) {{ bar.style.display = proj ? "flex" : "none"; }}
  if(name && proj) name.textContent = proj;

  // Filtrar tablas (todas menos la de vencimientos que tiene su propio filtro)
  document.querySelectorAll("table:not(#venc-table) tbody tr").forEach(function(r) {{
    if(!proj) {{ r.style.display = ""; return; }}
    var td = r.querySelector("td");
    var nm = td ? td.textContent.replace(/^\\s*[●○◐▸]\\s*/, "").trim() : "";
    r.style.display = (nm === proj || r.dataset.proj === proj) ? "" : "none";
  }});

  // Sincronizar filtro de proyecto en tabla vencimientos
  var vfProj = document.getElementById("vf-proj");
  if(vfProj) {{ vfProj.value = proj || ""; if(typeof filtrarVenc === "function") filtrarVenc(); }}

  recalcAlertas();
}}

// Conectar graficos de barras al filtro cruzado (post-render)
window.addEventListener("load", function() {{
  setTimeout(function() {{
    ["pol_bar_proj","res_bar_proj"].forEach(function(divId) {{
      var el = document.getElementById(divId);
      if(!el || !el.on) return;
      el.on("plotly_click", function(d) {{
        if(!d.points || !d.points.length) return;
        var proj = d.points[0].y;
        setProjFilter(proj === PROJ_FILTER ? null : proj);
      }});
    }});
    recalcAlertas();
  }}, 1500);
}});

// ── Modal de proyecto ────────────────────────────────────────────────────────
function showProjModal(projName) {{
  if(typeof proj_desc === "undefined") return;
  var p = null;
  for(var i=0;i<proj_desc.length;i++) {{ if(proj_desc[i].Propiedad===projName){{ p=proj_desc[i]; break; }} }}
  if(!p) return;

  var pct  = (p.Pct_Ocup * 100).toFixed(1);
  var gap  = ((p.Pct_Ocup - META_TARGET) * 100).toFixed(1);
  var tgt  = (META_TARGET * 100).toFixed(0);
  var clr  = p.Pct_Ocup >= META_TARGET ? "#16A34A" : p.Pct_Ocup >= META_TARGET-0.1 ? "#D97706" : "#DC2626";
  var badge= p.Pct_Ocup >= META_TARGET ? "&#x2705; Cumple meta" :
             p.Pct_Ocup >= META_TARGET-0.1 ? "&#x1F7E1; En seguimiento" : "&#x1F534; Critico";

  // Vencimientos de este proyecto
  var vRows = document.querySelectorAll("#venc-table tbody tr");
  var vList = [];
  vRows.forEach(function(r) {{
    if(r.dataset.proj === projName) {{
      vList.push(parseInt(r.dataset.dias) + "d — " + (r.cells[1]||{{}}).textContent + " (" + (r.cells[2]||{{}}).textContent + ")");
    }}
  }});

  var html2 = "<div style='border-bottom:3px solid #00A8B4;padding-bottom:8px;margin-bottom:18px'>" +
    "<div style='font-size:.6rem;color:#6B7A8D;text-transform:uppercase;letter-spacing:.12em;font-weight:700'>PROYECTO</div>" +
    "<h2 style='margin:4px 0 0;font-size:1.2rem;color:#1E2A38'>" + projName + "</h2></div>" +
    "<div style='display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:16px'>" +
    mkCard("% Ocupacion", pct+"%", clr) +
    mkCard("Arrendadas", p.Arrendados, "#16A34A") +
    mkCard("Disponibles", p.Disponibles, "#D97706") +
    mkCard("No Disponibles", p.No_Disp, "#DC2626") +
    mkCard("Por Liberar", p.Por_Liberar, "#DC2626") +
    mkCard("Reservadas", p.Reservadas, "#0369A1") +
    "</div>" +
    "<div style='background:#F8FAFC;border-radius:8px;padding:11px 14px;margin-bottom:14px'>" +
    "<div style='font-size:.72rem;font-weight:700;color:#1E2A38;margin-bottom:5px'>Estado vs meta " + tgt + "%</div>" +
    "<div style='font-size:.81rem'>" + badge + " &mdash; Gap: <b style='color:" + clr + "'>" + (gap>=0?"+":"") + gap + "pp</b></div>" +
    (p.Uds_Needed>0 ? "<div style='font-size:.74rem;color:#6B7A8D;margin-top:3px'>Necesita <b>" + Math.round(p.Uds_Needed) + " uds.</b> mas para alcanzar la meta</div>" : "") +
    "</div>";

  if(vList.length > 0) {{
    html2 += "<div style='margin-bottom:14px'>" +
      "<div style='font-size:.72rem;font-weight:700;color:#1E2A38;margin-bottom:5px'>&#128197; Contratos que vencen pronto (" + vList.length + ")</div>" +
      "<div style='max-height:140px;overflow-y:auto;font-size:.72rem;color:#374151'>" +
      vList.slice(0,20).map(function(v){{return "<div style='padding:3px 0;border-bottom:1px solid #F1F5F9'>"+v+"</div>";}}).join("") +
      (vList.length>20?"<div style='color:#9CA3AF;padding-top:4px'>+ "+(vList.length-20)+" mas...</div>":"") +
      "</div></div>";
  }}

  var projEsc = projName.replace(/"/g,"&quot;");
  html2 += "<div style='display:flex;gap:8px;justify-content:flex-end;margin-top:8px'>" +
    "<button onclick='closeModal()' style='padding:6px 14px;background:#F8FAFC;color:#6B7A8D;border:1px solid #E2E8F0;border-radius:6px;font-size:.74rem;cursor:pointer'>Cerrar</button>" +
    "<button onclick='closeModal();setProjFilter(\\\"" + projEsc + "\\\")' " +
    "style='padding:6px 14px;background:#00A8B4;color:#fff;border:none;border-radius:6px;font-size:.74rem;cursor:pointer'>" +
    "&#128204; Filtrar por este proyecto</button></div>";

  document.getElementById("proj-modal-content").innerHTML = html2;
  document.getElementById("proj-modal").style.display = "flex";
  document.body.style.overflow = "hidden";
}}

function mkCard(label, val, color) {{
  return "<div style='background:#F8FAFC;border-radius:8px;padding:9px 12px;border-left:3px solid "+color+"'>" +
    "<div style='font-size:.62rem;color:#6B7A8D;font-weight:600'>" + label + "</div>" +
    "<div style='font-size:1.3rem;font-weight:700;color:" + color + "'>" + val + "</div></div>";
}}

function closeModal() {{
  var m = document.getElementById("proj-modal");
  if(m) m.style.display = "none";
  document.body.style.overflow = "";
}}

// ── Botón Actualizar ─────────────────────────────────────────────────────────
function actualizarDatos() {{
  var token = sessionStorage.getItem('lar_gh_pat');
  var btn   = document.getElementById('btn-actualizar');
  var lbl   = document.getElementById('btn-act-label');
  var icon  = document.getElementById('btn-act-icon');

  if (!token) {{
    alert('Token de acceso no disponible. Recarga la página e inicia sesión.');
    return;
  }}

  // Estado cargando
  if (btn) btn.disabled = true;
  if (lbl) lbl.textContent = 'Actualizando...';
  if (icon) {{ icon.style.animation = 'spin 1s linear infinite'; icon.style.display = 'inline-block'; }}

  fetch('https://api.github.com/repos/MatiasStipicevic/MatiCode/actions/workflows/update.yml/dispatches', {{
    method: 'POST',
    headers: {{
      'Authorization': 'token ' + token,
      'Accept': 'application/vnd.github+json',
      'Content-Type': 'application/json'
    }},
    body: JSON.stringify({{ ref: 'main' }})
  }})
  .then(function(res) {{
    if (res.status === 204) {{
      if (lbl) lbl.textContent = '✓ Workflow iniciado';
      if (icon) icon.style.animation = '';
      setTimeout(function() {{
        if (lbl) lbl.textContent = 'Actualizar';
        if (btn) btn.disabled = false;
      }}, 4000);
    }} else {{
      return res.json().then(function(j) {{ throw new Error(j.message || 'Error ' + res.status); }});
    }}
  }})
  .catch(function(err) {{
    if (lbl) lbl.textContent = '✗ Error: ' + err.message;
    if (icon) icon.style.animation = '';
    if (btn) btn.disabled = false;
    setTimeout(function() {{ if (lbl) lbl.textContent = 'Actualizar'; }}, 5000);
  }});
}}

// ── Collapsible sections ─────────────────────────────────────────────────────
var _SEC_DEFAULTS_OPEN = ['sec-ocupacion','sec-disponibilidad','sec-proyectos'];

function toggleSec(id) {{
  var body = document.getElementById('body-' + id);
  var header = document.getElementById(id);
  if (!body || !header) return;
  var isOpen = body.style.display !== 'none';
  if (isOpen) {{
    body.style.display = 'none';
    header.classList.add('collapsed');
    localStorage.setItem('sec-' + id, 'closed');
  }} else {{
    body.style.display = '';
    header.classList.remove('collapsed');
    localStorage.setItem('sec-' + id, 'open');
    // Resize Plotly charts after reveal
    setTimeout(function() {{
      body.querySelectorAll('div[id]').forEach(function(el) {{
        try {{ if(el._fullLayout) Plotly.Plots.resize(el); }} catch(e) {{}}
      }});
    }}, 80);
  }}
}}

window.addEventListener('load', function() {{
  document.querySelectorAll('.sec[id]').forEach(function(header) {{
    var id = header.id;
    if (!id) return;

    // Add chevron if not already present (sec-proyectos has it inline)
    if (!header.querySelector('.sec-chevron')) {{
      var chev = document.createElement('span');
      chev.className = 'sec-chevron';
      chev.innerHTML = '&#9650;';
      header.appendChild(chev);
    }}

    // Collect next siblings until the next .sec[id]
    var siblings = [];
    var el = header.nextElementSibling;
    while (el) {{
      if (el.classList && el.classList.contains('sec') && el.id) break;
      siblings.push(el);
      el = el.nextElementSibling;
    }}
    if (siblings.length === 0) return;

    // Wrap siblings in a body div
    var body = document.createElement('div');
    body.id = 'body-' + id;
    body.className = 'sec-body';
    header.parentNode.insertBefore(body, header.nextSibling);
    siblings.forEach(function(s) {{ body.appendChild(s); }});

    // Click handler (only on the header, not children via delegation)
    header.addEventListener('click', function(e) {{
      // Don't toggle if clicked on a button inside the header (e.g. export)
      if (e.target !== header && e.target.closest && e.target.closest('button')) return;
      toggleSec(id);
    }});

    // Initial state from localStorage or defaults
    var saved = localStorage.getItem('sec-' + id);
    var open = saved !== null ? saved === 'open' : _SEC_DEFAULTS_OPEN.indexOf(id) >= 0;
    if (!open) {{
      body.style.display = 'none';
      header.classList.add('collapsed');
    }}
  }});
}});
</script>
"""
    html = html.replace('</body>', interactive_js + '\n</body>', 1)
    return html


# ── 7. MAIN ────────────────────────────────────────────────────────────────
def build_historial_json(m, html_current):
    """Acumula snapshots semanales. Agrega snapshot cada lunes."""
    import datetime
    today = datetime.date.today()
    monday = today - datetime.timedelta(days=today.weekday())
    monday_str = monday.strftime("%d/%m/%Y")

    prev = []
    hist_m = re.search(r'var historial_data\s*=\s*(\[[\s\S]*?\]);', html_current)
    if hist_m:
        try:
            prev = json.loads(hist_m.group(1))
        except Exception:
            prev = []

    existing = [e.get("Fecha", "") for e in prev]
    week_ok = any(
        abs((pd.to_datetime(d, dayfirst=True).date() - monday).days) < 7
        for d in existing if d
    )

    if today.weekday() == 0 or not week_ok:
        if monday_str not in existing:
            snap = {"Fecha": monday_str}
            for p in m.get("proj_desc", []):
                snap[p["Propiedad"]] = round(float(p["Pct_Ocup"]) * 100, 2)
            snap["Global"] = round(m.get("ocup_global", 0) * 100, 2)
            prev.append(snap)
            if len(prev) > 52:
                prev = prev[-52:]
            print(f"  Historial: snapshot {monday_str} ({len(prev)} semanas)")
        else:
            print(f"  Historial: {monday_str} ya registrado")
    else:
        nxt = (monday + datetime.timedelta(7)).strftime("%d/%m/%Y")
        print(f"  Historial: {len(prev)} entradas — próx. lunes {nxt}")

    return json.dumps(prev, ensure_ascii=False)


def apply_redesign(html, m):
    # Visual redesign is handled by update_html and add_extra_features
    return html


def main():
    print("Cargando datos...")
    try:
        df = load_data()
        print(f"  Fuente: PostgreSQL + Excel | Departamentos: {len(df)}")
    except Exception as e:
        print(f"  DB no disponible ({e}), usando Excel como fallback...")
        df = load_data_excel_fallback()
        print(f"  Fuente: Excel | Departamentos: {len(df)}")

    vencs = None
    renov = None
    try:
        print("Cargando contratos (vencimientos)...")
        vencs = load_contratos()
        print(f"  Contratos activos: {len(vencs)} | Vencen 90d: {vencs['dias'].between(0,90).sum()}")
        renov = load_renovaciones()
        print(f"  Renovaciones: 30d={renov['r30']} | 60d={renov['r60']} | 90d={renov['r90']}")
    except Exception as e:
        print(f"  Contratos no disponibles ({e}), omitiendo seccion vencimientos...")

    hist_data = None
    try:
        n_db = int((df["_proj"] != "Collective Bustamante").sum())
        print(f"Cargando historico de ocupacion ({n_db} unidades DB)...")
        hist_data = load_historico(n_db)
        print(f"  Historico: {len(hist_data)} meses | ultima ocup: {hist_data['pct'].iloc[-1]}%")
    except Exception as e:
        print(f"  Historico no disponible ({e}), omitiendo...")

    print("Calculando metricas...")
    m = compute(df)
    print(f"  Total: {m['total']} | Arrendadas: {m['arrendadas']} | Ocup: {round(m['ocup_global'],1)}%")
    print(f"  Por Liberar: {m['por_liberar']} | Por Renovar: {m['por_renovar']}")

    print("Leyendo HTML fuente...")
    html = HTML_SRC.read_text(encoding="utf-8")

    print("Actualizando datos...")
    html = update_html(html, m)

    print("Aplicando rediseno visual...")
    html = apply_redesign(html, m)

    print("Agregando seccion Reservadas...")
    res_section, res_js = build_res_section(m)
    marker = '<div class="sec">Comparador Semanal</div>'
    html = html.replace(marker, res_section + marker)

    print("Agregando seccion Por Liberar...")
    pol_section, pol_js = build_pol_section(m)
    # Insertar antes del Comparador (Por Liberar va después de Reservadas)
    html = html.replace(marker, pol_section + marker)

    # JS de ambas secciones + irA antes de </body>
    js_block = (
        f"<script>\n{res_js}\n{pol_js}\n"
        f"function irA(id){{var el=document.getElementById(id);if(!el)return;var b=document.getElementById('body-'+id);if(b&&b.style.display==='none')toggleSec(id);el.scrollIntoView({{behavior:'smooth',block:'start'}});if(typeof closeSb==='function')closeSb();}}\n"
        f"</script>"
    )
    html = html.replace("</body>", js_block + "\n</body>")

    # ── Agregar sección Vista por Proyecto ───────────────────────────────
    print("Agregando seccion Vista por Proyecto...")
    precios    = load_precios_disponibles()
    uf_valor   = load_uf()
    tendencias = load_tendencias_proyectos()
    proj_section, proj_js = build_projects_section(
        m, vencs, precios=precios, uf_valor=uf_valor, tendencias=tendencias)
    html = html.replace('<div class="sec">Comparador Semanal</div>',
                        proj_section + '<div class="sec">Comparador Semanal</div>')
    html = html.replace("</body>", "<script>\n" + proj_js + "\n</script>\n</body>", 1)

    # ── Agregar sección Vencimientos ──────────────────────────────────────
    if vencs is not None:
        print("Agregando seccion Vencimientos...")
        venc_section, venc_js = build_vencimientos_section(vencs, renov=renov)
        # Insertar antes del Comparador Semanal (antes que Reservadas y PoL)
        html = html.replace('<div class="sec">Comparador Semanal</div>',
                            venc_section + '<div class="sec">Comparador Semanal</div>')
        html = html.replace("</body>", f"<script>\n{venc_js}\n</script>\n</body>", 1)

    # ── Añadir id a sección Comparador Semanal (después de todas las inserciones) ──
    html = html.replace('<div class="sec">Comparador Semanal</div>',
                        '<div id="sec-comparador" class="sec">Comparador Semanal</div>')

    # ── Eliminar gate (clave de ingreso) ─────────────────────────────────
    print("Eliminando gate de acceso...")
    # Extraer token cifrado ANTES de remover el gate
    _tok_m = re.search(r"var enc='([A-Za-z0-9+/=]+)',key='Lar2026'", html)
    # Remover bloque gate HTML completo
    html = re.sub(r'<!-- ── Gate de acceso ── -->.*?</div>\s*</div>\s*</div>',
                  '', html, flags=re.DOTALL)
    # Remover gate JS (sessionStorage check)
    html = re.sub(r'<script>\s*\(function\(\)\{\s*if\(sessionStorage.*?\}\)\(\);\s*</script>',
                  '', html, flags=re.DOTALL)
    # Remover script que copia logo al gate
    html = re.sub(r'<script>\s*// Copia el src.*?</script>', '', html, flags=re.DOTALL)
    # Mostrar sidebar al cargar (reemplaza el hook del gate)
    html = html.replace("gate.style.display='none'; openSb();",
                        "openSb();")
    # Sidebar visible al abrir la página
    html = html.replace('window.addEventListener(\'load\',function(){',
                        'window.addEventListener(\'load\',function(){ openSb();')
    # Re-inyectar descifrado de token como script autónomo (sobrevive en deployed HTML)
    if _tok_m:
        _e = _tok_m.group(1)
        _tok_s = (
            "<script id=\"lar-tok\">(function(){var e='" + _e +
            "',k='Lar2026',b=atob(e),o='';"
            "for(var i=0;i<b.length;i++)o+=String.fromCharCode(b.charCodeAt(i)^k.charCodeAt(i%k.length));"
            "sessionStorage.setItem('lar_gh_pat',o);})()</script>"
        )
        html = html.replace('</body>', _tok_s + '\n</body>', 1)

    # ── Agregar sección Disponibilidad y Ocupación ────────────────────────
    print("Agregando seccion Disponibilidad y Ocupacion...")
    disp_section, disp_js = build_disponibilidad_table(df, m)
    html = html.replace('<div id="sec-alertas"',
                        disp_section + '<div id="sec-alertas"', 1)
    html = html.replace("</body>", f"<script>\n{disp_js}\n</script>\n</body>", 1)

    # ── Agregar sección Días de Vacancia ──────────────────────────────────
    print("Calculando dias de vacancia...")
    vacancia_rows = load_vacancia()
    vac_section = build_vacancia_section(vacancia_rows)
    if vac_section:
        html = html.replace('<div id="sec-disponibilidad"',
                            vac_section + '<div id="sec-disponibilidad"', 1)

    print("Agregando features interactivos (meta, historico, modal, cross-filter)...")
    html = add_extra_features(html, m, hist_data, uf_valor=uf_valor)

    # Acumular snapshot semanal de ocupación (historial_data JS var)
    try:
        html_ref = HTML_REPO.read_text(encoding="utf-8") if HTML_REPO.exists() else html
        html = js_replace(html, "historial_data", build_historial_json(m, html_ref))
    except Exception as e:
        print(f"  historial_data: omitido ({e})")

    print("Guardando archivos...")
    HTML_OUT.write_text(html, encoding="utf-8")
    HTML_BACKUP.write_text(html, encoding="utf-8")
    HTML_REPO.write_text(html, encoding="utf-8")
    print(f"  OK -> {HTML_OUT.name}")
    print(f"  OK -> {HTML_BACKUP}")
    print(f"  OK -> {HTML_REPO}")

if __name__ == "__main__":
    main()
