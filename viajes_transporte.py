import streamlit as st
import psycopg2
import pandas as pd
from datetime import datetime, timedelta, time, date
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pytz

# ==================== CONFIGURACIÓN ====================
st.set_page_config(
    page_title="Control de Viajes",
    layout="wide",
    page_icon="🚚",
    initial_sidebar_state="collapsed"
)

# ==================== CREDENCIALES ====================
SUPABASE_DB_URL = "postgresql://postgres.hhzuggxvdzzfmnvfulmp:Negritasantia@aws-1-us-east-1.pooler.supabase.com:6543/postgres"

# ==================== CATÁLOGO PLACAS / CONDUCTORES ====================
PLACA_CONDUCTOR = {
    "NOX459": "HABID CAMACHO",
    "NOX460": "JOSE ORTEGA PEREZ",
    "NOX461": "CARLOS TAFUR",
    "SON047": "ISAIAS VESGA",
    "SON048": "FLAVIO ROSENDO MALTE TUTALCHA",
    "SOP148": "SLITH JOSE ORTEGA PACHECO",
    "SOP149": "ABRAHAM SEGUNDO ALVAREZ VALLE",
    "SOP150": None,
    "SRO661": "JULIAN CALETH CORONADO",
    "SRO672": "PEDRO VILLAMIL",
    "TMW882": "JESUS DAVID MONTES MOSQUERA",
    "TRL282": "CHRISTIAN MARTINEZ NAVARRO",
    "TRL298": "YEIMI DUQUE ZULUAGA",
    "UYQ308": "REIMUR MANUEL",
    "UYV084": "RAMON TAFUR HERNANDEZ",
    "UYY788": "EDUARDO RAFAEL OLIVARES ALCAZAR",
    "PSX350": "EDGAR DE JESUS RAMIREZ",
}

TODOS_CONDUCTORES = sorted([
    "REIMUR MANUEL", "HABID CAMACHO", "JOSE ORTEGA PEREZ", "CARLOS TAFUR",
    "ISAIAS VESGA", "FLAVIO ROSENDO MALTE TUTALCHA", "SLITH JOSE ORTEGA PACHECO",
    "ABRAHAM SEGUNDO ALVAREZ VALLE", "RAMON TAFUR HERNANDEZ", "JULIAN CALETH CORONADO",
    "PEDRO VILLAMIL", "JESUS DAVID MONTES MOSQUERA", "CHRISTIAN MARTINEZ NAVARRO",
    "YEIMI DUQUE ZULUAGA", "EDGAR DE JESUS RAMIREZ", "EDUARDO RAFAEL OLIVARES ALCAZAR",
])

ESTADOS_VIAJE = ["✅ Completado", "❌ Anulado", "⚠️ Incumplido", "🔄 En Curso"]

RUTAS_FRECUENTES = [
    ("PUERTO PALERMO", "AGOFER"), ("PUERTO BARRANQUILLA", "VIA40"),
    ("PUERTO BARRANQUILLA", "PROCAR"), ("PUERTO BARRANQUILLA", "CIENAGA"),
    ("PUERTO BARRANQUILLA", "MEICO"), ("PUERTO BARRANQUILLA", "MEICO CIRCUNVALAR"),
    ("PUERTO BARRANQUILLA", "SOLEDAD"), ("PUERTO PALERMO", "ZF BAQ"),
    ("PUERTO BARRANQUILLA", "ZF BAQ"), ("ZF BAQ", "ZF BAQ"),
    ("ZF BAQ", "JUAN MINA"), ("ZF BAQ", "TRIANGULO"),
    ("PUERTO BARRANQUILLA", "JUAN MINA"), ("PUERTO BARRANQUILLA", "ALMAGRARIO"),
    ("PUERTO BARRANQUILLA", "ALPOPULAR"), ("PUERTO BARRANQUILLA", "AGOFER"),
    ("PUERTO BARRANQUILLA", "AGUACHICA"), ("PUERTO BARRANQUILLA", "IMPORTADO"),
    ("PUERTO BARRANQUILLA", "GALAPA"), ("PUERTO BARRANQUILLA", "CAYENAS"),
    ("PUERTO BARRANQUILLA", "OMEGA"), ("PUERTO BARRANQUILLA", "SANTA MARTA"),
    ("PUERTO BARRANQUILLA", "MEDELLIN"), ("PUERTO BARRANQUILLA", "MONTERIA"),
    ("PUERTO BARRANQUILLA", "PARAGUACHON"), ("PUERTO BARRANQUILLA", "SAN ROQUE"),
    ("PUERTO BARRANQUILLA", "VIA AEROPUERTO"), ("PUERTO BARRANQUILLA", "FRENTE AEROPUERTO"),
    ("PUERTO PALERMO", "CIRCUNVALAR"), ("PUERTO PALERMO", "MALAMBO"),
    ("PUERTO PALERMO", "MONTERIA"), ("CENTRO LOGISTICO CARTAGENA", "YARA"),
    ("CARTAGENA", "BARRANCABERMEJA"), ("PALMAR", "CARTAGENA"),
    ("MALAMBO", "MONTERIA"), ("PALERMO", "MALAMBO"),
]
ORIGENES_FRECUENTES = sorted(set(r[0] for r in RUTAS_FRECUENTES))
LABEL_MANUAL = "✏️ Escribir manualmente..."

CLIENTES_FRECUENTES = [
    "AGOFER", "MONOMEROS COLOMBO VENEZOLANOS S.A.", "PROCAR", "MEICO",
    "WORLD", "TRAIDING", "MAT2", "SULOGISTICS", "SUDECO", "TRIANGULO",
    "DELTA", "CARGO ANDINA", "TRANSOLICAR", "TLC", "TULUA MADERAS",
    "KBINA", "KABIBA", "PASIFIC", "MOTOTRANSPORTAMO",
]
LABEL_MANUAL_CLI = "✏️ Escribir manualmente..."

COORDENADAS = {
    "PUERTO BARRANQUILLA": (10.9831, -74.7894), "PUERTO PALERMO": (10.9125, -74.7489),
    "PALERMO": (10.9125, -74.7489), "ZF BAQ": (10.9700, -74.8100),
    "AGOFER": (10.9190, -74.8010), "MEICO": (10.9650, -74.8350),
    "MEICO CIRCUNVALAR": (10.9680, -74.8320), "PROCAR": (10.9550, -74.8200),
    "VIA40": (10.9900, -74.8000), "VIA AEROPUERTO": (10.9990, -74.7780),
    "FRENTE AEROPUERTO": (10.9990, -74.7780), "SOLEDAD": (10.9180, -74.7670),
    "MALAMBO": (10.8610, -74.7730), "GALAPA": (10.9060, -74.8880),
    "JUAN MINA": (10.9750, -74.9200), "ALMAGRARIO": (10.9620, -74.8150),
    "ALPOPULAR": (10.9600, -74.8180), "CAYENAS": (10.9580, -74.8220),
    "OMEGA": (10.9570, -74.8230), "CIRCUNVALAR": (10.9640, -74.8060),
    "TRIANGULO": (10.9660, -74.8080), "IMPORTADO": (10.9640, -74.8100),
    "CIENAGA": (11.0060, -74.2510), "SANTA MARTA": (11.2408, -74.1990),
    "SAN ROQUE": (8.5310, -73.5730), "AGUACHICA": (8.3097, -73.6197),
    "PARAGUACHON": (11.3320, -72.3820), "MONTERIA": (8.7575, -75.8812),
    "MEDELLIN": (6.2442, -75.5812), "BARRANCABERMEJA": (7.0653, -73.8547),
    "CARTAGENA": (10.3910, -75.4794), "CENTRO LOGISTICO CARTAGENA": (10.4061, -75.5100),
    "PALMAR": (10.7800, -75.1100), "YARA": (10.3850, -75.4950),
}

# ==================== CSS ====================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@400;600;700&family=Barlow:wght@300;400;500&display=swap');
    html, body, [class*="css"] { font-family: 'Barlow', sans-serif; }
    .main-header {
        background: linear-gradient(135deg, #0f2027, #203a43, #2c5364);
        padding: 1.5rem 2rem; border-radius: 12px; margin-bottom: 1.5rem;
    }
    .main-header h1 { font-family: 'Barlow Condensed', sans-serif; font-size: 2rem; font-weight: 700; color: white; margin: 0; letter-spacing: 1px; }
    .main-header p { color: #a0c4d8; margin: 0; font-size: 0.9rem; }
    .kpi-box { background: white; border-radius: 10px; padding: 1rem 1.2rem; border-left: 5px solid #2c5364; box-shadow: 0 2px 8px rgba(0,0,0,0.07); margin-bottom: 0.5rem; }
    .kpi-box .kpi-val { font-size: 2rem; font-weight: 700; color: #0f2027; }
    .kpi-box .kpi-lbl { font-size: 0.8rem; color: #666; text-transform: uppercase; letter-spacing: 1px; }
    div[data-testid="stTabs"] button { font-family: 'Barlow Condensed', sans-serif; font-weight: 600; font-size: 1rem; letter-spacing: 0.5px; }
    .conductor-auto { background: #e8f5e9; border-left: 4px solid #2ecc71; padding: 0.5rem 1rem; border-radius: 6px; margin: 0.3rem 0; font-weight: 600; color: #1a5c2a; }
    .conductor-manual { background: #fff3e0; border-left: 4px solid #f39c12; padding: 0.5rem 1rem; border-radius: 6px; margin: 0.3rem 0; font-weight: 600; color: #7d4600; }
    .correo-card { background: #f8f9fa; border-radius: 10px; border: 1px solid #dee2e6; padding: 1.2rem; margin-bottom: 1rem; }
</style>
""", unsafe_allow_html=True)


# ==================== BASE DE DATOS ====================
class DB:
    def __init__(self):
        self.url = SUPABASE_DB_URL
        self.init()

    def conn(self):
        return psycopg2.connect(self.url)

    def init(self):
        try:
            c = self.conn(); cur = c.cursor()
            cur.execute("""
                CREATE TABLE IF NOT EXISTS viajes_transporte (
                    id SERIAL PRIMARY KEY,
                    fecha_registro TIMESTAMP DEFAULT (now() AT TIME ZONE 'America/Bogota'),
                    fecha DATE NOT NULL,
                    placa TEXT NOT NULL,
                    conductor TEXT,
                    cliente TEXT,
                    origen TEXT,
                    destino TEXT,
                    hora_cita_cargue TIME,
                    hora_salida_cargue TIME,
                    hora_llegada_descargue TIME,
                    hora_salida_descargue TIME,
                    contenedor TEXT,
                    carga TEXT,
                    numero_importacion_bl TEXT,
                    manifiesto TEXT,
                    observacion TEXT,
                    estado TEXT DEFAULT 'Completado'
                )
            """)
            for col in [
                "ALTER TABLE viajes_transporte ADD COLUMN IF NOT EXISTS cliente TEXT",
                "ALTER TABLE viajes_transporte ADD COLUMN IF NOT EXISTS contenedor TEXT",
                "ALTER TABLE viajes_transporte ADD COLUMN IF NOT EXISTS numero_importacion_bl TEXT",
                "ALTER TABLE viajes_transporte ADD COLUMN IF NOT EXISTS manifiesto TEXT",
                "ALTER TABLE viajes_transporte ADD COLUMN IF NOT EXISTS estado TEXT DEFAULT 'Completado'",
            ]:
                try: cur.execute(col); c.commit()
                except Exception: 
                    try: c.rollback()
                    except: pass
            c.commit(); c.close()
        except Exception as e:
            st.error(f"Error DB init: {e}")

    def guardar_viaje(self, datos: dict) -> bool:
        try:
            c = self.conn(); cur = c.cursor()
            cur.execute("""
                INSERT INTO viajes_transporte
                (fecha, placa, conductor, cliente, origen, destino,
                 hora_cita_cargue, hora_salida_cargue, hora_llegada_descargue, hora_salida_descargue,
                 contenedor, carga, numero_importacion_bl, manifiesto, observacion, estado)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            """, (
                datos["fecha"], datos["placa"], datos["conductor"], datos["cliente"],
                datos["origen"], datos["destino"],
                datos["hora_cita_cargue"], datos["hora_salida_cargue"],
                datos["hora_llegada_descargue"], datos["hora_salida_descargue"],
                datos["contenedor"], datos["carga"],
                datos["numero_importacion_bl"], datos["manifiesto"],
                datos["observacion"], datos["estado"]
            ))
            c.commit(); c.close(); return True
        except Exception as e:
            st.error(f"Error guardando: {e}"); return False

    def actualizar_viaje(self, viaje_id: int, datos: dict) -> bool:
        try:
            c = self.conn(); cur = c.cursor()
            cur.execute("""
                UPDATE viajes_transporte SET
                fecha=%s, placa=%s, conductor=%s, cliente=%s, origen=%s, destino=%s,
                hora_cita_cargue=%s, hora_salida_cargue=%s,
                hora_llegada_descargue=%s, hora_salida_descargue=%s,
                contenedor=%s, carga=%s, numero_importacion_bl=%s,
                manifiesto=%s, observacion=%s, estado=%s
                WHERE id=%s
            """, (
                datos["fecha"], datos["placa"], datos["conductor"], datos["cliente"],
                datos["origen"], datos["destino"],
                datos["hora_cita_cargue"], datos["hora_salida_cargue"],
                datos["hora_llegada_descargue"], datos["hora_salida_descargue"],
                datos["contenedor"], datos["carga"],
                datos["numero_importacion_bl"], datos["manifiesto"],
                datos["observacion"], datos["estado"], viaje_id
            ))
            c.commit(); c.close(); return True
        except Exception as e:
            st.error(f"Error actualizando: {e}"); return False

    def eliminar_viaje(self, viaje_id: int) -> bool:
        try:
            c = self.conn(); cur = c.cursor()
            cur.execute("DELETE FROM viajes_transporte WHERE id=%s", (viaje_id,))
            c.commit(); c.close(); return True
        except Exception as e:
            st.error(f"Error eliminando: {e}"); return False

    def obtener_viajes(self, fecha_ini=None, fecha_fin=None, placa=None,
                       conductor=None, cliente=None, estado=None) -> pd.DataFrame:
        c = self.conn()
        q = """SELECT id, fecha, placa, conductor, cliente, origen, destino,
                      hora_cita_cargue, hora_salida_cargue,
                      hora_llegada_descargue, hora_salida_descargue,
                      contenedor, carga, numero_importacion_bl,
                      manifiesto, observacion, estado
               FROM viajes_transporte WHERE 1=1"""
        params = []
        if fecha_ini: q += " AND fecha >= %s"; params.append(fecha_ini)
        if fecha_fin: q += " AND fecha <= %s"; params.append(fecha_fin)
        if placa and placa != "Todas": q += " AND placa = %s"; params.append(placa)
        if conductor: q += " AND conductor ILIKE %s"; params.append(f"%{conductor}%")
        if cliente: q += " AND cliente ILIKE %s"; params.append(f"%{cliente}%")
        if estado and estado != "Todos": q += " AND estado = %s"; params.append(estado)
        q += " ORDER BY fecha DESC, id DESC"
        try:
            df = pd.read_sql(q, c, params=params); return df
        except: return pd.DataFrame()
        finally: c.close()

    def placas_unicas(self):
        c = self.conn()
        try:
            df = pd.read_sql("SELECT DISTINCT placa FROM viajes_transporte ORDER BY placa", c)
            return df["placa"].tolist()
        except: return []
        finally: c.close()

    def stats_dashboard(self, fecha_ini, fecha_fin):
        c = self.conn()
        try:
            df = pd.read_sql("""
                SELECT fecha, placa, conductor, cliente, estado,
                       hora_cita_cargue, hora_salida_cargue,
                       hora_llegada_descargue, hora_salida_descargue
                FROM viajes_transporte
                WHERE fecha >= %s AND fecha <= %s
                ORDER BY fecha
            """, c, params=[fecha_ini, fecha_fin])
            return df
        except: return pd.DataFrame()
        finally: c.close()


# ==================== HELPERS ====================
def hora_a_time(val):
    if val is None or (isinstance(val, float) and pd.isna(val)): return None
    if isinstance(val, time): return val
    try:
        s = str(val)[:5]; h, m = s.split(":"); return time(int(h), int(m))
    except: return None

def str_hora(val):
    t = hora_a_time(val)
    return t.strftime("%H:%M") if t else "—"

def calcular_duracion(h_ini, h_fin):
    t1 = hora_a_time(h_ini); t2 = hora_a_time(h_fin)
    if not t1 or not t2: return None
    d1 = timedelta(hours=t1.hour, minutes=t1.minute)
    d2 = timedelta(hours=t2.hour, minutes=t2.minute)
    diff = d2 - d1
    if diff.total_seconds() < 0: diff += timedelta(days=1)
    return int(diff.total_seconds() / 60)

def mins_a_str(mins):
    if mins is None: return "—"
    h, m = divmod(int(mins), 60)
    return f"{h}h {m:02d}m"


# ==================== EXCEL BASE ====================
def generar_excel(df: pd.DataFrame, titulo: str = "Control de Viajes") -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Viajes"

    ft_titulo  = Font(name="Calibri", bold=True, size=14, color="FFFFFF")
    ft_header  = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
    ft_normal  = Font(name="Calibri", size=9)
    ft_total   = Font(name="Calibri", bold=True, size=10)
    ft_anulado = Font(name="Calibri", size=9, color="C0392B")
    ft_incump  = Font(name="Calibri", size=9, color="D35400")

    fill_titulo  = PatternFill("solid", start_color="0F2027")
    fill_header  = PatternFill("solid", start_color="203A43")
    fill_alt     = PatternFill("solid", start_color="EBF5FB")
    fill_total   = PatternFill("solid", start_color="D5DBDB")
    fill_anulado = PatternFill("solid", start_color="FADBD8")
    fill_incump  = PatternFill("solid", start_color="FDEBD0")

    borde  = Border(left=Side(style="thin"), right=Side(style="thin"),
                    top=Side(style="thin"),  bottom=Side(style="thin"))
    centro = Alignment(horizontal="center", vertical="center", wrap_text=True)
    izq    = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    ws.merge_cells("A1:P1")
    now_col = datetime.now(pytz.timezone("America/Bogota"))
    ws["A1"] = f"🚚 {titulo}   |   Generado: {now_col.strftime('%d/%m/%Y %H:%M')} (COL)   |   Total: {len(df)} viajes"
    ws["A1"].font = ft_titulo; ws["A1"].fill = fill_titulo; ws["A1"].alignment = centro
    ws.row_dimensions[1].height = 30

    columnas = [
        ("fecha","FECHA",12), ("placa","PLACA",12), ("conductor","CONDUCTOR",26),
        ("cliente","CLIENTE",22), ("origen","ORIGEN",20), ("destino","DESTINO",20),
        ("hora_cita_cargue","H.CITA CARGUE",14), ("hora_salida_cargue","H.SALIDA CARGUE",14),
        ("hora_llegada_descargue","H.LLEGADA DESC.",14), ("hora_salida_descargue","H.SALIDA DESC.",14),
        ("contenedor","CONTENEDOR",18), ("carga","CARGA",14),
        ("numero_importacion_bl","IMP / BL",18), ("manifiesto","MANIFIESTO",12),
        ("observacion","OBSERVACIÓN",28), ("estado","ESTADO",14),
    ]
    cols_coord = ["LAT. ORIGEN","LON. ORIGEN","LAT. DESTINO","LON. DESTINO"]
    total_cols = len(columnas) + len(cols_coord)
    ws.merge_cells(f"A1:{get_column_letter(total_cols)}1")

    for idx, (key, nombre, ancho) in enumerate(columnas, start=1):
        cell = ws.cell(row=2, column=idx, value=nombre)
        cell.font = ft_header; cell.fill = fill_header; cell.alignment = centro; cell.border = borde
        ws.column_dimensions[get_column_letter(idx)].width = ancho
    for i, nombre in enumerate(cols_coord, start=len(columnas)+1):
        cell = ws.cell(row=2, column=i, value=nombre)
        cell.font = ft_header; cell.fill = PatternFill("solid", start_color="1A5276")
        cell.alignment = centro; cell.border = borde
        ws.column_dimensions[get_column_letter(i)].width = 14
    ws.row_dimensions[2].height = 28

    for row_idx, (_, fila) in enumerate(df.iterrows(), start=3):
        estado_val = str(fila.get("estado", ""))
        es_an = "Anulado" in estado_val; es_in = "Incumplido" in estado_val
        fill_f = fill_anulado if es_an else (fill_incump if es_in else (fill_alt if row_idx % 2 == 0 else None))
        for col_idx, (key, _, _) in enumerate(columnas, start=1):
            val = fila.get(key, "")
            if not isinstance(val, str) and pd.isna(val): val = ""
            if key.startswith("hora_") and val:
                try: val = str(val)[:5]
                except: val = ""
            cell = ws.cell(row=row_idx, column=col_idx, value=str(val) if val != "" else "")
            cell.border = borde
            cell.alignment = centro if key in ("fecha","placa","estado") or key.startswith("hora_") else izq
            cell.font = ft_anulado if es_an else (ft_incump if es_in else ft_normal)
            if fill_f: cell.fill = fill_f
        origen_v  = str(fila.get("origen",  "") or "").strip().upper()
        destino_v = str(fila.get("destino", "") or "").strip().upper()
        lat_o, lon_o = COORDENADAS.get(origen_v,  (None, None)) if origen_v  in COORDENADAS else (None, None)
        lat_d, lon_d = COORDENADAS.get(destino_v, (None, None)) if destino_v in COORDENADAS else (None, None)
        fill_coord = PatternFill("solid", start_color="D6EAF8") if fill_f is None and row_idx % 2 == 0 else fill_f
        for ci, val in enumerate([lat_o, lon_o, lat_d, lon_d], start=len(columnas)+1):
            cell = ws.cell(row=row_idx, column=ci, value=val if val is not None else "")
            cell.font = ft_normal; cell.border = borde; cell.alignment = centro
            if fill_coord: cell.fill = fill_coord
        ws.row_dimensions[row_idx].height = 18

    completados = len(df[df["estado"].str.contains("Completado", na=False)]) if "estado" in df.columns else 0
    anulados    = len(df[df["estado"].str.contains("Anulado",    na=False)]) if "estado" in df.columns else 0
    incumplidos = len(df[df["estado"].str.contains("Incumplido", na=False)]) if "estado" in df.columns else 0
    total_row   = len(df) + 3
    try: ws.merge_cells(f"A{total_row}:{get_column_letter(len(columnas))}{total_row}")
    except: pass
    ct = ws.cell(row=total_row, column=1, value=f"TOTAL VIAJES: {len(df)}   |   ✅ {completados}  ❌ {anulados}  ⚠️ {incumplidos}")
    ct.font = ft_total; ct.fill = fill_total; ct.alignment = centro

    # ---- Hoja Resumen ----
    ws2 = wb.create_sheet("Resumen")
    def hdr(ws, fila, col1, col2, texto):
        c = ws.cell(fila, col1, texto)
        c.font = ft_header; c.fill = PatternFill("solid", start_color="203A43")
        c.alignment = centro; c.border = borde; ws.row_dimensions[fila].height = 20
        for col in range(col1+1, col2+1):
            cx = ws.cell(fila, col, "")
            cx.fill = PatternFill("solid", start_color="203A43"); cx.border = borde

    ws2["A1"] = "Resumen General de Operaciones"
    ws2["A1"].font = Font(name="Calibri", bold=True, size=13, color="FFFFFF")
    ws2["A1"].fill = PatternFill("solid", start_color="0F2027"); ws2["A1"].alignment = centro
    ws2.row_dimensions[1].height = 26
    hdr(ws2, 2, 1, 2, "RESUMEN GENERAL")
    en_curso = len(df[df["estado"].str.contains("En Curso", na=False)]) if "estado" in df.columns else 0
    kpis = [
        ("Total Viajes", len(df)), ("Completados", completados), ("Anulados", anulados),
        ("Incumplidos", incumplidos), ("En Curso", en_curso),
        ("% Cumplimiento", f"{round(completados/len(df)*100,1)}%" if len(df) > 0 else "0%"),
    ]
    for i, (m, v) in enumerate(kpis, start=3):
        c1 = ws2.cell(i, 1, m); c2 = ws2.cell(i, 2, v)
        c1.font = ft_normal; c2.font = ft_total; c1.border = borde; c2.border = borde
        c1.alignment = izq; c2.alignment = centro
        if i % 2 == 0:
            c1.fill = PatternFill("solid", start_color="EBF5FB")
            c2.fill = PatternFill("solid", start_color="EBF5FB")
    if "cliente" in df.columns and df["cliente"].notna().any():
        hdr(ws2, 2, 4, 5, "VIAJES POR CLIENTE")
        por_cli = df.groupby("cliente").size().reset_index(name="v").sort_values("v", ascending=False)
        for i, row in enumerate(por_cli.itertuples(), start=3):
            c1 = ws2.cell(i, 4, row.cliente); c2 = ws2.cell(i, 5, int(row.v))
            c1.font = ft_normal; c2.font = ft_total; c1.border = borde; c2.border = borde
            c1.alignment = izq; c2.alignment = centro
            if i % 2 == 0:
                c1.fill = PatternFill("solid", start_color="EBF5FB")
                c2.fill = PatternFill("solid", start_color="EBF5FB")
    if "placa" in df.columns:
        hdr(ws2, 2, 7, 8, "VIAJES POR PLACA")
        por_placa = df.groupby("placa").size().reset_index(name="v").sort_values("v", ascending=False)
        for i, row in enumerate(por_placa.itertuples(), start=3):
            c1 = ws2.cell(i, 7, row.placa); c2 = ws2.cell(i, 8, int(row.v))
            c1.font = ft_normal; c2.font = ft_total; c1.border = borde; c2.border = borde
            c1.alignment = izq; c2.alignment = centro
            if i % 2 == 0:
                c1.fill = PatternFill("solid", start_color="EBF5FB")
                c2.fill = PatternFill("solid", start_color="EBF5FB")
    for col_l, w in zip(["A","B","C","D","E","F","G","H"], [22,10,3,24,8,3,12,8]):
        ws2.column_dimensions[col_l].width = w

    # ---- Hoja Conductores ----
    ws3 = wb.create_sheet("Conductores")
    ws3["A1"] = "Ranking de Conductores"
    ws3["A1"].font = Font(name="Calibri", bold=True, size=13, color="FFFFFF")
    ws3["A1"].fill = PatternFill("solid", start_color="0F2027"); ws3["A1"].alignment = centro
    ws3.row_dimensions[1].height = 26
    hdrs3_general = ["CONDUCTOR","TOTAL","COMPLET.","ANULADOS","INCUMPL.","EN CURSO","% CUMPL."]
    for ci, h in enumerate(hdrs3_general, start=1):
        c = ws3.cell(2, ci, h); c.font = ft_header
        c.fill = PatternFill("solid", start_color="203A43"); c.alignment = centro; c.border = borde
    ws3.row_dimensions[2].height = 20
    tiempos_por_conductor = {}
    if "conductor" in df.columns:
        for conductor_nombre, grupo in df.groupby("conductor"):
            if not conductor_nombre or str(conductor_nombre).strip() == "": continue
            esperas, transitos, descargues = [], [], []
            for _, r in grupo.iterrows():
                e = calcular_duracion(r.get("hora_cita_cargue"),       r.get("hora_salida_cargue"))
                t = calcular_duracion(r.get("hora_salida_cargue"),     r.get("hora_llegada_descargue"))
                d = calcular_duracion(r.get("hora_llegada_descargue"), r.get("hora_salida_descargue"))
                if e is not None: esperas.append(e)
                if t is not None: transitos.append(t)
                if d is not None: descargues.append(d)
            prom_e   = sum(esperas)   / len(esperas)   if esperas   else None
            prom_t   = sum(transitos) / len(transitos) if transitos else None
            prom_d   = sum(descargues)/ len(descargues)if descargues else None
            prom_tot = (prom_e + prom_t + prom_d) if all(x is not None for x in [prom_e, prom_t, prom_d]) else None
            tiempos_por_conductor[conductor_nombre] = {"espera": prom_e, "transito": prom_t, "descargue": prom_d, "total": prom_tot}
    if "conductor" in df.columns:
        df_cond = df.groupby("conductor").agg(
            total=("conductor","count"),
            comp=("estado", lambda x: x.str.contains("Completado", na=False).sum()),
            anul=("estado", lambda x: x.str.contains("Anulado",    na=False).sum()),
            incu=("estado", lambda x: x.str.contains("Incumplido", na=False).sum()),
            curs=("estado", lambda x: x.str.contains("En Curso",   na=False).sum()),
        ).reset_index().sort_values("total", ascending=False)
        for i, row in enumerate(df_cond.itertuples(), start=3):
            pct = f"{round(row.comp/row.total*100,1)}%" if row.total > 0 else "0%"
            vals = [row.conductor, row.total, row.comp, row.anul, row.incu, row.curs, pct]
            fill_c = PatternFill("solid", start_color="EBF5FB") if i % 2 == 0 else None
            for ci, v in enumerate(vals, start=1):
                c = ws3.cell(i, ci, v); c.font = ft_normal; c.border = borde
                c.alignment = izq if ci == 1 else centro
                if fill_c: c.fill = fill_c
    for col_l, w in zip(["A","B","C","D","E","F","G"], [32,8,10,10,10,10,10]):
        ws3.column_dimensions[col_l].width = w

    # ---- Hoja Tiempos ----
    ws4 = wb.create_sheet("Tiempos")
    ws4["A1"] = "Analisis de Tiempos por Viaje"
    ws4["A1"].font = Font(name="Calibri", bold=True, size=13, color="FFFFFF")
    ws4["A1"].fill = PatternFill("solid", start_color="0F2027"); ws4["A1"].alignment = centro
    ws4.row_dimensions[1].height = 26
    hdrs4 = ["FECHA","PLACA","CONDUCTOR","CLIENTE","ESPERA CARGUE","TRANSITO","DESCARGUE","TOTAL OPERACION"]
    for ci, h in enumerate(hdrs4, start=1):
        c = ws4.cell(2, ci, h); c.font = ft_header
        c.fill = PatternFill("solid", start_color="203A43"); c.alignment = centro; c.border = borde
    ws4.row_dimensions[2].height = 20
    tot_espera = tot_transito = tot_desc = tot_total = 0
    count_e = count_t = count_d = count_tot = 0
    for i, (_, row) in enumerate(df.iterrows(), start=3):
        t_espera  = calcular_duracion(row.get("hora_cita_cargue"),       row.get("hora_salida_cargue"))
        t_transit = calcular_duracion(row.get("hora_salida_cargue"),     row.get("hora_llegada_descargue"))
        t_desc    = calcular_duracion(row.get("hora_llegada_descargue"), row.get("hora_salida_descargue"))
        t_total   = (t_espera + t_transit + t_desc) if all(x is not None for x in [t_espera, t_transit, t_desc]) else None
        if t_espera  is not None: tot_espera  += t_espera;  count_e   += 1
        if t_transit is not None: tot_transito += t_transit; count_t   += 1
        if t_desc    is not None: tot_desc    += t_desc;    count_d   += 1
        if t_total   is not None: tot_total   += t_total;   count_tot += 1
        vals = [str(row.get("fecha","")), str(row.get("placa","")), str(row.get("conductor","")),
                str(row.get("cliente","")), mins_a_str(t_espera), mins_a_str(t_transit), mins_a_str(t_desc), mins_a_str(t_total)]
        fill_t = PatternFill("solid", start_color="EBF5FB") if i % 2 == 0 else None
        for ci, v in enumerate(vals, start=1):
            c = ws4.cell(i, ci, v); c.font = ft_normal; c.border = borde
            c.alignment = izq if ci in (1,2,3,4) else centro
            if fill_t: c.fill = fill_t
    fila_prom = len(df) + 3
    cp = ws4.cell(fila_prom, 1, "PROMEDIO")
    cp.font = ft_total; cp.fill = fill_total; cp.alignment = centro; cp.border = borde
    for ci, (tot, cnt) in enumerate([(tot_espera,count_e),(tot_transito,count_t),(tot_desc,count_d),(tot_total,count_tot)], start=5):
        c = ws4.cell(fila_prom, ci, mins_a_str(tot/cnt if cnt > 0 else None))
        c.font = ft_total; c.fill = fill_total; c.alignment = centro; c.border = borde
    for col_l, w in zip(["A","B","C","D","E","F","G","H"], [12,10,28,20,14,14,14,16]):
        ws4.column_dimensions[col_l].width = w
    ws4.freeze_panes = "A3"

    # ---- Hoja Gráfica ----
    try:
        from openpyxl.chart import PieChart, Reference
        from openpyxl.chart.series import DataPoint
        ws5 = wb.create_sheet("Grafica")
        ws5["A1"] = "Estado"; ws5["B1"] = "Cantidad"
        ws5["A1"].font = ft_header; ws5["B1"].font = ft_header
        ws5["A1"].fill = PatternFill("solid", start_color="203A43")
        ws5["B1"].fill = PatternFill("solid", start_color="203A43")
        estados_graf = ["Completado","Anulado","Incumplido","En Curso"]
        for i, est in enumerate(estados_graf, start=2):
            cnt = len(df[df["estado"].str.contains(est, na=False)]) if "estado" in df.columns else 0
            ws5.cell(i, 1, est).border = borde; ws5.cell(i, 2, cnt).border = borde
        pie = PieChart(); pie.title = "Distribucion de Viajes por Estado"; pie.style = 10
        labels = Reference(ws5, min_col=1, min_row=2, max_row=5)
        data   = Reference(ws5, min_col=2, min_row=1, max_row=5)
        pie.add_data(data, titles_from_data=True); pie.set_categories(labels)
        pie.width = 15; pie.height = 12
        colores = ["2ECC71","E74C3C","F39C12","3498DB"]
        for idx, color in enumerate(colores):
            pt = DataPoint(idx=idx); pt.graphicalProperties.solidFill = color; pie.series[0].dPt.append(pt)
        ws5.add_chart(pie, "D1")
        for col_l, w in zip(["A","B"], [16, 10]): ws5.column_dimensions[col_l].width = w
    except Exception: pass

    ws.freeze_panes = "A3"
    output = io.BytesIO(); wb.save(output); return output.getvalue()


# ==================== HOJA COMPARATIVO 3 MESES ====================
def agregar_hoja_comparativo(wb: Workbook, db, meses: int = 3):
    from dateutil.relativedelta import relativedelta

    ws = wb.create_sheet("Comparativo")

    ft_titulo = Font(name="Calibri", bold=True, size=13, color="FFFFFF")
    ft_header = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
    ft_normal = Font(name="Calibri", size=9)
    ft_total  = Font(name="Calibri", bold=True, size=10)
    ft_up     = Font(name="Calibri", bold=True, size=9, color="1E8449")
    ft_down   = Font(name="Calibri", bold=True, size=9, color="C0392B")
    ft_eq     = Font(name="Calibri", size=9, color="7F8C8D")

    fill_titulo = PatternFill("solid", start_color="0F2027")
    fill_header = PatternFill("solid", start_color="203A43")
    fill_alt    = PatternFill("solid", start_color="EBF5FB")
    fill_total  = PatternFill("solid", start_color="D5DBDB")
    fill_mejor  = PatternFill("solid", start_color="D5F5E3")
    fill_peor   = PatternFill("solid", start_color="FADBD8")

    borde  = Border(left=Side(style="thin"), right=Side(style="thin"),
                    top=Side(style="thin"),  bottom=Side(style="thin"))
    centro = Alignment(horizontal="center", vertical="center", wrap_text=True)
    izq    = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    col_bog = pytz.timezone("America/Bogota")
    hoy     = datetime.now(col_bog).date()

    periodos = []
    for i in range(meses - 1, -1, -1):
        primer_dia = hoy.replace(day=1) - relativedelta(months=i)
        ultimo_dia = hoy if i == 0 else (primer_dia + relativedelta(months=1)) - timedelta(days=1)
        periodos.append((primer_dia, ultimo_dia))

    datos_meses = []
    for ini, fin in periodos:
        df = db.obtener_viajes(fecha_ini=ini, fecha_fin=fin)
        datos_meses.append({"ini": ini, "fin": fin, "df": df})

    n_cols = meses + 3
    ws.merge_cells(f"A1:{get_column_letter(n_cols)}1")
    ws["A1"] = f"📊 COMPARATIVO ÚLTIMOS {meses} MESES  —  Generado: {datetime.now(col_bog).strftime('%d/%m/%Y %H:%M')} (COL)"
    ws["A1"].font = ft_titulo; ws["A1"].fill = fill_titulo; ws["A1"].alignment = centro
    ws.row_dimensions[1].height = 28

    fila = 3
    col_delta = meses + 2

    # Encabezados
    ws.cell(fila, 1, "INDICADOR").font = ft_header
    ws.cell(fila, 1).fill = fill_header; ws.cell(fila, 1).alignment = centro; ws.cell(fila, 1).border = borde
    ws.column_dimensions["A"].width = 30

    for col_idx, dm in enumerate(datos_meses, start=2):
        label = dm["ini"].strftime("%B %Y").upper()
        c = ws.cell(fila, col_idx, label)
        c.font = ft_header; c.fill = fill_header; c.alignment = centro; c.border = borde
        ws.column_dimensions[get_column_letter(col_idx)].width = 18

    c_d = ws.cell(fila, col_delta, "▲▼ vs MES ANT.")
    c_d.font = ft_header; c_d.fill = PatternFill("solid", start_color="1A5276")
    c_d.alignment = centro; c_d.border = borde
    ws.column_dimensions[get_column_letter(col_delta)].width = 16
    fila += 1

    def kpi_row(nombre, valores, formato=None, mayor_es_mejor=True):
        nonlocal fila
        fill_f = fill_alt if fila % 2 == 0 else None
        c = ws.cell(fila, 1, nombre); c.font = ft_normal; c.alignment = izq; c.border = borde
        if fill_f: c.fill = fill_f
        for ci, v in enumerate(valores, start=2):
            disp = f"{v:.1f}%" if formato == "pct" else (str(v) if v is not None else "—")
            cx = ws.cell(fila, ci, disp); cx.font = ft_normal; cx.alignment = centro; cx.border = borde
            if fill_f: cx.fill = fill_f
        if len(valores) >= 2 and valores[-1] is not None and valores[-2] is not None:
            diff = valores[-1] - valores[-2]
            delta_str = f"{'↑' if diff > 0 else ('↓' if diff < 0 else '=')} {abs(diff):.1f}pp" if formato == "pct" else f"{'↑' if diff > 0 else ('↓' if diff < 0 else '=')} {abs(int(diff))}"
            cd = ws.cell(fila, col_delta, delta_str)
            cd.font = (ft_up if (diff > 0) == mayor_es_mejor else ft_down) if diff != 0 else ft_eq
            cd.alignment = centro; cd.border = borde
            if fill_f: cd.fill = fill_f
        else:
            cd = ws.cell(fila, col_delta, "—"); cd.font = ft_eq; cd.alignment = centro; cd.border = borde
        nums = [v for v in valores if v is not None]
        if nums:
            mejor = max(nums) if mayor_es_mejor else min(nums)
            peor  = min(nums) if mayor_es_mejor else max(nums)
            for ci, v in enumerate(valores, start=2):
                if v == mejor and nums.count(mejor) == 1: ws.cell(fila, ci).fill = fill_mejor
                elif v == peor and nums.count(peor) == 1:  ws.cell(fila, ci).fill = fill_peor
        fila += 1

    totales  = [len(dm["df"]) for dm in datos_meses]
    complet  = [len(dm["df"][dm["df"]["estado"].str.contains("Completado", na=False)]) if not dm["df"].empty and "estado" in dm["df"].columns else 0 for dm in datos_meses]
    anulados = [len(dm["df"][dm["df"]["estado"].str.contains("Anulado",    na=False)]) if not dm["df"].empty and "estado" in dm["df"].columns else 0 for dm in datos_meses]
    incumpl  = [len(dm["df"][dm["df"]["estado"].str.contains("Incumplido", na=False)]) if not dm["df"].empty and "estado" in dm["df"].columns else 0 for dm in datos_meses]
    pct_cump = [round(c/t*100, 1) if t > 0 else 0.0 for c, t in zip(complet, totales)]

    kpi_row("Total Viajes",    totales,  mayor_es_mejor=True)
    kpi_row("✅ Completados",   complet,  mayor_es_mejor=True)
    kpi_row("❌ Anulados",      anulados, mayor_es_mejor=False)
    kpi_row("⚠️ Incumplidos",   incumpl,  mayor_es_mejor=False)
    kpi_row("% Cumplimiento",  pct_cump, formato="pct", mayor_es_mejor=True)

    def prom_tiempo_mes(df_m, col_ini, col_fin):
        if df_m.empty: return None
        vals = [calcular_duracion(r.get(col_ini), r.get(col_fin)) for _, r in df_m.iterrows()]
        vals = [v for v in vals if v is not None]
        return round(sum(vals)/len(vals)) if vals else None

    esperas    = [prom_tiempo_mes(dm["df"], "hora_cita_cargue",       "hora_salida_cargue")     for dm in datos_meses]
    transitos  = [prom_tiempo_mes(dm["df"], "hora_salida_cargue",     "hora_llegada_descargue") for dm in datos_meses]
    descargues = [prom_tiempo_mes(dm["df"], "hora_llegada_descargue", "hora_salida_descargue")  for dm in datos_meses]

    def kpi_tiempo_row(nombre, valores_min):
        nonlocal fila
        fill_f = fill_alt if fila % 2 == 0 else None
        c = ws.cell(fila, 1, nombre); c.font = ft_normal; c.alignment = izq; c.border = borde
        if fill_f: c.fill = fill_f
        for ci, v in enumerate(valores_min, start=2):
            cx = ws.cell(fila, ci, mins_a_str(v)); cx.font = ft_normal; cx.alignment = centro; cx.border = borde
            if fill_f: cx.fill = fill_f
        nums = [v for v in valores_min if v is not None]
        if nums:
            for ci, v in enumerate(valores_min, start=2):
                if v == min(nums) and nums.count(min(nums)) == 1: ws.cell(fila, ci).fill = fill_mejor
                elif v == max(nums) and nums.count(max(nums)) == 1: ws.cell(fila, ci).fill = fill_peor
        if len(valores_min) >= 2 and valores_min[-1] is not None and valores_min[-2] is not None:
            diff = valores_min[-1] - valores_min[-2]
            cd = ws.cell(fila, col_delta, f"{'↑' if diff > 0 else ('↓' if diff < 0 else '=')} {mins_a_str(abs(diff))}")
            cd.font = ft_down if diff > 0 else (ft_up if diff < 0 else ft_eq)
            cd.alignment = centro; cd.border = borde
        else:
            cd = ws.cell(fila, col_delta, "—"); cd.font = ft_eq; cd.alignment = centro; cd.border = borde
        fila += 1

    kpi_tiempo_row("⏱️ Espera Cargue (prom.)",  esperas)
    kpi_tiempo_row("🚛 Tránsito (prom.)",        transitos)
    kpi_tiempo_row("📦 Descargue (prom.)",        descargues)

    # ---- Ranking conductores ----
    fila += 1
    ws.merge_cells(f"A{fila}:{get_column_letter(n_cols)}{fila}")
    ws.cell(fila, 1, "🏆 RANKING DE CONDUCTORES POR MES").font = ft_titulo
    ws.cell(fila, 1).fill = PatternFill("solid", start_color="1A3A4A"); ws.cell(fila, 1).alignment = centro
    ws.row_dimensions[fila].height = 22; fila += 1

    col_start = 1
    for dm in datos_meses:
        label = dm["ini"].strftime("%B %Y").upper(); df_m = dm["df"]
        try:
            ws.merge_cells(f"{get_column_letter(col_start)}{fila}:{get_column_letter(col_start+3)}{fila}")
        except: pass
        ws.cell(fila, col_start, label).font = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
        ws.cell(fila, col_start).fill = PatternFill("solid", start_color="2C5364"); ws.cell(fila, col_start).alignment = centro
        for ci, h in enumerate(["CONDUCTOR","VIAJES","COMP.","% CUMPL."], start=col_start):
            c = ws.cell(fila+1, ci, h); c.font = ft_header; c.fill = fill_header; c.alignment = centro; c.border = borde
        if not df_m.empty and "conductor" in df_m.columns:
            cond_g = df_m.groupby("conductor").agg(
                total=("conductor","count"),
                comp=("estado", lambda x: x.str.contains("Completado", na=False).sum())
            ).reset_index().sort_values("total", ascending=False).head(8)
            for ri, row in enumerate(cond_g.itertuples(), start=fila+2):
                pct = f"{round(row.comp/row.total*100,1)}%" if row.total > 0 else "0%"
                fill_r = fill_alt if ri % 2 == 0 else None
                for ci, v in enumerate([row.conductor, row.total, row.comp, pct], start=col_start):
                    cx = ws.cell(ri, ci, v); cx.font = ft_normal; cx.border = borde
                    cx.alignment = izq if ci == col_start else centro
                    if fill_r: cx.fill = fill_r
        col_start += 5
        ws.column_dimensions[get_column_letter(col_start-1)].width = 2  # separador

    # ---- Top clientes ----
    fila += 12
    ws.merge_cells(f"A{fila}:{get_column_letter(n_cols)}{fila}")
    ws.cell(fila, 1, "🏢 TOP CLIENTES POR MES").font = ft_titulo
    ws.cell(fila, 1).fill = PatternFill("solid", start_color="1A3A4A"); ws.cell(fila, 1).alignment = centro
    ws.row_dimensions[fila].height = 22; fila += 1
    col_start = 1
    for dm in datos_meses:
        label = dm["ini"].strftime("%B %Y").upper(); df_m = dm["df"]
        try:
            ws.merge_cells(f"{get_column_letter(col_start)}{fila}:{get_column_letter(col_start+2)}{fila}")
        except: pass
        ws.cell(fila, col_start, label).font = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
        ws.cell(fila, col_start).fill = PatternFill("solid", start_color="2C5364"); ws.cell(fila, col_start).alignment = centro
        for ci, h in enumerate(["CLIENTE","VIAJES","%"], start=col_start):
            c = ws.cell(fila+1, ci, h); c.font = ft_header; c.fill = fill_header; c.alignment = centro; c.border = borde
        if not df_m.empty and "cliente" in df_m.columns and df_m["cliente"].notna().any():
            cli_g = df_m.groupby("cliente").size().reset_index(name="v").sort_values("v", ascending=False).head(6)
            total_m = len(df_m)
            for ri, row in enumerate(cli_g.itertuples(), start=fila+2):
                pct_cli = f"{round(row.v/total_m*100,1)}%" if total_m > 0 else "—"
                fill_r = fill_alt if ri % 2 == 0 else None
                for ci, v in enumerate([row.cliente, row.v, pct_cli], start=col_start):
                    cx = ws.cell(ri, ci, v); cx.font = ft_normal; cx.border = borde
                    cx.alignment = izq if ci == col_start else centro
                    if fill_r: cx.fill = fill_r
        col_start += 4

    ws.freeze_panes = "A3"
    return wb


def generar_excel_con_comparativo(db, fecha_ini, fecha_fin, titulo="Control de Viajes"):
    from openpyxl import load_workbook
    df = db.obtener_viajes(fecha_ini=fecha_ini, fecha_fin=fecha_fin)
    excel_bytes = generar_excel(df, titulo=titulo)
    wb = load_workbook(io.BytesIO(excel_bytes))
    wb = agregar_hoja_comparativo(wb, db, meses=3)
    output = io.BytesIO(); wb.save(output)
    return output.getvalue(), df


# ==================== ENVÍO CORREO ====================
def enviar_reporte_gmail(gmail_usuario, gmail_app_password, destinatarios,
                         excel_bytes, df, periodo_label="", asunto_extra=""):
    import smtplib, ssl
    from email.mime.multipart import MIMEMultipart
    from email.mime.base import MIMEBase
    from email.mime.text import MIMEText
    from email import encoders

    col_bog = pytz.timezone("America/Bogota"); ahora = datetime.now(col_bog)
    total = len(df)
    comp  = len(df[df["estado"].str.contains("Completado", na=False)]) if not df.empty and "estado" in df.columns else 0
    anul  = len(df[df["estado"].str.contains("Anulado",    na=False)]) if not df.empty and "estado" in df.columns else 0
    incum = len(df[df["estado"].str.contains("Incumplido", na=False)]) if not df.empty and "estado" in df.columns else 0
    pct   = round(comp/total*100, 1) if total > 0 else 0

    asunto = f"🚚 Reporte Diario Transporte — {ahora.strftime('%d/%m/%Y')}"
    if asunto_extra: asunto += f" | {asunto_extra}"

    cuerpo_html = f"""
    <html><body style="font-family:Calibri,Arial,sans-serif;color:#2c3e50;background:#f4f6f7;padding:20px;">
      <div style="max-width:620px;margin:auto;background:white;border-radius:10px;overflow:hidden;box-shadow:0 4px 15px rgba(0,0,0,0.1);">
        <div style="background:linear-gradient(135deg,#0f2027,#203a43,#2c5364);padding:24px 28px;">
          <h1 style="color:white;margin:0;font-size:22px;letter-spacing:1px;">🚚 REPORTE DIARIO DE TRANSPORTE</h1>
          <p style="color:#a0c4d8;margin:6px 0 0;">{ahora.strftime('%A, %d de %B de %Y').upper()} &nbsp;|&nbsp; {periodo_label}</p>
        </div>
        <div style="padding:24px 28px;">
          <table style="width:100%;border-collapse:collapse;margin-bottom:20px;">
            <tr>
              <td style="padding:12px;background:#ecf0f1;border-radius:8px;text-align:center;width:24%;">
                <div style="font-size:26px;font-weight:700;color:#2c3e50;">{total}</div>
                <div style="font-size:11px;color:#7f8c8d;text-transform:uppercase;">Total Viajes</div>
              </td>
              <td style="width:2%;"></td>
              <td style="padding:12px;background:#eafaf1;border-radius:8px;text-align:center;width:24%;">
                <div style="font-size:26px;font-weight:700;color:#27ae60;">{comp}</div>
                <div style="font-size:11px;color:#7f8c8d;text-transform:uppercase;">✅ Completados</div>
              </td>
              <td style="width:2%;"></td>
              <td style="padding:12px;background:#fdf2f2;border-radius:8px;text-align:center;width:24%;">
                <div style="font-size:26px;font-weight:700;color:#e74c3c;">{anul}</div>
                <div style="font-size:11px;color:#7f8c8d;text-transform:uppercase;">❌ Anulados</div>
              </td>
              <td style="width:2%;"></td>
              <td style="padding:12px;background:#fef9e7;border-radius:8px;text-align:center;width:24%;">
                <div style="font-size:26px;font-weight:700;color:#f39c12;">{incum}</div>
                <div style="font-size:11px;color:#7f8c8d;text-transform:uppercase;">⚠️ Incumplidos</div>
              </td>
            </tr>
          </table>
          <div style="background:linear-gradient(90deg,#203a43,#2c5364);border-radius:8px;padding:14px 20px;margin-bottom:20px;text-align:center;">
            <span style="color:white;font-size:15px;font-weight:600;">% Cumplimiento: </span>
            <span style="color:#2ecc71;font-size:24px;font-weight:700;">{pct}%</span>
          </div>
          <p style="color:#7f8c8d;font-size:12px;">
            📎 Se adjunta el reporte completo en Excel con detalle de viajes, tiempos, ranking de conductores
            y <strong>comparativo de los últimos 3 meses</strong>.
          </p>
        </div>
        <div style="background:#ecf0f1;padding:12px 28px;text-align:center;">
          <p style="color:#95a5a6;font-size:11px;margin:0;">Generado automáticamente · {ahora.strftime('%H:%M')} COL</p>
        </div>
      </div>
    </body></html>"""

    msg = MIMEMultipart("mixed")
    msg["Subject"] = asunto; msg["From"] = gmail_usuario; msg["To"] = ", ".join(destinatarios)
    msg.attach(MIMEText(cuerpo_html, "html"))

    nombre_archivo = f"Reporte_Transporte_{ahora.strftime('%Y%m%d')}.xlsx"
    part = MIMEBase("application", "octet-stream"); part.set_payload(excel_bytes)
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{nombre_archivo}"')
    msg.attach(part)

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(gmail_usuario, gmail_app_password)
        server.sendmail(gmail_usuario, destinatarios, msg.as_string())
    return True


# ==================== SCHEDULER ====================
def iniciar_scheduler(hora_envio: str, callback):
    import schedule, threading, time as time_mod
    schedule.clear()
    schedule.every().day.at(hora_envio).do(callback)
    def run():
        while True:
            schedule.run_pending(); time_mod.sleep(30)
    t = threading.Thread(target=run, daemon=True); t.start()
    return t

def detener_scheduler():
    try:
        import schedule; schedule.clear()
    except: pass


# ==================== MAIN ====================
def main():
    st.markdown("""
    <div class="main-header">
        <h1>🚚 CONTROL DE VIAJES</h1>
        <p>Registro y seguimiento de operaciones de transporte</p>
    </div>
    """, unsafe_allow_html=True)

    if "db" not in st.session_state: st.session_state.db = DB()
    if "editando_id" not in st.session_state: st.session_state.editando_id = None
    if "correo_config" not in st.session_state:
        st.session_state.correo_config = {
            "gmail_usuario": "", "gmail_password": "", "destinatarios": "",
            "hora_envio": "07:00", "scheduler_on": False, "ultimo_envio": None,
        }

    db = st.session_state.db
    cfg = st.session_state.correo_config

    tab1, tab2, tab3, tab4 = st.tabs([
        "📝 Nuevo Viaje", "🔍 Historial y Reportes", "📊 Dashboard", "✉️ Reportes por Correo"
    ])

    # ===================== TAB 1: NUEVO VIAJE =====================
    with tab1:
        st.markdown("### Registrar Nuevo Viaje")
        f1, f2, f3, f4 = st.columns(4)
        with f1: fecha_pre = st.date_input("📅 Fecha", datetime.now(), key="pre_fecha")
        with f2:
            placas_lista = list(PLACA_CONDUCTOR.keys())
            placa_pre = st.selectbox("🚛 Placa", placas_lista, key="pre_placa")
        with f3:
            conductor_fijo = PLACA_CONDUCTOR.get(placa_pre)
            cond_opts = ["— Seleccionar —"] + TODOS_CONDUCTORES
            cond_default = cond_opts.index(conductor_fijo) if conductor_fijo in cond_opts else 0
            conductor_sel = st.selectbox("👤 Conductor", cond_opts, index=cond_default, key="pre_conductor")
        with f4:
            cli_sel = st.selectbox("🏢 Cliente", CLIENTES_FRECUENTES + [LABEL_MANUAL_CLI], key="pre_cliente")
        if cli_sel == LABEL_MANUAL_CLI:
            cliente_pre = st.text_input("✏️ Escribir cliente manualmente", placeholder="Nombre del cliente...", key="pre_cli_manual")
        else:
            cliente_pre = cli_sel

        st.markdown("#### 🗺️ Ruta")
        ruta_opts = [f"{o}  →  {d}" for o, d in RUTAS_FRECUENTES] + [LABEL_MANUAL]
        ruta_sel = st.selectbox("🗺️ Ruta frecuente", ruta_opts, index=len(ruta_opts)-1, key="pre_ruta")
        c5, c6 = st.columns(2)
        if ruta_sel == LABEL_MANUAL:
            with c5: origen_pre  = st.text_input("📍 Origen",  placeholder="Escribe el origen...",  key="pre_origen")
            with c6: destino_pre = st.text_input("🏁 Destino", placeholder="Escribe el destino...", key="pre_destino")
        else:
            _o, _d = ruta_sel.split("  →  ")
            with c5: st.info(f"📍 **Origen:** {_o}")
            with c6: st.info(f"🏁 **Destino:** {_d}")
            origen_pre, destino_pre = _o, _d

        with st.form("form_viaje", clear_on_submit=True):
            fecha = fecha_pre; placa = placa_pre
            conductor = "" if conductor_sel == "— Seleccionar —" else conductor_sel
            cliente = cliente_pre; origen = origen_pre; destino = destino_pre
            st.markdown("#### ⏱️ Tiempos de Operación")
            h1, h2, h3, h4 = st.columns(4)
            with h1: hora_cita_cargue       = st.time_input("Cita Cargue",       value=None, step=300)
            with h2: hora_salida_cargue     = st.time_input("Salida Cargue",     value=None, step=300)
            with h3: hora_llegada_descargue = st.time_input("Llegada Descargue", value=None, step=300)
            with h4: hora_salida_descargue  = st.time_input("Salida Descargue",  value=None, step=300)
            st.markdown("#### 📦 Información de Carga")
            d1, d2, d3, d4 = st.columns(4)
            with d1: contenedor         = st.text_input("Contenedor")
            with d2: carga              = st.text_input("Carga")
            with d3: numero_importacion = st.text_input("Nº Importación / BL")
            with d4: manifiesto         = st.text_input("Manifiesto")
            e1, e2 = st.columns([1, 3])
            with e1: estado      = st.selectbox("🚦 Estado", ESTADOS_VIAJE)
            with e2: observacion = st.text_area("📝 Observaciones", height=80)
            submitted = st.form_submit_button("💾 Guardar Viaje", type="primary", use_container_width=True)

        if submitted:
            if not placa: st.error("⚠️ La placa es obligatoria.")
            else:
                datos = {
                    "fecha": fecha, "placa": placa, "conductor": conductor, "cliente": cliente,
                    "origen": origen, "destino": destino,
                    "hora_cita_cargue": hora_cita_cargue, "hora_salida_cargue": hora_salida_cargue,
                    "hora_llegada_descargue": hora_llegada_descargue, "hora_salida_descargue": hora_salida_descargue,
                    "contenedor": contenedor, "carga": carga,
                    "numero_importacion_bl": numero_importacion, "manifiesto": manifiesto,
                    "observacion": observacion,
                    "estado": estado.split(" ", 1)[1] if " " in estado else estado
                }
                if db.guardar_viaje(datos):
                    st.success(f"✅ Viaje guardado — {placa} | {conductor} | {origen} → {destino}")
                    st.balloons()

    # ===================== TAB 2: HISTORIAL =====================
    with tab2:
        st.markdown("### 🔍 Historial de Viajes")
        with st.expander("🛠️ Filtros", expanded=True):
            f1, f2, f3, f4, f5, f6 = st.columns(6)
            with f1: fi   = st.date_input("Desde", datetime.now() - timedelta(days=30), key="h_fi")
            with f2: ff   = st.date_input("Hasta", datetime.now(), key="h_ff")
            with f3:
                placas_h = ["Todas"] + list(PLACA_CONDUCTOR.keys())
                fp = st.selectbox("Placa", placas_h, key="h_fp")
            with f4: fc   = st.text_input("Conductor", key="h_fc")
            with f5: fcli = st.text_input("Cliente",   key="h_fcli")
            with f6:
                estados_f = ["Todos"] + [e.split(" ", 1)[1] for e in ESTADOS_VIAJE]
                fe = st.selectbox("Estado", estados_f, key="h_fe")

        df = db.obtener_viajes(fi, ff, fp, fc, fcli, fe if fe != "Todos" else None)

        if not df.empty:
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Total Viajes", len(df))
            k2.metric("✅ Completados", len(df[df["estado"].str.contains("Completado", na=False)]))
            k3.metric("❌ Anulados",    len(df[df["estado"].str.contains("Anulado",    na=False)]))
            k4.metric("⚠️ Incumplidos", len(df[df["estado"].str.contains("Incumplido", na=False)]))
            st.divider()

            col_exp1, col_exp2, col_exp3 = st.columns([2, 2, 3])
            with col_exp1:
                nombre_rep = st.text_input("Nombre del reporte", value="Control_Viajes", key="rep_nombre")
            with col_exp2:
                st.markdown("<br>", unsafe_allow_html=True)
                excel_data = generar_excel(df, titulo=nombre_rep)
                st.download_button("⬇️ Excel Simple", data=excel_data,
                    file_name=f"{nombre_rep}_{datetime.now(pytz.timezone('America/Bogota')).strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with col_exp3:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("📊 Descargar Excel + Comparativo 3 Meses", type="primary", use_container_width=True):
                    with st.spinner("Generando comparativo..."):
                        try:
                            excel_comp, _ = generar_excel_con_comparativo(db, fi, ff, titulo=nombre_rep)
                            nombre_comp = f"{nombre_rep}_comparativo_{datetime.now(pytz.timezone('America/Bogota')).strftime('%Y%m%d_%H%M')}.xlsx"
                            st.download_button("⬇️ Descargar Ahora", data=excel_comp,
                                file_name=nombre_comp,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_comp")
                        except Exception as e:
                            st.error(f"Error: {e}")

            st.divider()
            cols_tabla = ["id","fecha","placa","conductor","cliente","origen","destino",
                          "contenedor","carga","numero_importacion_bl","manifiesto","estado"]
            cols_ex = [c for c in cols_tabla if c in df.columns]
            st.dataframe(df[cols_ex], use_container_width=True, hide_index=True)
            st.divider()
            st.subheader("✏️ Ver Detalle / Editar")
            df["_label"] = df.apply(
                lambda r: f"ID {r['id']} | {r['fecha']} | {r['placa']} | {r.get('cliente','')} | {r.get('origen','')} → {r.get('destino','')} | {r.get('estado','')}",
                axis=1)
            sel = st.selectbox("Seleccionar viaje:", df["_label"].tolist(), key="h_sel")
            if sel:
                vid = int(sel.split(" | ")[0].replace("ID ", ""))
                row = df[df["id"] == vid].iloc[0]
                editando = st.session_state.editando_id == vid
                if not editando:
                    c1, c2, c3 = st.columns(3)
                    with c1:
                        st.info(f"**Placa:** {row['placa']}")
                        st.write(f"**Conductor:** {row.get('conductor','')}"); st.write(f"**Cliente:** {row.get('cliente','')}"); st.write(f"**Fecha:** {row['fecha']}")
                    with c2:
                        st.write(f"**Origen:** {row.get('origen','')}"); st.write(f"**Destino:** {row.get('destino','')}"); st.write(f"**Contenedor:** {row.get('contenedor','')}"); st.write(f"**Carga:** {row.get('carga','')}")
                    with c3:
                        st.write(f"**Importación/BL:** {row.get('numero_importacion_bl','')}"); st.write(f"**Manifiesto:** {row.get('manifiesto','')}")
                        estado_raw = str(row.get('estado',''))
                        color = "🟢" if "Completado" in estado_raw else ("🔴" if "Anulado" in estado_raw else "🟡")
                        st.write(f"**Estado:** {color} {estado_raw}"); st.write(f"**Observación:** {row.get('observacion','')}")
                    st.write(f"**Horas:** Cita: `{str_hora(row['hora_cita_cargue'])}` | Salida Cargue: `{str_hora(row['hora_salida_cargue'])}` | Llegada: `{str_hora(row['hora_llegada_descargue'])}` | Salida Desc: `{str_hora(row['hora_salida_descargue'])}`")
                    bc1, bc2 = st.columns(2)
                    with bc1:
                        if st.button("✏️ Editar", key=f"eb_{vid}"): st.session_state.editando_id = vid; st.rerun()
                    with bc2:
                        if st.button("🗑️ Eliminar", key=f"del_{vid}"): db.eliminar_viaje(vid); st.success("Eliminado."); st.rerun()
                else:
                    st.markdown("#### ✏️ Editando viaje")
                    with st.form(f"edit_{vid}"):
                        ec1, ec2, ec3, ec4 = st.columns(4)
                        with ec1: e_fecha = st.date_input("Fecha", value=row["fecha"], key=f"ef_{vid}")
                        with ec2:
                            placas_e = list(PLACA_CONDUCTOR.keys())
                            placa_idx = placas_e.index(row["placa"]) if row["placa"] in placas_e else 0
                            e_placa = st.selectbox("Placa", placas_e, index=placa_idx, key=f"ep_{vid}")
                        with ec3:
                            cond_fijo_e = PLACA_CONDUCTOR.get(e_placa)
                            cond_actual = str(row.get("conductor") or "")
                            cond_opts_e = ["— Seleccionar —"] + TODOS_CONDUCTORES
                            default_e = cond_opts_e.index(cond_fijo_e) if cond_fijo_e in cond_opts_e else (cond_opts_e.index(cond_actual) if cond_actual in cond_opts_e else 0)
                            e_cond_sel = st.selectbox("👤 Conductor", cond_opts_e, index=default_e, key=f"ec_{vid}")
                            e_conductor = "" if e_cond_sel == "— Seleccionar —" else e_cond_sel
                        with ec4:
                            cli_actual = str(row.get("cliente") or "")
                            cli_opts = CLIENTES_FRECUENTES + [LABEL_MANUAL_CLI]
                            cli_idx = cli_opts.index(cli_actual) if cli_actual in cli_opts else len(cli_opts)-1
                            e_cli_sel = st.selectbox("Cliente", cli_opts, index=cli_idx, key=f"ecl_{vid}")
                            e_cliente = st.text_input("Cliente (manual)", value=cli_actual if cli_actual not in CLIENTES_FRECUENTES else "", key=f"ecl_m_{vid}") if e_cli_sel == LABEL_MANUAL_CLI else e_cli_sel
                        er1, er2 = st.columns(2)
                        with er1: e_origen  = st.text_input("Origen",  value=str(row.get("origen") or ""),  key=f"eo_{vid}")
                        with er2: e_destino = st.text_input("Destino", value=str(row.get("destino") or ""), key=f"ed_{vid}")
                        st.markdown("#### ⏱️ Horas")
                        eh1, eh2, eh3, eh4 = st.columns(4)
                        with eh1: e_hcc = st.time_input("Cita Cargue",       value=hora_a_time(row["hora_cita_cargue"]),       step=300, key=f"ehcc_{vid}")
                        with eh2: e_hsc = st.time_input("Salida Cargue",     value=hora_a_time(row["hora_salida_cargue"]),     step=300, key=f"ehsc_{vid}")
                        with eh3: e_hld = st.time_input("Llegada Descargue", value=hora_a_time(row["hora_llegada_descargue"]), step=300, key=f"ehld_{vid}")
                        with eh4: e_hsd = st.time_input("Salida Descargue",  value=hora_a_time(row["hora_salida_descargue"]),  step=300, key=f"ehsd_{vid}")
                        ed1, ed2, ed3, ed4 = st.columns(4)
                        with ed1: e_cont  = st.text_input("Contenedor",  value=str(row.get("contenedor") or ""),            key=f"eco_{vid}")
                        with ed2: e_carga = st.text_input("Carga",       value=str(row.get("carga") or ""),                 key=f"eca_{vid}")
                        with ed3: e_bl    = st.text_input("Imp / BL",    value=str(row.get("numero_importacion_bl") or ""), key=f"ebl_{vid}")
                        with ed4: e_man   = st.text_input("Manifiesto",  value=str(row.get("manifiesto") or ""),            key=f"ema_{vid}")
                        estados_l = [e.split(" ", 1)[1] for e in ESTADOS_VIAJE]
                        est_idx = estados_l.index(str(row.get("estado") or "Completado")) if str(row.get("estado") or "Completado") in estados_l else 0
                        ee1, ee2 = st.columns([1, 3])
                        with ee1: e_estado = st.selectbox("Estado", ESTADOS_VIAJE, index=est_idx, key=f"est_{vid}")
                        with ee2: e_obs    = st.text_area("Observaciones", value=str(row.get("observacion") or ""), key=f"eob_{vid}", height=80)
                        sg1, sg2 = st.columns(2)
                        with sg1: guardar  = st.form_submit_button("💾 Guardar Cambios", type="primary")
                        with sg2: cancelar = st.form_submit_button("❌ Cancelar")
                    if guardar:
                        datos_edit = {
                            "fecha": e_fecha, "placa": e_placa, "conductor": e_conductor, "cliente": e_cliente,
                            "origen": e_origen, "destino": e_destino,
                            "hora_cita_cargue": e_hcc, "hora_salida_cargue": e_hsc,
                            "hora_llegada_descargue": e_hld, "hora_salida_descargue": e_hsd,
                            "contenedor": e_cont, "carga": e_carga,
                            "numero_importacion_bl": e_bl, "manifiesto": e_man, "observacion": e_obs,
                            "estado": e_estado.split(" ", 1)[1] if " " in e_estado else e_estado
                        }
                        if db.actualizar_viaje(vid, datos_edit):
                            st.success("✅ Viaje actualizado."); st.session_state.editando_id = None; st.rerun()
                    if cancelar:
                        st.session_state.editando_id = None; st.rerun()
        else:
            st.warning("No hay viajes con los filtros seleccionados.")

    # ===================== TAB 3: DASHBOARD =====================
    with tab3:
        st.markdown("### 📊 Dashboard de Operaciones")
        try:
            import plotly.express as px
            col_r1, col_r2 = st.columns([2, 4])
            with col_r1:
                rango = st.date_input("Período", value=(datetime.now().replace(day=1), datetime.now()), key="dash_rango")
            if not (isinstance(rango, (list, tuple)) and len(rango) == 2):
                st.info("Selecciona un rango de fechas completo."); return
            df_s = db.stats_dashboard(rango[0], rango[1])
            if df_s.empty: st.info("No hay datos en este período."); return
            total = len(df_s)
            comp  = len(df_s[df_s["estado"].str.contains("Completado", na=False)])
            anul  = len(df_s[df_s["estado"].str.contains("Anulado",    na=False)])
            incum = len(df_s[df_s["estado"].str.contains("Incumplido", na=False)])
            curso = len(df_s[df_s["estado"].str.contains("En Curso",   na=False)])
            pct   = round(comp / total * 100) if total > 0 else 0
            k1, k2, k3, k4, k5 = st.columns(5)
            k1.metric("🚚 Total", total); k2.metric("✅ Completados", comp, f"{pct}%")
            k3.metric("❌ Anulados", anul); k4.metric("⚠️ Incumplidos", incum); k5.metric("🔄 En Curso", curso)
            st.divider()
            g1, g2 = st.columns(2)
            with g1:
                st.markdown("#### Distribución por Estado")
                est_c = df_s["estado"].value_counts().reset_index(); est_c.columns = ["estado","cantidad"]
                fig1 = px.pie(est_c, values="cantidad", names="estado", hole=0.45,
                              color="estado", color_discrete_map={"Completado":"#2ecc71","Anulado":"#e74c3c","Incumplido":"#f39c12","En Curso":"#3498db"})
                fig1.update_layout(margin=dict(t=10,b=10), height=300); st.plotly_chart(fig1, use_container_width=True)
            with g2:
                st.markdown("#### Viajes por Día")
                df_dia = df_s.groupby("fecha").size().reset_index(name="viajes")
                fig2 = px.bar(df_dia, x="fecha", y="viajes", color_discrete_sequence=["#2c5364"], text="viajes")
                fig2.update_traces(textposition="outside"); fig2.update_layout(margin=dict(t=10,b=10), height=300)
                st.plotly_chart(fig2, use_container_width=True)
            st.divider()
            g3, g4 = st.columns(2)
            with g3:
                st.markdown("#### Viajes por Cliente")
                if "cliente" in df_s.columns and df_s["cliente"].notna().any():
                    df_cli = df_s.groupby("cliente").size().reset_index(name="viajes").sort_values("viajes")
                    fig3 = px.bar(df_cli, x="viajes", y="cliente", orientation="h", color="viajes", color_continuous_scale="Blues", text="viajes")
                    fig3.update_traces(textposition="outside"); fig3.update_layout(margin=dict(t=10,b=10), height=max(250,len(df_cli)*40), coloraxis_showscale=False)
                    st.plotly_chart(fig3, use_container_width=True)
            with g4:
                st.markdown("#### Viajes por Placa")
                df_placa = df_s.groupby("placa").size().reset_index(name="viajes").sort_values("viajes")
                fig4 = px.bar(df_placa, x="viajes", y="placa", orientation="h", color="viajes", color_continuous_scale="Teal", text="viajes")
                fig4.update_traces(textposition="outside"); fig4.update_layout(margin=dict(t=10,b=10), height=max(250,len(df_placa)*40), coloraxis_showscale=False)
                st.plotly_chart(fig4, use_container_width=True)
            st.divider()
            st.markdown("#### 🏆 Ranking de Conductores")
            df_cond = df_s[df_s["conductor"].notna() & (df_s["conductor"].str.strip() != "")].groupby("conductor").agg(
                viajes=("conductor","count"),
                completados=("estado", lambda x: x.str.contains("Completado", na=False).sum()),
                anulados=("estado", lambda x: x.str.contains("Anulado", na=False).sum()),
                incumplidos=("estado", lambda x: x.str.contains("Incumplido", na=False).sum()),
            ).reset_index().sort_values("viajes", ascending=False)
            df_cond["% Cumplimiento"] = (df_cond["completados"]/df_cond["viajes"]*100).round(1).astype(str) + "%"
            df_cond.columns = ["Conductor","Total","✅ Comp.","❌ Anul.","⚠️ Incump.","% Cumplimiento"]
            st.dataframe(df_cond, use_container_width=True, hide_index=True)
        except ImportError: st.warning("Instala plotly: `pip install plotly`")
        except Exception as e: st.error(f"Error en dashboard: {e}")

    # ===================== TAB 4: CORREO =====================
    with tab4:
        st.markdown("### ✉️ Reportes por Correo — Configuración y Envío")

        with st.expander("⚙️ Configuración de Cuenta Gmail", expanded=not cfg["gmail_usuario"]):
            st.info(
                "💡 **¿Cómo obtener una App Password de Gmail?**\n\n"
                "1. Ve a **myaccount.google.com** → Seguridad\n"
                "2. Activa **Verificación en 2 pasos** (si no la tienes)\n"
                "3. Busca **Contraseñas de aplicaciones** → Genera una para 'Correo'\n"
                "4. Copia las 16 letras y pégalas abajo"
            )
            c1, c2 = st.columns(2)
            with c1:
                nuevo_usuario = st.text_input("📧 Correo Gmail remitente", value=cfg["gmail_usuario"],
                                              placeholder="tuempresa@gmail.com", key="cfg_gmail")
            with c2:
                nuevo_password = st.text_input("🔑 App Password (16 caracteres)", value=cfg["gmail_password"],
                                               type="password", placeholder="xxxx xxxx xxxx xxxx", key="cfg_pwd")
            nuevos_dest = st.text_area("📬 Destinatarios (uno por línea o coma)", value=cfg["destinatarios"],
                                       placeholder="gerencia@empresa.com\noperaciones@empresa.com", height=100, key="cfg_dest")
            nueva_hora = st.time_input("⏰ Hora de envío diario (hora Colombia)",
                                       value=datetime.strptime(cfg["hora_envio"], "%H:%M").time(), step=300, key="cfg_hora")
            if st.button("💾 Guardar Configuración", type="primary"):
                cfg["gmail_usuario"]  = nuevo_usuario.strip()
                cfg["gmail_password"] = nuevo_password.strip()
                cfg["destinatarios"]  = nuevos_dest.strip()
                cfg["hora_envio"]     = nueva_hora.strftime("%H:%M")
                st.success("✅ Configuración guardada."); st.rerun()

        st.divider()

        # Panel scheduler
        col_sch1, col_sch2 = st.columns([3, 2])
        with col_sch1:
            st.markdown("#### 🤖 Envío Automático Diario")
            st.markdown(f"Estado: {'🟢 **ACTIVO**' if cfg['scheduler_on'] else '🔴 **INACTIVO**'}")
            if cfg["scheduler_on"]: st.caption(f"Envía todos los días a las **{cfg['hora_envio']}** (hora Colombia)")
            if cfg["ultimo_envio"]: st.caption(f"Último envío: **{cfg['ultimo_envio']}**")
        with col_sch2:
            st.markdown("<br>", unsafe_allow_html=True)
            if not cfg["scheduler_on"]:
                if st.button("▶️ Activar Envío Automático", type="primary", use_container_width=True):
                    if not cfg["gmail_usuario"] or not cfg["gmail_password"] or not cfg["destinatarios"]:
                        st.error("⚠️ Completa la configuración de correo primero.")
                    else:
                        def _callback():
                            try:
                                col_tz = pytz.timezone("America/Bogota")
                                hoy_cb = datetime.now(col_tz).date()
                                ini_cb = hoy_cb.replace(day=1)
                                excel_b, df_cb = generar_excel_con_comparativo(
                                    db, fecha_ini=ini_cb, fecha_fin=hoy_cb,
                                    titulo=f"Reporte {hoy_cb.strftime('%B %Y').upper()}"
                                )
                                dests_cb = [d.strip() for d in cfg["destinatarios"].replace(",","\n").split("\n") if d.strip()]
                                enviar_reporte_gmail(
                                    gmail_usuario=cfg["gmail_usuario"], gmail_app_password=cfg["gmail_password"],
                                    destinatarios=dests_cb, excel_bytes=excel_b, df=df_cb,
                                    periodo_label=f"Mes: {ini_cb.strftime('%B %Y').upper()}"
                                )
                                cfg["ultimo_envio"] = datetime.now(col_tz).strftime("%d/%m/%Y %H:%M")
                            except Exception: pass
                        iniciar_scheduler(cfg["hora_envio"], _callback)
                        cfg["scheduler_on"] = True
                        st.success(f"✅ Activado. Enviará a las **{cfg['hora_envio']}** cada día.")
                        st.rerun()
            else:
                if st.button("⏹️ Desactivar", use_container_width=True):
                    detener_scheduler(); cfg["scheduler_on"] = False
                    st.info("Envío automático desactivado."); st.rerun()

        st.divider()

        # Envío manual
        st.markdown("#### 📤 Envío Manual")
        _hoy = datetime.now(pytz.timezone("America/Bogota")).date()
        m1, m2, m3 = st.columns(3)
        with m1: man_ini = st.date_input("Desde", value=_hoy.replace(day=1), key="man_ini")
        with m2: man_fin = st.date_input("Hasta", value=_hoy, key="man_fin")
        with m3: man_titulo = st.text_input("Título", value="Control_Viajes", key="man_titulo")

        col_dl, col_send = st.columns(2)
        with col_dl:
            if st.button("📥 Descargar Excel + Comparativo", use_container_width=True):
                with st.spinner("Generando..."):
                    try:
                        excel_b, _ = generar_excel_con_comparativo(db, man_ini, man_fin, titulo=man_titulo)
                        st.download_button("⬇️ Descargar", data=excel_b,
                            file_name=f"{man_titulo}_comp_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_manual")
                    except Exception as e: st.error(f"Error: {e}")

        with col_send:
            if st.button("✉️ Generar y Enviar Ahora", type="primary", use_container_width=True):
                if not cfg["gmail_usuario"] or not cfg["gmail_password"] or not cfg["destinatarios"]:
                    st.error("⚠️ Configura el correo antes de enviar.")
                else:
                    with st.spinner("Enviando..."):
                        try:
                            excel_b, df_env = generar_excel_con_comparativo(db, man_ini, man_fin, titulo=man_titulo)
                            dests = [d.strip() for d in cfg["destinatarios"].replace(",","\n").split("\n") if d.strip()]
                            enviar_reporte_gmail(
                                gmail_usuario=cfg["gmail_usuario"], gmail_app_password=cfg["gmail_password"],
                                destinatarios=dests, excel_bytes=excel_b, df=df_env,
                                periodo_label=f"{man_ini.strftime('%d/%m/%Y')} — {man_fin.strftime('%d/%m/%Y')}",
                                asunto_extra=man_titulo
                            )
                            cfg["ultimo_envio"] = datetime.now(pytz.timezone("America/Bogota")).strftime("%d/%m/%Y %H:%M")
                            st.success(f"✅ Enviado a: {', '.join(dests)}")
                            st.balloons()
                        except Exception as e:
                            if "Authentication" in str(e) or "auth" in str(e).lower():
                                st.error("❌ Error de autenticación. Verifica correo y App Password.")
                            else:
                                st.error(f"❌ Error: {e}")

        # Vista previa destinatarios
        if cfg["destinatarios"]:
            st.divider()
            dests_p = [d.strip() for d in cfg["destinatarios"].replace(",","\n").split("\n") if d.strip()]
            st.markdown(f"**📬 Destinatarios ({len(dests_p)}):**")
            for d in dests_p: st.markdown(f"- `{d}`")


if __name__ == "__main__":
    main()
