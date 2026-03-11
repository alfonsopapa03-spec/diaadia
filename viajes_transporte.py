import streamlit as st
import psycopg2
import pandas as pd
from datetime import datetime, timedelta, time
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
# Placa -> conductor fijo (None = no tiene fijo, se elige manualmente)
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
    "REIMUR MANUEL",
    "HABID CAMACHO",
    "JOSE ORTEGA PEREZ",
    "CARLOS TAFUR",
    "ISAIAS VESGA",
    "FLAVIO ROSENDO MALTE TUTALCHA",
    "SLITH JOSE ORTEGA PACHECO",
    "ABRAHAM SEGUNDO ALVAREZ VALLE",
    "RAMON TAFUR HERNANDEZ",
    "JULIAN CALETH CORONADO",
    "PEDRO VILLAMIL",
    "JESUS DAVID MONTES MOSQUERA",
    "CHRISTIAN MARTINEZ NAVARRO",
    "YEIMI DUQUE ZULUAGA",
    "EDGAR DE JESUS RAMIREZ",
    "EDUARDO RAFAEL OLIVARES ALCAZAR",
])

ESTADOS_VIAJE = ["✅ Completado", "❌ Anulado", "⚠️ Incumplido", "🔄 En Curso"]

# ==================== RUTAS FRECUENTES ====================
RUTAS_FRECUENTES = [
    ("PUERTO PALERMO", "AGOFER"),
    ("PUERTO BARRANQUILLA", "VIA40"),
    ("PUERTO BARRANQUILLA", "PROCAR"),
    ("PUERTO BARRANQUILLA", "CIENAGA"),
    ("PUERTO BARRANQUILLA", "MEICO"),
    ("PUERTO BARRANQUILLA", "MEICO CIRCUNVALAR"),
    ("PUERTO BARRANQUILLA", "SOLEDAD"),
    ("PUERTO PALERMO", "ZF BAQ"),
    ("PUERTO BARRANQUILLA", "ZF BAQ"),
    ("ZF BAQ", "ZF BAQ"),
    ("ZF BAQ", "JUAN MINA"),
    ("ZF BAQ", "TRIANGULO"),
    ("PUERTO BARRANQUILLA", "JUAN MINA"),
    ("PUERTO BARRANQUILLA", "ALMAGRARIO"),
    ("PUERTO BARRANQUILLA", "ALPOPULAR"),
    ("PUERTO BARRANQUILLA", "AGOFER"),
    ("PUERTO BARRANQUILLA", "AGUACHICA"),
    ("PUERTO BARRANQUILLA", "IMPORTADO"),
    ("PUERTO BARRANQUILLA", "GALAPA"),
    ("PUERTO BARRANQUILLA", "CAYENAS"),
    ("PUERTO BARRANQUILLA", "OMEGA"),
    ("PUERTO BARRANQUILLA", "SANTA MARTA"),
    ("PUERTO BARRANQUILLA", "MEDELLIN"),
    ("PUERTO BARRANQUILLA", "MONTERIA"),
    ("PUERTO BARRANQUILLA", "PARAGUACHON"),
    ("PUERTO BARRANQUILLA", "SAN ROQUE"),
    ("PUERTO BARRANQUILLA", "VIA AEROPUERTO"),
    ("PUERTO BARRANQUILLA", "FRENTE AEROPUERTO"),
    ("PUERTO PALERMO", "CIRCUNVALAR"),
    ("PUERTO PALERMO", "MALAMBO"),
    ("PUERTO PALERMO", "MONTERIA"),
    ("CENTRO LOGISTICO CARTAGENA", "YARA"),
    ("CARTAGENA", "BARRANCABERMEJA"),
    ("PALMAR", "CARTAGENA"),
    ("MALAMBO", "MONTERIA"),
    ("PALERMO", "MALAMBO"),
]

ORIGENES_FRECUENTES = sorted(set(r[0] for r in RUTAS_FRECUENTES))
LABEL_MANUAL = "✏️ Escribir manualmente..."

# ==================== CLIENTES FRECUENTES ====================
CLIENTES_FRECUENTES = [
    "AGOFER",
    "MONOMEROS COLOMBO VENEZOLANOS S.A.",
    "PROCAR",
    "MEICO",
    "WORLD",
    "TRAIDING",
    "MAT2",
    "SULOGISTICS",
    "SUDECO",
    "TRIANGULO",
    "DELTA",
    "CARGO ANDINA",
    "TRANSOLICAR",
    "TLC",
    "TULUA MADERAS",
    "KBINA",
    "KABIBA",
    "PASIFIC",
    "MOTOTRANSPORTAMO",
]
LABEL_MANUAL_CLI = "✏️ Escribir manualmente..."

# ==================== COORDENADAS POR LUGAR ====================
COORDENADAS = {
    "PUERTO BARRANQUILLA":      (10.9831, -74.7894),
    "PUERTO PALERMO":           (10.9125, -74.7489),
    "PALERMO":                  (10.9125, -74.7489),
    "ZF BAQ":                   (10.9700, -74.8100),
    "AGOFER":                   (10.9190, -74.8010),
    "MEICO":                    (10.9650, -74.8350),
    "MEICO CIRCUNVALAR":        (10.9680, -74.8320),
    "PROCAR":                   (10.9550, -74.8200),
    "VIA40":                    (10.9900, -74.8000),
    "VIA AEROPUERTO":           (10.9990, -74.7780),
    "FRENTE AEROPUERTO":        (10.9990, -74.7780),
    "SOLEDAD":                  (10.9180, -74.7670),
    "MALAMBO":                  (10.8610, -74.7730),
    "GALAPA":                   (10.9060, -74.8880),
    "JUAN MINA":                (10.9750, -74.9200),
    "ALMAGRARIO":               (10.9620, -74.8150),
    "ALPOPULAR":                (10.9600, -74.8180),
    "CAYENAS":                  (10.9580, -74.8220),
    "OMEGA":                    (10.9570, -74.8230),
    "CIRCUNVALAR":              (10.9640, -74.8060),
    "TRIANGULO":                (10.9660, -74.8080),
    "IMPORTADO":                (10.9640, -74.8100),
    "CIENAGA":                  (11.0060, -74.2510),
    "SANTA MARTA":              (11.2408, -74.1990),
    "SAN ROQUE":                (8.5310,  -73.5730),
    "AGUACHICA":                (8.3097,  -73.6197),
    "PARAGUACHON":              (11.3320, -72.3820),
    "MONTERIA":                 (8.7575,  -75.8812),
    "MEDELLIN":                 (6.2442,  -75.5812),
    "BARRANCABERMEJA":          (7.0653,  -73.8547),
    "CARTAGENA":                (10.3910, -75.4794),
    "CENTRO LOGISTICO CARTAGENA": (10.4061, -75.5100),
    "PALMAR":                   (10.7800, -75.1100),
    "YARA":                     (10.3850, -75.4950),
    "PROCAR":                   (10.9550, -74.8200),
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
    .main-header h1 {
        font-family: 'Barlow Condensed', sans-serif;
        font-size: 2rem; font-weight: 700; color: white; margin: 0; letter-spacing: 1px;
    }
    .main-header p { color: #a0c4d8; margin: 0; font-size: 0.9rem; }
    .kpi-box {
        background: white; border-radius: 10px; padding: 1rem 1.2rem;
        border-left: 5px solid #2c5364; box-shadow: 0 2px 8px rgba(0,0,0,0.07);
        margin-bottom: 0.5rem;
    }
    .kpi-box .kpi-val { font-size: 2rem; font-weight: 700; color: #0f2027; }
    .kpi-box .kpi-lbl { font-size: 0.8rem; color: #666; text-transform: uppercase; letter-spacing: 1px; }
    div[data-testid="stTabs"] button {
        font-family: 'Barlow Condensed', sans-serif;
        font-weight: 600; font-size: 1rem; letter-spacing: 0.5px;
    }
    .conductor-auto {
        background: #e8f5e9; border-left: 4px solid #2ecc71;
        padding: 0.5rem 1rem; border-radius: 6px; margin: 0.3rem 0;
        font-weight: 600; color: #1a5c2a;
    }
    .conductor-manual {
        background: #fff3e0; border-left: 4px solid #f39c12;
        padding: 0.5rem 1rem; border-radius: 6px; margin: 0.3rem 0;
        font-weight: 600; color: #7d4600;
    }
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
            c = self.conn()
            cur = c.cursor()
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
                try:
                    cur.execute(col); c.commit()
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
            c.commit(); c.close()
            return True
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
        """Retorna datos agregados para el dashboard"""
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


# ==================== EXCEL ====================
def generar_excel(df: pd.DataFrame, titulo: str = "Control de Viajes") -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Viajes"

    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    # Estilos
    lado   = Side(style="thin", color="CCCCCC")
    borde  = Border(left=lado, right=lado, top=lado, bottom=lado)
    centro = Alignment(horizontal="center", vertical="center", wrap_text=True)
    izq    = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    COLOR_HEADER = "1A3A4A"
    COLOR_COMP   = "D5F5E3"
    COLOR_ANUL   = "FADBD8"
    COLOR_INCU   = "FDEBD0"
    COLOR_CURS   = "D6EAF8"

    ft_titulo  = Font(name="Calibri", bold=True,  size=12, color="FFFFFF")
    ft_header  = Font(name="Calibri", bold=True,  size=10, color="FFFFFF")
    ft_normal  = Font(name="Calibri", bold=False, size=10)
    ft_total   = Font(name="Calibri", bold=True,  size=10)
    fill_hdr   = PatternFill("solid", start_color=COLOR_HEADER)
    fill_total = PatternFill("solid", start_color="203A43")

    columnas = ["ID","FECHA","PLACA","CONDUCTOR","CLIENTE","ORIGEN","DESTINO",
                "CITA CARGUE","SALIDA CARGUE","LLEGADA DESC.","SALIDA DESC.",
                "CONTENEDOR","CARGA","IMPORTACION/BL","MANIFIESTO","ESTADO","OBSERVACION"]

    col_map = {
        "id":"ID","fecha":"FECHA","placa":"PLACA","conductor":"CONDUCTOR",
        "cliente":"CLIENTE","origen":"ORIGEN","destino":"DESTINO",
        "hora_cita_cargue":"CITA CARGUE","hora_salida_cargue":"SALIDA CARGUE",
        "hora_llegada_descargue":"LLEGADA DESC.","hora_salida_descargue":"SALIDA DESC.",
        "contenedor":"CONTENEDOR","carga":"CARGA","numero_importacion":"IMPORTACION/BL",
        "manifiesto":"MANIFIESTO","estado":"ESTADO","observacion":"OBSERVACION"
    }

    now_col = datetime.now(pytz.timezone("America/Bogota"))

    # Fila 1: título (sin merge)
    ws.cell(1, 1, f"Control de Viajes  |  Generado: {now_col.strftime('%d/%m/%Y %H:%M')} (COL)  |  Total: {len(df)} viajes")
    ws.cell(1, 1).font = ft_titulo
    ws.cell(1, 1).fill = PatternFill("solid", start_color="0F2027")
    ws.cell(1, 1).alignment = izq
    ws.row_dimensions[1].height = 24
    # Rellenar resto de fila 1 con mismo color
    for ci in range(2, len(columnas)+1):
        ws.cell(1, ci).fill = PatternFill("solid", start_color="0F2027")

    # Fila 2: headers
    for ci, col in enumerate(columnas, start=1):
        c = ws.cell(2, ci, col)
        c.font = ft_header; c.fill = fill_hdr
        c.alignment = centro; c.border = borde
    ws.row_dimensions[2].height = 22

    completados = anulados = incumplidos = 0

    # Filas de datos
    for ri, (_, row) in enumerate(df.iterrows(), start=3):
        estado_val = str(row.get("estado",""))
        if   "Completado"  in estado_val: fill_row = PatternFill("solid", start_color=COLOR_COMP); completados  += 1
        elif "Anulado"     in estado_val: fill_row = PatternFill("solid", start_color=COLOR_ANUL); anulados     += 1
        elif "Incumplido"  in estado_val: fill_row = PatternFill("solid", start_color=COLOR_INCU); incumplidos  += 1
        elif "En Curso"    in estado_val: fill_row = PatternFill("solid", start_color=COLOR_CURS)
        else:                              fill_row = None

        for ci, col_name in enumerate(col_map.keys(), start=1):
            val = row.get(col_name, "")
            val = "" if val is None or (isinstance(val, float) and pd.isna(val)) else val
            c = ws.cell(ri, ci, str(val) if not isinstance(val, (int, float)) else val)
            c.font = ft_normal; c.border = borde
            c.alignment = izq if ci in (4,5,6,7,17) else centro
            if fill_row: c.fill = fill_row
        ws.row_dimensions[ri].height = 18

    # Fila total (sin merge)
    total_row = len(df) + 3
    ws.cell(total_row, 1, f"TOTAL: {len(df)} viajes  |  Completados: {completados}  |  Anulados: {anulados}  |  Incumplidos: {incumplidos}")
    ws.cell(total_row, 1).font = ft_total
    ws.cell(total_row, 1).fill = fill_total
    ws.cell(total_row, 1).alignment = izq
    ws.cell(total_row, 1).border = borde
    for ci in range(2, len(columnas)+1):
        ws.cell(total_row, ci).fill = fill_total
        ws.cell(total_row, ci).border = borde

    # Anchos de columna
    anchos = [6,11,9,28,22,22,18,12,12,12,12,16,14,16,12,14,28]
    for ci, w in enumerate(anchos, start=1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = "A3"

    # ==================== HOJA RESUMEN ====================
    ws2 = wb.create_sheet("Resumen")

    def hdr_simple(wsh, fila, col, texto):
        c = wsh.cell(fila, col, texto)
        c.font = ft_header
        c.fill = PatternFill("solid", start_color=COLOR_HEADER)
        c.alignment = centro; c.border = borde
        wsh.row_dimensions[fila].height = 20

    ws2.cell(1, 1, "RESUMEN GENERAL DE OPERACIONES")
    ws2.cell(1, 1).font = ft_titulo
    ws2.cell(1, 1).fill = PatternFill("solid", start_color="0F2027")
    ws2.cell(1, 1).alignment = izq
    for ci in range(2, 9):
        ws2.cell(1, ci).fill = PatternFill("solid", start_color="0F2027")
    ws2.row_dimensions[1].height = 26

    # KPIs
    en_curso = len(df[df["estado"].str.contains("En Curso", na=False)]) if "estado" in df.columns else 0
    hdr_simple(ws2, 2, 1, "METRICA"); hdr_simple(ws2, 2, 2, "VALOR")
    kpis = [("Total Viajes", len(df)),("Completados", completados),("Anulados", anulados),
            ("Incumplidos", incumplidos),("En Curso", en_curso),
            ("% Cumplimiento", f"{round(completados/len(df)*100,1)}%" if len(df)>0 else "0%")]
    for i,(m,v) in enumerate(kpis, start=3):
        c1=ws2.cell(i,1,m); c2=ws2.cell(i,2,v)
        c1.font=ft_normal; c2.font=ft_total
        c1.border=borde; c2.border=borde
        c1.alignment=izq; c2.alignment=centro
        if i%2==0:
            c1.fill=PatternFill("solid",start_color="EBF5FB")
            c2.fill=PatternFill("solid",start_color="EBF5FB")

    # Por cliente
    if "cliente" in df.columns and df["cliente"].notna().any():
        hdr_simple(ws2, 2, 4, "CLIENTE"); hdr_simple(ws2, 2, 5, "VIAJES")
        por_cli = df.groupby("cliente").size().reset_index(name="v").sort_values("v", ascending=False)
        for i, row in enumerate(por_cli.itertuples(), start=3):
            c1=ws2.cell(i,4,row.cliente); c2=ws2.cell(i,5,int(row.v))
            c1.font=ft_normal; c2.font=ft_total
            c1.border=borde; c2.border=borde
            c1.alignment=izq; c2.alignment=centro
            if i%2==0:
                c1.fill=PatternFill("solid",start_color="EBF5FB")
                c2.fill=PatternFill("solid",start_color="EBF5FB")

    # Por placa
    if "placa" in df.columns:
        hdr_simple(ws2, 2, 7, "PLACA"); hdr_simple(ws2, 2, 8, "VIAJES")
        por_placa = df.groupby("placa").size().reset_index(name="v").sort_values("v", ascending=False)
        for i, row in enumerate(por_placa.itertuples(), start=3):
            c1=ws2.cell(i,7,row.placa); c2=ws2.cell(i,8,int(row.v))
            c1.font=ft_normal; c2.font=ft_total
            c1.border=borde; c2.border=borde
            c1.alignment=izq; c2.alignment=centro
            if i%2==0:
                c1.fill=PatternFill("solid",start_color="EBF5FB")
                c2.fill=PatternFill("solid",start_color="EBF5FB")

    for col_l,w in zip(["A","B","C","D","E","F","G","H"],[22,10,3,26,8,3,12,8]):
        ws2.column_dimensions[col_l].width = w

    # ==================== HOJA CONDUCTORES ====================
    ws3 = wb.create_sheet("Conductores")
    ws3.cell(1,1,"RANKING DE CONDUCTORES")
    ws3.cell(1,1).font = ft_titulo
    ws3.cell(1,1).fill = PatternFill("solid", start_color="0F2027")
    ws3.cell(1,1).alignment = izq
    for ci in range(2,8): ws3.cell(1,ci).fill = PatternFill("solid", start_color="0F2027")
    ws3.row_dimensions[1].height = 26

    hdrs3 = ["CONDUCTOR","TOTAL","COMPLET.","ANULADOS","INCUMPL.","EN CURSO","% CUMPL."]
    for ci,h in enumerate(hdrs3, start=1):
        c=ws3.cell(2,ci,h); c.font=ft_header
        c.fill=fill_hdr; c.alignment=centro; c.border=borde
    ws3.row_dimensions[2].height = 20

    if "conductor" in df.columns:
        df_cond = df[df["conductor"].notna() & (df["conductor"].str.strip()!="")].groupby("conductor").agg(
            total=("conductor","count"),
            comp=("estado", lambda x: x.str.contains("Completado",na=False).sum()),
            anul=("estado", lambda x: x.str.contains("Anulado",na=False).sum()),
            incu=("estado", lambda x: x.str.contains("Incumplido",na=False).sum()),
            curs=("estado", lambda x: x.str.contains("En Curso",na=False).sum()),
        ).reset_index().sort_values("total", ascending=False)
        for i,row in enumerate(df_cond.itertuples(), start=3):
            pct = f"{round(row.comp/row.total*100,1)}%" if row.total>0 else "0%"
            vals=[row.conductor,row.total,row.comp,row.anul,row.incu,row.curs,pct]
            fill_c=PatternFill("solid",start_color="EBF5FB") if i%2==0 else None
            for ci,v in enumerate(vals, start=1):
                c=ws3.cell(i,ci,v); c.font=ft_normal; c.border=borde
                c.alignment=izq if ci==1 else centro
                if fill_c: c.fill=fill_c

    for col_l,w in zip(["A","B","C","D","E","F","G"],[32,8,10,10,10,10,10]):
        ws3.column_dimensions[col_l].width = w
    ws3.freeze_panes = "A3"

    # ==================== HOJA TIEMPOS ====================
    ws4 = wb.create_sheet("Tiempos")
    ws4.cell(1,1,"ANALISIS DE TIEMPOS POR VIAJE")
    ws4.cell(1,1).font = ft_titulo
    ws4.cell(1,1).fill = PatternFill("solid", start_color="0F2027")
    ws4.cell(1,1).alignment = izq
    for ci in range(2,9): ws4.cell(1,ci).fill = PatternFill("solid", start_color="0F2027")
    ws4.row_dimensions[1].height = 26

    hdrs4=["FECHA","PLACA","CONDUCTOR","CLIENTE","ESPERA CARGUE","TRANSITO","DESCARGUE","TOTAL"]
    for ci,h in enumerate(hdrs4, start=1):
        c=ws4.cell(2,ci,h); c.font=ft_header
        c.fill=fill_hdr; c.alignment=centro; c.border=borde
    ws4.row_dimensions[2].height = 20

    tot_e=tot_t=tot_d=tot_tot=count_e=count_t=count_d=count_tot=0
    for i,(_, row) in enumerate(df.iterrows(), start=3):
        t_e = calcular_duracion(row.get("hora_cita_cargue"),       row.get("hora_salida_cargue"))
        t_t = calcular_duracion(row.get("hora_salida_cargue"),     row.get("hora_llegada_descargue"))
        t_d = calcular_duracion(row.get("hora_llegada_descargue"), row.get("hora_salida_descargue"))
        t_tot = (t_e+t_t+t_d) if all(x is not None for x in [t_e,t_t,t_d]) else None
        if t_e   is not None: tot_e   += t_e;   count_e   += 1
        if t_t   is not None: tot_t   += t_t;   count_t   += 1
        if t_d   is not None: tot_d   += t_d;   count_d   += 1
        if t_tot is not None: tot_tot += t_tot; count_tot += 1
        vals=[str(row.get("fecha","")),str(row.get("placa","")),str(row.get("conductor","")),
              str(row.get("cliente","")),mins_a_str(t_e),mins_a_str(t_t),mins_a_str(t_d),mins_a_str(t_tot)]
        fill_t=PatternFill("solid",start_color="EBF5FB") if i%2==0 else None
        for ci,v in enumerate(vals, start=1):
            c=ws4.cell(i,ci,v); c.font=ft_normal; c.border=borde
            c.alignment=izq if ci<=4 else centro
            if fill_t: c.fill=fill_t

    fp=len(df)+3
    promedios=["PROMEDIO","","","",
               mins_a_str(tot_e/count_e if count_e>0 else None),
               mins_a_str(tot_t/count_t if count_t>0 else None),
               mins_a_str(tot_d/count_d if count_d>0 else None),
               mins_a_str(tot_tot/count_tot if count_tot>0 else None)]
    for ci,v in enumerate(promedios, start=1):
        c=ws4.cell(fp,ci,v); c.font=ft_total
        c.fill=fill_total; c.alignment=centro; c.border=borde

    for col_l,w in zip(["A","B","C","D","E","F","G","H"],[12,10,28,22,14,14,14,14]):
        ws4.column_dimensions[col_l].width = w
    ws4.freeze_panes = "A3"

    # ==================== HOJA GRAFICA ====================
    try:
        from openpyxl.chart import PieChart, Reference
        from openpyxl.chart.series import DataPoint
        ws5 = wb.create_sheet("Grafica")
        ws5.cell(1,1,"Estado"); ws5.cell(1,2,"Cantidad")
        ws5.cell(1,1).font=ft_header; ws5.cell(1,2).font=ft_header
        ws5.cell(1,1).fill=fill_hdr; ws5.cell(1,2).fill=fill_hdr
        estados_g=["Completado","Anulado","Incumplido","En Curso"]
        for i,est in enumerate(estados_g, start=2):
            cnt=len(df[df["estado"].str.contains(est,na=False)]) if "estado" in df.columns else 0
            ws5.cell(i,1,est).border=borde; ws5.cell(i,2,cnt).border=borde
        pie=PieChart(); pie.title="Distribucion por Estado"; pie.style=10
        labels=Reference(ws5,min_col=1,min_row=2,max_row=5)
        data=Reference(ws5,min_col=2,min_row=1,max_row=5)
        pie.add_data(data,titles_from_data=True); pie.set_categories(labels)
        pie.width=15; pie.height=12
        colores=["2ECC71","E74C3C","F39C12","3498DB"]
        for idx,color in enumerate(colores):
            pt=DataPoint(idx=idx); pt.graphicalProperties.solidFill=color
            pie.series[0].dPt.append(pt)
        ws5.add_chart(pie,"D1")
        ws5.column_dimensions["A"].width=16; ws5.column_dimensions["B"].width=10
    except Exception:
        pass

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

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
    """Calcula minutos entre dos horas (puede cruzar medianoche)"""
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


# ==================== MAIN ====================
def main():
    st.markdown("""
    <div class="main-header">
        <h1>🚚 CONTROL DE VIAJES</h1>
        <p>Registro y seguimiento de operaciones de transporte</p>
    </div>
    """, unsafe_allow_html=True)

    if "db" not in st.session_state:
        st.session_state.db = DB()
    if "editando_id" not in st.session_state:
        st.session_state.editando_id = None
    if "placa_sel" not in st.session_state:
        st.session_state.placa_sel = list(PLACA_CONDUCTOR.keys())[0]

    db = st.session_state.db

    tab1, tab2, tab3 = st.tabs(["📝 Nuevo Viaje", "🔍 Historial y Reportes", "📊 Dashboard"])

    # ===================== TAB 1: NUEVO VIAJE =====================
    with tab1:
        st.markdown("### Registrar Nuevo Viaje")

        # Selectores fuera del form para reactividad
        f1, f2, f3, f4 = st.columns(4)
        with f1:
            fecha_pre = st.date_input("📅 Fecha", datetime.now(), key="pre_fecha")
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
            with c5: origen_pre  = st.text_input("📍 Origen", placeholder="Escribe el origen...", key="pre_origen")
            with c6: destino_pre = st.text_input("🏁 Destino", placeholder="Escribe el destino...", key="pre_destino")
        else:
            _o, _d = ruta_sel.split("  →  ")
            with c5: st.info(f"📍 **Origen:** {_o}")
            with c6: st.info(f"🏁 **Destino:** {_d}")
            origen_pre, destino_pre = _o, _d

        with st.form("form_viaje", clear_on_submit=True):
            fecha = fecha_pre
            placa = placa_pre
            conductor = "" if conductor_sel == "— Seleccionar —" else conductor_sel
            cliente = cliente_pre
            origen = origen_pre
            destino = destino_pre


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
            with e1: estado     = st.selectbox("🚦 Estado", ESTADOS_VIAJE)
            with e2: observacion = st.text_area("📝 Observaciones", height=80)

            submitted = st.form_submit_button("💾 Guardar Viaje", type="primary", use_container_width=True)

        if submitted:
            if not placa:
                st.error("⚠️ La placa es obligatoria.")
            else:
                datos = {
                    "fecha": fecha, "placa": placa, "conductor": conductor,
                    "cliente": cliente, "origen": origen, "destino": destino,
                    "hora_cita_cargue": hora_cita_cargue,
                    "hora_salida_cargue": hora_salida_cargue,
                    "hora_llegada_descargue": hora_llegada_descargue,
                    "hora_salida_descargue": hora_salida_descargue,
                    "contenedor": contenedor, "carga": carga,
                    "numero_importacion_bl": numero_importacion,
                    "manifiesto": manifiesto, "observacion": observacion,
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
            with f5: fcli = st.text_input("Cliente", key="h_fcli")
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

            col_exp1, col_exp2 = st.columns([2, 5])
            with col_exp1:
                nombre_rep = st.text_input("Nombre del reporte", value="Control_Viajes", key="rep_nombre")
            with col_exp2:
                st.markdown("<br>", unsafe_allow_html=True)
                excel_data = generar_excel(df, titulo=nombre_rep)
                st.download_button(
                    "⬇️ Descargar Excel",
                    data=excel_data,
                    file_name=f"{nombre_rep}_{datetime.now(pytz.timezone('America/Bogota')).strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

            st.divider()

            cols_tabla = ["id","fecha","placa","conductor","cliente","origen","destino",
                          "contenedor","carga","numero_importacion_bl","manifiesto","estado"]
            cols_ex = [c for c in cols_tabla if c in df.columns]
            st.dataframe(df[cols_ex], use_container_width=True, hide_index=True)

            st.divider()
            st.subheader("✏️ Ver Detalle / Editar")

            df["_label"] = df.apply(
                lambda r: f"ID {r['id']} | {r['fecha']} | {r['placa']} | {r.get('cliente','')} | {r.get('origen','')} → {r.get('destino','')} | {r.get('estado','')}",
                axis=1
            )
            sel = st.selectbox("Seleccionar viaje:", df["_label"].tolist(), key="h_sel")

            if sel:
                vid = int(sel.split(" | ")[0].replace("ID ", ""))
                row = df[df["id"] == vid].iloc[0]
                editando = st.session_state.editando_id == vid

                if not editando:
                    c1, c2, c3 = st.columns(3)
                    with c1:
                        st.info(f"**Placa:** {row['placa']}")
                        st.write(f"**Conductor:** {row.get('conductor','')}")
                        st.write(f"**Cliente:** {row.get('cliente','')}")
                        st.write(f"**Fecha:** {row['fecha']}")
                    with c2:
                        st.write(f"**Origen:** {row.get('origen','')}")
                        st.write(f"**Destino:** {row.get('destino','')}")
                        st.write(f"**Contenedor:** {row.get('contenedor','')}")
                        st.write(f"**Carga:** {row.get('carga','')}")
                    with c3:
                        st.write(f"**Importación/BL:** {row.get('numero_importacion_bl','')}")
                        st.write(f"**Manifiesto:** {row.get('manifiesto','')}")
                        estado_raw = str(row.get('estado',''))
                        color = "🟢" if "Completado" in estado_raw else ("🔴" if "Anulado" in estado_raw else "🟡")
                        st.write(f"**Estado:** {color} {estado_raw}")
                        st.write(f"**Observación:** {row.get('observacion','')}")
                    st.write(f"**Horas:** Cita: `{str_hora(row['hora_cita_cargue'])}` | Salida Cargue: `{str_hora(row['hora_salida_cargue'])}` | Llegada: `{str_hora(row['hora_llegada_descargue'])}` | Salida Desc: `{str_hora(row['hora_salida_descargue'])}`")

                    bc1, bc2 = st.columns(2)
                    with bc1:
                        if st.button("✏️ Editar", key=f"eb_{vid}"):
                            st.session_state.editando_id = vid; st.rerun()
                    with bc2:
                        if st.button("🗑️ Eliminar", key=f"del_{vid}"):
                            db.eliminar_viaje(vid); st.success("Eliminado."); st.rerun()
                else:
                    st.markdown("#### ✏️ Editando viaje")
                    with st.form(f"edit_{vid}"):
                        ec1, ec2, ec3, ec4 = st.columns(4)
                        with ec1: e_fecha     = st.date_input("Fecha", value=row["fecha"], key=f"ef_{vid}")
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
                            if e_cli_sel == LABEL_MANUAL_CLI:
                                e_cliente = st.text_input("Cliente (manual)", value=cli_actual if cli_actual not in CLIENTES_FRECUENTES else "", key=f"ecl_m_{vid}")
                            else:
                                e_cliente = e_cli_sel

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
                            "fecha": e_fecha, "placa": e_placa, "conductor": e_conductor,
                            "cliente": e_cliente, "origen": e_origen, "destino": e_destino,
                            "hora_cita_cargue": e_hcc, "hora_salida_cargue": e_hsc,
                            "hora_llegada_descargue": e_hld, "hora_salida_descargue": e_hsd,
                            "contenedor": e_cont, "carga": e_carga,
                            "numero_importacion_bl": e_bl, "manifiesto": e_man,
                            "observacion": e_obs,
                            "estado": e_estado.split(" ", 1)[1] if " " in e_estado else e_estado
                        }
                        if db.actualizar_viaje(vid, datos_edit):
                            st.success("✅ Viaje actualizado.")
                            st.session_state.editando_id = None; st.rerun()
                    if cancelar:
                        st.session_state.editando_id = None; st.rerun()
        else:
            st.warning("No hay viajes con los filtros seleccionados.")

    # ===================== TAB 3: DASHBOARD =====================
    with tab3:
        st.markdown("### 📊 Dashboard de Operaciones")

        try:
            import plotly.express as px
            import plotly.graph_objects as go

            col_r1, col_r2 = st.columns([2, 4])
            with col_r1:
                rango = st.date_input(
                    "Período",
                    value=(datetime.now().replace(day=1), datetime.now()),
                    key="dash_rango"
                )

            if not (isinstance(rango, (list, tuple)) and len(rango) == 2):
                st.info("Selecciona un rango de fechas completo.")
                return

            df_s = db.stats_dashboard(rango[0], rango[1])

            if df_s.empty:
                st.info("No hay datos en este período.")
                return

            # ---- KPIs principales ----
            total = len(df_s)
            comp  = len(df_s[df_s["estado"].str.contains("Completado", na=False)])
            anul  = len(df_s[df_s["estado"].str.contains("Anulado",    na=False)])
            incum = len(df_s[df_s["estado"].str.contains("Incumplido", na=False)])
            curso = len(df_s[df_s["estado"].str.contains("En Curso",   na=False)])
            pct   = round(comp / total * 100) if total > 0 else 0

            k1, k2, k3, k4, k5 = st.columns(5)
            k1.metric("🚚 Total Viajes", total)
            k2.metric("✅ Completados", comp, f"{pct}%")
            k3.metric("❌ Anulados", anul)
            k4.metric("⚠️ Incumplidos", incum)
            k5.metric("🔄 En Curso", curso)

            st.divider()

            # ---- Fila 1: Estado + Viajes por día ----
            g1, g2 = st.columns(2)

            with g1:
                st.markdown("#### Distribución por Estado")
                est_c = df_s["estado"].value_counts().reset_index()
                est_c.columns = ["estado", "cantidad"]
                colores_estado = {
                    "Completado": "#2ecc71", "Anulado": "#e74c3c",
                    "Incumplido": "#f39c12", "En Curso": "#3498db"
                }
                fig1 = px.pie(est_c, values="cantidad", names="estado", hole=0.45,
                              color="estado", color_discrete_map=colores_estado)
                fig1.update_layout(margin=dict(t=10, b=10), height=300)
                st.plotly_chart(fig1, use_container_width=True)

            with g2:
                st.markdown("#### Viajes por Día")
                df_dia = df_s.groupby("fecha").size().reset_index(name="viajes")
                fig2 = px.bar(df_dia, x="fecha", y="viajes",
                              color_discrete_sequence=["#2c5364"],
                              text="viajes")
                fig2.update_traces(textposition="outside")
                fig2.update_layout(margin=dict(t=10, b=10), height=300,
                                   xaxis_title="", yaxis_title="Viajes")
                st.plotly_chart(fig2, use_container_width=True)

            st.divider()

            # ---- Fila 2: Por cliente + Por placa ----
            g3, g4 = st.columns(2)

            with g3:
                st.markdown("#### Viajes por Cliente")
                if "cliente" in df_s.columns and df_s["cliente"].notna().any():
                    df_cli = df_s.groupby("cliente").size().reset_index(name="viajes").sort_values("viajes")
                    fig3 = px.bar(df_cli, x="viajes", y="cliente", orientation="h",
                                  color="viajes", color_continuous_scale="Blues",
                                  text="viajes")
                    fig3.update_traces(textposition="outside")
                    fig3.update_layout(margin=dict(t=10, b=10), height=max(250, len(df_cli)*40),
                                       coloraxis_showscale=False, yaxis_title="", xaxis_title="Viajes")
                    st.plotly_chart(fig3, use_container_width=True)
                else:
                    st.info("Sin datos de cliente.")

            with g4:
                st.markdown("#### Viajes por Placa")
                df_placa = df_s.groupby("placa").size().reset_index(name="viajes").sort_values("viajes")
                fig4 = px.bar(df_placa, x="viajes", y="placa", orientation="h",
                              color="viajes", color_continuous_scale="Teal",
                              text="viajes")
                fig4.update_traces(textposition="outside")
                fig4.update_layout(margin=dict(t=10, b=10), height=max(250, len(df_placa)*40),
                                   coloraxis_showscale=False, yaxis_title="", xaxis_title="Viajes")
                st.plotly_chart(fig4, use_container_width=True)

            st.divider()

            # ---- Fila 3: Tiempos promedio + Ranking semanal ----
            g5, g6 = st.columns(2)

            with g5:
                st.markdown("#### ⏱️ Tiempos Promedio de Operación")
                tiempos = []
                for _, r in df_s.iterrows():
                    t_cargue = calcular_duracion(r["hora_cita_cargue"], r["hora_salida_cargue"])
                    t_transito = calcular_duracion(r["hora_salida_cargue"], r["hora_llegada_descargue"])
                    t_descargue = calcular_duracion(r["hora_llegada_descargue"], r["hora_salida_descargue"])
                    tiempos.append({
                        "espera_cargue": t_cargue,
                        "transito": t_transito,
                        "descargue": t_descargue
                    })
                df_t = pd.DataFrame(tiempos)
                prom = {
                    "Espera en Cargue": df_t["espera_cargue"].dropna().mean(),
                    "Tránsito": df_t["transito"].dropna().mean(),
                    "Descargue": df_t["descargue"].dropna().mean(),
                }
                prom_df = pd.DataFrame([
                    {"Etapa": k, "Minutos": round(v) if not pd.isna(v) else 0, "Tiempo": mins_a_str(v)}
                    for k, v in prom.items()
                ])
                fig5 = px.bar(prom_df, x="Etapa", y="Minutos",
                              color="Etapa", text="Tiempo",
                              color_discrete_sequence=["#2c5364", "#2980b9", "#1abc9c"])
                fig5.update_traces(textposition="outside")
                fig5.update_layout(margin=dict(t=10, b=10), height=300,
                                   showlegend=False, xaxis_title="", yaxis_title="Minutos promedio")
                st.plotly_chart(fig5, use_container_width=True)

            with g6:
                st.markdown("#### 📅 Ranking por Día de la Semana")
                df_s["dia_semana"] = pd.to_datetime(df_s["fecha"]).dt.day_name()
                orden = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
                nombres_es = {"Monday":"Lunes","Tuesday":"Martes","Wednesday":"Miércoles",
                              "Thursday":"Jueves","Friday":"Viernes","Saturday":"Sábado","Sunday":"Domingo"}
                df_s["dia_es"] = df_s["dia_semana"].map(nombres_es)
                df_semana = df_s.groupby("dia_semana").size().reset_index(name="viajes")
                df_semana["orden"] = df_semana["dia_semana"].map({d: i for i, d in enumerate(orden)})
                df_semana = df_semana.sort_values("orden")
                df_semana["dia_es"] = df_semana["dia_semana"].map(nombres_es)
                fig6 = px.bar(df_semana, x="dia_es", y="viajes",
                              color="viajes", color_continuous_scale="Oranges",
                              text="viajes")
                fig6.update_traces(textposition="outside")
                fig6.update_layout(margin=dict(t=10, b=10), height=300,
                                   coloraxis_showscale=False, xaxis_title="", yaxis_title="Viajes")
                st.plotly_chart(fig6, use_container_width=True)

            st.divider()

            # ---- Tabla ranking conductores ----
            st.markdown("#### 🏆 Ranking de Conductores")
            df_cond = df_s[df_s["conductor"].notna() & (df_s["conductor"].str.strip() != "")].groupby("conductor").agg(
                viajes=("conductor", "count"),
                completados=("estado", lambda x: x.str.contains("Completado", na=False).sum()),
                anulados=("estado", lambda x: x.str.contains("Anulado", na=False).sum()),
                incumplidos=("estado", lambda x: x.str.contains("Incumplido", na=False).sum()),
            ).reset_index().sort_values("viajes", ascending=False).drop_duplicates(subset="conductor")
            df_cond["% Cumplimiento"] = (df_cond["completados"] / df_cond["viajes"] * 100).round(1).astype(str) + "%"
            df_cond.columns = ["Conductor", "Total", "✅ Comp.", "❌ Anul.", "⚠️ Incump.", "% Cumplimiento"]
            st.dataframe(df_cond, use_container_width=True, hide_index=True)

        except ImportError:
            st.warning("Instala plotly: `pip install plotly`")
        except Exception as e:
            st.error(f"Error en dashboard: {e}")


if __name__ == "__main__":
    main()
