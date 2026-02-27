import streamlit as st
import psycopg2
import pandas as pd
from datetime import datetime, timedelta, time
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==================== CONFIGURACIÓN ====================
st.set_page_config(
    page_title="Control de Viajes",
    layout="wide",
    page_icon="🚚",
    initial_sidebar_state="collapsed"
)

# ==================== CREDENCIALES ====================
SUPABASE_DB_URL = "postgresql://postgres.hhzuggxvdzzfmnvfulmp:Negritasantia@aws-1-us-east-1.pooler.supabase.com:6543/postgres"

# ==================== ESTADOS DEL VIAJE ====================
ESTADOS_VIAJE = ["✅ Completado", "❌ Anulado", "⚠️ Incumplido", "🔄 En Curso"]

# ==================== CSS PERSONALIZADO ====================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@400;600;700&family=Barlow:wght@300;400;500&display=swap');

    html, body, [class*="css"] {
        font-family: 'Barlow', sans-serif;
    }
    .main-header {
        background: linear-gradient(135deg, #0f2027, #203a43, #2c5364);
        padding: 1.5rem 2rem;
        border-radius: 12px;
        margin-bottom: 1.5rem;
        display: flex;
        align-items: center;
        gap: 1rem;
    }
    .main-header h1 {
        font-family: 'Barlow Condensed', sans-serif;
        font-size: 2rem;
        font-weight: 700;
        color: white;
        margin: 0;
        letter-spacing: 1px;
    }
    .main-header p {
        color: #a0c4d8;
        margin: 0;
        font-size: 0.9rem;
    }
    .estado-completado { color: #2ecc71; font-weight: 600; }
    .estado-anulado    { color: #e74c3c; font-weight: 600; }
    .estado-incumplido { color: #f39c12; font-weight: 600; }
    .estado-encurso    { color: #3498db; font-weight: 600; }
    .kpi-card {
        background: #f8fafc;
        border-left: 4px solid #2c5364;
        border-radius: 8px;
        padding: 1rem 1.2rem;
        margin-bottom: 0.5rem;
    }
    div[data-testid="stTabs"] button {
        font-family: 'Barlow Condensed', sans-serif;
        font-weight: 600;
        font-size: 1rem;
        letter-spacing: 0.5px;
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
            # Columnas que pueden faltar en tablas ya existentes
            for col in [
                "ALTER TABLE viajes_transporte ADD COLUMN IF NOT EXISTS cliente TEXT",
                "ALTER TABLE viajes_transporte ADD COLUMN IF NOT EXISTS contenedor TEXT",
                "ALTER TABLE viajes_transporte ADD COLUMN IF NOT EXISTS numero_importacion_bl TEXT",
                "ALTER TABLE viajes_transporte ADD COLUMN IF NOT EXISTS manifiesto TEXT",
                "ALTER TABLE viajes_transporte ADD COLUMN IF NOT EXISTS estado TEXT DEFAULT 'Completado'",
            ]:
                try:
                    cur.execute(col)
                    c.commit()
                except Exception:
                    try: c.rollback()
                    except: pass
            c.commit()
            c.close()
        except Exception as e:
            st.error(f"Error DB init: {e}")

    def guardar_viaje(self, datos: dict) -> bool:
        try:
            c = self.conn()
            cur = c.cursor()
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
            st.error(f"Error guardando: {e}")
            return False

    def actualizar_viaje(self, viaje_id: int, datos: dict) -> bool:
        try:
            c = self.conn()
            cur = c.cursor()
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
            c.commit(); c.close()
            return True
        except Exception as e:
            st.error(f"Error actualizando: {e}")
            return False

    def eliminar_viaje(self, viaje_id: int) -> bool:
        try:
            c = self.conn()
            cur = c.cursor()
            cur.execute("DELETE FROM viajes_transporte WHERE id=%s", (viaje_id,))
            c.commit(); c.close()
            return True
        except Exception as e:
            st.error(f"Error eliminando: {e}")
            return False

    def obtener_viajes(self, fecha_ini=None, fecha_fin=None, placa=None,
                       conductor=None, cliente=None, estado=None) -> pd.DataFrame:
        c = self.conn()
        q = """
            SELECT id, fecha, placa, conductor, cliente, origen, destino,
                   hora_cita_cargue, hora_salida_cargue,
                   hora_llegada_descargue, hora_salida_descargue,
                   contenedor, carga, numero_importacion_bl,
                   manifiesto, observacion, estado
            FROM viajes_transporte WHERE 1=1
        """
        params = []
        if fecha_ini:
            q += " AND fecha >= %s"; params.append(fecha_ini)
        if fecha_fin:
            q += " AND fecha <= %s"; params.append(fecha_fin)
        if placa and placa != "Todas":
            q += " AND placa = %s"; params.append(placa)
        if conductor:
            q += " AND conductor ILIKE %s"; params.append(f"%{conductor}%")
        if cliente:
            q += " AND cliente ILIKE %s"; params.append(f"%{cliente}%")
        if estado and estado != "Todos":
            q += " AND estado = %s"; params.append(estado)
        q += " ORDER BY fecha DESC, id DESC"
        try:
            df = pd.read_sql(q, c, params=params)
            return df
        except:
            return pd.DataFrame()
        finally:
            c.close()

    def placas_unicas(self):
        c = self.conn()
        try:
            df = pd.read_sql("SELECT DISTINCT placa FROM viajes_transporte ORDER BY placa", c)
            return df["placa"].tolist()
        except:
            return []
        finally:
            c.close()


# ==================== EXCEL ====================
def generar_excel(df: pd.DataFrame, titulo: str = "Control de Viajes") -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Viajes"

    # Estilos
    azul_oscuro = "0F2027"
    azul_medio  = "203A43"
    azul_claro  = "D6EAF8"
    verde       = "1E8449"
    rojo        = "C0392B"
    naranja     = "D35400"

    ft_titulo  = Font(name="Calibri", bold=True, size=14, color="FFFFFF")
    ft_header  = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
    ft_normal  = Font(name="Calibri", size=9)
    ft_total   = Font(name="Calibri", bold=True, size=10)
    ft_anulado = Font(name="Calibri", size=9, color=rojo)
    ft_incump  = Font(name="Calibri", size=9, color=naranja)

    fill_titulo  = PatternFill("solid", start_color=azul_oscuro)
    fill_header  = PatternFill("solid", start_color=azul_medio)
    fill_alt     = PatternFill("solid", start_color="EBF5FB")
    fill_total   = PatternFill("solid", start_color="D5DBDB")
    fill_anulado = PatternFill("solid", start_color="FADBD8")
    fill_incump  = PatternFill("solid", start_color="FDEBD0")

    borde = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    centro = Alignment(horizontal="center", vertical="center", wrap_text=True)
    izq    = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    # --- TÍTULO ---
    ws.merge_cells("A1:R1")
    ws["A1"] = f"🚚 {titulo}   |   Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}   |   Total viajes: {len(df)}"
    ws["A1"].font = ft_titulo
    ws["A1"].fill = fill_titulo
    ws["A1"].alignment = centro
    ws.row_dimensions[1].height = 30

    # --- ENCABEZADOS ---
    columnas = [
        ("fecha",                  "FECHA",             12),
        ("placa",                  "PLACA",             12),
        ("conductor",              "CONDUCTOR",         22),
        ("cliente",                "CLIENTE",           22),
        ("origen",                 "ORIGEN",            18),
        ("destino",                "DESTINO",           18),
        ("hora_cita_cargue",       "H. CITA CARGUE",   14),
        ("hora_salida_cargue",     "H. SALIDA CARGUE", 14),
        ("hora_llegada_descargue", "H. LLEGADA DESC.",  14),
        ("hora_salida_descargue",  "H. SALIDA DESC.",   14),
        ("contenedor",             "CONTENEDOR",        16),
        ("carga",                  "CARGA",             18),
        ("numero_importacion_bl",  "IMP / BL",          18),
        ("manifiesto",             "MANIFIESTO",        14),
        ("observacion",            "OBSERVACIÓN",       28),
        ("estado",                 "ESTADO",            14),
    ]

    col_keys  = [c[0] for c in columnas if c[0] in df.columns or c[0] in [x[0] for x in columnas]]
    col_names = {c[0]: c[1] for c in columnas}
    col_widths = {c[0]: c[2] for c in columnas}

    header_row = 2
    for idx, (key, nombre, ancho) in enumerate(columnas, start=1):
        cell = ws.cell(row=header_row, column=idx, value=nombre)
        cell.font = ft_header
        cell.fill = fill_header
        cell.alignment = centro
        cell.border = borde
        ws.column_dimensions[get_column_letter(idx)].width = ancho
    ws.row_dimensions[header_row].height = 30

    # --- DATOS ---
    for row_idx, (_, fila) in enumerate(df.iterrows(), start=3):
        estado_val = str(fila.get("estado", "")).strip()
        es_anulado   = "Anulado"   in estado_val
        es_incumplido = "Incumplido" in estado_val
        fill_fila = fill_anulado if es_anulado else (fill_incump if es_incumplido else (fill_alt if row_idx % 2 == 0 else None))

        for col_idx, (key, _, _) in enumerate(columnas, start=1):
            val = fila.get(key, "")
            if pd.isna(val) if not isinstance(val, str) else False:
                val = ""
            # Formatear horas
            if key.startswith("hora_") and val and val != "":
                try:
                    val = str(val)[:5]  # HH:MM
                except:
                    val = ""
            cell = ws.cell(row=row_idx, column=col_idx, value=str(val) if val != "" else "")
            cell.border = borde
            cell.alignment = centro if key in ("fecha","placa","estado") or key.startswith("hora_") else izq
            if es_anulado:
                cell.font = ft_anulado
            elif es_incumplido:
                cell.font = ft_incump
            else:
                cell.font = ft_normal
            if fill_fila:
                cell.fill = fill_fila

        ws.row_dimensions[row_idx].height = 18

    # --- FILA TOTALES ---
    total_row = len(df) + 3
    ws.merge_cells(f"A{total_row}:O{total_row}")
    cell_tot = ws.cell(row=total_row, column=1, value=f"TOTAL VIAJES: {len(df)}")
    cell_tot.font = ft_total
    cell_tot.fill = fill_total
    cell_tot.alignment = centro

    completados  = len(df[df["estado"].str.contains("Completado",  na=False)]) if "estado" in df.columns else 0
    anulados     = len(df[df["estado"].str.contains("Anulado",     na=False)]) if "estado" in df.columns else 0
    incumplidos  = len(df[df["estado"].str.contains("Incumplido",  na=False)]) if "estado" in df.columns else 0

    resumen_txt = f"✅ Completados: {completados}   |   ❌ Anulados: {anulados}   |   ⚠️ Incumplidos: {incumplidos}"
    ws.merge_cells(f"P{total_row}:R{total_row}")
    cell_res = ws.cell(row=total_row, column=16, value=resumen_txt)
    cell_res.font = ft_total
    cell_res.fill = fill_total
    cell_res.alignment = centro

    # --- HOJA RESUMEN ---
    ws2 = wb.create_sheet("Resumen")
    ws2.merge_cells("A1:D1")
    ws2["A1"] = "Resumen General de Viajes"
    ws2["A1"].font = Font(name="Calibri", bold=True, size=13, color="FFFFFF")
    ws2["A1"].fill = fill_titulo
    ws2["A1"].alignment = centro
    ws2.row_dimensions[1].height = 26

    resumen_datos = [
        ("MÉTRICA", "VALOR"),
        ("Total Viajes", len(df)),
        ("✅ Completados", completados),
        ("❌ Anulados", anulados),
        ("⚠️ Incumplidos", incumplidos),
        ("🔄 En Curso", len(df[df["estado"].str.contains("En Curso", na=False)]) if "estado" in df.columns else 0),
    ]
    for r_idx, (metrica, valor) in enumerate(resumen_datos, start=2):
        c1 = ws2.cell(r_idx, 1, metrica)
        c2 = ws2.cell(r_idx, 2, valor)
        c1.border = borde; c2.border = borde
        c1.alignment = izq; c2.alignment = centro
        if r_idx == 2:
            c1.font = ft_header; c2.font = ft_header
            c1.fill = fill_header; c2.fill = fill_header
        else:
            c1.font = ft_normal; c2.font = ft_total

    # Resumen por cliente
    if "cliente" in df.columns and df["cliente"].notna().any():
        ws2.cell(9, 1, "VIAJES POR CLIENTE").font = ft_header
        ws2.cell(9, 1).fill = fill_header
        ws2.cell(9, 1).alignment = centro
        ws2.merge_cells("A9:B9")
        por_cliente = df.groupby("cliente").size().reset_index(name="viajes").sort_values("viajes", ascending=False)
        for i, row in por_cliente.iterrows():
            r = i + 10
            ws2.cell(r, 1, row["cliente"]).border = borde
            ws2.cell(r, 2, int(row["viajes"])).border = borde
            ws2.cell(r, 1).font = ft_normal; ws2.cell(r, 2).font = ft_normal
            ws2.cell(r, 1).alignment = izq; ws2.cell(r, 2).alignment = centro

    for col_l, w in zip(["A","B","C","D"], [30, 14, 14, 14]):
        ws2.column_dimensions[col_l].width = w

    # Congelar primera fila de datos
    ws.freeze_panes = "A3"

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# ==================== HELPERS ====================
def hora_a_time(val):
    """Convierte valor de BD a objeto time o None"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, time):
        return val
    try:
        s = str(val)[:5]
        h, m = s.split(":")
        return time(int(h), int(m))
    except:
        return None

def str_hora(val):
    """Muestra hora como HH:MM o vacío"""
    t = hora_a_time(val)
    return t.strftime("%H:%M") if t else ""


# ==================== MAIN ====================
def main():
    st.markdown("""
    <div class="main-header">
        <div>
            <h1>🚚 CONTROL DE VIAJES</h1>
            <p>Registro y seguimiento de operaciones de transporte</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if "db" not in st.session_state:
        st.session_state.db = DB()
    if "editando_id" not in st.session_state:
        st.session_state.editando_id = None

    db = st.session_state.db

    tab1, tab2, tab3 = st.tabs(["📝 Nuevo Viaje", "🔍 Historial y Reportes", "📊 Estadísticas"])

    # ===================== TAB 1: NUEVO VIAJE =====================
    with tab1:
        st.markdown("### Registrar Nuevo Viaje")

        with st.form("form_viaje", clear_on_submit=True):
            # Fila 1: Datos básicos
            c1, c2, c3, c4 = st.columns(4)
            with c1: fecha      = st.date_input("📅 Fecha", datetime.now())
            with c2: placa      = st.text_input("🚛 Placa").upper().strip()
            with c3: conductor  = st.text_input("👤 Conductor")
            with c4: cliente    = st.text_input("🏢 Cliente")

            # Fila 2: Ruta
            c5, c6 = st.columns(2)
            with c5: origen  = st.text_input("📍 Origen")
            with c6: destino = st.text_input("🏁 Destino")

            # Fila 3: Horas
            st.markdown("#### ⏱️ Tiempos de Operación")
            h1, h2, h3, h4 = st.columns(4)
            with h1: hora_cita_cargue       = st.time_input("Hora Cita Cargue",       value=None, step=300)
            with h2: hora_salida_cargue     = st.time_input("Hora Salida Cargue",     value=None, step=300)
            with h3: hora_llegada_descargue = st.time_input("Hora Llegada Descargue", value=None, step=300)
            with h4: hora_salida_descargue  = st.time_input("Hora Salida Descargue",  value=None, step=300)

            # Fila 4: Info carga
            st.markdown("#### 📦 Información de Carga")
            d1, d2, d3, d4 = st.columns(4)
            with d1: contenedor          = st.text_input("Contenedor")
            with d2: carga               = st.text_input("Carga")
            with d3: numero_importacion  = st.text_input("Nº Importación / BL")
            with d4: manifiesto          = st.text_input("Manifiesto")

            # Fila 5: Estado y observación
            e1, e2 = st.columns([1, 3])
            with e1:
                estado = st.selectbox("🚦 Estado del Viaje", ESTADOS_VIAJE)
            with e2:
                observacion = st.text_area("📝 Observaciones", height=80)

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
                    st.success(f"✅ Viaje guardado: {placa} | {origen} → {destino} | {estado}")
                    st.balloons()

    # ===================== TAB 2: HISTORIAL =====================
    with tab2:
        st.markdown("### 🔍 Historial de Viajes")

        with st.expander("🛠️ Filtros", expanded=True):
            f1, f2, f3, f4, f5, f6 = st.columns(6)
            with f1: fi = st.date_input("Desde", datetime.now() - timedelta(days=30), key="h_fi")
            with f2: ff = st.date_input("Hasta", datetime.now(), key="h_ff")
            with f3:
                placas = ["Todas"] + db.placas_unicas()
                fp = st.selectbox("Placa", placas, key="h_fp")
            with f4: fc = st.text_input("Conductor", key="h_fc")
            with f5: fcli = st.text_input("Cliente", key="h_fcli")
            with f6:
                estados_filtro = ["Todos"] + [e.split(" ", 1)[1] for e in ESTADOS_VIAJE]
                fe = st.selectbox("Estado", estados_filtro, key="h_fe")

        df = db.obtener_viajes(fi, ff, fp, fc, fcli, fe if fe != "Todos" else None)

        if not df.empty:
            # KPIs rápidos
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Total Viajes", len(df))
            k2.metric("✅ Completados", len(df[df["estado"].str.contains("Completado", na=False)]))
            k3.metric("❌ Anulados",    len(df[df["estado"].str.contains("Anulado",    na=False)]))
            k4.metric("⚠️ Incumplidos", len(df[df["estado"].str.contains("Incumplido", na=False)]))

            st.divider()

            # Exportar Excel
            col_exp1, col_exp2 = st.columns([2, 5])
            with col_exp1:
                nombre_rep = st.text_input("Nombre del reporte", value="Control_Viajes", key="rep_nombre")
            with col_exp2:
                st.markdown("<br>", unsafe_allow_html=True)
                excel_data = generar_excel(df, titulo=nombre_rep)
                st.download_button(
                    "⬇️ Descargar Excel",
                    data=excel_data,
                    file_name=f"{nombre_rep}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

            st.divider()

            # Tabla con formato de estado
            df_display = df.copy()
            df_display["horas"] = df_display.apply(
                lambda r: f"{str_hora(r['hora_cita_cargue'])} / {str_hora(r['hora_salida_cargue'])} / {str_hora(r['hora_llegada_descargue'])} / {str_hora(r['hora_salida_descargue'])}", axis=1
            )

            cols_tabla = ["id","fecha","placa","conductor","cliente","origen","destino",
                          "contenedor","carga","numero_importacion_bl","manifiesto","estado"]
            cols_existentes = [c for c in cols_tabla if c in df_display.columns]
            st.dataframe(df_display[cols_existentes], use_container_width=True, hide_index=True)

            st.divider()

            # --- DETALLE / EDICIÓN ---
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
                    # Vista lectura
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

                    st.write(f"**Horas:** Cita Cargue: `{str_hora(row['hora_cita_cargue'])}` | Salida Cargue: `{str_hora(row['hora_salida_cargue'])}` | Llegada Desc: `{str_hora(row['hora_llegada_descargue'])}` | Salida Desc: `{str_hora(row['hora_salida_descargue'])}`")

                    bc1, bc2 = st.columns(2)
                    with bc1:
                        if st.button("✏️ Editar", key=f"eb_{vid}"):
                            st.session_state.editando_id = vid
                            st.rerun()
                    with bc2:
                        if st.button("🗑️ Eliminar", key=f"del_{vid}"):
                            db.eliminar_viaje(vid)
                            st.success("Eliminado.")
                            st.rerun()
                else:
                    # Formulario edición
                    st.markdown("#### ✏️ Editando viaje")
                    with st.form(f"edit_{vid}"):
                        ec1, ec2, ec3, ec4 = st.columns(4)
                        with ec1: e_fecha     = st.date_input("Fecha", value=row["fecha"], key=f"ef_{vid}")
                        with ec2: e_placa     = st.text_input("Placa", value=str(row["placa"] or ""), key=f"ep_{vid}").upper()
                        with ec3: e_conductor = st.text_input("Conductor", value=str(row.get("conductor") or ""), key=f"ec_{vid}")
                        with ec4: e_cliente   = st.text_input("Cliente", value=str(row.get("cliente") or ""), key=f"ecl_{vid}")

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
                        with ed1: e_cont  = st.text_input("Contenedor",   value=str(row.get("contenedor") or ""),           key=f"eco_{vid}")
                        with ed2: e_carga = st.text_input("Carga",        value=str(row.get("carga") or ""),                key=f"eca_{vid}")
                        with ed3: e_bl    = st.text_input("Imp / BL",     value=str(row.get("numero_importacion_bl") or ""), key=f"ebl_{vid}")
                        with ed4: e_man   = st.text_input("Manifiesto",   value=str(row.get("manifiesto") or ""),           key=f"ema_{vid}")

                        estado_actual = str(row.get("estado") or "Completado")
                        estados_limpios = [e.split(" ", 1)[1] for e in ESTADOS_VIAJE]
                        est_idx = estados_limpios.index(estado_actual) if estado_actual in estados_limpios else 0
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
                            st.session_state.editando_id = None
                            st.rerun()
                    if cancelar:
                        st.session_state.editando_id = None
                        st.rerun()
        else:
            st.warning("No hay viajes con los filtros seleccionados.")

    # ===================== TAB 3: ESTADÍSTICAS =====================
    with tab3:
        st.markdown("### 📊 Estadísticas")
        try:
            import plotly.express as px

            col_f1, col_f2 = st.columns([2, 4])
            with col_f1:
                rango = st.date_input("Rango", value=(datetime.now().replace(day=1), datetime.now()), key="stat_rango")

            if isinstance(rango, tuple) and len(rango) == 2:
                df_stat = db.obtener_viajes(rango[0], rango[1])
                if not df_stat.empty:
                    s1, s2, s3, s4 = st.columns(4)
                    s1.metric("Total", len(df_stat))
                    s2.metric("✅ Completados", len(df_stat[df_stat["estado"].str.contains("Completado", na=False)]))
                    s3.metric("❌ Anulados",    len(df_stat[df_stat["estado"].str.contains("Anulado",    na=False)]))
                    s4.metric("⚠️ Incumplidos", len(df_stat[df_stat["estado"].str.contains("Incumplido", na=False)]))

                    st.divider()
                    g1, g2 = st.columns(2)

                    with g1:
                        st.subheader("Viajes por Estado")
                        est_count = df_stat["estado"].value_counts().reset_index()
                        est_count.columns = ["estado", "cantidad"]
                        fig1 = px.pie(est_count, values="cantidad", names="estado", hole=0.4,
                                      color_discrete_map={"Completado":"#2ecc71","Anulado":"#e74c3c","Incumplido":"#f39c12","En Curso":"#3498db"})
                        st.plotly_chart(fig1, use_container_width=True)

                    with g2:
                        st.subheader("Viajes por Día")
                        df_dia = df_stat.groupby("fecha").size().reset_index(name="viajes")
                        fig2 = px.bar(df_dia, x="fecha", y="viajes", color_discrete_sequence=["#2c5364"])
                        st.plotly_chart(fig2, use_container_width=True)

                    if "cliente" in df_stat.columns and df_stat["cliente"].notna().any():
                        st.subheader("Viajes por Cliente")
                        df_cli = df_stat.groupby("cliente").size().reset_index(name="viajes").sort_values("viajes", ascending=True)
                        fig3 = px.bar(df_cli, x="viajes", y="cliente", orientation="h", color="viajes",
                                      color_continuous_scale="Blues")
                        st.plotly_chart(fig3, use_container_width=True)
                else:
                    st.info("No hay datos en este rango.")
        except ImportError:
            st.warning("Instala plotly para ver gráficas: `pip install plotly`")


if __name__ == "__main__":
    main()
