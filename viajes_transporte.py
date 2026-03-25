import streamlit as st
from supabase import create_client, Client
import pandas as pd
from datetime import datetime
import plotly.express as px
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

# ==================== CONFIGURACIÓN ====================
st.set_page_config(page_title="Gestión de Inventario Pro", layout="wide", page_icon="📦")

# Credenciales de Supabase (Sustituye con las tuyas)
SUPABASE_URL = "https://tu-proyecto.supabase.co"
SUPABASE_KEY = "tu-anon-key"

@st.cache_resource
def get_supabase() -> Client:
    return create_client(SUPABASE_URL, SUPABASE_KEY)

supabase = get_supabase()

# ==================== ESTILOS CSS ====================
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1e3c72, #2a5298);
        padding: 2rem; border-radius: 15px; color: white; margin-bottom: 2rem;
    }
    .stMetric { background: #f8f9fa; padding: 15px; border-radius: 10px; border: 1px solid #eee; }
</style>
""", unsafe_allow_html=True)

# ==================== LÓGICA DE AUTENTICACIÓN ====================
def login_ui():
    st.markdown("<div class='main-header'><h1>📦 Sistema de Inventario</h1><p>Accede para gestionar tus productos</p></div>", unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        menu = ["Login", "Registro"]
        choice = st.selectbox("Acción", menu)
        email = st.text_input("Correo Electrónico")
        password = st.text_input("Contraseña", type="password")

        if choice == "Registro":
            if st.button("Crear Cuenta"):
                try:
                    res = supabase.auth.sign_up({"email": email, "password": password})
                    st.success("Cuenta creada. Revisa tu correo para confirmar.")
                except Exception as e:
                    st.error(f"Error: {e}")
        else:
            if st.button("Entrar"):
                try:
                    res = supabase.auth.sign_in_with_password({"email": email, "password": password})
                    st.session_state.user = res.user
                    st.rerun()
                except Exception as e:
                    st.error("Credenciales inválidas")

# ==================== FUNCIONES DE DATOS ====================
def fetch_data():
    # Supabase filtra automáticamente por el RLS (User ID)
    res = supabase.table("inventario").select("*").order("fecha_registro").execute()
    return pd.DataFrame(res.data)

def save_product(data):
    data["user_id"] = st.session_state.user.id # Forzamos el ID del usuario actual
    return supabase.table("inventario").insert(data).execute()

# ==================== INTERFAZ PRINCIPAL ====================
def main_app():
    # Sidebar con info de usuario
    with st.sidebar:
        st.write(f"👤 **Usuario:**\n{st.session_state.user.email}")
        if st.button("Cerrar Sesión"):
            supabase.auth.sign_out()
            del st.session_state.user
            st.rerun()

    tab1, tab2, tab3 = st.tabs(["📊 Dashboard", "📦 Mi Inventario", "➕ Añadir Producto"])

    df = fetch_data()

    # --- TAB: DASHBOARD ---
    with tab1:
        if not df.empty:
            st.subheader("Estado General")
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Total Productos", len(df))
            k2.metric("Stock Total", df["cantidad"].sum())
            valor_inv = (df["cantidad"] * df["precio_compra"]).sum()
            k3.metric("Valor Inventario", f"${valor_inv:,.2f}")
            
            low_stock = len(df[df["cantidad"] < 5])
            k4.metric("Stock Bajo (<5)", low_stock, delta_color="inverse")

            c1, c2 = st.columns(2)
            with c1:
                fig = px.pie(df, names='categoria', values='cantidad', title="Distribución por Categoría")
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                fig2 = px.bar(df.nlargest(10, 'cantidad'), x='nombre', y='cantidad', title="Top 10 Productos en Stock")
                st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("No hay datos para mostrar. Agrega tu primer producto.")

    # --- TAB: INVENTARIO ---
    with tab2:
        st.subheader("Listado de Productos")
        if not df.empty:
            # Filtros
            search = st.text_input("🔍 Buscar por nombre o SKU")
            df_filtered = df[df['nombre'].str.contains(search, case=False) | df['sku'].str.contains(search, case=False)]
            
            st.dataframe(df_filtered.drop(columns=['user_id']), use_container_width=True, hide_index=True)
            
            # Exportar Excel (Lógica similar a tu código original)
            if st.button("📂 Descargar Excel"):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_filtered.to_excel(writer, index=False, sheet_name='Inventario')
                st.download_button("Descargar Archivo", data=output.getvalue(), file_name="inventario.xlsx")
        else:
            st.warning("Inventario vacío.")

    # --- TAB: AÑADIR PRODUCTO ---
    with tab3:
        st.subheader("Nuevo Item")
        with st.form("form_producto", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                nombre = st.text_input("Nombre del Producto*")
                sku = st.text_input("SKU / Código")
                categoria = st.selectbox("Categoría", ["Electrónica", "Ropa", "Alimentos", "Hogar", "Otros"])
            with col2:
                cantidad = st.number_input("Cantidad Inicial", min_value=0, step=1)
                p_compra = st.number_input("Precio Compra", min_value=0.0)
                p_venta = st.number_input("Precio Venta", min_value=0.0)
            
            ubicacion = st.text_input("Ubicación en Almacén")
            notas = st.text_area("Notas adicionales")
            
            if st.form_submit_button("Guardar Producto", type="primary"):
                if nombre:
                    nuevo_item = {
                        "nombre": nombre, "sku": sku, "categoria": categoria,
                        "cantidad": cantidad, "precio_compra": p_compra,
                        "precio_venta": p_venta, "ubicacion": ubicacion, "notas": notas
                    }
                    save_product(nuevo_item)
                    st.success("Producto guardado correctamente.")
                    st.rerun()
                else:
                    st.error("El nombre es obligatorio")

# ==================== PUNTO DE ENTRADA ====================
if __name__ == "__main__":
    if "user" not in st.session_state:
        login_ui()
    else:
        main_app()
