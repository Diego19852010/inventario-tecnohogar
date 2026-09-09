from datetime import datetime
import io
import sqlite3
import zoneinfo  # Librería nativa para manejo de zonas horarias
import cv2
import numpy as np
import pandas as pd
import plotly.express as px
import qrcode
import streamlit as st

# Librerías para PDF
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

DB_NAME = "inventario_v2.db"


def obtener_fecha_hora_actual():
  """Obtiene la fecha y hora actual ajustada a la zona horaria local."""
  tz = zoneinfo.ZoneInfo("America/Santiago")
  return datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")


def init_db():
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()

  cursor.execute("""
        CREATE TABLE IF NOT EXISTS productos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT UNIQUE,
            nombre TEXT NOT NULL,
            categoria TEXT NOT NULL,
            precio REAL NOT NULL,
            stock INTEGER NOT NULL,
            stock_minimo INTEGER NOT NULL
        )
    """)

  cursor.execute("""
        CREATE TABLE IF NOT EXISTS movimientos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            producto_id INTEGER NOT NULL,
            tipo TEXT NOT NULL,
            cantidad INTEGER NOT NULL,
            fecha_hora TEXT NOT NULL,
            FOREIGN KEY (producto_id) REFERENCES productos (id)
        )
    """)

  cursor.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            rol TEXT NOT NULL
        )
    """)

  cursor.execute("SELECT COUNT(*) FROM usuarios")
  if cursor.fetchone()[0] == 0:
    cursor.execute(
        "INSERT INTO usuarios (username, password, rol) VALUES (?, ?, ?)",
        ("admin", "admin123", "Administrador"),
    )
    cursor.execute(
        "INSERT INTO usuarios (username, password, rol) VALUES (?, ?, ?)",
        ("vendedor", "vendedor123", "Vendedor"),
    )

  conn.commit()
  conn.close()


def verificar_usuario(username, password):
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()
  cursor.execute(
      "SELECT username, rol FROM usuarios WHERE username = ? AND password = ?",
      (username, password),
  )
  user = cursor.fetchone()
  conn.close()
  return user


def obtener_productos():
  conn = sqlite3.connect(DB_NAME)
  df = pd.read_sql_query("SELECT * FROM productos", conn)
  conn.close()
  return df


def obtener_historial():
  conn = sqlite3.connect(DB_NAME)
  query = """
        SELECT 
            m.id, 
            COALESCE(p.codigo, 'N/A') AS codigo, 
            COALESCE(p.nombre, 'Producto Eliminado') AS producto, 
            COALESCE(p.categoria, 'Sin Categoría') AS categoria, 
            COALESCE(p.precio, 0) AS precio, 
            m.tipo, 
            m.cantidad, 
            m.fecha_hora 
        FROM movimientos m
        LEFT JOIN productos p ON m.producto_id = p.id
        ORDER BY m.id DESC
    """
  df = pd.read_sql_query(query, conn)
  conn.close()
  return df


def agregar_producto(codigo, nombre, categoria, precio, stock, stock_min):
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()
  cursor.execute(
      """
        INSERT INTO productos (codigo, nombre, categoria, precio, stock, stock_minimo)
        VALUES (?, ?, ?, ?, ?, ?)
    """,
      (codigo, nombre, categoria, float(precio), int(stock), int(stock_min)),
  )

  prod_id = cursor.lastrowid
  fecha = obtener_fecha_hora_actual()
  cursor.execute(
      """
        INSERT INTO movimientos (producto_id, tipo, cantidad, fecha_hora)
        VALUES (?, ?, ?, ?)
    """,
      (prod_id, "Registro Inicial", int(stock), fecha),
  )

  conn.commit()
  conn.close()


def registrar_movimiento(prod_id, tipo, cantidad_cambio, nuevo_stock):
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()

  cursor.execute(
      "UPDATE productos SET stock = ? WHERE id = ?",
      (int(nuevo_stock), int(prod_id)),
  )

  fecha = obtener_fecha_hora_actual()
  cursor.execute(
      """
        INSERT INTO movimientos (producto_id, tipo, cantidad, fecha_hora)
        VALUES (?, ?, ?, ?)
    """,
      (int(prod_id), str(tipo), int(cantidad_cambio), fecha),
  )

  conn.commit()
  conn.close()


def generar_excel():
  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df_prod = obtener_productos()
    df_hist = obtener_historial()
    df_prod.to_excel(writer, sheet_name="Inventario Actual", index=False)
    df_hist.to_excel(writer, sheet_name="Historial Movimientos", index=False)
  return output.getvalue()


def generar_pdf_boleta(
    producto_nombre, cantidad, precio_unitario, total, cliente_nombre
):
  buffer = io.BytesIO()
  doc = SimpleDocTemplate(
      buffer,
      pagesize=letter,
      rightMargin=40,
      leftMargin=40,
      topMargin=40,
      bottomMargin=40,
  )
  elements = []

  styles = getSampleStyleSheet()
  title_style = ParagraphStyle(
      "TitleStyle",
      parent=styles["Heading1"],
      fontSize=18,
      textColor=colors.HexColor("#1E3A8A"),
  )
  subtitle_style = ParagraphStyle(
      "SubTitleStyle",
      parent=styles["Normal"],
      fontSize=10,
      textColor=colors.gray,
  )

  elements.append(
      Paragraph("<b>TECNOHOGAR - COMPROBANTE DE VENTA</b>", title_style)
  )
  elements.append(
      Paragraph(f"Fecha: {obtener_fecha_hora_actual()}", subtitle_style)
  )
  elements.append(
      Paragraph(
          f"Cliente: {cliente_nombre if cliente_nombre else 'Cliente General'}",
          subtitle_style,
      )
  )
  elements.append(Spacer(1, 15))

  data = [
      ["Producto", "Cantidad", "Precio Unitario", "Total"],
      [
          producto_nombre,
          str(cantidad),
          f"${precio_unitario:.2f}",
          f"${total:.2f}",
      ],
  ]

  tabla = Table(data, colWidths=[200, 80, 100, 100])
  tabla.setStyle(
      TableStyle([
          ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E3A8A")),
          ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
          ("ALIGN", (0, 0), (-1, -1), "CENTER"),
          ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
          ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
          ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#F3F4F6")),
          ("GRID", (0, 0), (-1, -1), 1, colors.HexColor("#D1D5DB")),
      ])
  )

  elements.append(tabla)
  elements.append(Spacer(1, 20))
  elements.append(
      Paragraph(
          f"<b>TOTAL PAGADO: ${total:.2f}</b>",
          ParagraphStyle(
              "Total",
              parent=styles["Heading2"],
              fontSize=14,
              textColor=colors.HexColor("#059669"),
          ),
      )
  )
  elements.append(Spacer(1, 30))
  elements.append(
      Paragraph("<i>¡Gracias por tu compra en TecnoHogar!</i>", subtitle_style)
  )

  doc.build(elements)
  return buffer.getvalue()


init_db()

st.set_page_config(
    page_title="TecnoHogar - Sistema de Gestión", layout="wide", page_icon="📦"
)
st.markdown(
    '<meta name="google" content="notranslate">', unsafe_allow_html=True
)

# --- INYECCIÓN DE CSS PARA DISEÑO PROFESIONAL ---
st.markdown(
    """
    <style>
    .stApp {
        background-color: #F8FAFC;
    }
    h1, h2, h3 {
        font-family: 'Inter', sans-serif;
        color: #0F172A;
        font-weight: 700;
    }
    section[data-testid="stSidebar"] {
        background-color: #0F172A !important;
    }
    section[data-testid="stSidebar"] * {
        color: #F8FAFC !important;
    }
    .metric-card {
        background: #FFFFFF;
        border-radius: 12px;
        padding: 20px;
        border: 1px solid #E2E8F0;
        box-shadow: 0px 4px 12px rgba(0, 0, 0, 0.03);
        text-align: center;
    }
    .metric-title {
        color: #64748B;
        font-size: 0.9rem;
        font-weight: 600;
        text-transform: uppercase;
        margin-bottom: 8px;
    }
    .metric-value {
        color: #0F172A;
        font-size: 1.8rem;
        font-weight: 800;
    }
    .stButton>button {
        background-color: #2563EB;
        color: white;
        border-radius: 8px;
        font-weight: 600;
        border: none;
        padding: 0.5rem 1rem;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #1D4ED8;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.3);
    }
    div[data-testid="stForm"] {
        background: #FFFFFF;
        padding: 2rem;
        border-radius: 12px;
        border: 1px solid #E2E8F0;
        box-shadow: 0px 4px 12px rgba(0, 0, 0, 0.02);
    }
    </style>
""",
    unsafe_allow_html=True,
)

if "logged_in" not in st.session_state:
  st.session_state["logged_in"] = False
  st.session_state["username"] = ""
  st.session_state["rol"] = ""

if not st.session_state["logged_in"]:
  col_login1, col_login2, col_login3 = st.columns([1, 2, 1])

  with col_login2:
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown(
        "<h2 style='text-align: center;'>🔐 Iniciar Sesión</h2>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<p style='text-align: center; color: #64748B;'>TecnoHogar - Sistema"
        " de Gestión</p>",
        unsafe_allow_html=True,
    )

    with st.form("login_form"):
      user_input = st.text_input("Usuario")
      pass_input = st.text_input("Contraseña", type="password")
      submit = st.form_submit_button("Ingresar al Sistema", use_container_width=True)

      if submit:
        usuario = verificar_usuario(user_input, pass_input)
        if usuario:
          st.session_state["logged_in"] = True
          st.session_state["username"] = usuario[0]
          st.session_state["rol"] = usuario[1]
          st.success(f"¡Bienvenido {usuario[0]} ({usuario[1]})!")
          st.rerun()
        else:
          st.error("Usuario o contraseña incorrectos.")

    st.info(
        "**Cuentas de prueba:**\n- **Admin:** `admin` | `admin123`\n-"
        " **Vendedor:** `vendedor` | `vendedor123`"
    )

else:
  st.sidebar.markdown(
      f"### 👤 {st.session_state['username']}\n**Rol:**"
      f" {st.session_state['rol']}"
  )
  st.sidebar.markdown("---")

  if st.sidebar.button("🚪 Cerrar Sesión", use_container_width=True):
    st.session_state["logged_in"] = False
    st.rerun()

  st.sidebar.markdown("---")

  if st.session_state["rol"] == "Administrador":
    opciones = [
        "Ver Inventario",
        "Registrar Venta",
        "Escanear Código / QR",
        "Dashboard y Gráficos",
        "Agregar Producto",
        "Ajustar Stock",
        "Historial de Movimientos",
        "Exportar Reportes",
    ]
  else:
    opciones = ["Ver Inventario", "Registrar Venta", "Escanear Código / QR"]

  opcion = st.sidebar.radio("Navegación", opciones)

  st.markdown(
      "<h1 style='color: #0F172A;'>📦 TecnoHogar Management</h1>",
      unsafe_allow_html=True,
  )
  st.markdown("---")

  # --- 1. VER INVENTARIO ---
  if opcion == "Ver Inventario":
    st.subheader("📋 Lista de Productos en Stock")
    df = obtener_productos()

    if df.empty:
      st.info("El inventario está vacío. Agrega un producto desde el menú.")
    else:
      col_f1, col_f2 = st.columns(2)
      with col_f1:
        busqueda = st.text_input("🔍 Buscar por nombre o código")
      with col_f2:
        categorias = ["Todas"] + list(df["categoria"].unique())
        cat_filtro = st.selectbox("📂 Filtrar por categoría", categorias)

      df_filtrado = df.copy()
      if busqueda:
        df_filtrado = df_filtrado[
            df_filtrado["nombre"].str.contains(busqueda, case=False)
            | df_filtrado["codigo"].str.contains(busqueda, case=False)
        ]
      if cat_filtro != "Todas":
        df_filtrado = df_filtrado[df_filtrado["categoria"] == cat_filtro]

      bajo_stock = df[df["stock"] <= df["stock_minimo"]]
      if not bajo_stock.empty:
        st.warning(
            f"⚠️ Atención: Hay **{len(bajo_stock)}** producto(s) con stock"
            " bajo el mínimo."
        )

      st.dataframe(df_filtrado, use_container_width=True, hide_index=True)

  # --- 2. ESCANEAR CÓDIGO / QR ---
  elif opcion == "Escanear Código / QR":
    st.subheader("📷 Escáner de Código QR / Barras")
    st.write(
        "Escanea el código del producto utilizando la cámara del dispositivo."
    )

    img_file = st.camera_input("Capturar Código")

    if img_file:
      bytes_data = img_file.getvalue()
      cv_img = cv2.imdecode(
          np.frombuffer(bytes_data, np.uint8), cv2.IMREAD_COLOR
      )
      detector = cv2.QRCodeDetector()
      data, bbox, _ = detector.detectAndDecode(cv_img)

      if data:
        st.success(f"¡Código detectado correctamente!: **{data}**")
        df = obtener_productos()
        match = df[df["codigo"] == data]

        if not match.empty:
          prod = match.iloc[0]
          st.info(
              f"**Producto:** {prod['nombre']} | **Precio:** ${prod['precio']}"
              f" | **Stock Actual:** {prod['stock']}"
          )
        else:
          st.warning("No existe un producto registrado con ese código.")
      else:
        st.error(
            "No se reconoció un código válido. Intenta enfocar mejor la imagen."
        )

  # --- 3. REGISTRAR VENTA ---
  elif opcion == "Registrar Venta":
    st.subheader("🛒 Punto de Venta")
    df = obtener_productos()

    if df.empty:
      st.info("No existen productos registrados para realizar ventas.")
    else:
      col_v1, col_v2 = st.columns([2, 1])

      with col_v1:
        prod_nom = st.selectbox("Seleccionar Producto", df["nombre"].tolist())
        prod = df[df["nombre"] == prod_nom].iloc[0]
        stock_actual = int(prod["stock"])

        cliente = st.text_input("Nombre del Cliente", value="Cliente General")
        cant_venta = st.number_input(
            "Cantidad", min_value=1, max_value=max(1, stock_actual), value=1
        )

      with col_v2:
        total = cant_venta * prod["precio"]
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-title">RESUMEN DE VENTA</div>
                <div class="metric-value">${total:,.2f}</div>
                <p style="color:#64748B; font-size:0.85rem; margin-top:8px;">
                    {prod['nombre']}<br>Stock disp: {stock_actual}
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown("<br>", unsafe_allow_html=True)

        if stock_actual <= 0:
          st.error("Sin Stock Disponible")
        else:
          if st.button("🛍️ Confirmar Venta", use_container_width=True):
            nuevo_stock = stock_actual - cant_venta
            registrar_movimiento(prod["id"], "Venta", cant_venta, nuevo_stock)
            st.success("¡Venta realizada con éxito!")

            pdf_bytes = generar_pdf_boleta(
                prod["nombre"], cant_venta, prod["precio"], total, cliente
            )

            st.download_button(
                label="📄 Descargar Boleta PDF",
                data=pdf_bytes,
                file_name=(
                    f"boleta_{prod['nombre']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
                ),
                mime="application/pdf",
                use_container_width=True,
            )

  # --- 4. DASHBOARD Y GRÁFICOS ---
  elif opcion == "Dashboard y Gráficos":
    st.subheader("📊 Métricas de Rendimiento")
    df_hist = obtener_historial()

    if df_hist.empty or "tipo" not in df_hist.columns:
      st.info("Aún no existen transacciones registradas.")
    else:
      df_hist["tipo_clean"] = (
          df_hist["tipo"].astype(str).str.strip().str.lower()
      )
      df_ventas = df_hist[df_hist["tipo_clean"].str.contains("venta")].copy()

      if df_ventas.empty:
        st.info("No se han registrado ventas hasta el momento.")
      else:
        df_ventas["cantidad"] = pd.to_numeric(
            df_ventas["cantidad"], errors="coerce"
        ).fillna(0)
        df_ventas["precio"] = pd.to_numeric(
            df_ventas["precio"], errors="coerce"
        ).fillna(0)
        df_ventas["monto_total"] = df_ventas["cantidad"] * df_ventas["precio"]

        m1, m2, m3 = st.columns(3)
        with m1:
          st.markdown(
              f"""
                <div class="metric-card">
                    <div class="metric-title">TOTAL RECAUDADO</div>
                    <div class="metric-value">${df_ventas['monto_total'].sum():,.2f}</div>
                </div>
                """,
              unsafe_allow_html=True,
          )

        with m2:
          st.markdown(
              f"""
                <div class="metric-card">
                    <div class="metric-title">UNIDADES VENDIDAS</div>
                    <div class="metric-value">{int(df_ventas['cantidad'].sum()):,}</div>
                </div>
                """,
              unsafe_allow_html=True,
          )

        with m3:
          st.markdown(
              f"""
                <div class="metric-card">
                    <div class="metric-title">TRANSACCIONES</div>
                    <div class="metric-value">{len(df_ventas)}</div>
                </div>
                """,
              unsafe_allow_html=True,
          )

        st.markdown("<br>", unsafe_allow_html=True)
        col_g1, col_g2 = st.columns(2)

        with col_g1:
          st.markdown("### 🏆 Productos Más Vendidos")
          ventas_por_prod = (
              df_ventas.groupby("producto")["cantidad"].sum().reset_index()
          )
          fig_bar = px.bar(
              ventas_por_prod,
              x="producto",
              y="cantidad",
              labels={"producto": "Producto", "cantidad": "Unidades"},
              color="cantidad",
              color_continuous_scale="Blues",
          )
          fig_bar.update_layout(
              paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)"
          )
          st.plotly_chart(fig_bar, use_container_width=True)

        with col_g2:
          st.markdown("### 🏷️ Ventas por Categoría")
          ventas_por_cat = (
              df_ventas.groupby("categoria")["monto_total"].sum().reset_index()
          )
          fig_pie = px.pie(
              ventas_por_cat,
              values="monto_total",
              names="categoria",
              hole=0.4,
              color_discrete_sequence=px.colors.qualitative.Set2,
          )
          fig_pie.update_layout(paper_bgcolor="rgba(0,0,0,0)")
          st.plotly_chart(fig_pie, use_container_width=True)

  # --- 5. AGREGAR PRODUCTO ---
  elif opcion == "Agregar Producto":
    st.subheader("➕ Añadir Nuevo Producto")

    with st.form("form_agregar"):
      codigo = st.text_input(
          "Código SKU / QR / Barras",
          value=f"TH-{datetime.now().strftime('%M%S')}",
      )
      nombre = st.text_input("Nombre del Producto")
      categoria = st.selectbox(
          "Categoría",
          ["Cables y Adaptadores", "Audio", "Cargadores", "Fundas", "Otros"],
      )
      col1, col2, col3 = st.columns(3)
      with col1:
        precio = st.number_input(
            "Precio Unitario ($)", min_value=0.0, format="%.2f"
        )
      with col2:
        stock = st.number_input("Stock Inicial", min_value=0, step=1)
      with col3:
        stock_min = st.number_input("Alerta Stock Mínimo", min_value=1, value=5)

      if st.form_submit_button("Guardar en Inventario"):
        if nombre and codigo:
          try:
            agregar_producto(
                codigo, nombre, categoria, precio, stock, stock_min
            )
            st.success(
                f"¡Producto '{nombre}' registrado con código '{codigo}'!"
            )
          except Exception:
            st.error("El código especificado ya existe en la base de datos.")
        else:
          st.error("Por favor completa los campos obligatorios.")

  # --- 6. AJUSTAR STOCK ---
  elif opcion == "Ajustar Stock":
    st.subheader("🔄 Control e Ingreso de Stock")
    df = obtener_productos()

    if df.empty:
      st.info("No hay productos disponibles para ajustar.")
    else:
      prod_nom = st.selectbox("Seleccionar Producto", df["nombre"].tolist())
      prod = df[df["nombre"] == prod_nom].iloc[0]

      st.info(f"**Stock actual en sistema:** {prod['stock']} unidades")

      tipo_ajuste = st.radio(
          "Tipo de Operación", ["Entrada de Mercadería", "Ajuste Manual"]
      )
      cant_cambio = st.number_input("Cantidad", min_value=1, value=1)

      if st.button("Actualizar Stock"):
        if tipo_ajuste == "Entrada de Mercadería":
          nuevo_stock = prod["stock"] + cant_cambio
          tipo_log = "Entrada"
        else:
          nuevo_stock = cant_cambio
          tipo_log = "Ajuste Manual"

        registrar_movimiento(prod["id"], tipo_log, cant_cambio, nuevo_stock)
        st.success("¡Stock actualizado de forma exitosa!")

  # --- 7. HISTORIAL DE MOVIMIENTOS ---
  elif opcion == "Historial de Movimientos":
    st.subheader("📜 Registro Completo de Movimientos")
    df_hist = obtener_historial()

    if df_hist.empty:
      st.info("Aún no hay movimientos registrados.")
    else:
      st.dataframe(df_hist, use_container_width=True, hide_index=True)

  # --- 8. EXPORTAR REPORTES ---
  elif opcion == "Exportar Reportes":
    st.subheader("📥 Exportación de Datos")
    st.write(
        "Descarga una copia completa de la base de datos en formato Excel (.xlsx)"
    )
    excel_data = generar_excel()
    st.download_button(
        label="📊 Descargar Reporte Completo Excel",
        data=excel_data,
        file_name=(
            f"reporte_inventario_{datetime.now().strftime('%Y%m%d')}.xlsx"
        ),
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )
