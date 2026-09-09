from datetime import datetime
import io
import sqlite3
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
        SELECT m.id, p.codigo, p.nombre AS producto, p.categoria, p.precio, m.tipo, m.cantidad, m.fecha_hora 
        FROM movimientos m
        JOIN productos p ON m.producto_id = p.id
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
      (codigo, nombre, categoria, precio, stock, stock_min),
  )

  prod_id = cursor.lastrowid
  fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
  cursor.execute(
      """
        INSERT INTO movimientos (producto_id, tipo, cantidad, fecha_hora)
        VALUES (?, ?, ?, ?)
    """,
      (prod_id, "Registro Inicial", stock, fecha),
  )

  conn.commit()
  conn.close()


def registrar_movimiento(prod_id, tipo, cantidad_cambio, nuevo_stock):
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()
  cursor.execute(
      "UPDATE productos SET stock = ? WHERE id = ?", (nuevo_stock, prod_id)
  )

  fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
  cursor.execute(
      """
        INSERT INTO movimientos (producto_id, tipo, cantidad, fecha_hora)
        VALUES (?, ?, ?, ?)
    """,
      (prod_id, tipo, cantidad_cambio, fecha),
  )

  conn.commit()
  conn.close()


def generar_qr(codigo):
  qr = qrcode.QRCode(box_size=10, border=2)
  qr.add_data(codigo)
  qr.make(fit=True)
  img = qr.make_image(fill_color="black", back_color="white")
  buf = io.BytesIO()
  img.save(buf, format="PNG")
  return buf.getvalue()


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
      Paragraph(
          f"Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
          subtitle_style,
      )
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
    page_title="TecnoHogar - Inventario", layout="wide", page_icon="📦"
)
st.markdown(
    '<meta name="google" content="notranslate">', unsafe_allow_html=True
)

if "logged_in" not in st.session_state:
  st.session_state["logged_in"] = False
  st.session_state["username"] = ""
  st.session_state["rol"] = ""

if not st.session_state["logged_in"]:
  st.title("🔐 Iniciar Sesión - TecnoHogar")

  with st.form("login_form"):
    user_input = st.text_input("Usuario")
    pass_input = st.text_input("Contraseña", type="password")
    submit = st.form_submit_button("Ingresar")

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
      "**Cuentas creadas por defecto:**\n- **Admin:** Usuario: `admin` |"
      " Clave: `admin123`\n- **Vendedor:** Usuario: `vendedor` | Clave:"
      " `vendedor123`"
  )

else:
  st.sidebar.write(
      f"👤 **Usuario:** {st.session_state['username']} | **Rol:**"
      f" {st.session_state['rol']}"
  )
  if st.sidebar.button("Cerrar Sesión"):
    st.session_state["logged_in"] = False
    st.rerun()

  st.title("📦 Sistema de Gestión de Inventario - TecnoHogar")

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

  opcion = st.sidebar.selectbox("Menú de Opciones", opciones)

  # --- VER INVENTARIO ---
  if opcion == "Ver Inventario":
    st.subheader("📋 Lista de Productos")
    df = obtener_productos()

    if df.empty:
      st.info("El inventario está vacío. Agrega un producto desde el menú.")
    else:
      col_f1, col_f2 = st.columns(2)
      with col_f1:
        busqueda = st.text_input("🔍 Buscar por nombre o código del producto")
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
        st.warning(f"⚠️ Hay {len(bajo_stock)} producto(s) con stock crítico.")

      st.dataframe(df_filtrado, use_container_width=True, hide_index=True)

  # --- ESCANEAR CÓDIGO / QR ---
  elif opcion == "Escanear Código / QR":
    st.subheader("📷 Escáner de Código QR o de Barras")
    st.write(
      "Toma una foto al código QR o de barras del producto usando la cámara."
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
        st.success(f"¡Código detectado!: **{data}**")
        df = obtener_productos()
        match = df[df["codigo"] == data]

        if not match.empty:
          prod = match.iloc[0]
          st.write(
              f"**Producto:** {prod['nombre']} | **Precio:** ${prod['precio']} |"
              f" **Stock:** {prod['stock']}"
          )
        else:
          st.warning("No se encontró ningún producto con ese código.")
      else:
        st.error("No se pudo detectar un código válido en la imagen.")

  # --- REGISTRAR VENTA ---
  elif opcion == "Registrar Venta":
    st.subheader("🛒 Registrar Venta")
    df = obtener_productos()

    if df.empty:
      st.info("No hay productos registrados para vender.")
    else:
      prod_nom = st.selectbox("Selecciona Producto", df["nombre"].tolist())
      prod = df[df["nombre"] == prod_nom].iloc[0]
      stock_actual = int(prod["stock"])

      st.write(
          f"**Código:** {prod['codigo']} | **Precio:** ${prod['precio']:.2f} |"
          f" **Stock disponible:** {stock_actual}"
      )

      if stock_actual <= 0:
        st.error(
            "⚠️ Este producto no tiene unidades disponibles en stock para"
            " vender."
        )
      else:
        cliente = st.text_input(
            "Nombre del Cliente (Opcional)", value="Cliente General"
        )
        cant_venta = st.number_input(
            "Cantidad a vender",
            min_value=1,
            max_value=stock_actual,
            value=1,
            step=1,
        )
        total = cant_venta * prod["precio"]
        st.write(f"### **Total a cobrar:** ${total:.2f}")

        if st.button("Confirmar Venta"):
          nuevo_stock = stock_actual - cant_venta
          registrar_movimiento(prod["id"], "Venta", cant_venta, nuevo_stock)
          st.success(
              f"¡Venta registrada! Se descontaron {cant_venta} unidad(es) de"
              f" '{prod['nombre']}'."
          )

          pdf_bytes = generar_pdf_boleta(
              prod["nombre"], cant_venta, prod["precio"], total, cliente
          )

          st.download_button(
              label="📄 Descargar Boleta en PDF",
              data=pdf_bytes,
              file_name=(
                  f"boleta_{prod['nombre']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
              ),
              mime="application/pdf",
          )

  # --- DASHBOARD Y GRÁFICOS ---
  elif opcion == "Dashboard y Gráficos":
    st.subheader("📊 Métricas y Análisis de Ventas")
    df_hist = obtener_historial()
    df_ventas = df_hist[df_hist["tipo"] == "Venta"].copy()

    if df_ventas.empty:
      st.info("Aún no hay registro de ventas para mostrar gráficos.")
    else:
      df_ventas["monto_total"] = df_ventas["cantidad"] * df_ventas["precio"]
      m1, m2, m3 = st.columns(3)
      m1.metric(
          "Total Recaudado ($)", f"${df_ventas['monto_total'].sum():,.2f}"
      )
      m2.metric("Unidades Vendidas", df_ventas["cantidad"].sum())
      m3.metric("Transacciones de Venta", len(df_ventas))

      st.markdown("---")
      col_g1, col_g2 = st.columns(2)

      with col_g1:
        st.markdown("### 🏆 Más Vendidos (Unidades)")
        ventas_por_prod = (
            df_ventas.groupby("producto")["cantidad"].sum().reset_index()
        )
        fig_bar = px.bar(
            ventas_por_prod,
            x="producto",
            y="cantidad",
            color="cantidad",
            color_continuous_scale="Viridis",
        )
        st.plotly_chart(fig_bar, use_container_width=True)

      with col_g2:
        st.markdown("### 🏷️ Ventas por Categoría ($)")
        ventas_por_cat = (
            df_ventas.groupby("categoria")["monto_total"].sum().reset_index()
        )
        fig_pie = px.pie(
            ventas_por_cat, values="monto_total", names="categoria", hole=0.4
        )
        st.plotly_chart(fig_pie, use_container_width=True)

  # --- AGREGAR PRODUCTO ---
  elif opcion == "Agregar Producto":
    st.subheader("➕ Registrar Nuevo Producto")

    with st.form("form_agregar"):
      codigo = st.text_input(
          "Código del Producto (SKU / Barras / QR)",
          value=f"TH-{datetime.now().strftime('%M%S')}",
      )
      nombre = st.text_input("Nombre del Producto")
      categoria = st.selectbox(
          "Categoría",
          ["Cables y Adaptadores", "Audio", "Cargadores", "Fundas", "Otros"],
      )
      col1, col2, col3 = st.columns(3)
      with col1:
        precio = st.number_input("Precio ($)", min_value=0.0, format="%.2f")
      with col2:
        stock = st.number_input("Stock Inicial", min_value=0, step=1)
      with col3:
        stock_min = st.number_input("Stock Mínimo Alerta", min_value=1, value=5)

      if st.form_submit_button("Guardar Producto"):
        if nombre and codigo:
          try:
            agregar_producto(
                codigo, nombre, categoria, precio, stock, stock_min
            )
            st.success(f"¡Producto '{nombre}' agregado con código '{codigo}'!")
          except Exception as e:
            st.error("El código ya existe. Ingresa un código único.")
        else:
          st.error("Completa todos los campos obligatorios.")

  # --- AJUSTAR STOCK ---
  elif opcion == "Ajustar Stock":
    st.subheader("🔄 Modificar Stock Manualmente")
    df = obtener_productos()

    if df.empty:
      st.info("No hay productos registrados.")
    else:
      prod_nom = st.selectbox("Selecciona un producto", df["nombre"].tolist())
      prod = df[df["nombre"] == prod_nom].iloc[0]

      st.write(f"**Stock actual:** {prod['stock']}")
      tipo_ajuste = st.radio(
          "Acción", ["Entrada de Mercadería", "Ajuste Manual"]
      )
      cant_cambio = st.number_input("Cantidad", min_value=1, value=1)

      if st.button("Guardar Cambios"):
        if tipo_ajuste == "Entrada de Mercadería":
          nuevo_stock = prod["stock"] + cant_cambio
          tipo_log = "Entrada"
        else:
          nuevo_stock = cant_cambio
          tipo_log = "Ajuste Manual"

        registrar_movimiento(prod["id"], tipo_log, cant_cambio, nuevo_stock)
        st.success("¡Stock actualizado correctamente!")

  # --- HISTORIAL DE MOVIMIENTOS ---
  elif opcion == "Historial de Movimientos":
    st.subheader("📜 Historial de Entradas, Salidas y Ventas")
    df_hist = obtener_historial()

    if df_hist.empty:
      st.info("Aún no se han registrado movimientos.")
    else:
      st.dataframe(df_hist, use_container_width=True, hide_index=True)

  # --- EXPORTAR REPORTES ---
  elif opcion == "Exportar Reportes":
    st.subheader("📥 Descargar Reportes en Excel")
    excel_data = generar_excel()
    st.download_button(
        label="📥 Descargar Reporte (.xlsx)",
        data=excel_data,
        file_name=(
            f"reporte_inventario_{datetime.now().strftime('%Y%m%d')}.xlsx"
        ),
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )
