"""
Bot de Telegram - Taller Automotriz del Mayo (con agenda de citas y control en Excel)

- Cualquier mensaje que escriba el cliente muestra el menú (no necesita /start).
- Cada servicio tiene precio aproximado y duración (1 a 2 horas).
- El cliente elige servicio, deja sus datos y elige DÍA y HORA con botones.
- Una cita ocupa 1 o 2 horas según el servicio. Solo se muestran horas libres.
- Cada cita confirmada se guarda en citas.db Y se agrega sola al Excel citas.xlsx
  (junto a este script). El Excel trae el precio base de cada servicio; el costo real lo registras con
  /costo NÚMERO MONTO (o escribiéndolo tú en el Excel).

Requisitos:
    pip install python-telegram-bot openpyxl

Antes de ejecutar, pega tu TOKEN y tu ID más abajo.
"""

import logging
import os
import sqlite3
from contextlib import closing
from copy import copy
from datetime import date, datetime, time as hora_t, timedelta, timezone

from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import (
    ApplicationBuilder,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    filters,
)

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
except ImportError:
    raise SystemExit(
        "Falta instalar openpyxl (para el Excel). En la terminal ejecuta:\n"
        "    python -m pip install openpyxl"
    )

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
# Evita que el token aparezca en la terminal
logging.getLogger("httpx").setLevel(logging.WARNING)

# ----------------------------------------------------------------------
# CONFIGURACIÓN (edita estos datos)
# ----------------------------------------------------------------------
TOKEN = "8783603618:AAHmQ5c8ApqtIdWX485hwCg_owTQkB8YMmw"
ADMIN_CHAT_ID = "7900745053"  # tu ID de Telegram, para recibir avisos y usar /citas

NOMBRE_TALLER = "Taller Automotriz del Mayo"
DIRECCION = "Jiménez y Cuarta, col. Tierra Blanca, Navojoa, Sonora"
TELEFONO = "6421167272"

# Horas en que EMPIEZA una cita.
# Lunes a viernes: de 9:00 a 17:00. Sábado: de 9:00 a 13:00. Domingo: cerrado.
# Una cita de 2 horas que empieza a las 5 pm termina a las 7 pm, así que el sistema
# solo ofrece horas en las que el servicio alcanza a terminar antes de cerrar.
HORAS_SEMANA = [f"{h:02d}:00" for h in range(9, 18)]
HORAS_SABADO = [f"{h:02d}:00" for h in range(9, 14)]
DIAS_ADELANTE = 7  # cuántos días hacia adelante se puede agendar
HORARIO = "Lunes a Viernes 9:00 am - 6:00 pm | Sábado 9:00 am - 2:00 pm | Domingo cerrado"

# Navojoa, Sonora: UTC-7 todo el año (no hay horario de verano)
TZ = timezone(timedelta(hours=-7))

NOTA_PRECIOS = (
    "ℹ️ Precios aproximados para autos y pickups. "
    "El costo final se confirma al revisar tu vehículo."
)

# "horas" = cuántas horas se bloquean en la agenda (1 o 2).
# Cambia los precios y tiempos a tu gusto.
SERVICIOS = {
    "aceite": {
        "nombre": "🛢️ Cambio de aceite",
        "descripcion": (
            "Cambio de aceite y filtro de aceite.\n"
            "Revisamos niveles y te recomendamos el aceite adecuado para tu vehículo."
        ),
        "precio": (
            "• Aceite mineral: desde $450\n"
            "• Aceite semisintético: desde $650\n"
            "• Aceite sintético: desde $850\n"
            "(incluye filtro de aceite y mano de obra)"
        ),
        "tiempo": "1 hora",
        "horas": 1,
        "precio_desde": 450,
    },
    "afinacion": {
        "nombre": "🔧 Afinación",
        "descripcion": (
            "Afinación completa: cambio de bujías, filtros y revisión del motor "
            "para mejorar el rendimiento y ahorrar combustible."
        ),
        "precio": (
            "• Autos y pickups de modelo anterior: desde $1,200\n"
            "• Modelos recientes: de $1,800 a $3,000"
        ),
        "tiempo": "2 horas",
        "horas": 2,
        "precio_desde": 1200,
    },
    "frenos": {
        "nombre": "🛑 Revisión de frenos",
        "descripcion": (
            "Revisión de balatas, discos, tambores y líquido de frenos.\n"
            "Tu seguridad es lo primero."
        ),
        "precio": (
            "• Revisión de frenos: desde $250\n"
            "• Cambio de balatas delanteras: desde $900 (mano de obra y balatas)\n"
            "• Discos, rectificado y líquido de frenos: se cotizan aparte"
        ),
        "tiempo": "1 a 2 horas",
        "horas": 2,
        "precio_desde": 250,
    },
    "llantas": {
        "nombre": "🛞 Cambio de llantas",
        "descripcion": (
            "Cambio de llantas, balanceo y alineación.\n"
            "Contamos con varias marcas y medidas."
        ),
        "precio": (
            "• Montaje y balanceo: desde $100 por llanta\n"
            "• Alineación: desde $400 (auto) / $550 (camioneta)\n"
            "• Llantas nuevas: según medida y marca"
        ),
        "tiempo": "1 a 2 horas",
        "horas": 2,
        "precio_desde": 400,
    },
    "diagnostico": {
        "nombre": "🔍 Diagnóstico general",
        "descripcion": (
            "Revisión completa con escáner y checklist de puntos clave "
            "para detectar fallas a tiempo."
        ),
        "precio": "• Diagnóstico con escáner y revisión general: desde $350",
        "tiempo": "1 hora",
        "horas": 1,
        "precio_desde": 350,
    },
    "suspension": {
        "nombre": "🚗 Servicio de suspensión",
        "descripcion": (
            "Revisión y cambio de amortiguadores, rótulas, terminales, "
            "bujes y demás componentes de suspensión y dirección."
        ),
        "precio": (
            "• Revisión de suspensión y dirección: desde $300\n"
            "• Cambio de amortiguadores: mano de obra desde $600 por eje (piezas aparte)\n"
            "• Reparación completa: de $5,000 a $10,000 según vehículo"
        ),
        "tiempo": "2 horas (las reparaciones grandes pueden llevar más; te lo confirmamos)",
        "horas": 2,
        "precio_desde": 300,
    },
}

# Estados de la conversación para agendar cita
NOMBRE, TELEFONO_CLIENTE, VEHICULO, DIA, HORA = range(5)

DIAS_SEMANA = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
MESES = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
]

CARPETA = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(CARPETA, "citas.db")
EXCEL_PATH = os.path.join(CARPETA, "citas.xlsx")


# ----------------------------------------------------------------------
# BASE DE DATOS DE CITAS
# ----------------------------------------------------------------------
def init_db():
    with closing(sqlite3.connect(DB_PATH)) as con:
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS citas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha TEXT NOT NULL,
                hora TEXT NOT NULL,
                duracion INTEGER NOT NULL DEFAULT 1,
                servicio TEXT,
                nombre TEXT,
                telefono TEXT,
                vehiculo TEXT,
                telegram_id INTEGER,
                usuario_tg TEXT,
                creada TEXT,
                excel_ok INTEGER NOT NULL DEFAULT 0
            )
            """
        )
        # Si ya tenías un citas.db de una versión anterior, le agrega las columnas nuevas
        columnas = {fila[1] for fila in con.execute("PRAGMA table_info(citas)")}
        if "duracion" not in columnas:
            con.execute("ALTER TABLE citas ADD COLUMN duracion INTEGER NOT NULL DEFAULT 1")
        if "usuario_tg" not in columnas:
            con.execute("ALTER TABLE citas ADD COLUMN usuario_tg TEXT")
        if "excel_ok" not in columnas:
            con.execute("ALTER TABLE citas ADD COLUMN excel_ok INTEGER NOT NULL DEFAULT 0")
        con.commit()


def horas_que_ocupa(hora: str, duracion: int) -> list:
    inicio = int(hora[:2])
    return [f"{inicio + i:02d}:00" for i in range(duracion)]


def _ocupadas(con, fecha_iso: str) -> set:
    filas = con.execute(
        "SELECT hora, duracion FROM citas WHERE fecha = ?", (fecha_iso,)
    ).fetchall()
    ocupadas = set()
    for hora, duracion in filas:
        ocupadas.update(horas_que_ocupa(hora, duracion or 1))
    return ocupadas


def horas_ocupadas(fecha_iso: str) -> set:
    with closing(sqlite3.connect(DB_PATH)) as con:
        return _ocupadas(con, fecha_iso)


def guardar_cita(
    fecha_iso: str, hora: str, duracion: int, datos: dict, telegram_id: int, usuario_tg: str
) -> bool:
    """Guarda la cita. Devuelve False si esa hora (o parte de ella) ya estaba reservada."""
    necesarias = horas_que_ocupa(hora, duracion)
    con = sqlite3.connect(DB_PATH, isolation_level=None)
    try:
        con.execute("BEGIN IMMEDIATE")
        ocupadas = _ocupadas(con, fecha_iso)
        if any(h in ocupadas for h in necesarias):
            con.execute("ROLLBACK")
            return False
        con.execute(
            """
            INSERT INTO citas
                (fecha, hora, duracion, servicio, nombre, telefono, vehiculo,
                 telegram_id, usuario_tg, creada, excel_ok)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0)
            """,
            (
                fecha_iso,
                hora,
                duracion,
                datos.get("servicio"),
                datos.get("nombre"),
                datos.get("telefono"),
                datos.get("vehiculo"),
                telegram_id,
                usuario_tg,
                ahora().isoformat(timespec="seconds"),
            ),
        )
        con.execute("COMMIT")
        return True
    except Exception:
        try:
            con.execute("ROLLBACK")
        except sqlite3.Error:
            pass
        raise
    finally:
        con.close()


# ----------------------------------------------------------------------
# EXCEL (citas.xlsx)
# ----------------------------------------------------------------------
ENCABEZADOS = [
    "ID", "Fecha", "Día", "Hora", "Duración (h)", "Servicio", "Cliente", "Teléfono",
    "Vehículo", "Usuario Telegram", "Registrada", "Estado",
    "Precio base (desde)", "Costo cobrado", "Notas",
]
ANCHOS = [6, 12, 12, 10, 12, 24, 22, 14, 22, 20, 20, 14, 18, 14, 32]
COL_ESTADO = 12
COL_PRECIO_BASE = 13
COL_COSTO = 14
ESTADOS = "Confirmada,Atendida,Cancelada,No llegó"
FORMATO_DINERO = '"$"#,##0.00'


def sin_emoji(texto: str) -> str:
    """'🛢️ Cambio de aceite' -> 'Cambio de aceite'"""
    texto = texto or ""
    partes = texto.split(" ", 1)
    return partes[1] if len(partes) == 2 else texto


# Precio base ("desde") de cada servicio, buscado por su nombre sin emoji
PRECIO_BASE_POR_NOMBRE = {
    sin_emoji(s["nombre"]): s.get("precio_desde") for s in SERVICIOS.values()
}


def _migrar_excel_viejo(ws) -> None:
    """
    Si el Excel es de la versión anterior (sin la columna 'Precio base'), la inserta
    sin perder nada de lo que ya tenías anotado (costos, notas, estados).
    """
    if ws.cell(row=1, column=COL_PRECIO_BASE).value != "Costo cobrado":
        return
    for fila in range(1, ws.max_row + 1):
        # Mover Notas (col 14 -> 15) y luego Costo cobrado (col 13 -> 14)
        for col in (14, 13):
            origen = ws.cell(row=fila, column=col)
            destino = ws.cell(row=fila, column=col + 1)
            destino.value = origen.value
            destino._style = copy(origen._style)
        nueva = ws.cell(row=fila, column=COL_PRECIO_BASE)
        if fila == 1:
            nueva.value = "Precio base (desde)"
        else:
            id_ = ws.cell(row=fila, column=1).value
            servicio = ws.cell(row=fila, column=6).value
            hay_cita = id_ not in (None, "")
            nueva.value = PRECIO_BASE_POR_NOMBRE.get(servicio) if hay_cita else None
            nueva.number_format = FORMATO_DINERO
    ws.cell(row=1, column=COL_PRECIO_BASE)._style = copy(
        ws.cell(row=1, column=COL_ESTADO)._style
    )
    for col, ancho in enumerate(ANCHOS, start=1):
        ws.column_dimensions[get_column_letter(col)].width = ancho
    ws.auto_filter.ref = f"A1:{get_column_letter(len(ENCABEZADOS))}2000"


def _abrir_excel():
    """Abre citas.xlsx (o lo crea con encabezados y formato si todavía no existe)."""
    if os.path.exists(EXCEL_PATH):
        wb = load_workbook(EXCEL_PATH)
        ws = wb.active
        _migrar_excel_viejo(ws)
        return wb, ws

    wb = Workbook()
    ws = wb.active
    ws.title = "Citas"
    ws.append(ENCABEZADOS)
    for col, ancho in enumerate(ANCHOS, start=1):
        ws.column_dimensions[get_column_letter(col)].width = ancho
        celda = ws.cell(row=1, column=col)
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor="1F4E78")
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(ENCABEZADOS))}2000"
    # Lista desplegable en la columna "Estado"
    validacion = DataValidation(type="list", formula1=f'"{ESTADOS}"', allow_blank=True)
    letra = get_column_letter(COL_ESTADO)
    validacion.add(f"{letra}2:{letra}2000")
    ws.add_data_validation(validacion)
    return wb, ws


def _ultima_fila(ws) -> int:
    """Última fila que realmente tiene datos (ignora filas vacías con formato)."""
    fila = ws.max_row
    while fila > 1 and ws.cell(row=fila, column=1).value in (None, ""):
        fila -= 1
    return fila


def _valores_fila(c) -> list:
    id_, fecha, hora, duracion, servicio, nombre, telefono, vehiculo, usuario_tg, creada = c
    d = date.fromisoformat(fecha)
    try:
        registrada = datetime.fromisoformat(creada).replace(tzinfo=None) if creada else None
    except ValueError:
        registrada = None
    return [
        id_,
        d,
        DIAS_SEMANA[d.weekday()].capitalize(),
        hora_t(int(hora[:2]), 0),
        duracion or 1,
        sin_emoji(servicio),
        nombre or "",
        telefono or "",
        vehiculo or "",
        usuario_tg or "",
        registrada,
        "Confirmada",
        PRECIO_BASE_POR_NOMBRE.get(sin_emoji(servicio)),  # Precio base (desde)
        None,  # Costo cobrado (lo registras con /costo o lo escribes tú)
        None,  # Notas
    ]


def sincronizar_excel() -> bool:
    """
    Agrega al Excel las citas que todavía no están registradas ahí.
    Devuelve False si no se pudo escribir (por ejemplo, si el Excel está abierto);
    en ese caso las citas siguen guardadas y se agregan en el siguiente intento.
    """
    with closing(sqlite3.connect(DB_PATH)) as con:
        pendientes = con.execute(
            """
            SELECT id, fecha, hora, duracion, servicio, nombre, telefono, vehiculo,
                   usuario_tg, creada
            FROM citas WHERE excel_ok = 0 ORDER BY fecha, hora, id
            """
        ).fetchall()
        if not pendientes:
            return True
        try:
            wb, ws = _abrir_excel()
            siguiente = _ultima_fila(ws) + 1
            for c in pendientes:
                for col, valor in enumerate(_valores_fila(c), start=1):
                    ws.cell(row=siguiente, column=col, value=valor)
                ws.cell(row=siguiente, column=2).number_format = "dd/mm/yyyy"
                ws.cell(row=siguiente, column=4).number_format = "h:mm AM/PM"
                ws.cell(row=siguiente, column=11).number_format = "dd/mm/yyyy h:mm AM/PM"
                ws.cell(row=siguiente, column=COL_PRECIO_BASE).number_format = FORMATO_DINERO
                ws.cell(row=siguiente, column=COL_COSTO).number_format = FORMATO_DINERO
                siguiente += 1
            wb.save(EXCEL_PATH)
        except Exception:
            logging.exception("No se pudo escribir en el Excel (¿está abierto?)")
            return False
        con.executemany(
            "UPDATE citas SET excel_ok = 1 WHERE id = ?", [(c[0],) for c in pendientes]
        )
        con.commit()
    return True


def actualizar_excel(cita_id: int, cambios: dict) -> str:
    """
    Cambia celdas de la fila de esa cita. `cambios` = {número_de_columna: valor}.
    Devuelve "ok", "no_encontrada" o "error" (por ejemplo, si el Excel está abierto).
    """
    if not os.path.exists(EXCEL_PATH):
        return "no_encontrada"
    try:
        wb, ws = _abrir_excel()
        for fila in range(2, ws.max_row + 1):
            if ws.cell(row=fila, column=1).value == cita_id:
                for col, valor in cambios.items():
                    celda = ws.cell(row=fila, column=col)
                    celda.value = valor
                    if col == COL_COSTO:
                        celda.number_format = FORMATO_DINERO
                wb.save(EXCEL_PATH)
                return "ok"
        return "no_encontrada"
    except Exception:
        logging.exception("No se pudo actualizar el Excel (¿está abierto?)")
        return "error"


def marcar_cancelada_excel(cita_id: int) -> bool:
    """Pone 'Cancelada' en la columna Estado de esa cita. False si no se pudo guardar."""
    return actualizar_excel(cita_id, {COL_ESTADO: "Cancelada"}) != "error"


# ----------------------------------------------------------------------
# HORARIOS Y FORMATO
# ----------------------------------------------------------------------
def ahora() -> datetime:
    return datetime.now(TZ)


def horas_del_dia(d: date) -> list:
    if d.weekday() == 6:  # domingo
        return []
    if d.weekday() == 5:  # sábado
        return HORAS_SABADO
    return HORAS_SEMANA


def horas_disponibles(d: date, duracion: int = 1) -> list:
    """Horas de inicio en las que cabe un servicio de `duracion` horas sin chocar con otra cita."""
    todas = horas_del_dia(d)
    ocupadas = horas_ocupadas(d.isoformat())
    ahora_ = ahora()
    minutos_ahora = ahora_.hour * 60 + ahora_.minute
    libres = []
    for h in todas:
        necesarias = horas_que_ocupa(h, duracion)
        # Todas las horas del servicio deben estar dentro del horario y libres
        if any(x not in todas or x in ocupadas for x in necesarias):
            continue
        # Si es hoy, no ofrecer horas que ya pasaron (o faltan menos de 30 min)
        if d == ahora_.date() and int(h[:2]) * 60 <= minutos_ahora + 30:
            continue
        libres.append(h)
    return libres


def hora_bonita(h: str) -> str:
    h24 = int(h[:2])
    sufijo = "am" if h24 < 12 else "pm"
    return f"{h24 % 12 or 12}:00 {sufijo}"


def fecha_bonita(d: date) -> str:
    return f"{DIAS_SEMANA[d.weekday()]} {d.day} de {MESES[d.month - 1]}"


def texto_servicio(s: dict) -> str:
    return (
        f"{s['nombre']}\n\n"
        f"{s['descripcion']}\n\n"
        f"💲 Precios aproximados:\n{s['precio']}\n\n"
        f"⏱️ Tiempo aproximado: {s['tiempo']}\n\n"
        f"{NOTA_PRECIOS}"
    )


# ----------------------------------------------------------------------
# TECLADOS
# ----------------------------------------------------------------------
def teclado_menu() -> InlineKeyboardMarkup:
    botones = [
        [InlineKeyboardButton(s["nombre"], callback_data=f"servicio:{clave}")]
        for clave, s in SERVICIOS.items()
    ]
    botones.append([InlineKeyboardButton("📍 Contacto y horario", callback_data="contacto")])
    return InlineKeyboardMarkup(botones)


def texto_menu() -> str:
    return (
        f"👋 ¡Bienvenido a {NOMBRE_TALLER}!\n\n"
        "Elige el servicio que te interesa:"
    )


def teclado_dias(duracion: int = 1):
    """Botones con los próximos días que tienen horas libres. None si no hay ninguno."""
    hoy = ahora().date()
    botones = []
    for i in range(DIAS_ADELANTE + 1):
        d = hoy + timedelta(days=i)
        libres = horas_disponibles(d, duracion)
        if not libres:
            continue
        etiqueta = f"{fecha_bonita(d).capitalize()} ({len(libres)} libres)"
        botones.append([InlineKeyboardButton(etiqueta, callback_data=f"dia:{d.isoformat()}")])
    if not botones:
        return None
    botones.append([InlineKeyboardButton("❌ Cancelar", callback_data="cancelar_cita")])
    return InlineKeyboardMarkup(botones)


def teclado_horas(libres: list) -> InlineKeyboardMarkup:
    filas = []
    fila = []
    for h in libres:
        fila.append(InlineKeyboardButton(hora_bonita(h), callback_data=f"hora:{h}"))
        if len(fila) == 3:
            filas.append(fila)
            fila = []
    if fila:
        filas.append(fila)
    filas.append([InlineKeyboardButton("⬅️ Elegir otro día", callback_data="otro_dia")])
    filas.append([InlineKeyboardButton("❌ Cancelar", callback_data="cancelar_cita")])
    return InlineKeyboardMarkup(filas)


# ----------------------------------------------------------------------
# MENÚ
# ----------------------------------------------------------------------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text(texto_menu(), reply_markup=teclado_menu())
    return ConversationHandler.END


async def mensaje_libre(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Cualquier mensaje (hola, buenas, un sticker, etc.) muestra el menú, como /start."""
    await update.message.reply_text(texto_menu(), reply_markup=teclado_menu())


async def volver_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    await query.edit_message_text(texto_menu(), reply_markup=teclado_menu())


async def mostrar_servicio(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    clave = query.data.split(":")[1]
    servicio = SERVICIOS.get(clave)
    if not servicio:
        return
    teclado = InlineKeyboardMarkup(
        [
            [InlineKeyboardButton("📅 Agendar cita", callback_data=f"agendar:{clave}")],
            [InlineKeyboardButton("⬅️ Volver al menú", callback_data="menu")],
        ]
    )
    await query.edit_message_text(texto_servicio(servicio), reply_markup=teclado)


async def mostrar_contacto(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    teclado = InlineKeyboardMarkup(
        [[InlineKeyboardButton("⬅️ Volver al menú", callback_data="menu")]]
    )
    await query.edit_message_text(
        f"📍 {NOMBRE_TALLER}\n\n"
        f"Dirección: {DIRECCION}\n"
        f"Teléfono: {TELEFONO}\n"
        f"Horario: {HORARIO}",
        reply_markup=teclado,
    )


# ----------------------------------------------------------------------
# FLUJO PARA AGENDAR CITA
# ----------------------------------------------------------------------
async def iniciar_cita(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    clave = query.data.split(":")[1]
    servicio = SERVICIOS[clave]
    context.user_data.clear()
    context.user_data["servicio"] = servicio["nombre"]
    context.user_data["duracion"] = servicio["horas"]
    context.user_data["tiempo"] = servicio["tiempo"]
    await query.message.reply_text(
        f"Perfecto, vamos a agendar: {servicio['nombre']}\n\n"
        "¿Cuál es tu nombre?\n\n(Escribe /cancelar para salir)"
    )
    return NOMBRE


async def recibir_nombre(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["nombre"] = update.message.text.strip()
    await update.message.reply_text("📱 ¿Cuál es tu número de teléfono?")
    return TELEFONO_CLIENTE


async def recibir_telefono(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["telefono"] = update.message.text.strip()
    await update.message.reply_text(
        "🚗 ¿Qué vehículo tienes? (marca, modelo y año)\nEjemplo: Nissan Versa 2018"
    )
    return VEHICULO


async def recibir_vehiculo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    datos = context.user_data
    datos["vehiculo"] = update.message.text.strip()
    teclado = teclado_dias(datos.get("duracion", 1))
    if teclado is None:
        await update.message.reply_text(
            "😕 Por el momento no hay horarios libres en los próximos días.\n"
            f"Llámanos al {TELEFONO} para agendar tu cita."
        )
        datos.clear()
        return ConversationHandler.END
    await update.message.reply_text(
        f"📅 ¿Qué día quieres venir?\n\n⏱️ Este servicio dura aprox.: {datos.get('tiempo', '1 hora')}",
        reply_markup=teclado,
    )
    return DIA


async def elegir_dia(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    datos = context.user_data
    duracion = datos.get("duracion", 1)
    d = date.fromisoformat(query.data.split(":", 1)[1])
    libres = horas_disponibles(d, duracion)

    if not libres:
        teclado = teclado_dias(duracion)
        if teclado is None:
            await query.edit_message_text(
                f"😕 Ya no hay horarios libres. Llámanos al {TELEFONO}."
            )
            datos.clear()
            return ConversationHandler.END
        await query.edit_message_text(
            f"😕 {fecha_bonita(d).capitalize()} ya no tiene horarios libres.\n\nElige otro día:",
            reply_markup=teclado,
        )
        return DIA

    datos["fecha"] = d.isoformat()
    await query.edit_message_text(
        f"🕐 {fecha_bonita(d).capitalize()}\n\nEstas son las horas disponibles. Elige una:",
        reply_markup=teclado_horas(libres),
    )
    return HORA


async def cambiar_dia(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    datos = context.user_data
    teclado = teclado_dias(datos.get("duracion", 1))
    if teclado is None:
        await query.edit_message_text(f"😕 Ya no hay horarios libres. Llámanos al {TELEFONO}.")
        datos.clear()
        return ConversationHandler.END
    await query.edit_message_text("📅 ¿Qué día quieres venir?", reply_markup=teclado)
    return DIA


async def elegir_hora(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    datos = context.user_data

    requeridos = ("servicio", "nombre", "telefono", "vehiculo", "fecha")
    if not all(k in datos for k in requeridos):
        await query.edit_message_text("Esta solicitud expiró. Escribe cualquier mensaje para empezar de nuevo.")
        datos.clear()
        return ConversationHandler.END

    hora = query.data.split(":", 1)[1]
    duracion = datos.get("duracion", 1)
    d = date.fromisoformat(datos["fecha"])
    usuario = update.effective_user
    usuario_tg = f"@{usuario.username}" if usuario.username else "sin username"

    # Verifica que siga libre y la guarda (la base de datos revisa que no se empalme)
    if hora not in horas_disponibles(d, duracion) or not guardar_cita(
        d.isoformat(), hora, duracion, datos, usuario.id, usuario_tg
    ):
        libres = horas_disponibles(d, duracion)
        if libres:
            await query.edit_message_text(
                f"⚠️ Lo sentimos, las {hora_bonita(hora)} del {fecha_bonita(d)} "
                "ya están reservadas.\n\n"
                "Estas horas sí están libres ese día:",
                reply_markup=teclado_horas(libres),
            )
            return HORA
        teclado = teclado_dias(duracion)
        if teclado is None:
            await query.edit_message_text(
                "⚠️ Esa hora ya está reservada y no quedan más horarios libres.\n"
                f"Llámanos al {TELEFONO}."
            )
            datos.clear()
            return ConversationHandler.END
        await query.edit_message_text(
            f"⚠️ Las {hora_bonita(hora)} ya están reservadas y ese día ya no tiene "
            "más horarios.\n\nElige otro día:",
            reply_markup=teclado,
        )
        return DIA

    resumen = (
        "📋 Cita agendada\n\n"
        f"Servicio: {datos['servicio']}\n"
        f"Nombre: {datos['nombre']}\n"
        f"Teléfono: {datos['telefono']}\n"
        f"Vehículo: {datos['vehiculo']}\n"
        f"Día: {fecha_bonita(d).capitalize()}\n"
        f"Hora: {hora_bonita(hora)} (duración aprox.: {datos.get('tiempo', '1 hora')})"
    )
    await query.edit_message_text(
        f"{resumen}\n\n✅ ¡Listo! Tu cita quedó reservada.\n"
        f"📍 {DIRECCION}\n\n"
        "Escribe cualquier mensaje para volver al menú."
    )

    # Registrar la cita confirmada en el Excel
    excel_ok = sincronizar_excel()

    # Avisar al dueño/encargado del taller
    if ADMIN_CHAT_ID:
        aviso = f"{resumen}\n\nUsuario de Telegram: {usuario_tg} (id {usuario.id})"
        base = PRECIO_BASE_POR_NOMBRE.get(sin_emoji(datos["servicio"]))
        if base:
            aviso += f"\n💲 Precio base: desde ${base:,}"
        if excel_ok:
            aviso += "\n📊 Registrada en Excel"
        else:
            aviso += (
                "\n⚠️ No pude escribir en el Excel (¿está abierto?). La cita sí quedó "
                "guardada y se agregará sola en la próxima cita o con /excel."
            )
        try:
            await context.bot.send_message(chat_id=ADMIN_CHAT_ID, text=aviso)
        except Exception:
            logging.exception("No se pudo enviar el aviso al administrador")

    datos.clear()
    return ConversationHandler.END


async def usa_botones(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Por favor elige una opción con los botones de arriba 👆\n(o escribe /cancelar para salir)"
    )


async def cancelar(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text(
        "Cita cancelada. Escribe cualquier mensaje para ver el menú."
    )
    return ConversationHandler.END


async def cancelar_boton(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    context.user_data.clear()
    await query.edit_message_text("Cita cancelada. Escribe cualquier mensaje para ver el menú.")
    return ConversationHandler.END


async def boton_expirado(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Botones viejos de una conversación que ya terminó o que se perdió al reiniciar el bot."""
    query = update.callback_query
    await query.answer()
    await query.edit_message_text(
        "Este botón ya expiró. Escribe cualquier mensaje para volver al menú."
    )


# ----------------------------------------------------------------------
# COMANDOS SOLO PARA EL ADMINISTRADOR
# ----------------------------------------------------------------------
def es_admin(update: Update) -> bool:
    return bool(ADMIN_CHAT_ID) and str(update.effective_user.id) == str(ADMIN_CHAT_ID)


async def ver_citas(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/citas -> lista las citas de hoy en adelante."""
    if not es_admin(update):
        return
    hoy = ahora().date().isoformat()
    with closing(sqlite3.connect(DB_PATH)) as con:
        filas = con.execute(
            """
            SELECT id, fecha, hora, duracion, servicio, nombre, telefono, vehiculo
            FROM citas WHERE fecha >= ? ORDER BY fecha, hora
            """,
            (hoy,),
        ).fetchall()
    if not filas:
        await update.message.reply_text("No hay citas próximas.")
        return
    lineas = []
    for id_, fecha, hora, duracion, servicio, nombre, tel, veh in filas:
        d = date.fromisoformat(fecha)
        lineas.append(
            f"#{id_} · {fecha_bonita(d)} {hora_bonita(hora)} ({duracion} h)\n"
            f"   {servicio} — {nombre} ({tel}) — {veh}"
        )
    texto = "📅 Próximas citas:\n\n" + "\n\n".join(lineas)
    await update.message.reply_text(texto[:4000])


async def enviar_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/excel -> te manda el archivo citas.xlsx por Telegram (ya actualizado)."""
    if not es_admin(update):
        return
    if not sincronizar_excel():
        await update.message.reply_text(
            "⚠️ No pude actualizar el Excel porque está abierto en la computadora. "
            "Ciérralo y vuelve a escribir /excel."
        )
        return
    if not os.path.exists(EXCEL_PATH):
        await update.message.reply_text("Todavía no hay citas registradas.")
        return
    try:
        with open(EXCEL_PATH, "rb") as archivo:
            await update.message.reply_document(document=archivo, filename="citas_taller.xlsx")
    except Exception:
        logging.exception("No se pudo enviar el Excel")
        await update.message.reply_text(
            "No pude enviar el archivo. Si lo tienes abierto en Excel, ciérralo e inténtalo otra vez."
        )


async def registrar_costo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/costo 3 650 -> anota que la cita #3 costó $650 y la marca como Atendida en el Excel."""
    if not es_admin(update):
        return
    args = context.args
    if len(args) != 2 or not args[0].isdigit():
        await update.message.reply_text(
            "Uso: /costo NÚMERO MONTO\nEjemplo: /costo 3 650\n"
            "El número de la cita aparece en /citas (por ejemplo #3) y en la columna ID del Excel."
        )
        return
    try:
        monto = float(args[1].replace("$", "").replace(",", ""))
    except ValueError:
        await update.message.reply_text("El monto debe ser un número. Ejemplo: /costo 3 650")
        return
    monto = int(monto) if monto.is_integer() else monto
    cita_id = int(args[0])

    with closing(sqlite3.connect(DB_PATH)) as con:
        existe = con.execute("SELECT 1 FROM citas WHERE id = ?", (cita_id,)).fetchone()
    if not existe:
        await update.message.reply_text("No encontré una cita con ese número.")
        return
    if not sincronizar_excel():
        await update.message.reply_text(
            "⚠️ No pude escribir en el Excel porque está abierto. Ciérralo y vuelve a intentar."
        )
        return
    resultado = actualizar_excel(cita_id, {COL_COSTO: monto, COL_ESTADO: "Atendida"})
    if resultado == "ok":
        await update.message.reply_text(
            f"✅ Cita #{cita_id}: costo cobrado ${monto:,} registrado y marcada como Atendida."
        )
    elif resultado == "no_encontrada":
        await update.message.reply_text("No encontré esa cita en el Excel.")
    else:
        await update.message.reply_text(
            "⚠️ No pude escribir en el Excel (¿está abierto?). Ciérralo y vuelve a intentar."
        )


async def borrar_cita(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/borrar 12 -> elimina la cita #12, libera sus horas y la marca 'Cancelada' en el Excel."""
    if not es_admin(update):
        return
    if not context.args or not context.args[0].isdigit():
        await update.message.reply_text(
            "Uso: /borrar NÚMERO\nEl número de cada cita aparece en /citas (por ejemplo #12)."
        )
        return
    cita_id = int(context.args[0])
    with closing(sqlite3.connect(DB_PATH)) as con:
        cur = con.execute("DELETE FROM citas WHERE id = ?", (cita_id,))
        con.commit()
        borradas = cur.rowcount
    if not borradas:
        await update.message.reply_text("No encontré una cita con ese número.")
        return
    if marcar_cancelada_excel(cita_id):
        await update.message.reply_text(
            "✅ Cita eliminada. Esas horas ya quedaron libres y en el Excel quedó como Cancelada."
        )
    else:
        await update.message.reply_text(
            "✅ Cita eliminada y sus horas quedaron libres.\n"
            f"⚠️ No pude marcarla en el Excel (¿está abierto?). Cambia el Estado de la cita #{cita_id} "
            "a Cancelada manualmente."
        )


# ----------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------
def main():
    if "PEGA_AQUI" in TOKEN:
        raise SystemExit("Falta pegar tu token en la línea TOKEN = \"...\" al inicio del archivo.")

    init_db()
    # Pone al día el Excel con cualquier cita que aún no esté ahí
    if not sincronizar_excel():
        print("Aviso: no pude actualizar citas.xlsx (¿está abierto?). Se actualizará más tarde.")

    app = ApplicationBuilder().token(TOKEN).build()

    texto = filters.TEXT & ~filters.COMMAND

    conversacion_cita = ConversationHandler(
        entry_points=[CallbackQueryHandler(iniciar_cita, pattern=r"^agendar:")],
        states={
            NOMBRE: [MessageHandler(texto, recibir_nombre)],
            TELEFONO_CLIENTE: [MessageHandler(texto, recibir_telefono)],
            VEHICULO: [MessageHandler(texto, recibir_vehiculo)],
            DIA: [
                CallbackQueryHandler(elegir_dia, pattern=r"^dia:"),
                MessageHandler(texto, usa_botones),
            ],
            HORA: [
                CallbackQueryHandler(elegir_hora, pattern=r"^hora:"),
                CallbackQueryHandler(cambiar_dia, pattern=r"^otro_dia$"),
                MessageHandler(texto, usa_botones),
            ],
        },
        fallbacks=[
            CommandHandler("cancelar", cancelar),
            CommandHandler("start", start),
            CallbackQueryHandler(cancelar_boton, pattern=r"^cancelar_cita$"),
        ],
        allow_reentry=True,
    )

    # La conversación va primero para que /start y /cancelar funcionen bien dentro de ella
    app.add_handler(conversacion_cita)
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("menu", start))
    app.add_handler(CommandHandler("citas", ver_citas))
    app.add_handler(CommandHandler("excel", enviar_excel))
    app.add_handler(CommandHandler("costo", registrar_costo))
    app.add_handler(CommandHandler("borrar", borrar_cita))
    app.add_handler(CallbackQueryHandler(mostrar_servicio, pattern=r"^servicio:"))
    app.add_handler(CallbackQueryHandler(mostrar_contacto, pattern=r"^contacto$"))
    app.add_handler(CallbackQueryHandler(volver_menu, pattern=r"^menu$"))
    app.add_handler(
        CallbackQueryHandler(boton_expirado, pattern=r"^(dia:|hora:|otro_dia$|cancelar_cita$)")
    )
    # Cualquier otro mensaje (texto, sticker, foto, comando desconocido) muestra el menú
    app.add_handler(MessageHandler(filters.ChatType.PRIVATE, mensaje_libre))

    print("Bot del taller en marcha... (Ctrl+C para detener)")
    print(f"Excel de citas: {EXCEL_PATH}")
    app.run_polling()


