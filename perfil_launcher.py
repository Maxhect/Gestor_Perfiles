"""
Asistente de Perfiles - Lanzador de aplicaciones y archivos
=============================================================
Permite crear perfiles (ej: "Trabajo", "Diseño", "Streaming") que agrupan
programas y archivos. Con un clic, se abren todos los elementos de un perfil.

Funciones extra:
  - Elegir programas desde una lista de "instalados" (menú Inicio + registro).
  - Asignar un icono (.ico/.png/.jpg) a cada perfil (se ajusta automáticamente
    a cada tamaño sin deformarse).
  - Hacer que un perfil se inicie automáticamente al encender el equipo.
  - Mostrar un icono flotante en el escritorio (sin fondo, solo el icono)
    que al hacer clic lanza el perfil.
  - Mostrar un icono en la barra de tareas (área de notificación) por cada
    perfil, con clic para lanzarlo.

Requisitos: Python 3.8+ (tkinter incluido en Windows). Las dependencias
opcionales (pillow, pystray, pywin32) se instalan solas automáticamente la
primera vez que ejecutas la aplicación (necesitas conexión a internet ese
primer arranque). No hace falta instalarlas manualmente.

Ejecutar con:       python perfil_launcher.py
Sin ventana de consola: usa pythonw.exe en vez de python.exe

Los datos se guardan en:  %APPDATA%\\PerfilLauncher\\perfiles.json
"""

import argparse
import datetime
import importlib
import json
import os
import subprocess
import sys
import threading
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

try:
    import winreg  # solo existe en Windows, es parte de la librería estándar
except ImportError:
    winreg = None


# ---------------------------------------------------------------------------
# Dependencias opcionales: se instalan solas automáticamente la primera vez
# ---------------------------------------------------------------------------

TIENE_PIL = False
TIENE_PYWIN32 = False
TIENE_TRAY = False
Image = ImageTk = ImageDraw = None
pystray = None
win32com = None

_MODULOS_OPCIONALES = [
    ("PIL", "pillow"),
    ("pystray", "pystray"),
]
if sys.platform == "win32":
    _MODULOS_OPCIONALES.append(("win32com.client", "pywin32"))


def _dependencias_faltantes():
    """Devuelve los nombres de pip (ej. 'pillow') que todavía no están instalados."""
    faltantes = []
    for modulo, pip_nombre in _MODULOS_OPCIONALES:
        try:
            importlib.import_module(modulo)
        except ImportError:
            if pip_nombre not in faltantes:
                faltantes.append(pip_nombre)
    return faltantes


def _pip_instalar(pip_nombre):
    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "--quiet", "--disable-pip-version-check", pip_nombre]
        )
        return True
    except Exception:
        return False


def _cargar_dependencias_opcionales():
    """Importa (o vuelve a intentar importar tras instalar) las dependencias
    opcionales y actualiza las banderas globales TIENE_*."""
    global TIENE_PIL, TIENE_PYWIN32, TIENE_TRAY, Image, ImageTk, ImageDraw, pystray, win32com

    importlib.invalidate_caches()

    try:
        from PIL import Image as _Image, ImageTk as _ImageTk, ImageDraw as _ImageDraw
        Image, ImageTk, ImageDraw = _Image, _ImageTk, _ImageDraw
        TIENE_PIL = True
    except ImportError:
        TIENE_PIL = False

    try:
        import pystray as _pystray
        pystray = _pystray
        TIENE_TRAY = TIENE_PIL  # pystray necesita Pillow para construir la imagen
    except ImportError:
        TIENE_TRAY = False

    if sys.platform == "win32":
        try:
            import win32com.client as _win32com_client
            win32com = _win32com_client
            TIENE_PYWIN32 = True
        except ImportError:
            TIENE_PYWIN32 = False


def _instalar_dependencias_con_progreso(faltantes):
    """Ventana simple con barra de progreso mientras se instalan (con pip)
    las dependencias opcionales, la primera vez que hacen falta."""
    splash = tk.Tk()
    splash.title("Preparando el asistente")
    splash.geometry("380x140")
    splash.resizable(False, False)
    splash.attributes("-topmost", True)

    tk.Label(
        splash, text="Instalando componentes necesarios (solo la primera vez)...",
        font=("Segoe UI", 10), wraplength=340, justify="center",
    ).pack(pady=(18, 6))
    var_estado = tk.StringVar(value="Preparando...")
    tk.Label(splash, textvariable=var_estado, font=("Segoe UI", 9), fg="#555").pack()
    barra = ttk.Progressbar(splash, mode="determinate", maximum=max(len(faltantes), 1), length=300)
    barra.pack(pady=12)
    splash.update_idletasks()

    for i, pip_nombre in enumerate(faltantes, start=1):
        var_estado.set(f"Instalando {pip_nombre} ({i}/{len(faltantes)})...")
        splash.update()
        _pip_instalar(pip_nombre)
        barra["value"] = i
        splash.update()

    splash.destroy()
    _cargar_dependencias_opcionales()


# Primer intento (por si ya estaban instaladas de antes); si faltan, se
# instalan automáticamente al arrancar la interfaz gráfica en main().
_cargar_dependencias_opcionales()


# ---------------------------------------------------------------------------
# Almacenamiento
# ---------------------------------------------------------------------------

APPDATA = os.environ.get("APPDATA", os.path.expanduser("~"))
CONFIG_DIR = os.path.join(APPDATA, "PerfilLauncher")
CONFIG_FILE = os.path.join(CONFIG_DIR, "perfiles.json")

DIAS_SEMANA = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
DIAS_ABREV = ["L", "M", "X", "J", "V", "S", "D"]

def perfil_vacio():
    """Nuevo diccionario de perfil con valores por defecto (listas independientes)."""
    return {"items": [], "icono": None, "auto_inicio": False, "dias_activos": list(range(7))}


def dia_permitido(perfil):
    """True si hoy es uno de los días configurados para este perfil (todos por defecto)."""
    dias = perfil.get("dias_activos")
    if dias is None:
        return True
    return datetime.date.today().weekday() in dias


def cargar_perfiles():
    if not os.path.exists(CONFIG_FILE):
        return {}
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            datos = json.load(f)
    except Exception:
        return {}

    # Compatibilidad con el formato anterior (lista simple de items)
    normalizado = {}
    for nombre, valor in datos.items():
        if isinstance(valor, list):
            normalizado[nombre] = {"items": valor, "icono": None, "auto_inicio": False}
        else:
            perfil = perfil_vacio()
            perfil.update(valor)
            normalizado[nombre] = perfil
    return normalizado


def guardar_perfiles(perfiles):
    os.makedirs(CONFIG_DIR, exist_ok=True)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(perfiles, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# Anti-doble-clic: no relanzar un perfil si ya se inició hace poco
# ---------------------------------------------------------------------------

ESTADO_FILE = os.path.join(CONFIG_DIR, "estado.json")
COOLDOWN_SEGUNDOS = 120  # 2 minutos


def _cargar_estado():
    """Guarda cuándo se lanzó cada perfil por última vez (persiste entre
    ejecuciones, así el botón, el icono flotante, la bandeja y los accesos
    directos de Windows comparten el mismo control)."""
    if not os.path.exists(ESTADO_FILE):
        return {}
    try:
        with open(ESTADO_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _guardar_estado(estado):
    os.makedirs(CONFIG_DIR, exist_ok=True)
    try:
        with open(ESTADO_FILE, "w", encoding="utf-8") as f:
            json.dump(estado, f)
    except Exception:
        pass


def segundos_desde_ultimo_lanzamiento(nombre_perfil):
    """None si nunca se ha lanzado; si no, segundos transcurridos desde la última vez."""
    ultimo = _cargar_estado().get(nombre_perfil)
    if ultimo is None:
        return None
    return time.time() - ultimo


def registrar_lanzamiento(nombre_perfil):
    estado = _cargar_estado()
    estado[nombre_perfil] = time.time()
    _guardar_estado(estado)


# ---------------------------------------------------------------------------
# Lanzamiento de elementos
# ---------------------------------------------------------------------------

def abrir_item(ruta, argumentos=""):
    """Abre un archivo o programa usando la asociación predeterminada de Windows."""
    try:
        if argumentos.strip():
            subprocess.Popen(f'"{ruta}" {argumentos}', shell=True)
        else:
            os.startfile(ruta)  # noqa: usa la app/asociación predeterminada de Windows
    except Exception as e:
        messagebox.showerror("Error al abrir", f"No se pudo abrir:\n{ruta}\n\n{e}")


def lanzar_items(items):
    for item in items:
        abrir_item(item["ruta"], item.get("args", ""))


def lanzar_perfil_con_cooldown(nombre_perfil, perfil, cooldown=COOLDOWN_SEGUNDOS, forzar=False):
    """Lanza los items del perfil, salvo que:
      - hoy no sea uno de sus 'días activos' configurados, o
      - ya se haya lanzado hace menos de `cooldown` segundos
        (esto último evita aperturas duplicadas por doble clic, el icono
        flotante, la bandeja del sistema y los accesos directos de Windows a
        la vez).
    Devuelve 'ok', 'dia_no_permitido' o 'cooldown'."""
    if not forzar and not dia_permitido(perfil):
        return "dia_no_permitido"
    transcurrido = segundos_desde_ultimo_lanzamiento(nombre_perfil)
    if not forzar and transcurrido is not None and transcurrido < cooldown:
        return "cooldown"
    lanzar_items(perfil.get("items", []))
    registrar_lanzamiento(nombre_perfil)
    return "ok"


# ---------------------------------------------------------------------------
# Detección de programas instalados
# ---------------------------------------------------------------------------

def _listar_accesos_directos():
    """Busca accesos directos (.lnk/.url) en el Menú Inicio, común y del usuario."""
    carpetas = [
        os.path.join(os.environ.get("ProgramData", ""), r"Microsoft\Windows\Start Menu\Programs"),
        os.path.join(os.environ.get("APPDATA", ""), r"Microsoft\Windows\Start Menu\Programs"),
    ]
    resultados = []
    for carpeta in carpetas:
        if not carpeta or not os.path.isdir(carpeta):
            continue
        for root, _dirs, files in os.walk(carpeta):
            for f in files:
                if f.lower().endswith((".lnk", ".url")):
                    nombre = os.path.splitext(f)[0]
                    resultados.append((nombre, os.path.join(root, f)))
    return resultados


def _leer_uninstall_key(hive, subkey, flags=0):
    """Lee entradas de 'Agregar o quitar programas' desde el registro de Windows."""
    resultados = []
    try:
        with winreg.OpenKey(hive, subkey, 0, winreg.KEY_READ | flags) as key:
            total = winreg.QueryInfoKey(key)[0]
            for i in range(total):
                try:
                    sub = winreg.EnumKey(key, i)
                    with winreg.OpenKey(key, sub) as subk:
                        try:
                            nombre = winreg.QueryValueEx(subk, "DisplayName")[0]
                        except FileNotFoundError:
                            continue
                        if not nombre:
                            continue
                        ruta = None
                        try:
                            icono = winreg.QueryValueEx(subk, "DisplayIcon")[0]
                            posible = icono.split(",")[0].strip('"')
                            if posible.lower().endswith(".exe") and os.path.exists(posible):
                                ruta = posible
                        except FileNotFoundError:
                            pass
                        if ruta:
                            resultados.append((nombre, ruta))
                except OSError:
                    continue
    except FileNotFoundError:
        pass
    return resultados


def listar_programas_instalados():
    """Combina accesos directos del menú Inicio + registro de Windows en una sola lista."""
    encontrados = {}

    for nombre, ruta in _listar_accesos_directos():
        encontrados.setdefault(nombre.strip().lower(), (nombre, ruta))

    if winreg is not None:
        claves = [
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", 0),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall", 0),
            (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", 0),
        ]
        for hive, subkey, flags in claves:
            for nombre, ruta in _leer_uninstall_key(hive, subkey, flags):
                clave = nombre.strip().lower()
                if clave not in encontrados:
                    encontrados[clave] = (nombre, ruta)

    return sorted(encontrados.values(), key=lambda x: x[0].lower())


# ---------------------------------------------------------------------------
# Inicio automático con Windows
# ---------------------------------------------------------------------------

def _ejecutable_y_script():
    """Devuelve (pythonw.exe, ruta_absoluta_del_script) para construir accesos directos."""
    python_exe = sys.executable
    pythonw = python_exe
    posible = python_exe[:-len("python.exe")] + "pythonw.exe" if python_exe.lower().endswith("python.exe") else None
    if posible and os.path.exists(posible):
        pythonw = posible
    script = os.path.abspath(sys.argv[0])
    return pythonw, script


def _carpeta_inicio_automatico():
    return os.path.join(APPDATA, r"Microsoft\Windows\Start Menu\Programs\Startup")


def crear_acceso_directo(ruta_destino_lnk, nombre_perfil, icono=None):
    """Crea un acceso directo que ejecuta `perfil_launcher.py --lanzar <perfil>`."""
    pythonw, script = _ejecutable_y_script()
    argumentos = f'"{script}" --lanzar "{nombre_perfil}"'

    if TIENE_PYWIN32:
        shell = win32com.Dispatch("WScript.Shell")
        acceso = shell.CreateShortCut(ruta_destino_lnk)
        acceso.TargetPath = pythonw
        acceso.Arguments = argumentos
        acceso.WorkingDirectory = os.path.dirname(script)
        if icono and os.path.exists(icono):
            acceso.IconLocation = icono
        acceso.save()
        return ruta_destino_lnk
    else:
        # Alternativa sin pywin32: un .bat (funciona igual, sin icono personalizado)
        ruta_bat = os.path.splitext(ruta_destino_lnk)[0] + ".bat"
        with open(ruta_bat, "w", encoding="utf-8") as f:
            f.write(f'@echo off\r\nstart "" "{pythonw}" {argumentos}\r\n')
        return ruta_bat


def activar_inicio_automatico(nombre_perfil, icono=None):
    carpeta = _carpeta_inicio_automatico()
    os.makedirs(carpeta, exist_ok=True)
    destino = os.path.join(carpeta, f"{nombre_perfil}.lnk")
    return crear_acceso_directo(destino, nombre_perfil, icono)


def desactivar_inicio_automatico(nombre_perfil):
    carpeta = _carpeta_inicio_automatico()
    for ext in (".lnk", ".bat"):
        ruta = os.path.join(carpeta, f"{nombre_perfil}{ext}")
        if os.path.exists(ruta):
            os.remove(ruta)


# ---------------------------------------------------------------------------
# Iconos: se ajustan automáticamente al tamaño solicitado, sin deformarse
# ---------------------------------------------------------------------------

COLOR_TRANSPARENTE = "#ff00ff"  # magenta poco común, usado como "llave" de transparencia


def _icono_pil_ajustado(ruta, tamano_max=64):
    """Abre una imagen y la reduce/amplía manteniendo su proporción original
    (nunca la deforma), devolviendo un objeto PIL.Image con su tamaño real
    resultante (puede ser rectangular si el icono original no es cuadrado)."""
    if not TIENE_PIL or not ruta or not os.path.exists(ruta):
        return None
    try:
        img = Image.open(ruta).convert("RGBA")
    except Exception:
        return None
    img.thumbnail((tamano_max, tamano_max), Image.LANCZOS)
    return img


def _icono_generico_pil(nombre_perfil, tamano=64):
    """Círculo con la inicial del perfil, usado cuando no hay Pillow o no hay icono."""
    if not TIENE_PIL:
        return None
    img = Image.new("RGBA", (tamano, tamano), (0, 0, 0, 0))
    dibujo = ImageDraw.Draw(img)
    dibujo.ellipse((1, 1, tamano - 1, tamano - 1), fill=(59, 111, 214, 255))
    letra = (nombre_perfil or "?")[:1].upper()
    dibujo.text((tamano / 2, tamano / 2), letra, fill="white", anchor="mm")
    return img


def _icono_cuadrado_pil(ruta, nombre_perfil, tamano=64):
    """Versión centrada en lienzo cuadrado (tamano x tamano) para la barra de
    tareas, que exige una imagen de proporción fija sin recortar el icono."""
    base = _icono_pil_ajustado(ruta, tamano) if ruta else None
    if base is None:
        return _icono_generico_pil(nombre_perfil, tamano)
    lienzo = Image.new("RGBA", (tamano, tamano), (0, 0, 0, 0))
    x = (tamano - base.width) // 2
    y = (tamano - base.height) // 2
    lienzo.paste(base, (x, y), base)
    return lienzo


def cargar_imagen_icono(ruta, tamano_max=64):
    """Devuelve (PhotoImage, ancho, alto) ajustado proporcionalmente al
    tamaño real del icono, sin estirarlo ni deformarlo. Usa Pillow si está
    disponible; si no, hace un ajuste básico con subsample() para .png."""
    if not ruta or not os.path.exists(ruta):
        return None, tamano_max, tamano_max
    if TIENE_PIL:
        img = _icono_pil_ajustado(ruta, tamano_max)
        if img is None:
            return None, tamano_max, tamano_max
        return ImageTk.PhotoImage(img), img.width, img.height
    if ruta.lower().endswith(".png"):
        try:
            img = tk.PhotoImage(file=ruta)
            factor = max(1, img.width() // tamano_max, img.height() // tamano_max)
            if factor > 1:
                img = img.subsample(factor, factor)
            return img, img.width(), img.height()
        except Exception:
            return None, tamano_max, tamano_max
    return None, tamano_max, tamano_max


class IconoFlotante(tk.Toplevel):
    """Ventana sin bordes ni fondo: solo el icono del perfil, flotando y arrastrable.
    El tamaño de la ventana se ajusta automáticamente al tamaño real del icono
    (respetando su proporción original, sin estirarlo)."""

    def __init__(self, master, nombre_perfil, ruta_icono, al_hacer_clic, x=120, y=120, tamano_max=64):
        super().__init__(master)
        self.nombre_perfil = nombre_perfil
        self.overrideredirect(True)          # sin bordes ni barra de título
        self.attributes("-topmost", True)     # siempre visible
        try:
            self.attributes("-transparentcolor", COLOR_TRANSPARENTE)  # solo Windows
        except tk.TclError:
            pass
        self.configure(bg=COLOR_TRANSPARENTE)

        imagen, ancho, alto = cargar_imagen_icono(ruta_icono, tamano_max)
        self._imagen_ref = imagen  # evitar que el garbage collector la borre

        # La ventana se dimensiona según el icono real (ancho x alto), no un
        # cuadro fijo, para que iconos rectangulares no queden deformados.
        self.geometry(f"{max(ancho, 40)+16}x{alto+34}+{x}+{y}")

        contenedor = tk.Frame(self, bg=COLOR_TRANSPARENTE)
        contenedor.pack(expand=True, fill="both")

        if imagen is not None:
            lbl_icono = tk.Label(contenedor, image=imagen, bg=COLOR_TRANSPARENTE, cursor="hand2")
        else:
            # Sin icono elegido: círculo simple con la inicial del perfil, igual sin fondo
            tamano = tamano_max
            canvas = tk.Canvas(contenedor, width=tamano, height=tamano,
                                bg=COLOR_TRANSPARENTE, highlightthickness=0, cursor="hand2")
            canvas.create_oval(2, 2, tamano - 2, tamano - 2, fill="#3b6fd6", outline="")
            canvas.create_text(tamano // 2, tamano // 2, text=nombre_perfil[:1].upper(),
                                fill="white", font=("Segoe UI", int(tamano / 2.4), "bold"))
            lbl_icono = canvas

        lbl_icono.pack(pady=(2, 0))
        lbl_texto = tk.Label(contenedor, text=nombre_perfil, bg=COLOR_TRANSPARENTE,
                              fg="white", font=("Segoe UI", 8, "bold"), cursor="hand2")
        lbl_texto.pack()

        for widget in (lbl_icono, lbl_texto):
            widget.bind("<Button-1>", self._iniciar_arrastre)
            widget.bind("<B1-Motion>", self._arrastrar)
            widget.bind("<ButtonRelease-1>", lambda e: self._quizas_click(al_hacer_clic))
            widget.bind("<Double-Button-1>", lambda e: al_hacer_clic())
            widget.bind("<Button-3>", lambda e: self.destroy())  # clic derecho: ocultar

        self._origen = None
        self._movido = False

    def _iniciar_arrastre(self, event):
        self._origen = (event.x_root, event.y_root, self.winfo_x(), self.winfo_y())
        self._movido = False

    def _arrastrar(self, event):
        if not self._origen:
            return
        ox, oy, wx, wy = self._origen
        dx, dy = event.x_root - ox, event.y_root - oy
        if abs(dx) > 3 or abs(dy) > 3:
            self._movido = True
        self.geometry(f"+{wx+dx}+{wy+dy}")

    def _quizas_click(self, callback):
        if not self._movido:
            callback()
        self._origen = None


# ---------------------------------------------------------------------------
# Icono en la barra de tareas (área de notificación de Windows)
# ---------------------------------------------------------------------------

class BarraTareaPerfil:
    """Administra los iconos de bandeja (uno por perfil) usando pystray.
    Cada uno corre en su propio hilo en segundo plano."""

    _instancias = {}  # nombre_perfil -> pystray.Icon

    @classmethod
    def activo(cls, nombre_perfil):
        return nombre_perfil in cls._instancias

    @classmethod
    def mostrar(cls, nombre_perfil, ruta_icono, al_hacer_clic):
        if not TIENE_TRAY or nombre_perfil in cls._instancias:
            return
        imagen = _icono_cuadrado_pil(ruta_icono, nombre_perfil, tamano=64)

        def _lanzar(icon, item):
            al_hacer_clic()

        def _quitar(icon, item):
            cls.ocultar(nombre_perfil)

        menu = pystray.Menu(
            pystray.MenuItem("Iniciar perfil", _lanzar, default=True),
            pystray.MenuItem("Quitar de la barra de tareas", _quitar),
        )
        icono = pystray.Icon(nombre_perfil, imagen, nombre_perfil, menu)
        cls._instancias[nombre_perfil] = icono
        threading.Thread(target=icono.run, daemon=True).start()

    @classmethod
    def ocultar(cls, nombre_perfil):
        icono = cls._instancias.pop(nombre_perfil, None)
        if icono:
            try:
                icono.stop()
            except Exception:
                pass

    @classmethod
    def ocultar_todos(cls):
        for nombre in list(cls._instancias.keys()):
            cls.ocultar(nombre)


# ---------------------------------------------------------------------------
# Interfaz gráfica
# ---------------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Asistente de Perfiles")
        self.geometry("820x520")
        self.minsize(680, 420)

        self.perfiles = cargar_perfiles()
        self.perfil_actual = None
        self.iconos_flotantes = {}  # nombre_perfil -> IconoFlotante abierto

        self._construir_ui()
        self._refrescar_lista_perfiles()
        self.protocol("WM_DELETE_WINDOW", self._on_cerrar)

        faltantes = []
        if not TIENE_TRAY:
            faltantes.append("icono en la barra de tareas")
        if not TIENE_PYWIN32:
            faltantes.append("icono propio en 'Iniciar con Windows'")
        if faltantes:
            texto = (
                "No se pudieron instalar automáticamente algunos componentes "
                "(revisa tu conexión a internet), así que estas funciones no "
                "estarán disponibles por ahora:\n\n- " + "\n- ".join(faltantes) +
                "\n\nSe volverá a intentar la próxima vez que abras la aplicación."
            )
            self.after(600, lambda: self._avisar_una_vez("Aviso", texto))

    def _on_cerrar(self):
        BarraTareaPerfil.ocultar_todos()
        self.destroy()

    def _avisar_una_vez(self, titulo, texto):
        messagebox.showinfo(titulo, texto)

    # ---------------- UI ----------------
    def _construir_ui(self):
        izq = ttk.Frame(self, padding=10)
        izq.pack(side="left", fill="y")

        ttk.Label(izq, text="Perfiles", font=("Segoe UI", 12, "bold")).pack(anchor="w")

        self.lista_perfiles = tk.Listbox(izq, width=24, height=20, exportselection=False)
        self.lista_perfiles.pack(fill="y", expand=True, pady=(5, 5))
        self.lista_perfiles.bind("<<ListboxSelect>>", self._on_seleccionar_perfil)

        btns_perfil = ttk.Frame(izq)
        btns_perfil.pack(fill="x", pady=(5, 0))
        ttk.Button(btns_perfil, text="+ Nuevo", command=self._nuevo_perfil).pack(side="left", expand=True, fill="x")
        ttk.Button(btns_perfil, text="Renombrar", command=self._renombrar_perfil).pack(side="left", expand=True, fill="x")
        ttk.Button(btns_perfil, text="Eliminar", command=self._eliminar_perfil).pack(side="left", expand=True, fill="x")

        # Panel derecho
        der = ttk.Frame(self, padding=10)
        der.pack(side="left", fill="both", expand=True)

        cabecera = ttk.Frame(der)
        cabecera.pack(fill="x")

        self.canvas_icono = tk.Canvas(cabecera, width=40, height=40, highlightthickness=0)
        self.canvas_icono.pack(side="left", padx=(0, 8))

        self.lbl_perfil = ttk.Label(cabecera, text="Selecciona o crea un perfil", font=("Segoe UI", 13, "bold"))
        self.lbl_perfil.pack(side="left")

        self.btn_lanzar = ttk.Button(cabecera, text="▶ Iniciar perfil", command=self._lanzar_perfil, state="disabled")
        self.btn_lanzar.pack(side="right")

        # Barra de integración con Windows
        barra_win = ttk.LabelFrame(der, text="Integración con Windows", padding=8)
        barra_win.pack(fill="x", pady=(12, 5))

        ttk.Button(barra_win, text="🖼️ Elegir icono...", command=self._elegir_icono).pack(side="left")

        self.var_auto_inicio = tk.BooleanVar(value=False)
        self.chk_auto_inicio = ttk.Checkbutton(
            barra_win, text="🚀 Iniciar con Windows", variable=self.var_auto_inicio,
            command=self._toggle_auto_inicio, state="disabled",
        )
        self.chk_auto_inicio.pack(side="left", padx=6)

        self.var_flotante = tk.BooleanVar(value=False)
        self.chk_flotante = ttk.Checkbutton(
            barra_win, text="🖥️ Icono flotante en escritorio", variable=self.var_flotante,
            command=self._toggle_icono_flotante, state="disabled",
        )
        self.chk_flotante.pack(side="left", padx=6)

        self.var_bandeja = tk.BooleanVar(value=False)
        self.chk_bandeja = ttk.Checkbutton(
            barra_win, text="🗂️ Icono en la barra de tareas", variable=self.var_bandeja,
            command=self._toggle_icono_bandeja, state="disabled",
        )
        self.chk_bandeja.pack(side="left", padx=6)

        # Días en que se activa el perfil
        barra_dias = ttk.LabelFrame(der, text="Días en que se activa este perfil", padding=8)
        barra_dias.pack(fill="x", pady=(5, 5))

        self.vars_dias = []
        for abrev in DIAS_ABREV:
            var = tk.BooleanVar(value=True)
            chk = ttk.Checkbutton(barra_dias, text=abrev, variable=var, command=self._guardar_dias_activos)
            chk.pack(side="left", padx=5)
            self.vars_dias.append(var)

        ttk.Button(barra_dias, text="Todos", command=lambda: self._marcar_dias(True)).pack(side="left", padx=(12, 3))
        ttk.Button(barra_dias, text="Ninguno", command=lambda: self._marcar_dias(False)).pack(side="left")

        ttk.Label(der, text="Programas y archivos incluidos:").pack(anchor="w", pady=(12, 3))

        cols = ("nombre", "ruta", "args")
        self.tabla = ttk.Treeview(der, columns=cols, show="headings", height=12)
        self.tabla.heading("nombre", text="Nombre")
        self.tabla.heading("ruta", text="Ruta")
        self.tabla.heading("args", text="Argumentos (opcional)")
        self.tabla.column("nombre", width=140)
        self.tabla.column("ruta", width=380)
        self.tabla.column("args", width=140)
        self.tabla.pack(fill="both", expand=True)

        btns_items = ttk.Frame(der)
        btns_items.pack(fill="x", pady=(8, 0))
        ttk.Button(btns_items, text="📋 Elegir de programas instalados", command=self._abrir_selector_instalados).pack(side="left")
        ttk.Button(btns_items, text="+ Agregar manualmente", command=self._agregar_item).pack(side="left", padx=6)
        ttk.Button(btns_items, text="Quitar seleccionado", command=self._quitar_item).pack(side="left")

    # ---------------- Lógica de perfiles ----------------
    def _refrescar_lista_perfiles(self):
        self.lista_perfiles.delete(0, tk.END)
        for nombre in sorted(self.perfiles.keys()):
            self.lista_perfiles.insert(tk.END, nombre)

    def _perfil(self):
        """Diccionario del perfil actualmente seleccionado (o None)."""
        if not self.perfil_actual:
            return None
        return self.perfiles.setdefault(self.perfil_actual, perfil_vacio())

    def _on_seleccionar_perfil(self, event=None):
        sel = self.lista_perfiles.curselection()
        if not sel:
            return
        nombre = self.lista_perfiles.get(sel[0])
        self.perfil_actual = nombre
        self.lbl_perfil.config(text=nombre)
        self.btn_lanzar.config(state="normal")
        self.chk_auto_inicio.config(state="normal")
        self.chk_flotante.config(state="normal")
        self.chk_bandeja.config(state="normal")
        self.var_auto_inicio.set(bool(self._perfil().get("auto_inicio")))
        self.var_flotante.set(nombre in self.iconos_flotantes)
        self.var_bandeja.set(BarraTareaPerfil.activo(nombre))
        dias_activos = set(self._perfil().get("dias_activos", list(range(7))))
        for i, var in enumerate(self.vars_dias):
            var.set(i in dias_activos)
        self._refrescar_tabla()
        self._refrescar_icono_cabecera()

    def _guardar_dias_activos(self):
        perfil = self._perfil()
        if not perfil:
            return
        perfil["dias_activos"] = [i for i, var in enumerate(self.vars_dias) if var.get()]
        guardar_perfiles(self.perfiles)

    def _marcar_dias(self, valor):
        if not self.perfil_actual:
            return
        for var in self.vars_dias:
            var.set(valor)
        self._guardar_dias_activos()

    def _nuevo_perfil(self):
        nombre = simpledialog.askstring("Nuevo perfil", "Nombre del perfil (ej: Trabajo):", parent=self)
        if not nombre:
            return
        nombre = nombre.strip()
        if not nombre:
            return
        if nombre in self.perfiles:
            messagebox.showwarning("Ya existe", "Ya existe un perfil con ese nombre.")
            return
        self.perfiles[nombre] = perfil_vacio()
        self.perfiles[nombre]["items"] = []
        guardar_perfiles(self.perfiles)
        self._refrescar_lista_perfiles()

    def _renombrar_perfil(self):
        if not self.perfil_actual:
            return
        nuevo = simpledialog.askstring("Renombrar perfil", "Nuevo nombre:", initialvalue=self.perfil_actual, parent=self)
        if not nuevo or nuevo == self.perfil_actual:
            return
        if nuevo in self.perfiles:
            messagebox.showwarning("Ya existe", "Ya existe un perfil con ese nombre.")
            return
        self.perfiles[nuevo] = self.perfiles.pop(self.perfil_actual)
        self.perfil_actual = nuevo
        guardar_perfiles(self.perfiles)
        self._refrescar_lista_perfiles()
        self.lbl_perfil.config(text=nuevo)

    def _eliminar_perfil(self):
        if not self.perfil_actual:
            return
        if not messagebox.askyesno("Confirmar", f"¿Eliminar el perfil '{self.perfil_actual}'?"):
            return
        desactivar_inicio_automatico(self.perfil_actual)
        if self.perfil_actual in self.iconos_flotantes:
            self.iconos_flotantes.pop(self.perfil_actual).destroy()
        BarraTareaPerfil.ocultar(self.perfil_actual)
        del self.perfiles[self.perfil_actual]
        guardar_perfiles(self.perfiles)
        self.perfil_actual = None
        self.lbl_perfil.config(text="Selecciona o crea un perfil")
        self.btn_lanzar.config(state="disabled")
        self.chk_auto_inicio.config(state="disabled")
        self.chk_flotante.config(state="disabled")
        self.chk_bandeja.config(state="disabled")
        self._refrescar_lista_perfiles()
        self._refrescar_tabla()
        self._refrescar_icono_cabecera()

    # ---------------- Icono / integración con Windows ----------------
    def _refrescar_icono_cabecera(self):
        self.canvas_icono.delete("all")
        perfil = self._perfil()
        if not perfil:
            return
        img, _a, _h = cargar_imagen_icono(perfil.get("icono"), tamano_max=36)
        if img is not None:
            self._icono_cabecera_ref = img  # evitar garbage collection
            self.canvas_icono.create_image(20, 20, image=img)
        else:
            self.canvas_icono.create_oval(2, 2, 38, 38, fill="#3b6fd6", outline="")
            self.canvas_icono.create_text(20, 20, text=self.perfil_actual[:1].upper(),
                                           fill="white", font=("Segoe UI", 14, "bold"))

    def _elegir_icono(self):
        perfil = self._perfil()
        if not perfil:
            messagebox.showinfo("Selecciona un perfil", "Primero selecciona o crea un perfil.")
            return
        ruta = filedialog.askopenfilename(
            title="Selecciona un icono para el perfil",
            filetypes=[("Imágenes", "*.ico *.png *.jpg *.jpeg *.bmp"), ("Todos los archivos", "*.*")],
        )
        if not ruta:
            return
        perfil["icono"] = ruta
        guardar_perfiles(self.perfiles)
        self._refrescar_icono_cabecera()
        if not TIENE_PIL and not ruta.lower().endswith(".png"):
            messagebox.showinfo(
                "Aviso",
                "No se pudo instalar automáticamente el componente para leer iconos .ico/.jpg "
                "(revisa tu conexión a internet). Mientras tanto se usará un icono genérico. "
                "Se reintentará la próxima vez que abras la aplicación.",
            )

    def _toggle_auto_inicio(self):
        perfil = self._perfil()
        if not perfil:
            return
        activar = self.var_auto_inicio.get()
        try:
            if activar:
                activar_inicio_automatico(self.perfil_actual, perfil.get("icono"))
            else:
                desactivar_inicio_automatico(self.perfil_actual)
            perfil["auto_inicio"] = activar
            guardar_perfiles(self.perfiles)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo actualizar el inicio automático:\n{e}")
            self.var_auto_inicio.set(not activar)

    def _toggle_icono_flotante(self):
        nombre = self.perfil_actual
        if not nombre:
            return
        if nombre in self.iconos_flotantes:
            self.iconos_flotantes.pop(nombre).destroy()
            self.var_flotante.set(False)
            return
        perfil = self._perfil()
        ventana = IconoFlotante(
            self, nombre, perfil.get("icono"),
            al_hacer_clic=lambda: self._lanzar_con_cooldown_silencioso(nombre),
        )
        self.iconos_flotantes[nombre] = ventana
        self.var_flotante.set(True)

    def _toggle_icono_bandeja(self):
        nombre = self.perfil_actual
        if not nombre:
            return
        if not TIENE_TRAY:
            self.var_bandeja.set(False)
            messagebox.showinfo(
                "No disponible por ahora",
                "El icono en la barra de tareas necesita un componente que no se pudo "
                "instalar automáticamente (revisa tu conexión a internet y vuelve a "
                "abrir la aplicación).",
            )
            return
        if BarraTareaPerfil.activo(nombre):
            BarraTareaPerfil.ocultar(nombre)
            self.var_bandeja.set(False)
            return
        perfil = self._perfil()
        BarraTareaPerfil.mostrar(
            nombre, perfil.get("icono"),
            al_hacer_clic=lambda: self._lanzar_con_cooldown_silencioso(nombre),
        )
        self.var_bandeja.set(True)

    def _lanzar_con_cooldown_silencioso(self, nombre_perfil):
        """Usado por el icono flotante y el de la bandeja: si hoy no es un
        día activo del perfil, o si ya se lanzó hace menos de 2 minutos,
        simplemente no hace nada (sin ventanas emergentes que interrumpan)."""
        perfil = self.perfiles.get(nombre_perfil)
        if perfil:
            lanzar_perfil_con_cooldown(nombre_perfil, perfil)

    # ---------------- Lógica de items ----------------
    def _refrescar_tabla(self):
        self.tabla.delete(*self.tabla.get_children())
        perfil = self._perfil()
        if not perfil:
            return
        for item in perfil.get("items", []):
            self.tabla.insert("", tk.END, values=(item["nombre"], item["ruta"], item.get("args", "")))

    def _agregar_item(self):
        perfil = self._perfil()
        if not perfil:
            messagebox.showinfo("Selecciona un perfil", "Primero selecciona o crea un perfil.")
            return
        ruta = filedialog.askopenfilename(title="Selecciona un programa o archivo")
        if not ruta:
            return
        nombre_sugerido = os.path.splitext(os.path.basename(ruta))[0]
        nombre = simpledialog.askstring("Nombre para mostrar", "Nombre para este elemento:", initialvalue=nombre_sugerido, parent=self)
        if not nombre:
            nombre = nombre_sugerido
        args = simpledialog.askstring("Argumentos (opcional)", "Argumentos de línea de comandos (dejar vacío si no aplica):", parent=self)
        perfil["items"].append({"nombre": nombre, "ruta": ruta, "args": args or ""})
        guardar_perfiles(self.perfiles)
        self._refrescar_tabla()

    def _abrir_selector_instalados(self):
        perfil = self._perfil()
        if not perfil:
            messagebox.showinfo("Selecciona un perfil", "Primero selecciona o crea un perfil.")
            return

        ventana = tk.Toplevel(self)
        ventana.title("Programas instalados")
        ventana.geometry("500x460")
        ventana.transient(self)
        ventana.grab_set()

        ttk.Label(ventana, text="Buscando programas instalados...", padding=10).pack(anchor="w")
        ventana.update_idletasks()

        programas = listar_programas_instalados()

        cabecera = ttk.Frame(ventana, padding=(10, 0))
        cabecera.pack(fill="x")
        ttk.Label(cabecera, text="Buscar:").pack(side="left")
        var_filtro = tk.StringVar()
        entrada = ttk.Entry(cabecera, textvariable=var_filtro)
        entrada.pack(side="left", fill="x", expand=True, padx=(6, 0))
        entrada.focus_set()

        cont_lista = ttk.Frame(ventana, padding=10)
        cont_lista.pack(fill="both", expand=True)
        lista = tk.Listbox(cont_lista)
        scroll = ttk.Scrollbar(cont_lista, orient="vertical", command=lista.yview)
        lista.config(yscrollcommand=scroll.set)
        lista.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        def poblar(filtro=""):
            lista.delete(0, tk.END)
            filtro = filtro.strip().lower()
            for nombre, _ruta in programas:
                if filtro in nombre.lower():
                    lista.insert(tk.END, nombre)

        poblar()

        if not programas:
            ttk.Label(
                ventana,
                text="No se encontraron programas automáticamente.\nUsa 'Agregar manualmente' en su lugar.",
                padding=10, foreground="#a33",
            ).pack()

        def al_escribir(*_args):
            poblar(var_filtro.get())

        var_filtro.trace_add("write", al_escribir)

        mapa_nombre_ruta = {nombre: ruta for nombre, ruta in programas}

        def agregar_seleccion(_event=None):
            sel = lista.curselection()
            if not sel:
                return
            nombre = lista.get(sel[0])
            ruta = mapa_nombre_ruta.get(nombre)
            if not ruta:
                return
            args = simpledialog.askstring(
                "Argumentos (opcional)",
                "Argumentos de línea de comandos (dejar vacío si no aplica):",
                parent=ventana,
            )
            perfil["items"].append({"nombre": nombre, "ruta": ruta, "args": args or ""})
            guardar_perfiles(self.perfiles)
            self._refrescar_tabla()
            ventana.destroy()

        lista.bind("<Double-Button-1>", agregar_seleccion)

        botones = ttk.Frame(ventana, padding=10)
        botones.pack(fill="x")
        ttk.Button(botones, text="Agregar seleccionado", command=agregar_seleccion).pack(side="left")
        ttk.Button(botones, text="Cancelar", command=ventana.destroy).pack(side="left", padx=6)

    def _quitar_item(self):
        sel = self.tabla.selection()
        perfil = self._perfil()
        if not sel or not perfil:
            return
        idx = self.tabla.index(sel[0])
        del perfil["items"][idx]
        guardar_perfiles(self.perfiles)
        self._refrescar_tabla()

    # ---------------- Lanzar perfil ----------------
    def _lanzar_perfil(self):
        perfil = self._perfil()
        if not perfil:
            return
        if not perfil.get("items"):
            messagebox.showinfo("Perfil vacío", "Este perfil no tiene archivos ni programas agregados.")
            return

        estado = lanzar_perfil_con_cooldown(self.perfil_actual, perfil)
        if estado == "ok":
            return

        if estado == "dia_no_permitido":
            dias_activos = sorted(perfil.get("dias_activos", []))
            texto_dias = ", ".join(DIAS_SEMANA[d] for d in dias_activos) if dias_activos else "ningún día"
            hoy = DIAS_SEMANA[datetime.date.today().weekday()]
            if messagebox.askyesno(
                "No programado para hoy",
                f"'{self.perfil_actual}' está configurado para abrirse solo: {texto_dias}.\n"
                f"Hoy es {hoy}.\n\n¿Quieres iniciarlo de todas formas?",
            ):
                lanzar_perfil_con_cooldown(self.perfil_actual, perfil, forzar=True)
            return

        if estado == "cooldown":
            transcurrido = segundos_desde_ultimo_lanzamiento(self.perfil_actual) or 0
            restante = max(1, int(COOLDOWN_SEGUNDOS - transcurrido))
            if messagebox.askyesno(
                "Ya se inició hace poco",
                f"'{self.perfil_actual}' ya se abrió hace menos de 2 minutos "
                f"(faltan ~{restante} seg. para poder repetirlo automáticamente).\n\n"
                "¿Quieres iniciarlo de todas formas?",
            ):
                lanzar_perfil_con_cooldown(self.perfil_actual, perfil, forzar=True)


# ---------------------------------------------------------------------------
# Punto de entrada
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Asistente de Perfiles")
    parser.add_argument("--lanzar", metavar="PERFIL", help="Abre todos los items del perfil indicado y sale (sin GUI). Usado por accesos directos de Windows.")
    args = parser.parse_args()

    if args.lanzar:
        # Modo headless: usado desde el Menú Inicio o el inicio automático de Windows
        perfiles = cargar_perfiles()
        perfil = perfiles.get(args.lanzar)
        if perfil:
            lanzar_perfil_con_cooldown(args.lanzar, perfil)
        return

    if sys.platform != "win32":
        print("Este script está pensado para Windows (usa os.startfile y accesos directos .lnk).")

    # Instala automáticamente pillow / pystray / pywin32 si faltan, mostrando
    # una ventana de progreso. Solo ocurre en el primer arranque (o si se
    # borraron las dependencias); las siguientes veces se salta este paso.
    faltantes = _dependencias_faltantes()
    if faltantes:
        _instalar_dependencias_con_progreso(faltantes)

    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
