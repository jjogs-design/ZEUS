import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import webbrowser
import urllib.parse
from datetime import datetime
import os
import tempfile
import threading
import json
import base64

# El motor Chromium va empaquetado dentro del .exe (ver CREAR_EJECUTABLE_WINDOWS.bat).
# Esta variable debe fijarse ANTES de importar Playwright, tambien al ejecutar el .exe.
os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", "0")

try:
    from fpdf import FPDF
    from docx import Document
    from docx.shared import Inches
    import openpyxl
    from PIL import Image, ImageGrab, ImageTk
    EXPORT_READY = True
except ImportError:
    EXPORT_READY = False

try:
    from playwright.sync_api import sync_playwright
    PLAYWRIGHT_READY = True
except ImportError:
    PLAYWRIGHT_READY = False


class SelectorRecorte(tk.Toplevel):
    """Ventana transparente a pantalla completa para recortar una zona con el ratón."""
    def __init__(self, master, callback):
        super().__init__(master)
        self.callback = callback
        self.attributes("-fullscreen", True)
        self.attributes("-alpha", 0.3)
        self.configure(bg="black", cursor="cross")
        self.canvas = tk.Canvas(self, bg="black", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)

        self.start_x = None
        self.start_y = None
        self.rect = None

        self.canvas.bind("<ButtonPress-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.bind("<Escape>", lambda e: self.destroy())

    def on_press(self, event):
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline="#00a8ff", width=2)

    def on_drag(self, event):
        cur_x = self.canvas.canvasx(event.x)
        cur_y = self.canvas.canvasy(event.y)
        self.canvas.coords(self.rect, self.start_x, self.start_y, cur_x, cur_y)

    def on_release(self, event):
        end_x = self.canvas.canvasx(event.x)
        end_y = self.canvas.canvasy(event.y)
        x1 = int(min(self.start_x, end_x))
        y1 = int(min(self.start_y, end_y))
        x2 = int(max(self.start_x, end_x))
        y2 = int(max(self.start_y, end_y))
        self.destroy()
        if (x2 - x1) > 10 and (y2 - y1) > 10:
            self.after(150, lambda: self.callback((x1, y1, x2, y2)))


class ZeusOSINTApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ZEUS – Centro de Inteligencia OSINT")
        self.root.geometry("1060x800")
        self.root.configure(bg="#121417")

        self.resultados_sesion = []
        self.cronologia_eventos = []
        self.capturas_evidencia = []
        self.temp_dir = tempfile.mkdtemp(prefix="zeus_osint_")

        self._configurar_estilos()
        self._construir_interfaz()

    def _configurar_estilos(self):
        style = ttk.Style()
        style.theme_use("clam")

        bg_dark = "#121417"
        bg_panel = "#1e2229"
        accent = "#0078d4"
        success = "#107c41"
        text_light = "#f0f2f5"

        style.configure("TFrame", background=bg_dark)
        style.configure("Panel.TFrame", background=bg_panel, relief="flat")
        style.configure("TLabel", background=bg_dark, foreground=text_light, font=("Segoe UI", 10))
        style.configure("Panel.TLabel", background=bg_panel, foreground=text_light, font=("Segoe UI", 10))
        style.configure("Header.TLabel", background=bg_dark, foreground="#ffffff", font=("Segoe UI", 18, "bold"))
        style.configure("SubHeader.TLabel", background=bg_dark, foreground="#8b949e", font=("Segoe UI", 10))
        style.configure("TCheckbutton", background=bg_panel, foreground=text_light, font=("Segoe UI", 9))
        style.map("TCheckbutton", background=[("active", bg_panel)])

        style.configure("Accent.TButton", background=accent, foreground="white", font=("Segoe UI", 11, "bold"), padding=8)
        style.map("Accent.TButton", background=[("active", "#005a9e")])

        style.configure("Success.TButton", background=success, foreground="white", font=("Segoe UI", 10, "bold"), padding=7)
        style.map("Success.TButton", background=[("active", "#0b5a30")])

        style.configure("Action.TButton", background="#2d333b", foreground="white", font=("Segoe UI", 9, "bold"), padding=6)
        style.map("Action.TButton", background=[("active", "#444c56")])

    def _construir_interfaz(self):
        # --- CABECERA ---
        header_frame = ttk.Frame(self.root, padding="20 15 20 5")
        header_frame.pack(fill="x")

        ttk.Label(header_frame, text="ZEUS – CENTRO DE INTELIGENCIA OSINT", style="Header.TLabel").pack(anchor="w")
        ttk.Label(header_frame, text="Investigación manual en fuentes abiertas, captura visual (URL y pantalla) y generación de expedientes", style="SubHeader.TLabel").pack(anchor="w")

        # --- PANEL DE BÚSQUEDA ---
        search_panel = ttk.Frame(self.root, style="Panel.TFrame", padding="15")
        search_panel.pack(fill="x", padx=20, pady=5)

        ttk.Label(search_panel, text="Objetivo de investigación (Nombre, apellidos o alias):", style="Panel.TLabel", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 5))

        self.entry_objetivo = ttk.Entry(search_panel, font=("Segoe UI", 12), width=52)
        self.entry_objetivo.grid(row=1, column=0, sticky="we", padx=(0, 15))
        self.entry_objetivo.insert(0, "@usuario123")
        self.entry_objetivo.bind("<FocusIn>", lambda e: self._limpiar_placeholder())

        tipo_frame = ttk.Frame(search_panel, style="Panel.TFrame")
        tipo_frame.grid(row=0, column=1, sticky="e", pady=(0, 5))
        ttk.Label(tipo_frame, text="Tipo:", style="Panel.TLabel", font=("Segoe UI", 10, "bold")).pack(side="left", padx=(0, 6))
        self.combo_tipo = ttk.Combobox(tipo_frame, values=["Persona", "Entidad / Empresa"], width=18, state="readonly")
        self.combo_tipo.current(0)
        self.combo_tipo.pack(side="left")

        self.btn_buscar = ttk.Button(search_panel, text="🔍 INICIAR BÚSQUEDA OSINT", style="Accent.TButton", command=self.ejecutar_busqueda)
        self.btn_buscar.grid(row=1, column=1, sticky="e")

        fuentes_frame = ttk.Frame(search_panel, style="Panel.TFrame")
        fuentes_frame.grid(row=2, column=0, columnspan=2, sticky="w", pady=(12, 2))

        ttk.Label(fuentes_frame, text="Fuentes:", style="Panel.TLabel", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0, 8))

        self.fuentes_vars = {
            "X (Twitter)": tk.BooleanVar(value=True),
            "Facebook": tk.BooleanVar(value=True),
            "Instagram": tk.BooleanVar(value=True),
            "TikTok": tk.BooleanVar(value=True),
            "YouTube": tk.BooleanVar(value=True),
            "Telegram": tk.BooleanVar(value=True),
            "Web y Prensa": tk.BooleanVar(value=True),
            "Documentos (PDF/DOC/XLS)": tk.BooleanVar(value=True)
        }

        for nombre, var in self.fuentes_vars.items():
            ttk.Checkbutton(fuentes_frame, text=nombre, variable=var).pack(side="left", padx=5)

        # --- AVISO DE LIMITACIONES ---
        info_frame = ttk.Frame(self.root, padding="20 4 20 4")
        info_frame.pack(fill="x")
        aviso_texto = (
            "ℹ️ Nota operativa: Las redes sociales con acceso restringido requieren sesión activa en su navegador. "
            "Para capturar publicaciones cerradas use el Recorte de Pantalla; para web abierta, prensa o perfiles públicos puede usar la Captura directa desde URL."
        )
        ttk.Label(info_frame, text=aviso_texto, foreground="#a3b3c2", font=("Segoe UI", 8, "italic"), wraplength=1000).pack(anchor="w")

        # --- PESTAÑAS DE RESULTADOS ---
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=20, pady=6)

        # Pestaña 1: Ficha de Investigación
        tab_ficha = ttk.Frame(notebook, style="Panel.TFrame", padding="15")
        notebook.add(tab_ficha, text=" 📋 Ficha de Investigación ")

        ttk.Label(tab_ficha, text="Notas del Analista / Perfil del Objetivo / Hallazgos Clave:", style="Panel.TLabel", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 5))
        self.txt_ficha = tk.Text(tab_ficha, height=12, bg="#121417", fg="#f0f2f5", insertbackground="white", font=("Segoe UI", 10), relief="flat", padx=10, pady=10)
        self.txt_ficha.pack(fill="both", expand=True)

        # Pestaña 2: Publicaciones y Enlaces
        tab_enlaces = ttk.Frame(notebook, style="Panel.TFrame", padding="15")
        notebook.add(tab_enlaces, text=" 🔗 Publicaciones y Enlaces ")

        form_enlace = ttk.Frame(tab_enlaces, style="Panel.TFrame")
        form_enlace.pack(fill="x", pady=(0, 10))

        ttk.Label(form_enlace, text="Fuente:", style="Panel.TLabel").pack(side="left", padx=(0, 5))
        self.combo_fuente = ttk.Combobox(form_enlace, values=list(self.fuentes_vars.keys()), width=22, state="readonly")
        self.combo_fuente.current(0)
        self.combo_fuente.pack(side="left", padx=(0, 10))

        ttk.Label(form_enlace, text="URL / Evidencia:", style="Panel.TLabel").pack(side="left", padx=(0, 5))
        self.entry_url = ttk.Entry(form_enlace, width=48)
        self.entry_url.pack(side="left", padx=(0, 10))

        ttk.Button(form_enlace, text="+ Añadir Hallazgo", style="Action.TButton", command=self.agregar_hallazgo).pack(side="left")

        self.lista_enlaces = tk.Listbox(tab_enlaces, bg="#121417", fg="#f0f2f5", font=("Segoe UI", 9), relief="flat", selectbackground="#0078d4")
        self.lista_enlaces.pack(fill="both", expand=True)

        # Pestaña 3: Cronología
        tab_crono = ttk.Frame(notebook, style="Panel.TFrame", padding="15")
        notebook.add(tab_crono, text=" 🕒 Cronología ")

        form_crono = ttk.Frame(tab_crono, style="Panel.TFrame")
        form_crono.pack(fill="x", pady=(0, 10))

        ttk.Label(form_crono, text="Fecha (DD/MM/AAAA):", style="Panel.TLabel").pack(side="left", padx=(0, 5))
        self.entry_fecha = ttk.Entry(form_crono, width=15)
        self.entry_fecha.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self.entry_fecha.pack(side="left", padx=(0, 10))

        ttk.Label(form_crono, text="Evento / Actividad:", style="Panel.TLabel").pack(side="left", padx=(0, 5))
        self.entry_evento = ttk.Entry(form_crono, width=52)
        self.entry_evento.pack(side="left", padx=(0, 10))

        ttk.Button(form_crono, text="+ Registrar Hito", style="Action.TButton", command=self.agregar_cronologia).pack(side="left")

        self.lista_crono = tk.Listbox(tab_crono, bg="#121417", fg="#f0f2f5", font=("Segoe UI", 9), relief="flat", selectbackground="#0078d4")
        self.lista_crono.pack(fill="both", expand=True)

        # Pestaña 4: Evidencia Visual (URL + Recorte + Pantalla)
        tab_capturas = ttk.Frame(notebook, style="Panel.TFrame", padding="15")
        notebook.add(tab_capturas, text=" 📸 Evidencia Visual ")

        # Fila 1: Captura directa desde URL
        url_cap_frame = ttk.Frame(tab_capturas, style="Panel.TFrame")
        url_cap_frame.pack(fill="x", pady=(0, 8))

        ttk.Label(url_cap_frame, text="URL de Publicación:", style="Panel.TLabel", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0, 5))
        self.entry_url_captura = ttk.Entry(url_cap_frame, width=58)
        self.entry_url_captura.pack(side="left", padx=(0, 8))

        self.btn_cap_url = ttk.Button(url_cap_frame, text="🌐 Capturar desde URL", style="Accent.TButton", command=self.iniciar_captura_url)
        self.btn_cap_url.pack(side="left")

        # Fila 2: Captura manual / pantalla / recorte
        botones_cap_frame = ttk.Frame(tab_capturas, style="Panel.TFrame")
        botones_cap_frame.pack(fill="x", pady=(0, 8))

        ttk.Button(botones_cap_frame, text="✂️ Recortar Zona con Ratón", style="Action.TButton", command=self.iniciar_recorte_raton).pack(side="left", padx=(0, 6))
        ttk.Button(botones_cap_frame, text="🖥️ Pantalla Completa", style="Action.TButton", command=self.capturar_pantalla).pack(side="left", padx=(0, 6))
        ttk.Button(botones_cap_frame, text="📋 Pegar Portapapeles", style="Action.TButton", command=self.pegar_portapapeles).pack(side="left", padx=(0, 6))
        ttk.Button(botones_cap_frame, text="📂 Cargar Archivo", style="Action.TButton", command=self.cargar_imagen_local).pack(side="left", padx=(0, 6))
        ttk.Button(botones_cap_frame, text="🗑️ Eliminar", style="Action.TButton", command=self.eliminar_captura).pack(side="right")

        desc_cap_frame = ttk.Frame(tab_capturas, style="Panel.TFrame")
        desc_cap_frame.pack(fill="x", pady=(0, 8))
        ttk.Label(desc_cap_frame, text="Pie de foto / Descripción para el informe:", style="Panel.TLabel").pack(side="left", padx=(0, 8))
        self.entry_desc_captura = ttk.Entry(desc_cap_frame, width=72)
        self.entry_desc_captura.pack(side="left", fill="x", expand=True)

        visor_frame = ttk.Frame(tab_capturas, style="Panel.TFrame")
        visor_frame.pack(fill="both", expand=True)

        self.lista_capturas = tk.Listbox(visor_frame, width=48, bg="#121417", fg="#f0f2f5", font=("Segoe UI", 9), relief="flat", selectbackground="#0078d4")
        self.lista_capturas.pack(side="left", fill="both", expand=True, padx=(0, 10))
        self.lista_capturas.bind("<<ListboxSelect>>", self.mostrar_miniatura)

        self.lbl_preview = tk.Label(visor_frame, text="Vista previa de evidencia", bg="#121417", fg="#8b949e", width=45, height=12)
        self.lbl_preview.pack(side="right", fill="both", expand=True)

        # Pestaña 5: Generador de Informe
        tab_informe = ttk.Frame(notebook, style="Panel.TFrame", padding="15")
        notebook.add(tab_informe, text=" 📑 Generador de Informe ")

        top_informe = ttk.Frame(tab_informe, style="Panel.TFrame")
        top_informe.pack(fill="x", pady=(0, 8))

        ttk.Label(top_informe, text="Borrador del Informe Ejecutivo (incluye Anexo Gráfico automático):", style="Panel.TLabel", font=("Segoe UI", 10, "bold")).pack(side="left")
        ttk.Button(top_informe, text="⚡ Redactar Informe Automático", style="Accent.TButton", command=self.generar_borrador_informe).pack(side="right")

        self.txt_informe = tk.Text(tab_informe, height=12, bg="#121417", fg="#f0f2f5", insertbackground="white", font=("Segoe UI", 10), relief="flat", padx=10, pady=10)
        self.txt_informe.pack(fill="both", expand=True)

        # --- BARRA DE EXPORTACIÓN ---
        export_frame = ttk.Frame(self.root, padding="20 8 20 15")
        export_frame.pack(fill="x")

        ttk.Button(export_frame, text="💾 Guardar expediente", style="Action.TButton", command=self.guardar_expediente).pack(side="left", padx=(0, 4))
        ttk.Button(export_frame, text="📂 Abrir expediente", style="Action.TButton", command=self.abrir_expediente).pack(side="left", padx=(0, 18))

        ttk.Label(export_frame, text="Exportar:", font=("Segoe UI", 10, "bold")).pack(side="left", padx=(0, 10))

        ttk.Button(export_frame, text="📄 PDF", style="Action.TButton", command=self.exportar_pdf).pack(side="left", padx=4)
        ttk.Button(export_frame, text="📝 Word (.docx)", style="Action.TButton", command=self.exportar_word).pack(side="left", padx=4)
        ttk.Button(export_frame, text="📊 Excel (.xlsx)", style="Action.TButton", command=self.exportar_excel).pack(side="left", padx=4)

        ttk.Button(export_frame, text="🚀 EXPORTAR TODO (PDF + WORD + EXCEL)", style="Success.TButton", command=self.exportar_todo_simultaneo).pack(side="right")

    def _limpiar_placeholder(self):
        if self.entry_objetivo.get() == "@usuario123":
            self.entry_objetivo.delete(0, tk.END)

    def ejecutar_busqueda(self):
        objetivo = self.entry_objetivo.get().strip()
        if not objetivo or objetivo == "@usuario123":
            messagebox.showwarning("Dato requerido", "Por favor, introduce un nombre, apellidos o alias válido.")
            return

        query_encoded = urllib.parse.quote(objetivo)

        urls_busqueda = {
            "X (Twitter)": f"https://x.com/search?q={query_encoded}&f=user",
            "Facebook": f"https://www.facebook.com/search/top/?q={query_encoded}",
            "Instagram": f"https://www.google.com/search?q=site:instagram.com+%22{query_encoded}%22",
            "TikTok": f"https://www.tiktok.com/search?q={query_encoded}",
            "YouTube": f"https://www.youtube.com/results?search_query={query_encoded}",
            "Telegram": f"https://www.google.com/search?q=site:t.me+%22{query_encoded}%22",
            "Web y Prensa": f"https://www.google.com/search?q=%22{query_encoded}%22&tbm=nws",
            "Documentos (PDF/DOC/XLS)": f"https://www.google.com/search?q=%22{query_encoded}%22+(filetype:pdf+OR+filetype:doc+OR+filetype:docx+OR+filetype:xls+OR+filetype:xlsx)"
        }

        seleccionadas = [f for f, var in self.fuentes_vars.items() if var.get()]
        if not seleccionadas:
            messagebox.showwarning("Sin fuentes", "Selecciona al menos una fuente de búsqueda.")
            return

        timestamp = datetime.now().strftime("%d/%m/%Y %H:%M")
        self.txt_ficha.insert(tk.END, f"[{timestamp}] Inicio de investigación sobre: {objetivo}\n")
        self.txt_ficha.insert(tk.END, f"Fuentes consultadas: {', '.join(seleccionadas)}\n{'-'*55}\n")

        for fuente in seleccionadas:
            url = urls_busqueda[fuente]
            webbrowser.open_new_tab(url)
            registro = f"[{fuente}] Consulta lanzada: {url}"
            self.resultados_sesion.append((fuente, url, timestamp))
            self.lista_enlaces.insert(tk.END, registro)

        messagebox.showinfo("Búsqueda Iniciada", f"Se han abierto {len(seleccionadas)} pestañas en tu navegador para analizar '{objetivo}'.\n\nPuedes capturar publicaciones desde la pestaña 'Evidencia Visual'.")

    def agregar_hallazgo(self):
        fuente = self.combo_fuente.get()
        url = self.entry_url.get().strip()
        if not url:
            return
        timestamp = datetime.now().strftime("%d/%m/%Y %H:%M")
        self.resultados_sesion.append((fuente, url, timestamp))
        self.lista_enlaces.insert(tk.END, f"[{fuente}] ({timestamp}) -> {url}")
        self.entry_url.delete(0, tk.END)

    def agregar_cronologia(self):
        fecha = self.entry_fecha.get().strip()
        evento = self.entry_evento.get().strip()
        if not fecha or not evento:
            return
        self.cronologia_eventos.append((fecha, evento))
        self.lista_crono.insert(tk.END, f"{fecha} | {evento}")
        self.entry_evento.delete(0, tk.END)

    # --- GESTIÓN DE CAPTURAS (URL, RECORTE, PANTALLA, PORTAPAPELES) ---
    def _registrar_imagen(self, img_pil, origen_default):
        ts = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        desc = self.entry_desc_captura.get().strip() or origen_default
        nombre_archivo = f"evidencia_{len(self.capturas_evidencia)+1}_{datetime.now().strftime('%H%M%S')}.png"
        ruta_guardado = os.path.join(self.temp_dir, nombre_archivo)
        img_pil.save(ruta_guardado, "PNG")

        self.capturas_evidencia.append((ruta_guardado, desc, ts))
        self.lista_capturas.insert(tk.END, f"[{ts}] {desc}")
        self.entry_desc_captura.delete(0, tk.END)
        self.lista_capturas.selection_clear(0, tk.END)
        self.lista_capturas.selection_set(tk.END)
        self.mostrar_miniatura(None)

    def iniciar_captura_url(self):
        if not PLAYWRIGHT_READY:
            messagebox.showerror("Motor Web no disponible", "El módulo Playwright no está instalado. Ejecuta CREAR_EJECUTABLE_WINDOWS.bat para habilitar la captura directa desde URL.")
            return
        url = self.entry_url_captura.get().strip()
        if not url:
            messagebox.showwarning("URL vacía", "Introduce la dirección web (URL) de la publicación o noticia a capturar.")
            return
        if not url.startswith(("http://", "https://")):
            url = "https://" + url

        self.btn_cap_url.config(state="disabled", text="⏳ Capturando...")
        threading.Thread(target=self._proceso_captura_url, args=(url,), daemon=True).start()

    def _proceso_captura_url(self, url):
        try:
            ruta_tmp = os.path.join(self.temp_dir, f"url_cap_{datetime.now().strftime('%H%M%S')}.png")
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=True)
                context = browser.new_context(viewport={"width": 1280, "height": 900}, locale="es-ES")
                page = context.new_page()
                page.goto(url, wait_until="networkidle", timeout=25000)
                page.wait_for_timeout(1500)
                page.screenshot(path=ruta_tmp, full_page=False)
                browser.close()

            img = Image.open(ruta_tmp)
            self.root.after(0, lambda: self._finalizar_captura_url(img, url, None))
        except Exception as e:
            self.root.after(0, lambda: self._finalizar_captura_url(None, url, str(e)))

    def _finalizar_captura_url(self, img, url, error):
        self.btn_cap_url.config(state="normal", text="🌐 Capturar desde URL")
        if error:
            messagebox.showerror("Error al capturar URL", f"No se pudo renderizar la página automáticamente:\n{error}\n\nSugerencia: Ábrela en tu navegador y usa 'Recortar Zona con Ratón'.")
            return
        self.entry_url_captura.delete(0, tk.END)
        self._registrar_imagen(img, f"Captura web directa: {url}")
        messagebox.showinfo("Captura Web Completada", "La publicación se ha capturado e incorporado a la Evidencia Visual.")

    def iniciar_recorte_raton(self):
        self.root.iconify()
        self.root.after(300, lambda: SelectorRecorte(self.root, self._completar_recorte))

    def _completar_recorte(self, bbox):
        img = ImageGrab.grab(bbox=bbox)
        self.root.deiconify()
        self._registrar_imagen(img, "Recorte de pantalla seleccionado")

    def capturar_pantalla(self):
        if not EXPORT_READY: return
        self.root.iconify()
        self.root.after(450, self._ejecutar_grab_pantalla)

    def _ejecutar_grab_pantalla(self):
        img = ImageGrab.grab()
        self.root.deiconify()
        self._registrar_imagen(img, "Captura de pantalla completa")

    def pegar_portapapeles(self):
        if not EXPORT_READY: return
        img = ImageGrab.grabclipboard()
        if isinstance(img, Image.Image):
            self._registrar_imagen(img, "Recorte pegado desde portapapeles")
        else:
            messagebox.showwarning("Portapapeles vacío", "No hay ninguna imagen en el portapapeles.")

    def cargar_imagen_local(self):
        if not EXPORT_READY: return
        ruta = filedialog.askopenfilename(title="Seleccionar imagen de evidencia", filetypes=[("Imágenes", "*.png *.jpg *.jpeg *.bmp")])
        if not ruta: return
        img = Image.open(ruta)
        self._registrar_imagen(img, f"Archivo adjunto: {os.path.basename(ruta)}")

    def eliminar_captura(self):
        sel = self.lista_capturas.curselection()
        if not sel: return
        idx = sel[0]
        self.lista_capturas.delete(idx)
        self.capturas_evidencia.pop(idx)
        self.lbl_preview.config(image="", text="Vista previa de evidencia")

    def mostrar_miniatura(self, event):
        sel = self.lista_capturas.curselection()
        if not sel: return
        ruta_img = self.capturas_evidencia[sel[0]][0]
        img = Image.open(ruta_img)
        img.thumbnail((340, 220))
        photo = ImageTk.PhotoImage(img)
        self.lbl_preview.config(image=photo, text="")
        self.lbl_preview.image = photo

    # --- EXPEDIENTES (GUARDAR / ABRIR / HISTORIAL) ---
    def _carpeta_expedientes(self):
        carpeta = os.path.join(os.path.expanduser("~"), "Documents", "ZEUS_Expedientes")
        os.makedirs(carpeta, exist_ok=True)
        return carpeta

    def guardar_expediente(self):
        objetivo = self.entry_objetivo.get().strip()
        if not objetivo or objetivo == "@usuario123":
            messagebox.showwarning("Dato requerido", "Introduce primero el objetivo de la investigación.")
            return
        seguro = "".join(c for c in objetivo if c.isalnum() or c in ("_", "-")) or "Objetivo"
        ruta = filedialog.asksaveasfilename(
            title="Guardar expediente", initialdir=self._carpeta_expedientes(),
            initialfile=f"{seguro}_{datetime.now().strftime('%Y%m%d_%H%M')}.zeus",
            defaultextension=".zeus", filetypes=[("Expediente ZEUS", "*.zeus")])
        if not ruta:
            return

        capturas = []
        for ruta_img, desc, ts in self.capturas_evidencia:
            ext = os.path.splitext(ruta_img)[1].lower()
            mime = "image/jpeg" if ext in (".jpg", ".jpeg") else "image/png"
            with open(ruta_img, "rb") as fh:
                capturas.append({"desc": desc, "fecha": ts, "mime": mime,
                                 "imagen_b64": base64.b64encode(fh.read()).decode("ascii")})

        datos = {
            "version": 1,
            "tipo": self.combo_tipo.get(),
            "objetivo": objetivo,
            "ficha": self.txt_ficha.get("1.0", "end-1c"),
            "enlaces": [{"fuente": f, "url": u, "fecha": t} for f, u, t in self.resultados_sesion],
            "cronologia": [{"fecha": f, "evento": e} for f, e in self.cronologia_eventos],
            "capturas": capturas,
            "informe": self.txt_informe.get("1.0", "end-1c"),
            "guardado": datetime.now().strftime("%d/%m/%Y %H:%M"),
        }
        with open(ruta, "w", encoding="utf-8") as fh:
            json.dump(datos, fh, ensure_ascii=False, indent=1)
        messagebox.showinfo("Expediente guardado", f"Expediente guardado en:\n{ruta}")

    def abrir_expediente(self):
        ruta = filedialog.askopenfilename(
            title="Abrir expediente", initialdir=self._carpeta_expedientes(),
            filetypes=[("Expediente ZEUS", "*.zeus *.json")])
        if not ruta:
            return
        try:
            with open(ruta, "r", encoding="utf-8") as fh:
                datos = json.load(fh)
        except Exception as e:
            messagebox.showerror("Archivo no válido", f"No se pudo leer el expediente:\n{e}")
            return

        self.entry_objetivo.delete(0, tk.END)
        self.entry_objetivo.insert(0, datos.get("objetivo", ""))
        tipo = datos.get("tipo", "Persona")
        self.combo_tipo.set(tipo if tipo in self.combo_tipo["values"] else "Persona")

        self.txt_ficha.delete("1.0", tk.END)
        self.txt_ficha.insert("1.0", datos.get("ficha", ""))
        self.txt_informe.delete("1.0", tk.END)
        self.txt_informe.insert("1.0", datos.get("informe", ""))

        self.resultados_sesion = []
        self.lista_enlaces.delete(0, tk.END)
        for e in datos.get("enlaces", []):
            self.resultados_sesion.append((e.get("fuente", ""), e.get("url", ""), e.get("fecha", "")))
            self.lista_enlaces.insert(tk.END, f"[{e.get('fuente', '')}] ({e.get('fecha', '')}) -> {e.get('url', '')}")

        self.cronologia_eventos = []
        self.lista_crono.delete(0, tk.END)
        for c in datos.get("cronologia", []):
            self.cronologia_eventos.append((c.get("fecha", ""), c.get("evento", "")))
            self.lista_crono.insert(tk.END, f"{c.get('fecha', '')} | {c.get('evento', '')}")

        self.capturas_evidencia = []
        self.lista_capturas.delete(0, tk.END)
        self.lbl_preview.config(image="", text="Vista previa de evidencia")
        for i, c in enumerate(datos.get("capturas", []), 1):
            try:
                raw = base64.b64decode(c.get("imagen_b64", ""))
            except Exception:
                continue
            ext = ".jpg" if c.get("mime") == "image/jpeg" else ".png"
            destino = os.path.join(self.temp_dir, f"abierta_{i}_{datetime.now().strftime('%H%M%S')}{ext}")
            with open(destino, "wb") as fh:
                fh.write(raw)
            self.capturas_evidencia.append((destino, c.get("desc", ""), c.get("fecha", "")))
            self.lista_capturas.insert(tk.END, f"[{c.get('fecha', '')}] {c.get('desc', '')}")

        messagebox.showinfo("Expediente abierto", f"Expediente '{datos.get('objetivo', '')}' cargado correctamente.")

    # --- GENERADOR DE INFORME EJECUTIVO ---
    def generar_borrador_informe(self):
        objetivo = self.entry_objetivo.get().strip() or "No especificado"
        fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")
        notas = self.txt_ficha.get("1.0", tk.END).strip()

        enlaces_texto = "\n".join([f"  - [{ts}] {fuente}: {url}" for fuente, url, ts in self.resultados_sesion]) or "  - Sin enlaces registrados."
        crono_texto = "\n".join([f"  - {fecha}: {ev}" for fecha, ev in self.cronologia_eventos]) or "  - Sin hitos cronológicos registrados."
        capturas_texto = "\n".join([f"  - Figura {i+1} [{ts}]: {desc}" for i, (_, desc, ts) in enumerate(self.capturas_evidencia)]) or "  - Sin capturas adjuntas."

        plantilla = f"""====================================================================
INFORME DE INTELIGENCIA EN FUENTES ABIERTAS (OSINT)
GENERADO POR: ZEUS – CENTRO DE INTELIGENCIA OSINT
====================================================================

1. DATOS GENERALES DEL EXPEDIENTE
--------------------------------------------------------------------
Objetivo Investigado : {objetivo}
Tipo de Objetivo     : {self.combo_tipo.get()}
Fecha y Hora Emisión : {fecha_actual}
Clasificación        : USO INTERNO / INVESTIGACIÓN OSINT

2. RESUMEN EJECUTIVO Y NOTAS DEL ANALISTA
--------------------------------------------------------------------
{notas if notas else "No se han introducido observaciones manuales en la ficha de investigación."}

3. ANÁLISIS DE FUENTES, PUBLICACIONES Y DOCUMENTOS RASTREADOS
--------------------------------------------------------------------
{enlaces_texto}

4. CRONOLOGÍA DE EVENTOS IDENTIFICADOS
--------------------------------------------------------------------
{crono_texto}

5. ÍNDICE DE EVIDENCIA VISUAL (ANEXO GRÁFICO)
--------------------------------------------------------------------
{capturas_texto}

6. OBSERVACIONES METODOLÓGICAS Y LIMITACIONES
--------------------------------------------------------------------
La información contenida en este informe procede exclusivamente de fuentes abiertas públicamente accesibles o consultadas mediante sesión autorizada del analista. Las capturas visuales adjuntas reflejan el estado de la fuente en la fecha y hora indicadas.
====================================================================
"""
        self.txt_informe.delete("1.0", tk.END)
        self.txt_informe.insert("1.0", plantilla)
        messagebox.showinfo("Informe Redactado", "El borrador se ha generado correctamente. Puedes editarlo antes de exportar.")

    def _obtener_contenido_informe(self):
        contenido = self.txt_informe.get("1.0", tk.END).strip()
        if not contenido:
            self.generar_borrador_informe()
            contenido = self.txt_informe.get("1.0", tk.END).strip()
        return contenido

    def _verificar_librerias(self):
        if not EXPORT_READY:
            messagebox.showerror("Faltan dependencias", "Ejecuta CREAR_EJECUTABLE_WINDOWS.bat para instalar las librerías necesarias.")
            return False
        return True

    def _construir_pdf(self, ruta, objetivo):
        texto_informe = self._obtener_contenido_informe()
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", "B", 14)
        pdf.cell(0, 10, "ZEUS - EXPEDIENTE DE INTELIGENCIA OSINT", ln=True, align="C")
        pdf.ln(3)
        pdf.set_font("Courier", "", 9)
        texto_limpio = texto_informe.encode('latin-1', 'replace').decode('latin-1')
        pdf.multi_cell(0, 5, texto_limpio)

        if self.capturas_evidencia:
            pdf.add_page()
            pdf.set_font("Arial", "B", 13)
            pdf.cell(0, 10, "ANEXO GRAFICO - EVIDENCIA VISUAL CAPTURADA", ln=True, align="C")
            pdf.ln(5)
            for idx, (ruta_img, desc, ts) in enumerate(self.capturas_evidencia, 1):
                if pdf.get_y() > 210:
                    pdf.add_page()
                pdf.set_font("Arial", "B", 10)
                titulo_fig = f"Figura {idx} [{ts}]: {desc}".encode('latin-1', 'replace').decode('latin-1')
                pdf.multi_cell(0, 6, titulo_fig)
                pdf.ln(2)
                pdf.image(ruta_img, w=165)
                pdf.ln(8)

        pdf.output(ruta)

    def _construir_word(self, ruta, objetivo):
        texto_informe = self._obtener_contenido_informe()
        doc = Document()
        doc.add_heading("ZEUS – EXPEDIENTE DE INTELIGENCIA OSINT", 0)
        doc.add_paragraph(f"Objetivo investigado: {objetivo} | Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M')}")

        doc.add_heading("Informe Ejecutivo y Hallazgos", level=1)
        doc.add_paragraph(texto_informe)

        if self.capturas_evidencia:
            doc.add_page_break()
            doc.add_heading("Anexo Gráfico: Evidencia Visual Capturada", level=1)
            for idx, (ruta_img, desc, ts) in enumerate(self.capturas_evidencia, 1):
                doc.add_paragraph(f"Figura {idx} [{ts}] – {desc}", style="Heading 2")
                doc.add_picture(ruta_img, width=Inches(6.0))
                doc.add_paragraph("")

        doc.save(ruta)

    def _construir_excel(self, ruta, objetivo):
        wb = openpyxl.Workbook()

        ws1 = wb.active
        ws1.title = "Enlaces y Fuentes"
        ws1.append(["Fecha/Hora", "Fuente", "URL / Evidencia"])
        for fuente, url, ts in self.resultados_sesion:
            ws1.append([ts, fuente, url])

        ws2 = wb.create_sheet(title="Cronología")
        ws2.append(["Fecha", "Evento / Hito"])
        for fecha, ev in self.cronologia_eventos:
            ws2.append([fecha, ev])

        ws3 = wb.create_sheet(title="Evidencia Visual")
        ws3.append(["Nº Figura", "Fecha/Hora", "Descripción", "Archivo Temporal"])
        for idx, (ruta_img, desc, ts) in enumerate(self.capturas_evidencia, 1):
            ws3.append([f"Figura {idx}", ts, desc, os.path.basename(ruta_img)])

        ws4 = wb.create_sheet(title="Informe Completo")
        ws4.append(["Objetivo", objetivo])
        ws4.append(["Contenido del Informe", self._obtener_contenido_informe()])

        wb.save(ruta)

    def exportar_pdf(self):
        if not self._verificar_librerias(): return
        objetivo = self.entry_objetivo.get().strip() or "Sin_Especificar"
        ruta = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=f"Informe_ZEUS_{objetivo}.pdf", filetypes=[("PDF", "*.pdf")])
        if not ruta: return
        self._construir_pdf(ruta, objetivo)
        messagebox.showinfo("Éxito", "Informe PDF con evidencia visual generado correctamente.")

    def exportar_word(self):
        if not self._verificar_librerias(): return
        objetivo = self.entry_objetivo.get().strip() or "Sin_Especificar"
        ruta = filedialog.asksaveasfilename(defaultextension=".docx", initialfile=f"Informe_ZEUS_{objetivo}.docx", filetypes=[("Word", "*.docx")])
        if not ruta: return
        self._construir_word(ruta, objetivo)
        messagebox.showinfo("Éxito", "Informe Word (.docx) con evidencia visual generado correctamente.")

    def exportar_excel(self):
        if not self._verificar_librerias(): return
        objetivo = self.entry_objetivo.get().strip() or "Sin_Especificar"
        ruta = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=f"Datos_ZEUS_{objetivo}.xlsx", filetypes=[("Excel", "*.xlsx")])
        if not ruta: return
        self._construir_excel(ruta, objetivo)
        messagebox.showinfo("Éxito", "Libro Excel (.xlsx) generado correctamente.")

    def exportar_todo_simultaneo(self):
        if not self._verificar_librerias(): return
        carpeta = filedialog.askdirectory(title="Selecciona la carpeta donde guardar los 3 archivos")
        if not carpeta: return

        objetivo = "".join(c for c in (self.entry_objetivo.get().strip() or "Objetivo") if c.isalnum() or c in ("_", "-"))
        marca = datetime.now().strftime("%Y%m%d_%H%M")

        ruta_pdf = os.path.join(carpeta, f"Informe_ZEUS_{objetivo}_{marca}.pdf")
        ruta_word = os.path.join(carpeta, f"Informe_ZEUS_{objetivo}_{marca}.docx")
        ruta_excel = os.path.join(carpeta, f"Datos_ZEUS_{objetivo}_{marca}.xlsx")

        self._construir_pdf(ruta_pdf, objetivo)
        self._construir_word(ruta_word, objetivo)
        self._construir_excel(ruta_excel, objetivo)

        messagebox.showinfo("Exportación Completa", f"Se han generado los 3 archivos con éxito en:\n{carpeta}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ZeusOSINTApp(root)
    root.mainloop()
