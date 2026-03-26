import db_manager
import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import webbrowser
import urllib.parse
import requests
import io
import os
import unicodedata

# Constantes de estados
ESTADO_NO_ENTREGADA = 'No Entregada'
ESTADO_LISTA = 'Entregada - Lista'
ESTADO_FALTA_REVISION = 'Entregada - Falta Revisión'

def eliminar_acentos(s):
    if not s or pd.isna(s): return ""
    return "".join(c for c in unicodedata.normalize('NFD', str(s)) if unicodedata.category(c) != 'Mn')

def normalizar_nombre(nombre):
    if not nombre or pd.isna(nombre): 
        return ""
    # Quitar acentos, convertir a minúsculas y quitar comas
    n = eliminar_acentos(str(nombre)).lower().replace(",", " ")
    # Dividir en palabras, limpiar y ordenar alfabéticamente para ignorar el orden
    palabras = sorted([p.strip() for p in n.split() if p.strip()])
    return " ".join(palabras)

class AppPlanificaciones(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Tablero de Control - Planificaciones 2026")
        self.geometry("1100x700")
        
        # Tema
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # Variables para los archivos
        self.archivo_materias = ""
        self.archivo_planificaciones = ""
        self.df_base = None
        self.df_mostrado = None
        self.modo_seleccion = False
        self.seleccionados = set() # Set de índices del dataframe
        self.vista_actual = "Materias" # "Materias" o "Docentes"
        db_manager.init_db()

        self._create_widgets()

    def _create_widgets(self):
        # Frame principal
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=2)
        self.grid_rowconfigure(0, weight=1)

        # Panel Izquierdo (Controles y Gráfico)
        self.left_panel = ctk.CTkFrame(self)
        self.left_panel.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # --- Controles ---
        ctk.CTkLabel(self.left_panel, text="Carga de Archivos", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=10)

        self.btn_planis = ctk.CTkButton(self.left_panel, text="Sincronizar SharePoint", command=self.sincronizar_sharepoint, fg_color="#ff8c00", hover_color="#e67e00")
        self.btn_planis.pack(pady=5, padx=10, fill="x")

        ctk.CTkLabel(self.left_panel, text="URL SharePoint:", font=ctk.CTkFont(size=12)).pack(pady=(5, 0))
        self.sharepoint_url_var = ctk.StringVar(value=db_manager.get_config("sharepoint_url"))
        self.ent_sharepoint_url = ctk.CTkEntry(self.left_panel, textvariable=self.sharepoint_url_var, placeholder_text="Pegar link de SharePoint...")
        self.ent_sharepoint_url.pack(pady=2, padx=10, fill="x")
        self.sharepoint_url_var.trace_add("write", lambda *args: db_manager.set_config("sharepoint_url", self.sharepoint_url_var.get()))

        self.btn_planis_local = ctk.CTkButton(self.left_panel, text="Cargar Local (Excel)", command=self.load_planificaciones, fg_color="#3b3b3b", hover_color="#2b2b2b")
        self.btn_planis_local.pack(pady=5, padx=10, fill="x")

        self.btn_procesar = ctk.CTkButton(self.left_panel, text="Cruzar Datos", command=self.procesar_datos, fg_color="green", hover_color="darkgreen")
        self.btn_procesar.pack(pady=(20, 5), padx=10, fill="x")

        self.btn_db = ctk.CTkButton(self.left_panel, text="BD Local: Materias/Docentes", command=self.abrir_gestor_db, fg_color="#6b4c9a", hover_color="#4f3873")
        self.btn_db.pack(pady=5, padx=10, fill="x")

        self.btn_vista_docentes = ctk.CTkButton(self.left_panel, text="Vista por Docentes (Ventana)", command=self.mostrar_vista_docentes, fg_color="#1f538d", hover_color="#14375e")
        # self.btn_vista_docentes.pack(pady=5, padx=10, fill="x") # La mantenemos comentada o la quitamos si ya no es necesaria

        # --- Botones de Correo ---
        self.container_correo = ctk.CTkFrame(self.left_panel, fg_color="transparent")
        self.container_correo.pack(pady=5, padx=10, fill="x")

        # self.btn_iniciar_correo = ctk.CTkButton(self.container_correo, text="✉ Enviar Correo", command=self.activar_modo_correo, fg_color="#d9534f", hover_color="#c9302c")
        # self.btn_iniciar_correo.pack(fill="x")

        self.frame_correo_opciones = ctk.CTkFrame(self.container_correo, fg_color="transparent")
        # Se pack-ea solo cuando se activa el modo

        self.btn_sel_todos = ctk.CTkButton(self.frame_correo_opciones, text="Seleccionar Todos", command=self.seleccionar_todos_visibles, font=ctk.CTkFont(size=11))
        self.btn_sel_todos.grid(row=0, column=0, columnspan=2, padx=2, pady=2, sticky="ew")

        self.btn_ok_correo = ctk.CTkButton(self.frame_correo_opciones, text="OK (Abrir Outlook)", command=self.confirmar_envio_correo, fg_color="green", font=ctk.CTkFont(size=11))
        self.btn_ok_correo.grid(row=1, column=0, padx=2, pady=2, sticky="ew")

        self.btn_cancel_correo = ctk.CTkButton(self.frame_correo_opciones, text="Cancelar", command=self.desactivar_modo_correo, fg_color="gray", font=ctk.CTkFont(size=11))
        self.btn_cancel_correo.grid(row=1, column=1, padx=2, pady=2, sticky="ew")
        
        self.frame_correo_opciones.grid_columnconfigure(0, weight=1)
        self.frame_correo_opciones.grid_columnconfigure(1, weight=1)

        # --- Filtros ---
        ctk.CTkLabel(self.left_panel, text="Filtros por Estado", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(10, 0))

        self.var_v = ctk.BooleanVar(value=True)
        self.var_a = ctk.BooleanVar(value=True)
        self.var_r = ctk.BooleanVar(value=True)

        self.chk_v = ctk.CTkCheckBox(self.left_panel, text="Listas (Verde)", variable=self.var_v, command=self.aplicar_filtros, fg_color="#335c33", hover_color="#244024")
        self.chk_v.pack(pady=2, padx=10, anchor="w")
        self.chk_a = ctk.CTkCheckBox(self.left_panel, text="Falta Rev (Amarillo)", variable=self.var_a, command=self.aplicar_filtros, fg_color="#8a6d26", hover_color="#614c1b")
        self.chk_a.pack(pady=2, padx=10, anchor="w")
        self.chk_r = ctk.CTkCheckBox(self.left_panel, text="No Entreg. (Rojo)", variable=self.var_r, command=self.aplicar_filtros, fg_color="#823434", hover_color="#5c2525")
        self.chk_r.pack(pady=2, padx=10, anchor="w")

        self.docente_var = ctk.StringVar()
        self.filtro_docente = ctk.CTkEntry(self.left_panel, placeholder_text="Buscar docente...", textvariable=self.docente_var)
        self.filtro_docente.pack(pady=5, padx=10, fill="x")
        self.filtro_docente.bind("<KeyRelease>", lambda e: self.aplicar_filtros())

        ctk.CTkLabel(self.left_panel, text="Coordinador", font=ctk.CTkFont(size=12)).pack(pady=(5, 0))
        self.coord_var = ctk.StringVar(value="Todos")
        self.filtro_coord = ctk.CTkOptionMenu(self.left_panel, variable=self.coord_var, values=["Todos", "Isi", "Ale", "Sin Asignar"], command=self.aplicar_filtros)
        self.filtro_coord.pack(pady=5, padx=10, fill="x")

        # --- Contenedor Gráfico ---
        self.graph_frame = ctk.CTkFrame(self.left_panel)
        self.graph_frame.pack(pady=20, padx=10, fill="both", expand=True)

        # Panel Derecho (Tabla y Counters)
        self.right_panel = ctk.CTkFrame(self)
        self.right_panel.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")

        # --- Contenedor Counters ---
        self.counters_frame = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        self.counters_frame.pack(pady=10, padx=10, fill="x")
        
        # Grid para alinear los 3 contadores
        self.counters_frame.grid_columnconfigure(0, weight=1)
        self.counters_frame.grid_columnconfigure(1, weight=1)
        self.counters_frame.grid_columnconfigure(2, weight=1)

        font_counter = ctk.CTkFont(size=18, weight="bold")

        # Verde
        self.frame_verdes = ctk.CTkFrame(self.counters_frame, fg_color="#335c33", corner_radius=8)
        self.frame_verdes.grid(row=0, column=0, padx=5, sticky="ew")
        self.lbl_verdes = ctk.CTkLabel(self.frame_verdes, text="Listas (Verde): 0", font=font_counter, text_color="white")
        self.lbl_verdes.pack(pady=10)

        # Amarillo
        self.frame_amarillos = ctk.CTkFrame(self.counters_frame, fg_color="#8a6d26", corner_radius=8)
        self.frame_amarillos.grid(row=0, column=1, padx=5, sticky="ew")
        self.lbl_amarillos = ctk.CTkLabel(self.frame_amarillos, text="Falta Rev (Amarillo): 0", font=font_counter, text_color="white")
        self.lbl_amarillos.pack(pady=10)

        # Rojo
        self.frame_rojos = ctk.CTkFrame(self.counters_frame, fg_color="#823434", corner_radius=8)
        self.frame_rojos.grid(row=0, column=2, padx=5, sticky="ew")
        self.lbl_rojos = ctk.CTkLabel(self.frame_rojos, text="No Entreg. (Rojo): 0", font=font_counter, text_color="white")
        self.lbl_rojos.pack(pady=10)

        # --- Selector de Vista ---
        self.view_selector_frame = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        self.view_selector_frame.pack(pady=10, padx=10, fill="x")
        
        ctk.CTkLabel(self.view_selector_frame, text="Vista Actual:", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=10)
        
        self.btn_switch_view = ctk.CTkSegmentedButton(self.view_selector_frame, values=["Materias", "Docentes"], command=self.cambiar_vista_principal)
        self.btn_switch_view.set("Materias")
        self.btn_switch_view.pack(side="left", padx=10)

        self.lbl_resultados = ctk.CTkLabel(self.right_panel, text="Resultados", font=ctk.CTkFont(size=16, weight="bold"))
        self.lbl_resultados.pack(pady=5)

        # Configurar Treeview (usando ttk porque CTk no tiene tabla integrada standard)
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background="#2b2b2b", foreground="white", fieldbackground="#2b2b2b", rowheight=25)
        style.map('Treeview', background=[('selected', '#1f538d')])

        # Container para Tablas y Scrollbar
        self.frame_tablas = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        self.frame_tablas.pack(fill="both", expand=True, padx=10, pady=5)

        # Scrollbar única para ambas tablas
        self.tree_scroll = ttk.Scrollbar(self.frame_tablas, orient="vertical")
        
        # Treeview Materias (Plano)
        self.tree = ttk.Treeview(self.frame_tablas, columns=("Sel", "Materia", "Docente", "Estado"), show="headings")
        self.tree.heading("Sel", text="[ ]")
        self.tree.heading("Materia", text="Materia")
        self.tree.heading("Docente", text="Docente")
        self.tree.heading("Estado", text="Estado")
        
        self.tree.column("Sel", width=30, minwidth=30, anchor="center")
        self.tree.column("Materia", minwidth=250, width=350)
        self.tree.column("Docente", minwidth=150, width=200)
        self.tree.column("Estado", minwidth=150, width=150)
        
        # Treeview Docentes (Jerárquico)
        self.tree_hier = ttk.Treeview(self.frame_tablas, columns=("Verdes", "Amarillos", "Rojos"))
        self.tree_hier.heading("#0", text="Docente / Materias")
        self.tree_hier.heading("Verdes", text="Listas")
        self.tree_hier.heading("Amarillos", text="Falta Rev")
        self.tree_hier.heading("Rojos", text="No Entreg.")
        
        self.tree_hier.column("#0", width=400, minwidth=250)
        self.tree_hier.column("Verdes", width=100, anchor="center")
        self.tree_hier.column("Amarillos", width=100, anchor="center")
        self.tree_hier.column("Rojos", width=100, anchor="center")

        # Eventos de clic en la tabla para selección
        self.tree.bind("<ButtonRelease-1>", self.on_tree_click)
        self.tree_hier.bind("<ButtonRelease-1>", self.on_tree_click)

        # Configurar scroll
        self.tree.configure(yscrollcommand=self.tree_scroll.set)
        self.tree_hier.configure(yscrollcommand=self.tree_scroll.set)
        self.tree_scroll.configure(command=self.on_scroll_sync)
        
        # Inicialmente mostramos materias en el container
        self.tree.pack(fill="both", expand=True, side="left")
        self.tree_scroll.pack(side="right", fill="y")
        
        # --- Botones de Acción para Docentes (Panel Principal) ---
        self.frame_acciones_docentes = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        # Se mostrará solo en vista docentes
        
        self.btn_email_doc = ctk.CTkButton(self.frame_acciones_docentes, text="📧 Enviar Email al Docente", 
                                            command=self.enviar_email_docente, fg_color="#3b5998", hover_color="#2d4373")
        self.btn_email_doc.pack(side="left", padx=10)
        
        self.btn_ws_doc = ctk.CTkButton(self.frame_acciones_docentes, text="📋 Copiar Mensaje (WhatsApp)", 
                                         command=self.copiar_mensaje_docente, fg_color="#25D366", hover_color="#128C7E", text_color="black")
        self.btn_ws_doc.pack(side="left", padx=10)

        # Configurar ordenamiento para la tabla de resultados
        for col in ("Materia", "Docente", "Estado"):
            self.tree.heading(col, text=col, command=lambda _col=col: self.treeview_sort_column(self.tree, _col, False))

    def treeview_sort_column(self, tv, col, reverse):
        """Ordena el contenido de una columna en un Treeview."""
        l = [(tv.set(k, col), k) for k in tv.get_children('')]
        
        # Intentar ordenar numéricamente si es posible
        try:
            l.sort(key=lambda t: float(t[0]), reverse=reverse)
        except ValueError:
            l.sort(reverse=reverse)

        # Reordenar items
        for index, (val, k) in enumerate(l):
            tv.move(k, '', index)

        # Cambiar el comando para la próxima vez (alternar orden)
        tv.heading(col, command=lambda: self.treeview_sort_column(tv, col, not reverse))

    def sincronizar_sharepoint(self):
        url = self.sharepoint_url_var.get().strip()
        if not url:
            messagebox.showwarning("Atención", "Por favor, ingresa una URL de SharePoint primero.")
            return

        # Transformación básica para SharePoint (intentar forzar descarga)
        # Si es un link de edición de Office Online, a menudo funciona agregar &download=1
        download_url = url
        if "sharepoint.com" in url and "download=1" not in url:
            if "?" in url:
                download_url = url + "&download=1"
            else:
                download_url = url + "?download=1"

        try:
            self.btn_planis.configure(state="disabled", text="Sincronizando...")
            self.update()
            
            response = requests.get(download_url, timeout=30)
            response.raise_for_status()
            
            # Guardar temporalmente para que pandas lo lea o usar BytesIO
            data = io.BytesIO(response.content)
            
            # Verificar si es un Excel válido
            pd.read_excel(data, nrows=1) 
            
            # Si llegó aquí, es válido. Guardamos en memoria como si fuera un archivo cargado
            self.archivo_planificaciones = data
            messagebox.showinfo("Éxito", "Planificaciones sincronizadas correctamente desde SharePoint.")
            
        except Exception as e:
            messagebox.showerror("Error de Sincronización", 
                f"No se pudo descargar el archivo.\n\nDetalles: {str(e)}\n\n"
                "Asegúrate de que el link sea público (Cualquiera con el link -> Ver únicamente).")
        finally:
            self.btn_planis.configure(state="normal", text="Sincronizar SharePoint")

    def load_planificaciones(self):
        filepath = filedialog.askopenfilename(title="Seleccionar Planificaciones", filetypes=[("Excel files", "*.xlsx *.xls")])
        if filepath:
            self.archivo_planificaciones = filepath
            messagebox.showinfo("Archivo Cargado", f"Planificaciones: {os.path.basename(filepath)}")

    def procesar_datos(self):
        if not self.archivo_planificaciones:
            messagebox.showwarning("Advertencia", "Por favor seleccione el archivo de Planificaciones antes de procesar.")
            return

        try:
            # 1. Leer las materias (Nombres exactos y docentes si existen) desde la BD
            rels = db_manager.obtener_materias_con_docentes()
            if not rels:
                messagebox.showwarning("Advertencia", "Debe ingresar las Materias en la BD Local (o importarlas desde un Excel) primero.")
                return
            
            filas_mat = []
            for m_nombre, d_list in rels.items():
                docs_str = " - ".join(d_list) if d_list else "Desconocido"
                filas_mat.append({"Materia_Base": m_nombre, "Materia_Clean": m_nombre.lower().strip(), "Docentes_Asignados": docs_str})
            
            df_mat = pd.DataFrame(filas_mat)

            # 2. Leer Planificaciones Anuales
            if hasattr(self.archivo_planificaciones, 'seek'):
                self.archivo_planificaciones.seek(0)
            df_plan = pd.read_excel(self.archivo_planificaciones)
            
            # Limpieza y preparación de planificaciones
            if 'NOMBRE DE LA MATERIA' in df_plan.columns:
                df_plan['Materia_Plan'] = df_plan['NOMBRE DE LA MATERIA'].astype(str).str.strip().str.lower()
            else:
                messagebox.showerror("Error", "No se encontró la columna 'NOMBRE DE LA MATERIA' en las planificaciones.")
                return

            if 'DOCENTE TITULAR' not in df_plan.columns:
                df_plan['DOCENTE TITULAR'] = "Desconocido"
                
            col_revision = 'La planificación, ¿está lista para revisión del coordinador?'
            if col_revision not in df_plan.columns:
                # Intento buscar con posibles variaciones de espacios
                found = False
                for c in df_plan.columns:
                    if 'revisión del coordinador' in str(c).lower():
                        col_revision = c
                        found = True
                        break
                if not found:
                    df_plan[col_revision] = "Desconocido"

            # Buscar duplicados en las planificaciones
            duplicados = df_plan[df_plan.duplicated(subset=['Materia_Plan'], keep=False)]
            if not duplicados.empty:
                materias_dup = duplicados['NOMBRE DE LA MATERIA'].dropna().unique()
                if len(materias_dup) > 0:
                    lista_m = "\n- ".join(materias_dup)
                    messagebox.showinfo("Múltiples Entregas Detectadas", 
                                        f"Las siguientes materias tienen más de una planificación entregada:\n\n- {lista_m}\n\nEl sistema tomará la última entregada (por hora de inicio).")

            # 3. Cruce de datos: Left Join desde las materias del formulario
            # Como pueden haber varias filas por materia si hubo multiples envios, tomaremos el más reciente.
            if 'Hora de inicio' in df_plan.columns:
                df_plan = df_plan.sort_values('Hora de inicio').drop_duplicates(subset=['Materia_Plan'], keep='last')
            else:
                df_plan = df_plan.drop_duplicates(subset=['Materia_Plan'], keep='last')

            df_merged = pd.merge(df_mat, df_plan, left_on='Materia_Clean', right_on='Materia_Plan', how='left')

            # 4. Determinar Estado
            def obtener_estado(row):
                if pd.isna(row['Materia_Plan']):
                    return ESTADO_NO_ENTREGADA
                
                revision_val = str(row[col_revision]).lower().strip()
                if revision_val == 'sí' or revision_val == 'si':
                    return ESTADO_LISTA
                else:
                    return ESTADO_FALTA_REVISION

            df_merged['Estado'] = df_merged.apply(obtener_estado, axis=1)
            
            # Asignar Docente Final (raw, antes de unificar)
            def obtener_docente_final(row):
                if row['Estado'] == ESTADO_NO_ENTREGADA:
                    val = row['Docentes_Asignados']
                    return val if pd.notna(val) else "Sin Docente Asignado"
                else:
                    val = row['DOCENTE TITULAR']
                    return val if pd.notna(val) else "Sin Docente"

            df_merged['Docente_Raw'] = df_merged.apply(obtener_docente_final, axis=1)

            # 5. Unificar nombres de docentes usando la BD como referencia
            docentes_db = db_manager.obtener_todos_docentes()
            # Mapa de nombre_normalizado -> Nombre Oficial (d[1] es el nombre en la BD)
            map_oficial = {normalizar_nombre(d[1]): d[1] for d in docentes_db}

            def unificar_docentes(doc_str):
                if not doc_str or pd.isna(doc_str):
                    return "Sin Docente"
                partes = [p.strip() for p in str(doc_str).split('-')]
                unificados = []
                for p in partes:
                    p_norm = normalizar_nombre(p)
                    if p_norm in map_oficial:
                        unificados.append(map_oficial[p_norm])
                    else:
                        unificados.append(p)
                return " - ".join(unificados)

            df_merged['Docente'] = df_merged['Docente_Raw'].apply(unificar_docentes)
            
            # Guardamos la base procesada (Materia Original para visualización)
            self.df_base = df_merged[['Materia_Base', 'Docente', 'Estado']].copy()
            self.df_base.rename(columns={'Materia_Base': 'Materia'}, inplace=True)
            
            self.actualizar_contadores_globales()
            self.aplicar_filtros()
            messagebox.showinfo("Éxito", "Datos cruzados y nombres unificados correctamente.")

        except Exception as e:
            messagebox.showerror("Error al procesar", f"Ocurrió un error: {str(e)}")

    def actualizar_contadores_globales(self):
        if self.df_base is None:
            return
            
        conteo = self.df_base['Estado'].value_counts()
        
        verdes = conteo.get(ESTADO_LISTA, 0)
        amarillos = conteo.get(ESTADO_FALTA_REVISION, 0)
        rojos = conteo.get(ESTADO_NO_ENTREGADA, 0)
        
        self.lbl_verdes.configure(text=f"Listas (Verde): {verdes}")
        self.lbl_amarillos.configure(text=f"Falta Rev (Amarillo): {amarillos}")
        self.lbl_rojos.configure(text=f"No Entreg. (Rojo): {rojos}")

    def aplicar_filtros(self, *args):
        if self.df_base is None:
            return

        docente = self.docente_var.get().lower().strip()
        coord_f = self.coord_var.get()
        
        estados_permitidos = []
        if self.var_v.get(): estados_permitidos.append(ESTADO_LISTA)
        if self.var_a.get(): estados_permitidos.append(ESTADO_FALTA_REVISION)
        if self.var_r.get(): estados_permitidos.append(ESTADO_NO_ENTREGADA)

        df_filtered = self.df_base.copy()

        # Filtro de Estado
        df_filtered = df_filtered[df_filtered['Estado'].isin(estados_permitidos)]
        
        if docente:
            docente_norm = eliminar_acentos(docente).lower()
            mask = df_filtered['Docente'].apply(lambda d: docente_norm in eliminar_acentos(d).lower())
            df_filtered = df_filtered[mask]

        # Filtro de Coordinador (necesitamos cruzar con la BD)
        if coord_f != "Todos":
            docentes_db = db_manager.obtener_todos_docentes()
            # Usamos el nombre normalizado como clave para el mapa
            map_coord = {normalizar_nombre(d[1]): (d[3] if d[3] else "") for d in docentes_db}
            
            def check_coord(doc_str):
                # El docente puede ser "Doc 1 - Doc 2"
                nombres_raw = [n.strip() for n in str(doc_str).split('-')]
                for n_raw in nombres_raw:
                    n_norm = normalizar_nombre(n_raw)
                    if n_norm in map_coord:
                        c = map_coord[n_norm]
                        if coord_f == "Sin Asignar" and (c == "" or c is None):
                            return True
                        if c == coord_f:
                            return True
                return False

            mask = df_filtered['Docente'].apply(check_coord)
            df_filtered = df_filtered[mask]

        self.df_mostrado = df_filtered

        self.actualizar_tabla()
        self.actualizar_grafico()

    def activar_modo_correo(self):
        self.modo_seleccion = True
        self.seleccionados = set()
        self.btn_iniciar_correo.pack_forget()
        self.frame_correo_opciones.pack(pady=5, padx=10, fill="x")
        self.actualizar_tabla()

    def desactivar_modo_correo(self):
        self.modo_seleccion = False
        self.seleccionados = set()
        self.frame_correo_opciones.pack_forget()
        self.btn_iniciar_correo.pack(fill="x")
        self.actualizar_tabla()

    def seleccionar_todos_visibles(self):
        if self.df_mostrado is not None:
            self.seleccionados = set(self.df_mostrado.index)
            self.actualizar_tabla()

    def on_tree_click(self, event):
        if not self.modo_seleccion:
            return
        
        # Identificar en qué tabla se hizo clic
        tv = event.widget
        item_id = tv.identify_row(event.y)
        if not item_id:
            return
        
        # IMPORTANTE: Ignorar clics en el icono de expansión (+/-)
        # El elemento 'treeindicator' es el encargado de abrir/cerrar
        element = tv.identify_element(event.x, event.y)
        if element == "treeindicator":
            return
        
        tags = tv.item(item_id, "tags")
        if not tags:
            return

        tag_val = tags[0]
        
        if tag_val.startswith("docente:"):
            # Es un nodo de docente en la vista jerárquica
            doc_name = tag_val.replace("docente:", "")
            # Encontrar todos los índices de materias de este docente
            indices = self.df_mostrado[self.df_mostrado['Docente'] == doc_name].index
            
            # Si todos están seleccionados, deseleccionar. Si no, seleccionar todos.
            if all(idx in self.seleccionados for idx in indices):
                for idx in indices:
                    self.seleccionados.discard(idx)
            else:
                for idx in indices:
                    self.seleccionados.add(idx)
        else:
            # Es un índice de materia (tag numérico o materia en jerarquía)
            try:
                idx = int(tag_val)
                if idx in self.seleccionados:
                    self.seleccionados.remove(idx)
                else:
                    self.seleccionados.add(idx)
            except ValueError:
                pass
                
        self.actualizar_tabla()

    def confirmar_envio_correo(self):
        if not self.seleccionados:
            messagebox.showwarning("Atención", "No ha seleccionado ninguna materia/docente.")
            return
        
        # Obtener emails de la BD
        docentes_db = db_manager.obtener_todos_docentes()
        # Mapeo nombre -> email
        emails_dict = {d[1]: d[2] for d in docentes_db if d[2] and d[2].strip()}
        
        correos_destino = set()
        docentes_sin_email = set()

        for idx in self.seleccionados:
            docente_nombre = self.df_base.loc[idx, 'Docente']
            # El nombre puede venir como "Doc 1 - Doc 2" si hay varios
            nombres = [n.strip() for n in docente_nombre.split('-')]
            for n in nombres:
                if n in emails_dict:
                    correos_destino.add(emails_dict[n])
                elif n not in ["Desconocido", "Sin Docente Asignado", "Sin Docente"]:
                    docentes_sin_email.add(n)

        if not correos_destino:
            messagebox.showerror("Error", "Ninguno de los docentes seleccionados tiene un email registrado en la BD Local.")
            return

        if docentes_sin_email:
            lista = "\n- ".join(list(docentes_sin_email)[:10])
            if len(docentes_sin_email) > 10: lista += "\n...y otros."
            res = messagebox.askyesno("Docentes sin Email", 
                                       f"Los siguientes docentes seleccionados no tienen email en la BD:\n{lista}\n\n¿Desea continuar enviando a los que sí tienen correo?")
            if not res: return

        # Abrir Outlook / Mailto
        emails_str = ",".join(correos_destino)
        # Mailto: se puede abrir con el navegador o shell
        subject = urllib.parse.quote("Consulta Planificación 2026")
        body = urllib.parse.quote("Hola, le escribo para consultarle sobre la planificación...")
        mailto_url = f"mailto:{emails_str}?subject={subject}&body={body}"
        
        try:
            webbrowser.open(mailto_url)
            self.desactivar_modo_correo()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el cliente de correo: {e}")

    def cambiar_vista_principal(self, view):
        self.vista_actual = view
        if view == "Materias":
            self.tree_hier.pack_forget()
            self.frame_acciones_docentes.pack_forget()
            self.tree.pack(fill="both", expand=True, side="left")
            self.tree_scroll.configure(command=self.tree.yview)
        else:
            self.tree.pack_forget()
            self.tree_hier.pack(fill="both", expand=True, side="left")
            self.frame_acciones_docentes.pack(fill="x", padx=10, pady=10)
            self.tree_scroll.configure(command=self.tree_hier.yview)
        
        self.actualizar_tabla()

    def on_scroll_sync(self, *args):
        if self.vista_actual == "Materias":
            self.tree.yview(*args)
        else:
            self.tree_hier.yview(*args)

    def actualizar_tabla(self):
        if self.vista_actual == "Materias":
            self.actualizar_tabla_materias()
        else:
            self.actualizar_tabla_docentes()

    def actualizar_tabla_materias(self):
        # Limpiar tabla materias
        for item in self.tree.get_children():
            self.tree.delete(item)

        if self.df_mostrado is not None:
            for index, row in self.df_mostrado.iterrows():
                sel_mark = "[X]" if index in self.seleccionados else "[ ]"
                if not self.modo_seleccion:
                    sel_mark = "-"
                
                self.tree.insert("", "end", values=(sel_mark, row['Materia'], row['Docente'], row['Estado']), tags=(str(index),))

    def actualizar_tabla_docentes(self):
        # Guardar qué docentes estaban expandidos para restaurar el estado
        expandidos = set()
        if self.tree_hier.get_children():
            for item_id in self.tree_hier.get_children():
                # Obtenemos el nombre del docente del tag para identificarlo de forma única
                tags = self.tree_hier.item(item_id, "tags")
                if tags and tags[0].startswith("docente:"):
                    doc_name = tags[0].replace("docente:", "")
                    if self.tree_hier.item(item_id, "open"):
                        expandidos.add(doc_name)

        # Limpiar tabla jerárquica
        for item in self.tree_hier.get_children():
            self.tree_hier.delete(item)

        if self.df_mostrado is not None and not self.df_mostrado.empty:
            agrupado = self.df_mostrado.groupby('Docente')
            for docente, df_doc in agrupado:
                conteo = df_doc['Estado'].value_counts()
                v = conteo.get(ESTADO_LISTA, 0)
                a = conteo.get(ESTADO_FALTA_REVISION, 0)
                r = conteo.get(ESTADO_NO_ENTREGADA, 0)
                
                # Determinar marca de selección para el docente
                sel_mark = ""
                if self.modo_seleccion:
                    indices = df_doc.index
                    if all(idx in self.seleccionados for idx in indices) and not df_doc.empty:
                        sel_mark = "[X] "
                    else:
                        sel_mark = "[ ] "
                
                # Restaurar estado de expansión
                is_open = docente in expandidos
                
                doc_id = self.tree_hier.insert("", "end", text=f"{sel_mark}{docente}", values=(v, a, r), open=is_open, tags=(f"docente:{docente}",))
                
                for idx, row in df_doc.iterrows():
                    est = row['Estado']
                    mat = row['Materia']
                    indicador = "🔴 "
                    if est == ESTADO_LISTA: indicador = "🟢 "
                    elif est == ESTADO_FALTA_REVISION: indicador = "🟡 "
                    
                    sel_mat = ""
                    if self.modo_seleccion:
                        sel_mat = "[X] " if idx in self.seleccionados else "[ ] "
                    
                    self.tree_hier.insert(doc_id, "end", text=f"{sel_mat}{indicador} {mat}  ({est})", values=("", "", ""), tags=(str(idx),))

    def _get_data_docente_seleccionado(self):
        sel = self.tree_hier.selection()
        if not sel:
            return None
        item = self.tree_hier.item(sel[0])
        tags = item.get("tags", [])
        if not tags or not tags[0].startswith("docente:"):
            # Podría ser una materia, intentar subir al padre
            parent = self.tree_hier.parent(sel[0])
            if parent:
                item = self.tree_hier.item(parent)
                tags = item.get("tags", [])
            else:
                return None
        
        if not tags or not tags[0].startswith("docente:"):
            return None
            
        docente = tags[0].replace("docente:", "")
        
        # Obtener materias y estados del df_base para este docente
        df_doc = self.df_base[self.df_base['Docente'] == docente]
        materias = []
        for _, row in df_doc.iterrows():
            materias.append((row['Materia'], row['Estado']))
            
        # Buscar email
        docentes_db = db_manager.obtener_todos_docentes()
        n_norm = normalizar_nombre(docente)
        email = ""
        for d in docentes_db:
            if normalizar_nombre(d[1]) == n_norm:
                email = d[2] if d[2] else ""
                break
        
        return {'docente': docente, 'email': email, 'materias': materias}

    def generar_mensaje_docente(self, data):
        docente = data['docente']
        materias = data['materias']
        
        cuerpo = f"{docente},\n\nMe pongo en contacto contigo para actualizarte sobre el estado de recepción de tus planificaciones.\n\nA continuación, se detalla el estado actual de tus materias:\n\n"
        
        items_detalle = []
        tiene_lista = tiene_rev = tiene_urg = False
        
        for m, est in materias:
            est_msg = ""
            if est == ESTADO_LISTA:
                est_msg = "Entregada (Final)"
                tiene_lista = True
            elif est == ESTADO_FALTA_REVISION:
                est_msg = "Entregada para revisión"
                tiene_rev = True
            else:
                est_msg = "🚨 URGENTE (No entregada)"
                tiene_urg = True
            items_detalle.append(f"{m}: {est_msg}")
        
        cuerpo += "\n".join(items_detalle)
        cuerpo += "\n\nObservaciones según el estado:\n"
        
        if tiene_lista:
            cuerpo += "Entregada (Final): ¡Muchas gracias por tu compromiso y por cumplir con los plazos establecidos! Es un gran aporte para la organización del ciclo lectivo. ¡Felicitaciones!\n\n"
        if tiene_rev:
            cuerpo += "Entregada para revisión: Recibida correctamente. En la brevedad te estaré enviando las devoluciones o sugerencias pertinentes para finalizar el proceso.\n\n"
        if tiene_urg:
            cuerpo += "🚨 URGENTE (No entregada): No contamos con registro de entrega de este documento. Por favor, regulariza esta situación a la brevedad o comunícate conmigo si has tenido algún inconveniente técnico o personal para la carga.\n\n"
        
        cuerpo += "Quedamos a tu entera disposición para lo que necesites o para despejar cualquier duda técnica sobre el formato de entrega.\n\nAtentamente,\nIsi, Ale y Luis"
        return cuerpo

    def enviar_email_docente(self):
        data = self._get_data_docente_seleccionado()
        if not data:
            messagebox.showwarning("Aviso", "Seleccione un docente en la lista.")
            return
            
        if not data['email']:
            messagebox.showwarning("Falta Email", f"El docente {data['docente']} no tiene email registrado en la base de datos.")
            return
            
        cuerpo = self.generar_mensaje_docente(data)
        asunto = "Estado de Planificaciones Docentes"
        
        para = data['email']
        cc = "atombesi@etrr.edu.ar, lcornaglia@etrr.edu.ar"
        
        # Respaldo: Copiar al portapapeles por si Outlook falla (limitación de la "Nueva Outlook")
        self.clipboard_clear()
        self.clipboard_append(cuerpo)
        self.update()
        
        mailto_url = f"mailto:{para}?cc={urllib.parse.quote(cc)}&subject={urllib.parse.quote(asunto)}&body={urllib.parse.quote(cuerpo)}"
        
        try:
            webbrowser.open(mailto_url)
            messagebox.showinfo("Email y Respaldo", 
                                f"Se ha intentado abrir el correo para {data['docente']}.\n\n"
                                "NOTA: Si Outlook no se abre o da error, el mensaje ya ha sido COPIADO al portapapeles. "
                                "Solo tienes que crear el correo manualmente y presionar PEGAR.")
        except Exception as e:
            messagebox.showwarning("Error al abrir Correo", 
                                    f"No se pudo abrir el cliente de correo automáticamente.\n\n"
                                    "Sin embargo, el mensaje ha sido COPIADO al portapapeles. "
                                    "Puedes pegarlo manualmente en tu programa de correo.")

    def copiar_mensaje_docente(self):
        data = self._get_data_docente_seleccionado()
        if not data:
            messagebox.showwarning("Aviso", "Seleccione un docente en la lista.")
            return
            
        cuerpo = self.generar_mensaje_docente(data)
        self.clipboard_clear()
        self.clipboard_append(cuerpo)
        self.update()
        messagebox.showinfo("Copiado", f"Mensaje para {data['docente']} copiado al portapapeles.")

    def actualizar_grafico(self):
        # Limpiar frame del gráfico
        for widget in self.graph_frame.winfo_children():
            widget.destroy()

        if self.df_mostrado is None or self.df_mostrado.empty:
            return

        conteo = self.df_mostrado['Estado'].value_counts()
        labels = conteo.index.tolist()
        sizes = conteo.values.tolist()

        # Colores personalizados
        colores = {
            ESTADO_NO_ENTREGADA: '#d9534f', # Rojo
            ESTADO_LISTA: '#5cb85c', # Verde
            ESTADO_FALTA_REVISION: '#f0ad4e' # Naranja
        }
        cores_lista = [colores.get(l, '#999999') for l in labels]


        fig, ax = plt.subplots(figsize=(4, 4), dpi=100)
        fig.patch.set_facecolor('#2b2b2b') # Fondo oscuro igual a CTk
        ax.set_facecolor('#2b2b2b')

        wedges, texts, autotexts = ax.pie(sizes, labels=labels, colors=cores_lista, autopct='%1.1f%%', startangle=140, 
                                          textprops=dict(color="w", fontsize=9))
        
        ax.axis('equal')  # Círculo perfecto

        # Integrar en Tkinter
        canvas = FigureCanvasTkAgg(fig, master=self.graph_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    def mostrar_vista_docentes(self):
        if self.df_base is None or self.df_base.empty:
            messagebox.showinfo("Sin Datos", "Primero debes cruzar los datos.")
            return

        agrupado = self.df_base.groupby('Docente')

        top = ctk.CTkToplevel(self)
        top.title("Vista por Docentes")
        top.geometry("900x650")
        top.grab_set() # Focus en esta ventana

        lbl_titulo = ctk.CTkLabel(top, text="Estado de Entregas por Docente", font=ctk.CTkFont(size=18, weight="bold"))
        lbl_titulo.pack(pady=10)

        # Frame superior para filtros
        filtros_frame = ctk.CTkFrame(top, fg_color="transparent")
        filtros_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(filtros_frame, text="Mostrar:", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=(0, 10))
        
        var_verde = ctk.BooleanVar(value=True)
        var_amarillo = ctk.BooleanVar(value=True)
        var_rojo = ctk.BooleanVar(value=True)
        
        chk_verde = ctk.CTkCheckBox(filtros_frame, text="Listas (Verde)", variable=var_verde, fg_color="#335c33", hover_color="#244024", command=lambda: cargar_arbol())
        chk_verde.pack(side="left", padx=5)
        
        chk_amarillo = ctk.CTkCheckBox(filtros_frame, text="Falta Rev (Amarillo)", variable=var_amarillo, fg_color="#8a6d26", hover_color="#614c1b", command=lambda: cargar_arbol())
        chk_amarillo.pack(side="left", padx=5)
        
        chk_rojo = ctk.CTkCheckBox(filtros_frame, text="No Entreg. (Rojo)", variable=var_rojo, fg_color="#823434", hover_color="#5c2525", command=lambda: cargar_arbol())
        chk_rojo.pack(side="left", padx=5)

        ctk.CTkLabel(filtros_frame, text="Coordinador:").pack(side="left", padx=(20, 5))
        self.var_coord_filtro = ctk.StringVar(value="Todos")
        self.cmb_coord_filtro = ctk.CTkOptionMenu(filtros_frame, values=["Todos", "Isi", "Ale", "Sin Asignar"], 
                                                 variable=self.var_coord_filtro, command=lambda _: cargar_arbol(), width=120)
        self.cmb_coord_filtro.pack(side="left", padx=5)

        # Variables de estado para ordenamiento
        self.sort_col = None
        self.sort_rev = False

        # Container para el Treeview y Scrollbar
        frame_tree = ctk.CTkFrame(top, fg_color="transparent")
        frame_tree.pack(fill="both", expand=True, padx=10, pady=5)

        # Treeview para mostrar la jerarquía
        columns = ("Verdes", "Amarillos", "Rojos")
        tree_doc = ttk.Treeview(frame_tree, columns=columns)
        tree_doc.heading("#0", text="Docente / Materias", command=lambda: sort_tree("#0"))
        tree_doc.heading("Verdes", text="Listas (Verde)", command=lambda: sort_tree("Verdes"))
        tree_doc.heading("Amarillos", text="Falta Rev (Amarillo)", command=lambda: sort_tree("Amarillos"))
        tree_doc.heading("Rojos", text="No Entreg. (Rojo)", command=lambda: sort_tree("Rojos"))

        tree_doc.column("#0", width=400, minwidth=250)
        tree_doc.column("Verdes", width=120, anchor="center")
        tree_doc.column("Amarillos", width=140, anchor="center")
        tree_doc.column("Rojos", width=140, anchor="center")

        # Scrollbar
        scroll = ttk.Scrollbar(frame_tree, orient="vertical", command=tree_doc.yview)
        tree_doc.configure(yscrollcommand=scroll.set)
        
        tree_doc.pack(fill="both", expand=True, side="left")
        scroll.pack(side="right", fill="y")

        # Obtener coordinadores de la BD para mapeo
        docentes_db = db_manager.obtener_todos_docentes()
        # Mapa de nombre_normalizado -> (coordinador, email)
        map_info_doc = {normalizar_nombre(d[1]): (d[3] if d[3] else "", d[2] if d[2] else "") for d in docentes_db}

        # Guardar la data para ordenar
        datos_docentes = []
        for docente, df_doc in agrupado:
            conteo = df_doc['Estado'].value_counts()
            v = conteo.get(ESTADO_LISTA, 0)
            a = conteo.get(ESTADO_FALTA_REVISION, 0)
            r = conteo.get(ESTADO_NO_ENTREGADA, 0)
            materias = []
            for _, row in df_doc.iterrows():
                materias.append((row['Materia'], row['Estado']))
            
            n_norm = normalizar_nombre(docente)
            if n_norm in map_info_doc:
                coord, email = map_info_doc[n_norm]
            else:
                coord, email = "N/A", ""

            datos_docentes.append({'docente': docente, 'email': email, 'v': v, 'a': a, 'r': r, 'materias': materias, 'coordinador': coord})

        def cargar_arbol():
            # Limpiar árbol
            for item in tree_doc.get_children():
                tree_doc.delete(item)
                
            mostrar_v = var_verde.get()
            mostrar_a = var_amarillo.get()
            mostrar_r = var_rojo.get()
            filtro_c = self.var_coord_filtro.get()
            
            # Ordenar datos si hay sort activo
            if self.sort_col:
                if self.sort_col == "#0":
                    datos_docentes.sort(key=lambda x: str(x['docente']).lower(), reverse=self.sort_rev)
                elif self.sort_col == "Verdes":
                    datos_docentes.sort(key=lambda x: x['v'], reverse=self.sort_rev)
                elif self.sort_col == "Amarillos":
                    datos_docentes.sort(key=lambda x: x['a'], reverse=self.sort_rev)
                elif self.sort_col == "Rojos":
                    datos_docentes.sort(key=lambda x: x['r'], reverse=self.sort_rev)
            
            for data in datos_docentes:
                docente = data['docente']
                coord = data.get('coordinador', "")
                v, a, r = data['v'], data['a'], data['r']
                
                # Filtro por Coordinador
                if filtro_c != "Todos":
                    if filtro_c == "Sin Asignar":
                        if coord != "" or coord == "N/A": continue
                    elif filtro_c != coord:
                        continue
                
                # Filtrar si el docente tiene algo que mostrar según los checkboxes globales
                # Para un docente, vemos si tiene al menos 1 de los estados chequeados
                tiene_v = v > 0 and mostrar_v
                tiene_a = a > 0 and mostrar_a
                tiene_r = r > 0 and mostrar_r
                
                # Si no tiene ninguno de los seleccionados, lo omitimos completamente en la vista (opcional)
                # O si quieres mostrarlo igual, quita este if. Por pedido "filtros", se asume excluir los que no tienen nada
                if not (tiene_v or tiene_a or tiene_r):
                    if not (v == 0 and a == 0 and r == 0): # si todo es 0, mostrar igual
                        continue

                doc_id = tree_doc.insert("", "end", text=f"{docente}", values=(v, a, r), open=False)

                # Insertar Materias filtradas
                for mat, est in data['materias']:
                    if est == ESTADO_LISTA and mostrar_v:
                        indicador = "🟢 "
                        tree_doc.insert(doc_id, "end", text=f"{indicador} {mat}  ({est})", values=("", "", ""))
                    elif est == ESTADO_FALTA_REVISION and mostrar_a:
                        indicador = "🟡 "
                        tree_doc.insert(doc_id, "end", text=f"{indicador} {mat}  ({est})", values=("", "", ""))
                    elif est == ESTADO_NO_ENTREGADA and mostrar_r:
                        indicador = "🔴 "
                        tree_doc.insert(doc_id, "end", text=f"{indicador} {mat}  ({est})", values=("", "", ""))

        def sort_tree(col):
            if self.sort_col == col:
                self.sort_rev = not self.sort_rev
            else:
                self.sort_rev = True # Por defecto descendente (mayor cantidad primero) al clickear numero, o Z-A para nombre
                self.sort_col = col
            cargar_arbol()

        def generar_mensaje_docente(data):
            docente = data['docente']
            materias = data['materias']
            
            cuerpo = f"{docente},\n\nEspero que estés teniendo una excelente semana. Me pongo en contacto contigo para actualizarte sobre el estado de recepción de tus planificaciones.\n\nA continuación, se detalla el estado actual de tus materias:\n\n"
            
            items_detalle = []
            tiene_lista = tiene_rev = tiene_urg = False
            
            for m, est in materias:
                est_msg = ""
                if est == ESTADO_LISTA:
                    est_msg = "Entregada (Final)"
                    tiene_lista = True
                elif est == ESTADO_FALTA_REVISION:
                    est_msg = "Entregada para revisión"
                    tiene_rev = True
                else:
                    est_msg = "🚨 URGENTE (No entregada)"
                    tiene_urg = True
                items_detalle.append(f"{m}: {est_msg}")
            
            cuerpo += "\n".join(items_detalle)
            cuerpo += "\n\nObservaciones según el estado:\n"
            
            if tiene_lista:
                cuerpo += "Entregada (Final): ¡Muchas gracias por tu compromiso y por cumplir con los plazos establecidos! Es un gran aporte para la organización del ciclo lectivo. ¡Felicitaciones!\n\n"
            if tiene_rev:
                cuerpo += "Entregada para revisión: Recibida correctamente. En la brevedad te estaré enviando las devoluciones o sugerencias pertinentes para finalizar el proceso.\n\n"
            if tiene_urg:
                cuerpo += "🚨 URGENTE (No entregada): No contamos con registro de entrega de este documento. Por favor, regulariza esta situación a la brevedad o comunícate conmigo si has tenido algún inconveniente técnico o personal para la carga.\n\n"
            
            cuerpo += "Quedamos a tu entera disposición para lo que necesites o para despejar cualquier duda técnica sobre el formato de entrega.\n\nAtentamente,\nIsi. Ale y Luis"
            return cuerpo

        def enviar_email():
            sel = tree_doc.selection()
            if not sel:
                messagebox.showwarning("Aviso", "Seleccione un docente.")
                return
            index = tree_doc.index(sel[0])
            # Como el tree puede estar filtrado, debemos buscar por nombre
            nombre_buscado = tree_doc.item(sel[0])['text']
            data = next((d for d in datos_docentes if d['docente'] == nombre_buscado), None)
            
            if not data or not data['email']:
                messagebox.showwarning("Falta Email", f"El docente {nombre_buscado} no tiene email registrado en la base de datos.")
                return
            
            cuerpo = generar_mensaje_docente(data)
            asunto = "Estado de Planificaciones Docentes"
            
            para = data['email']
            cc = "atombesi@etrr.edu.ar, lcornaglia@etrr.edu.ar"
            
            # Codificar url
            mailto_url = f"mailto:{para}?cc={cc}&subject={urllib.parse.quote(asunto)}&body={urllib.parse.quote(cuerpo)}"
            webbrowser.open(mailto_url)

        def copiar_para_ws():
            sel = tree_doc.selection()
            if not sel:
                messagebox.showwarning("Aviso", "Seleccione un docente.")
                return
            nombre_buscado = tree_doc.item(sel[0])['text']
            data = next((d for d in datos_docentes if d['docente'] == nombre_buscado), None)
            
            if data:
                cuerpo = generar_mensaje_docente(data)
                self.clipboard_clear()
                self.clipboard_append(cuerpo)
                self.update() # Refrescar portapapeles
                messagebox.showinfo("Copiado", f"Mensaje para {nombre_buscado} copiado al portapapeles.")

        # Frame de botones de acción
        frame_acciones = ctk.CTkFrame(top, fg_color="transparent")
        frame_acciones.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkButton(frame_acciones, text="📧 Enviar Email al Docente", command=enviar_email, fg_color="#3b5998", hover_color="#2d4373").pack(side="left", padx=10)
        ctk.CTkButton(frame_acciones, text="📋 Copiar Mensaje (WhatsApp)", command=copiar_para_ws, fg_color="#25D366", hover_color="#128C7E", text_color="black").pack(side="left", padx=10)

        cargar_arbol()

    def abrir_gestor_db(self):
        top = ctk.CTkToplevel(self)
        top.title("Gestión de Base de Datos Local")
        top.geometry("750x600")
        top.grab_set()

        # --- Frame Superior (Persistente) ---
        frame_top = ctk.CTkFrame(top)
        frame_top.pack(fill="x", padx=20, pady=(20, 0))

        # Definir actualizar_combos primero para que sea accesible
        def actualizar_combos():
            materias = db_manager.obtener_todas_materias()
            docentes = db_manager.obtener_todos_docentes()
            
            nonlocal materias_dict, docentes_dict
            materias_dict = {m[1]: m[0] for m in materias}
            docentes_dict = {d[1]: d[0] for d in docentes}
            
            m_vals = list(materias_dict.keys()) if materias_dict else ["Sin Materias"]
            d_vals = ["Ninguno"] + (list(docentes_dict.keys()) if docentes_dict else ["Sin Docentes"])
            
            cmb_mat.configure(values=m_vals)
            cmb_doc1.configure(values=d_vals)
            cmb_doc2.configure(values=d_vals)
            cmb_doc3.configure(values=d_vals)
            cmb_doc4.configure(values=d_vals)

            for item in self.tree_mat.get_children(): self.tree_mat.delete(item)
            for item in self.tree_doc.get_children(): self.tree_doc.delete(item)
            for item in self.tree_rel.get_children(): self.tree_rel.delete(item)

            for m in materias: self.tree_mat.insert("", "end", values=(m[0], m[1]))
            for d in docentes: 
                email_str = d[2] if d[2] else "-"
                self.tree_doc.insert("", "end", values=(d[0], d[1], email_str))
            
            rels = db_manager.obtener_materias_con_docentes()
            for mat, docs in rels.items():
                d1 = docs[0] if len(docs) > 0 else "-"
                d2 = docs[1] if len(docs) > 1 else "-"
                d3 = docs[2] if len(docs) > 2 else "-"
                d4 = docs[3] if len(docs) > 3 else "-"
                self.tree_rel.insert("", "end", values=(mat, d1, d2, d3, d4))

        # Cargar por Excel
        def cargar_excel_bd():
            filepath = filedialog.askopenfilename(title="Seleccionar archivo de Materias", filetypes=[("Excel files", "*.xlsx *.xls")])
            if not filepath: return
            
            opcion = messagebox.askyesnocancel(
                "Importar a Base de Datos", 
                "¿Desea borrar todo el contenido actual de la base de datos antes de importar?\n\n- Sí: Borra todos los datos y carga los del Excel.\n- No: Agrega los datos del Excel conservando los actuales.\n- Cancelar: Aborta la operación."
            )
            
            if opcion is None: return
            if opcion: db_manager.limpiar_db()
                
            try:
                df_import = pd.read_excel(filepath)
                df_import.columns = [str(c).strip() for c in df_import.columns]
                col_materia = df_import.columns[0]
                
                # Buscar columnas de docentes y sus posibles atributos (Email/Coordinador)
                doc_cols = []
                for i, col in enumerate(df_import.columns):
                    if 'docente' in col.lower():
                        # Intentar encontrar email y coordinador que le sigan o tengan el mismo indice
                        email_col = None
                        coord_col = None
                        # Buscar en las siguientes columnas o por nombre
                        suffix = col.lower().replace("docente", "").strip()
                        for c_other in df_import.columns:
                            c_low = c_other.lower()
                            if 'email' in c_low and (suffix in c_low or not suffix):
                                email_col = c_other
                            if ('coord' in c_low or 'coordinador' in c_low) and (suffix in c_low or not suffix):
                                coord_col = c_other
                        
                        doc_cols.append({
                            "doc": col,
                            "email": email_col,
                            "coord": coord_col
                        })
                
                count_m = count_r = 0
                for index, row in df_import.iterrows():
                    mat_str = str(row[col_materia]).strip()
                    if mat_str.lower() == 'nan' or not mat_str: continue
                    mat_id = db_manager.agregar_materia(mat_str)
                    count_m += 1
                    
                    for d_info in doc_cols:
                        doc_str = str(row[d_info["doc"]]).strip()
                        if doc_str.lower() != 'nan' and doc_str:
                            email_str = str(row[d_info["email"]]).strip() if d_info["email"] and str(row[d_info["email"]]).lower() != 'nan' else ""
                            coord_str = str(row[d_info["coord"]]).strip() if d_info["coord"] and str(row[d_info["coord"]]).lower() != 'nan' else ""
                            
                            doc_id = db_manager.agregar_docente(doc_str, email_str, coord_str)
                            db_manager.asignar_docente_a_materia(mat_id, doc_id)
                            count_r += 1
                                
                messagebox.showinfo("Importación Completa", f"Registros guardados/actualizados:\n- {count_m} Materias\n- {count_r} Asignaciones")
                actualizar_combos()
            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al importar a la BD: {str(e)}")

        def exportar_excel_bd():
            rels = db_manager.obtener_materias_con_docentes_detallado()
            if not rels:
                messagebox.showwarning("Advertencia", "No hay datos para exportar.")
                return
            
            save_path = filedialog.asksaveasfilename(
                title="Guardar Exportación",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )
            if not save_path: return

            filas = []
            for mat, docs in rels.items():
                row_data = {"Materia": mat}
                # Asegurar 4 docentes con toda su info
                for i in range(1, 5):
                    if len(docs) >= i:
                        d = docs[i-1]
                        row_data[f"Docente {i}"] = d["nombre"]
                        row_data[f"Email {i}"] = d["email"]
                        row_data[f"Coordinador {i}"] = d["coordinador"]
                    else:
                        row_data[f"Docente {i}"] = "-"
                        row_data[f"Email {i}"] = "-"
                        row_data[f"Coordinador {i}"] = "-"
                filas.append(row_data)
            
            df_export = pd.DataFrame(filas)
            try:
                df_export.to_excel(save_path, index=False)
                messagebox.showinfo("Exportación Exitosa", f"Se exportaron {len(df_export)} registros a:\n{save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo exportar: {str(e)}")

        ctk.CTkButton(frame_top, text="Importar por Excel", command=cargar_excel_bd, fg_color="#b8860b", hover_color="#8b6508").pack(side="left", padx=10, pady=10)
        ctk.CTkButton(frame_top, text="Exportar a Excel", command=exportar_excel_bd, fg_color="#2e8b57", hover_color="#1e5d3e").pack(side="left", padx=10, pady=10)

        # Tabs
        tabview = ctk.CTkTabview(top)
        tabview.pack(padx=20, pady=10, fill="both", expand=True)

        tab_mat = tabview.add("Materias")
        tab_doc = tabview.add("Docentes")
        tab_rel = tabview.add("Asignaciones")

        # Tab Materias
        frame_acciones_mat = ctk.CTkFrame(tab_mat, fg_color="transparent")
        frame_acciones_mat.pack(fill="x", pady=10)

        # Cargar a mano
        frame_manual = ctk.CTkFrame(frame_acciones_mat, fg_color="transparent")
        frame_manual.pack(fill="x", padx=10)
        ctk.CTkLabel(frame_manual, text="Agregar Materia Manual:").pack(side="left", padx=5)
        ent_mat = ctk.CTkEntry(frame_manual, width=300)
        ent_mat.pack(side="left", padx=5)
        def add_mat():
            val = ent_mat.get()
            if val:
                db_manager.agregar_materia(val)
                ent_mat.delete(0, 'end')
                messagebox.showinfo("OK", f"Materia {val} guardada.")
                actualizar_combos()
        ctk.CTkButton(frame_manual, text="Guardar", command=add_mat, width=100).pack(side="left", padx=5)
        
        def del_mat():
            selected = self.tree_mat.selection()
            if not selected:
                messagebox.showwarning("Aviso", "Seleccione una materia para eliminar.")
                return
            item = self.tree_mat.item(selected[0])
            mat_id = item['values'][0]
            mat_nom = item['values'][1]
            if messagebox.askyesno("Confirmar", f"¿Desea eliminar la materia '{mat_nom}'?\nEsto también eliminará sus asignaciones."):
                db_manager.eliminar_materia(mat_id)
                actualizar_combos()
        
        ctk.CTkButton(frame_acciones_mat, text="Eliminar Materia Seleccionada", command=del_mat, fg_color="#d9534f", hover_color="#c9302c").pack(pady=5)

        self.tree_mat = ttk.Treeview(tab_mat, columns=("ID", "Nombre"), show="headings", height=8)
        self.tree_mat.heading("ID", text="ID")
        self.tree_mat.heading("Nombre", text="Nombre")
        self.tree_mat.column("Nombre", width=400)
        self.tree_mat.pack(pady=10, fill="both", expand=True)

        # Configurar ordenamiento para tabla materias
        for col in ("ID", "Nombre"):
            self.tree_mat.heading(col, text=col, command=lambda _col=col: self.treeview_sort_column(self.tree_mat, _col, False))

        # Edición in-line con doble click para materias
        def on_tree_mat_double_click(event):
            selected = self.tree_mat.selection()
            if not selected:
                return
            item = self.tree_mat.item(selected[0])
            values = item['values']
            mat_id = values[0]
            
            col_id = self.tree_mat.identify_column(event.x)
            bbox = self.tree_mat.bbox(selected[0], col_id)
            if not bbox or col_id == '#1': return # No editar ID
            
            col_index = int(col_id.replace('#', '')) - 1
            current_value = values[col_index] if values[col_index] != "-" else ""
            
            entry_edit = ctk.CTkEntry(self.tree_mat, width=bbox[2], height=bbox[3], corner_radius=0)
            entry_edit.place(x=bbox[0], y=bbox[1])
            entry_edit.insert(0, current_value)
            entry_edit.focus_set()
            
            def save_edit(e):
                new_val = entry_edit.get()
                entry_edit.destroy()
                if new_val != current_value and new_val:
                    db_manager.actualizar_materia_por_id(mat_id, new_val)
                    actualizar_combos()

            entry_edit.bind("<Return>", save_edit)
            entry_edit.bind("<FocusOut>", lambda e: entry_edit.destroy())

        self.tree_mat.bind("<Double-1>", on_tree_mat_double_click)

        # Tab Docentes
        ctk.CTkLabel(tab_doc, text="Agregar Docente").pack(pady=10)
        frame_ent_doc = ctk.CTkFrame(tab_doc, fg_color="transparent")
        frame_ent_doc.pack(pady=5)
        
        ent_doc_nombre = ctk.CTkEntry(frame_ent_doc, placeholder_text="Nombre del docente", width=200)
        ent_doc_nombre.pack(side="left", padx=5)
        ent_doc_email = ctk.CTkEntry(frame_ent_doc, placeholder_text="Email (opcional)", width=200)
        ent_doc_email.pack(side="left", padx=5)
        
        ctk.CTkLabel(frame_ent_doc, text="Coordinador:").pack(side="left", padx=5)
        cmb_doc_coord = ctk.CTkOptionMenu(frame_ent_doc, values=["Sin Asignar", "Isi", "Ale"], width=120)
        cmb_doc_coord.pack(side="left", padx=5)
        
        def add_doc():
            nom = ent_doc_nombre.get()
            eml = ent_doc_email.get()
            crd = cmb_doc_coord.get()
            if crd == "Sin Asignar": crd = ""
            if nom:
                db_manager.agregar_docente(nom, eml, crd)
                ent_doc_nombre.delete(0, 'end')
                ent_doc_email.delete(0, 'end')
                cmb_doc_coord.set("Sin Asignar")
                messagebox.showinfo("OK", f"Docente {nom} guardado.")
                actualizar_combos()
        def del_doc():
            selected = self.tree_doc.selection()
            if not selected:
                messagebox.showwarning("Aviso", "Seleccione un docente para eliminar.")
                return
            item = self.tree_doc.item(selected[0])
            doc_id = item['values'][0]
            doc_nom = item['values'][1]
            if messagebox.askyesno("Confirmar", f"¿Desea eliminar al docente '{doc_nom}'?\nSe quitará de todas sus materias."):
                db_manager.eliminar_docente(doc_id)
                ent_doc_nombre.delete(0, 'end')
                ent_doc_email.delete(0, 'end')
                cmb_doc_coord.set("Sin Asignar")
                actualizar_combos()

        btn_frame_doc = ctk.CTkFrame(tab_doc, fg_color="transparent")
        btn_frame_doc.pack(pady=10)
        ctk.CTkButton(btn_frame_doc, text="Guardar Docente", command=add_doc).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame_doc, text="Eliminar Seleccionado", command=del_doc, fg_color="#d9534f", hover_color="#c9302c").pack(side="left", padx=10)

        self.tree_doc = ttk.Treeview(tab_doc, columns=("ID", "Nombre", "Email", "Coordinador"), show="headings", height=8)
        self.tree_doc.heading("ID", text="ID")
        self.tree_doc.heading("Nombre", text="Nombre")
        self.tree_doc.heading("Email", text="Email")
        self.tree_doc.heading("Coordinador", text="Coordinador")
        self.tree_doc.column("Email", width=180)
        self.tree_doc.column("Coordinador", width=100, anchor="center")
        self.tree_doc.pack(pady=10, fill="both", expand=True)

        # Configurar ordenamiento para tabla docentes
        for col in ("ID", "Nombre", "Email", "Coordinador"):
            self.tree_doc.heading(col, text=col, command=lambda _col=col: self.treeview_sort_column(self.tree_doc, _col, False))

        def on_tree_doc_select(event):
            selected = self.tree_doc.selection()
            if not selected:
                return
            item = self.tree_doc.item(selected[0])
            values = item['values']
            if values:
                nom = values[1]
                eml = values[2] if values[2] != "-" else ""
                crd = values[3] if values[3] != "-" else "Sin Asignar"
                ent_doc_nombre.delete(0, 'end')
                ent_doc_nombre.insert(0, nom)
                ent_doc_email.delete(0, 'end')
                ent_doc_email.insert(0, eml)
                cmb_doc_coord.set(crd)

        self.tree_doc.bind("<<TreeviewSelect>>", on_tree_doc_select)

        # Edición in-line con doble click
        def on_tree_doc_double_click(event):
            selected = self.tree_doc.selection()
            if not selected:
                return
            item = self.tree_doc.item(selected[0])
            values = item['values']
            doc_id = values[0]
            
            # Determinar qué columna se clickeó
            col_id = self.tree_doc.identify_column(event.x)
            bbox = self.tree_doc.bbox(selected[0], col_id)
            if not bbox or col_id == '#1': return # No editar ID
            
            col_index = int(col_id.replace('#', '')) - 1
            current_value = values[col_index] if values[col_index] != "-" else ""
            
            # Crear entry por encima
            entry_edit = ctk.CTkEntry(self.tree_doc, width=bbox[2], height=bbox[3], corner_radius=0)
            entry_edit.place(x=bbox[0], y=bbox[1])
            entry_edit.insert(0, current_value)
            entry_edit.focus_set()
            
            def save_edit(e):
                new_val = entry_edit.get()
                entry_edit.destroy()
                if new_val != current_value:
                    n_id, n_nom, n_eml, n_crd = values[0], values[1], values[2], values[3]
                    if col_index == 1: n_nom = new_val
                    elif col_index == 2: n_eml = new_val
                    elif col_index == 3: n_crd = new_val
                    
                    db_manager.actualizar_docente_por_id(n_id, n_nom, n_eml if n_eml != "-" else "", n_crd if n_crd != "-" else "")
                    actualizar_combos()

            entry_edit.bind("<Return>", save_edit)
            entry_edit.bind("<FocusOut>", lambda e: entry_edit.destroy())

        self.tree_doc.bind("<Double-1>", on_tree_doc_double_click)

        # Tab Relacion (Asignar hasta 4 docentes por materia)
        ctk.CTkLabel(tab_rel, text="Asignar Docentes a Materia").pack(pady=5)
        
        cmb_mat = ctk.CTkOptionMenu(tab_rel, values=[])
        cmb_mat.pack(pady=5)
        
        frame_docentes = ctk.CTkFrame(tab_rel, fg_color="transparent")
        frame_docentes.pack(pady=5)
        
        cmb_doc1 = ctk.CTkOptionMenu(frame_docentes, values=[])
        cmb_doc1.grid(row=0, column=0, padx=5, pady=5)
        cmb_doc2 = ctk.CTkOptionMenu(frame_docentes, values=[])
        cmb_doc2.grid(row=0, column=1, padx=5, pady=5)
        cmb_doc3 = ctk.CTkOptionMenu(frame_docentes, values=[])
        cmb_doc3.grid(row=1, column=0, padx=5, pady=5)
        cmb_doc4 = ctk.CTkOptionMenu(frame_docentes, values=[])
        cmb_doc4.grid(row=1, column=1, padx=5, pady=5)

        def add_rel():
            m_nom = cmb_mat.get()
            seleccionados = [
                cmb_doc1.get(),
                cmb_doc2.get(),
                cmb_doc3.get(),
                cmb_doc4.get()
            ]
            
            if m_nom in materias_dict:
                mat_id = materias_dict[m_nom]
                guardados = 0
                for d_nom in seleccionados:
                    if d_nom in docentes_dict and d_nom != "Ninguno" and d_nom != "Sin Docentes":
                        db_manager.asignar_docente_a_materia(mat_id, docentes_dict[d_nom])
                        guardados += 1
                        
                if guardados > 0:
                    messagebox.showinfo("OK", f"{guardados} relación/es creada/s para {m_nom}.")
                    actualizar_combos()
                else:
                    messagebox.showwarning("Aviso", "No se seleccionó ningún docente válido.")

        def del_rel():
            m_nom = cmb_mat.get()
            if m_nom in materias_dict:
                mat_id = materias_dict[m_nom]
                if messagebox.askyesno("Confirmar", f"¿Quitar todos los docentes asignados a '{m_nom}'?"):
                    db_manager.eliminar_relaciones_materia(mat_id)
                    actualizar_combos()

        btn_frame_rel = ctk.CTkFrame(tab_rel, fg_color="transparent")
        btn_frame_rel.pack(pady=10)
        ctk.CTkButton(btn_frame_rel, text="Asignar", command=add_rel).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame_rel, text="Quitar Asignaciones", command=del_rel, fg_color="#d9534f", hover_color="#c9302c").pack(side="left", padx=10)

        self.tree_rel = ttk.Treeview(tab_rel, columns=("Materia", "Docente 1", "Docente 2", "Docente 3", "Docente 4"), show="headings", height=8)
        self.tree_rel.heading("Materia", text="Materia")
        self.tree_rel.heading("Docente 1", text="Docente 1")
        self.tree_rel.heading("Docente 2", text="Docente 2")
        self.tree_rel.heading("Docente 3", text="Docente 3")
        self.tree_rel.heading("Docente 4", text="Docente 4")
        
        self.tree_rel.column("Docente 4", width=120)
        
        self.tree_rel.pack(pady=10, fill="both", expand=True)

        # Configurar ordenamiento para tabla asignaciones
        for col in ("Materia", "Docente 1", "Docente 2", "Docente 3", "Docente 4"):
            self.tree_rel.heading(col, text=col, command=lambda _col=col: self.treeview_sort_column(self.tree_rel, _col, False))
        
        materias_dict = {}
        docentes_dict = {}

        def actualizar_combos():
            materias = db_manager.obtener_todas_materias()
            docentes = db_manager.obtener_todos_docentes()
            
            nonlocal materias_dict, docentes_dict
            materias_dict = {m[1]: m[0] for m in materias}
            docentes_dict = {d[1]: d[0] for d in docentes}
            
            m_vals = list(materias_dict.keys()) if materias_dict else ["Sin Materias"]
            d_vals = ["Ninguno"] + (list(docentes_dict.keys()) if docentes_dict else ["Sin Docentes"])
            
            cmb_mat.configure(values=m_vals)
            cmb_doc1.configure(values=d_vals)
            cmb_doc2.configure(values=d_vals)
            cmb_doc3.configure(values=d_vals)
            cmb_doc4.configure(values=d_vals)
            
            if materias_dict: cmb_mat.set(m_vals[0])
            cmb_doc1.set("Ninguno")
            cmb_doc2.set("Ninguno")
            cmb_doc3.set("Ninguno")
            cmb_doc4.set("Ninguno")

            # Actualizar vistas
            for item in self.tree_mat.get_children(): self.tree_mat.delete(item)
            for item in self.tree_doc.get_children(): self.tree_doc.delete(item)
            for item in self.tree_rel.get_children(): self.tree_rel.delete(item)

            for m in materias: self.tree_mat.insert("", "end", values=(m[0], m[1]))
            for d in docentes: 
                email_str = d[2] if d[2] else "-"
                coord_str = d[3] if d[3] else "-"
                self.tree_doc.insert("", "end", values=(d[0], d[1], email_str, coord_str))
            
            rels = db_manager.obtener_materias_con_docentes()
            for mat, docs in rels.items():
                d1 = docs[0] if len(docs) > 0 else "-"
                d2 = docs[1] if len(docs) > 1 else "-"
                d3 = docs[2] if len(docs) > 2 else "-"
                d4 = docs[3] if len(docs) > 3 else "-"
                self.tree_rel.insert("", "end", values=(mat, d1, d2, d3, d4))

        actualizar_combos()

        def on_tree_rel_select(event):
            selected = self.tree_rel.selection()
            if not selected:
                return
            item = self.tree_rel.item(selected[0])
            values = item['values']
            if values:
                mat = values[0]
                d1, d2, d3, d4 = values[1], values[2], values[3], values[4]
                
                # Seleccionar materia
                if mat in materias_dict:
                    cmb_mat.set(mat)
                
                # Seleccionar docentes
                cmb_doc1.set(d1 if d1 != "-" and d1 in docentes_dict else "Ninguno")
                cmb_doc2.set(d2 if d2 != "-" and d2 in docentes_dict else "Ninguno")
                cmb_doc3.set(d3 if d3 != "-" and d3 in docentes_dict else "Ninguno")
                cmb_doc4.set(d4 if d4 != "-" and d4 in docentes_dict else "Ninguno")

        self.tree_rel.bind("<<TreeviewSelect>>", on_tree_rel_select)

if __name__ == "__main__":
    app = AppPlanificaciones()
    app.mainloop()
