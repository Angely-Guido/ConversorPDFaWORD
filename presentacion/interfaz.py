import flet as ft
import os

# --- LÓGICA CON PROTECCIÓN A FALLOS ---
try:
    from logica.conversor_pdf import convertir_pdf_a_word, unir_pdfs
    from logica.conversor_word import convertir_word_a_pdf, unir_words
except ImportError:
    print("⚠️ MODO VISUAL: Backend no encontrado.")
    def convertir_pdf_a_word(a, b): return "Error: Backend no conectado"
    def convertir_word_a_pdf(a, b): return "Error: Backend no conectado"
    def unir_pdfs(a, b): return "Error: Backend no conectado"
    def unir_words(a, b): return "Error: Backend no conectado"

def iniciar_interfaz():
    def main(page: ft.Page):
        # --- Configuración Ventana ---
        page.title = "Conversor Pro - UCR"
        page.window_width = 950
        page.window_height = 750
        page.theme_mode = ft.ThemeMode.LIGHT
        page.padding = 20

        # --- Datos en Memoria ---
        archivos_convertir = []
        archivos_unir = []

        # --- Elementos UI Contenedores ---
        lista_convertir_ui = ft.Column(scroll=ft.ScrollMode.AUTO, expand=True, spacing=5)
        lista_unir_ui = ft.Column(scroll=ft.ScrollMode.AUTO, expand=True, spacing=5)
        
        # --- Notificaciones ---
        def mostrar_mensaje(texto, color="green"):
            page.snack_bar = ft.SnackBar(ft.Text(texto), bgcolor=color)
            page.snack_bar.open = True
            page.update()

        # ==========================================
        # 🟢 LÓGICA DRAG & DROP (DESDE WINDOWS)
        # ==========================================
        def drop_windows(e: ft.FilePickerResultEvent):
            tab_index = tabs.selected_index
            rutas = [f.path for f in e.files]
            
            if tab_index == 0: # Tab Convertir
                for r in rutas:
                    if r not in archivos_convertir:
                        archivos_convertir.append(r)
                render_lista_convertir()
                mostrar_mensaje(f"Se agregaron {len(rutas)} archivos", "blue")
            
            elif tab_index == 1: # Tab Unir
                pdfs = [r for r in rutas if r.lower().endswith(".pdf")]
                for r in pdfs:
                    archivos_unir.append(r)
                render_lista_unir()
                if pdfs: mostrar_mensaje(f"Se agregaron {len(pdfs)} PDFs", "blue")
                else: mostrar_mensaje("Aquí solo PDFs", "orange")

        page.on_file_drop = drop_windows

        # ==========================================
        # 1️⃣ PESTAÑA CONVERTIR
        # ==========================================
        def render_lista_convertir():
            lista_convertir_ui.controls.clear()
            for i, ruta in enumerate(archivos_convertir):
                nombre = os.path.basename(ruta)
                item = ft.Container(
                    content=ft.Row([
                        ft.Icon("description", color="blue"),
                        ft.Text(f"{i+1}. {nombre}", expand=1, size=14),
                        ft.IconButton("delete", icon_color="red", on_click=lambda e, x=i: borrar_convertir(x))
                    ]),
                    bgcolor="#F0F4F8", padding=10, border_radius=5
                )
                lista_convertir_ui.controls.append(item)
            page.update()

        def borrar_convertir(idx):
            archivos_convertir.pop(idx)
            render_lista_convertir()

        def cargar_carpeta(path):
            if not path: return
            es_pdf_word = switch_modo.value
            exts = [".pdf"] if es_pdf_word else [".docx", ".doc"]
            
            contador = 0
            for archivo in os.listdir(path):
                if any(archivo.lower().endswith(ext) for ext in exts):
                    ruta_completa = os.path.join(path, archivo)
                    if ruta_completa not in archivos_convertir:
                        archivos_convertir.append(ruta_completa)
                        contador += 1
            
            render_lista_convertir()
            mostrar_mensaje(f"Se cargaron {contador} archivos de la carpeta", "blue")

        def ejecutar_conversion(e):
            if not archivos_convertir:
                mostrar_mensaje("Lista vacía", "red")
                return
            
            es_pdf_word = switch_modo.value
            barra_progreso.visible = True
            page.update()

            errores = 0
            for ruta in archivos_convertir:
                carpeta = os.path.dirname(ruta)
                nombre = os.path.splitext(os.path.basename(ruta))[0]
                
                if es_pdf_word:
                    salida = os.path.join(carpeta, f"{nombre}_edit.docx")
                    res = convertir_pdf_a_word(ruta, salida)
                else:
                    salida = os.path.join(carpeta, f"{nombre}_doc.pdf")
                    res = convertir_word_a_pdf(ruta, salida)
                
                if "Error" in str(res): errores += 1

            barra_progreso.visible = False
            if errores == 0: mostrar_mensaje("¡Todo convertido con éxito!", "green")
            else: mostrar_mensaje(f"Terminado con {errores} errores", "orange")
            page.update()

        # Componentes Tab 1
        picker_conv = ft.FilePicker(on_result=lambda e: [archivos_convertir.append(f.path) for f in e.files] and render_lista_convertir() if e.files else None)
        picker_folder = ft.FilePicker(on_result=lambda e: cargar_carpeta(e.path))
        page.overlay.extend([picker_conv, picker_folder])
        
        # --- CORRECCIÓN AQUÍ: Quitamos el label interno del Switch ---
        switch_modo = ft.Switch(value=True) 
        
        barra_progreso = ft.ProgressBar(visible=False, color="blue")
        
        contenido_tab_conv = ft.Column([
            ft.Text("Conversor Masivo", size=20, weight="bold"),
            # Ahora la fila se verá limpia: Texto - Switch - Texto
            ft.Row([ft.Text("Word a PDF", weight="bold"), switch_modo, ft.Text("PDF a Word", weight="bold")]),
            ft.Row([
                ft.ElevatedButton("Seleccionar Archivos", icon="upload_file", on_click=lambda _: picker_conv.pick_files(allow_multiple=True)),
                ft.ElevatedButton("Seleccionar Carpeta", icon="folder_open", bgcolor="#EEEEEE", color="black", on_click=lambda _: picker_folder.get_directory_path())
            ]),
            ft.Divider(),
            ft.Text("Archivos en cola:", color="grey"),
            lista_convertir_ui, 
            ft.Divider(),
            ft.Row([ft.ElevatedButton("CONVERTIR AHORA", icon="play_arrow", bgcolor="blue", color="white", on_click=ejecutar_conversion)], alignment="end"),
            barra_progreso
        ], expand=True)

        # ==========================================
        # 2️⃣ PESTAÑA UNIR (Drag & Drop Corregido)
        # ==========================================
        
        def render_lista_unir():
            lista_unir_ui.controls.clear()
            for i, ruta in enumerate(archivos_unir):
                nombre = os.path.basename(ruta)
                
                # Diseño de la Tarjeta
                card_content = ft.Container(
                    content=ft.Row([
                        ft.Icon("drag_handle", color="grey"), 
                        ft.Icon("picture_as_pdf", color="red"),
                        ft.Text(f"{i+1}. {nombre}", expand=1, weight="bold"),
                        ft.IconButton("delete", icon_color="red", on_click=lambda e, x=i: borrar_unir(x))
                    ], alignment="center"),
                    bgcolor="white",
                    padding=10,
                    border=ft.border.all(1, "#DDDDDD"),
                    border_radius=8,
                    shadow=ft.BoxShadow(blur_radius=2, color="#E0E0E0")
                )

                # Draggable
                draggable = ft.Draggable(
                    group="unir_items",
                    content=card_content,
                    content_when_dragging=ft.Container(bgcolor="blue", width=50, height=5, border_radius=5),
                    data=i 
                )

                # Target
                target = ft.DragTarget(
                    group="unir_items",
                    content=draggable,
                    on_accept=drop_interno_unir,
                    data=i
                )
                
                lista_unir_ui.controls.append(target)
            page.update()

        def drop_interno_unir(e):
            try:
                src_index = -1
                for idx, control in enumerate(lista_unir_ui.controls):
                    if control.content.uid == e.src_id:
                        src_index = idx
                        break
                
                dest_index = int(e.control.data)

                if src_index != -1 and src_index != dest_index:
                    item = archivos_unir.pop(src_index)
                    archivos_unir.insert(dest_index, item)
                    render_lista_unir()
            except Exception as ex:
                print(f"Error al mover: {ex}")

        def borrar_unir(idx):
            archivos_unir.pop(idx)
            render_lista_unir()

        picker_unir = ft.FilePicker(on_result=lambda e: [archivos_unir.append(f.path) for f in e.files] and render_lista_unir() if e.files else None)
        picker_guardar = ft.FilePicker(on_result=lambda e: unir_pdfs(archivos_unir, e.path) if e.path else None)
        page.overlay.extend([picker_unir, picker_guardar])

        contenido_tab_unir = ft.Column([
            ft.Text("Unir PDFs", size=20, weight="bold"),
            ft.Text("Arrastra las tarjetas (desde el icono gris) para cambiar el orden.", size=12, color="grey"),
            ft.ElevatedButton("Agregar PDFs", icon="add", on_click=lambda _: picker_unir.pick_files(allow_multiple=True, allowed_extensions=["pdf"])),
            ft.Divider(),
            
            # CUERPO (LISTA)
            ft.Container(
                content=lista_unir_ui,
                expand=True,
                bgcolor="#FAFAFA",
                border_radius=10,
                padding=10
            ),
            
            # FOOTER (BOTÓN FIJO)
            ft.Container(
                content=ft.Row([
                    ft.ElevatedButton("UNIR PDFs", icon="merge_type", bgcolor="green", color="white", height=45, 
                                    on_click=lambda _: picker_guardar.save_file(file_name="unido.pdf") if len(archivos_unir) > 1 else mostrar_mensaje("Agrega al menos 2 PDFs", "orange"))
                ], alignment="end"),
                padding=ft.padding.only(top=10)
            )
        ], expand=True)

        # ==========================================
        # 🚀 ARMADO FINAL
        # ==========================================
        tabs = ft.Tabs(
            selected_index=0,
            animation_duration=300,
            tabs=[
                ft.Tab(text="Convertir", icon="transform", content=ft.Container(contenido_tab_conv, padding=20)),
                ft.Tab(text="Unir", icon="call_merge", content=ft.Container(contenido_tab_unir, padding=20)),
            ],
            expand=True
        )

        page.add(tabs)

    ft.app(target=main)