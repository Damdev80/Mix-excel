import flet as ft
import pandas as pd
from pathlib import Path


class ExcelMixerApp:
    def __init__(self, page: ft.Page):
        self.page = page
        self.page.title = "Mix Excel - Transferir Columnas"
        self.page.window_width = 600
        self.page.window_height = 800
        self.page.theme_mode = ft.ThemeMode.LIGHT
        self.page.padding = 30
        
        # Variables
        self.source_file = None
        self.dest_file = None
        self.source_columns = []
        self.dest_columns = []
        
        # UI Components
        self.source_file_text = ft.Text("No seleccionado", size=12, color="grey700")
        self.dest_file_text = ft.Text("No seleccionado", size=12, color="grey700")
        self.source_column_dropdown = ft.Dropdown(
            label="Columna con valores a copiar",
            hint_text="Selecciona una columna",
            disabled=True,
            width=300
        )
        self.source_ref_column_dropdown = ft.Dropdown(
            label="Columna de referencia origen",
            hint_text="Selecciona columna de comparación",
            disabled=True,
            width=300
        )
        self.dest_column_dropdown = ft.Dropdown(
            label="Columna donde pegar valores",
            hint_text="Selecciona una columna",
            disabled=True,
            width=300
        )
        self.dest_ref_column_dropdown = ft.Dropdown(
            label="Columna de referencia destino",
            hint_text="Selecciona columna de comparación",
            disabled=True,
            width=300
        )
        self.tolerance_field = ft.TextField(
            label="Tolerancia (%)",
            hint_text="Ej: 5 para ±5%",
            value="5",
            width=150,
            keyboard_type=ft.KeyboardType.NUMBER
        )
        self.status_text = ft.Text("", size=12, text_align=ft.TextAlign.CENTER)
        
        # File picker
        self.file_picker = ft.FilePicker(on_result=self.on_file_picked)
        self.page.overlay.append(self.file_picker)
        
        self.current_picker_type = None
        
        self.setup_ui()
    
    def setup_ui(self):
        """Configurar la interfaz de usuario"""
        
        # Título
        title = ft.Container(
            content=ft.Column([
                ft.Text(
                    "📊 Mix Excel",
                    size=28,
                    weight=ft.FontWeight.BOLD,
                    color="blue700"
                ),
                ft.Text(
                    "Distribuye valores basándose en coincidencias",
                    size=12,
                    color="grey600",
                    text_align=ft.TextAlign.CENTER
                ),
            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
            alignment=ft.alignment.center,
            margin=ft.margin.only(bottom=20)
        )
        
        # Sección archivo origen
        source_section = ft.Container(
            content=ft.Column([
                ft.Text("1. Archivo Origen", size=16, weight=ft.FontWeight.W_500),
                ft.Row([
                    ft.ElevatedButton(
                        "Seleccionar archivo",
                        icon=ft.Icons.UPLOAD_FILE,
                        on_click=lambda _: self.pick_file("source")
                    ),
                ]),
                self.source_file_text,
                ft.Divider(height=10, color="transparent"),
                self.source_ref_column_dropdown,
                self.source_column_dropdown,
            ], spacing=10),
            padding=20,
            border=ft.border.all(1, "grey300"),
            border_radius=10,
            margin=ft.margin.only(bottom=20)
        )
        
        # Sección archivo destino
        dest_section = ft.Container(
            content=ft.Column([
                ft.Text("2. Archivo Destino", size=16, weight=ft.FontWeight.W_500),
                ft.Row([
                    ft.ElevatedButton(
                        "Seleccionar archivo",
                        icon=ft.Icons.UPLOAD_FILE,
                        on_click=lambda _: self.pick_file("dest")
                    ),
                ]),
                self.dest_file_text,
                ft.Divider(height=10, color="transparent"),
                self.dest_ref_column_dropdown,
                self.dest_column_dropdown,
            ], spacing=10),
            padding=20,
            border=ft.border.all(1, "grey300"),
            border_radius=10,
            margin=ft.margin.only(bottom=20)
        )
        
        # Sección de tolerancia
        tolerance_section = ft.Container(
            content=ft.Column([
                ft.Text("3. Configuración", size=16, weight=ft.FontWeight.W_500),
                ft.Row([
                    self.tolerance_field,
                    ft.Text("Para valores numéricos cercanos", size=12, color="grey700")
                ], alignment=ft.MainAxisAlignment.START),
            ], spacing=10),
            padding=20,
            border=ft.border.all(1, "grey300"),
            border_radius=10,
            margin=ft.margin.only(bottom=20)
        )
        
        # Botón de transferir
        transfer_button = ft.Container(
            content=ft.ElevatedButton(
                "✨ Transferir Datos",
                icon=ft.Icons.ARROW_FORWARD,
                on_click=self.transfer_data,
                style=ft.ButtonStyle(
                    color="white",
                    bgcolor="blue700",
                    padding=20,
                ),
                width=300
            ),
            alignment=ft.alignment.center,
            margin=ft.margin.only(top=10, bottom=20)
        )
        
        # Status
        status_container = ft.Container(
            content=self.status_text,
            alignment=ft.alignment.center,
            padding=10,
            border_radius=5,
        )
        
        # Main container
        main_container = ft.Container(
            content=ft.Column([
                title,
                source_section,
                dest_section,
                tolerance_section,
                transfer_button,
                status_container
            ], scroll=ft.ScrollMode.AUTO),
            expand=True
        )
        
        self.page.add(main_container)
    
    def pick_file(self, picker_type):
        """Abrir el selector de archivos"""
        self.current_picker_type = picker_type
        self.file_picker.pick_files(
            allowed_extensions=["xlsx", "xls"],
            dialog_title="Selecciona un archivo Excel"
        )
    
    def on_file_picked(self, e: ft.FilePickerResultEvent):
        """Manejar la selección de archivo"""
        if e.files and len(e.files) > 0:
            file_path = e.files[0].path
            
            if self.current_picker_type == "source":
                self.source_file = file_path
                self.source_file_text.value = Path(file_path).name
                self.source_file_text.color = "green700"
                self.load_columns("source")
            elif self.current_picker_type == "dest":
                self.dest_file = file_path
                self.dest_file_text.value = Path(file_path).name
                self.dest_file_text.color = "green700"
                self.load_columns("dest")
            
            self.page.update()
    
    def load_columns(self, file_type):
        """Cargar las columnas del archivo Excel"""
        try:
            file_path = self.source_file if file_type == "source" else self.dest_file
            df = pd.read_excel(file_path)
            columns = df.columns.tolist()
            
            if file_type == "source":
                self.source_columns = columns
                self.source_column_dropdown.options = [
                    ft.dropdown.Option(col) for col in columns
                ]
                self.source_column_dropdown.disabled = False
                self.source_ref_column_dropdown.options = [
                    ft.dropdown.Option(col) for col in columns
                ]
                self.source_ref_column_dropdown.disabled = False
            else:
                self.dest_columns = columns
                self.dest_column_dropdown.options = [
                    ft.dropdown.Option(col) for col in columns
                ]
                self.dest_column_dropdown.disabled = False
                self.dest_ref_column_dropdown.options = [
                    ft.dropdown.Option(col) for col in columns
                ]
                self.dest_ref_column_dropdown.disabled = False
            
            self.page.update()
            
        except PermissionError:
            file_name = "origen" if file_type == "source" else "destino"
            self.show_status(
                f"❌ Error: El archivo {file_name} está abierto. Ciérralo e intenta de nuevo.",
                "red700"
            )
        except Exception as e:
            self.show_status(f"❌ Error al cargar columnas: {str(e)}", "red700")
    
    def transfer_data(self, e):
        """Transferir datos de una columna a otra"""
        try:
            # Validaciones
            if not self.source_file or not self.dest_file:
                self.show_status("⚠️ Selecciona ambos archivos", "orange700")
                return
            
            if not self.source_column_dropdown.value or not self.source_ref_column_dropdown.value:
                self.show_status("⚠️ Selecciona ambas columnas del archivo origen", "orange700")
                return
            
            if not self.dest_column_dropdown.value or not self.dest_ref_column_dropdown.value:
                self.show_status("⚠️ Selecciona ambas columnas del archivo destino", "orange700")
                return
            
            # Leer datos
            source_df = pd.read_excel(self.source_file)
            dest_df = pd.read_excel(self.dest_file)
            
            source_col = self.source_column_dropdown.value
            source_ref_col = self.source_ref_column_dropdown.value
            dest_col = self.dest_column_dropdown.value
            dest_ref_col = self.dest_ref_column_dropdown.value
            
            # Obtener tolerancia
            try:
                tolerance = float(self.tolerance_field.value) / 100.0
            except:
                tolerance = 0.05  # 5% por defecto
            
            # Distribuir datos basándose en coincidencias
            matches_found = 0
            
            for idx, dest_ref_value in dest_df[dest_ref_col].items():
                if pd.isna(dest_ref_value):
                    continue
                
                # Buscar coincidencia en origen
                match_found = False
                
                for src_idx, src_ref_value in source_df[source_ref_col].items():
                    if pd.isna(src_ref_value):
                        continue
                    
                    # Verificar si coinciden
                    if self.values_match(dest_ref_value, src_ref_value, tolerance):
                        # Asignar el valor
                        dest_df.at[idx, dest_col] = source_df.at[src_idx, source_col]
                        matches_found += 1
                        match_found = True
                        break
            
            # Guardar
            try:
                dest_df.to_excel(self.dest_file, index=False)
                self.show_status(
                    f"✅ Transferencia completada: {matches_found} coincidencias encontradas",
                    "green700"
                )
            except PermissionError:
                self.show_status(
                    "❌ Error: El archivo de destino está abierto. Ciérralo e intenta de nuevo.",
                    "red700"
                )
            
        except PermissionError:
            self.show_status(
                "❌ Error: Uno de los archivos está abierto. Ciérralos e intenta de nuevo.",
                "red700"
            )
        except Exception as e:
            error_msg = str(e)
            if "Permission denied" in error_msg or "Errno 13" in error_msg:
                self.show_status(
                    "❌ Error: El archivo está abierto en otra aplicación. Ciérralo primero.",
                    "red700"
                )
            else:
                self.show_status(f"❌ Error: {error_msg}", "red700")
    
    def values_match(self, val1, val2, tolerance):
        """Verificar si dos valores coinciden o son cercanos"""
        # Si son exactamente iguales (comparación de strings)
        if str(val1).strip().lower() == str(val2).strip().lower():
            return True
        
        # Si ambos son numéricos, comparar con tolerancia
        try:
            num1 = float(val1)
            num2 = float(val2)
            
            # Calcular el rango de tolerancia
            lower_bound = num2 * (1 - tolerance)
            upper_bound = num2 * (1 + tolerance)
            
            return lower_bound <= num1 <= upper_bound
        except (ValueError, TypeError):
            # Si no son numéricos, solo comparación exacta
            return False
    
    def show_status(self, message, color):
        """Mostrar mensaje de estado"""
        self.status_text.value = message
        self.status_text.color = color
        self.page.update()


def main(page: ft.Page):
    ExcelMixerApp(page)


if __name__ == "__main__":
    ft.app(target=main)
