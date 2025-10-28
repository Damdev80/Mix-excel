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
        self.page.padding = 40
        self.page.bgcolor = "grey50"
        
        # Variables
        self.source_file = None
        self.dest_file = None
        self.source_columns = []
        self.dest_columns = []
        
        # UI Components
        self.source_file_text = ft.Text("Ningún archivo seleccionado", size=12, color="grey500", weight=ft.FontWeight.W_400)
        self.dest_file_text = ft.Text("Ningún archivo seleccionado", size=12, color="grey500", weight=ft.FontWeight.W_400)
        self.source_column_dropdown = ft.Dropdown(
            label="Columna con valores",
            hint_text="Selecciona una columna",
            disabled=True,
            width=300,
            border_radius=8,
            border_color="grey300"
        )
        self.source_ref_column_dropdown = ft.Dropdown(
            label="Columna de referencia",
            hint_text="Número de factura",
            disabled=True,
            width=300,
            border_radius=8,
            border_color="grey300"
        )
        self.dest_column_dropdown = ft.Dropdown(
            label="Columna destino",
            hint_text="Donde pegar valores",
            disabled=True,
            width=300,
            border_radius=8,
            border_color="grey300"
        )
        self.dest_ref_column_dropdown = ft.Dropdown(
            label="Columna de referencia",
            hint_text="Número de factura",
            disabled=True,
            width=300,
            border_radius=8,
            border_color="grey300"
        )
        self.dest_adjacent_column_dropdown = ft.Dropdown(
            label="Columna adyacente (para proporción)",
            hint_text="Columna para calcular distribución",
            disabled=True,
            width=300,
            border_radius=8,
            border_color="grey300"
        )
        self.tolerance_field = ft.TextField(
            label="Tolerancia (%)",
            hint_text="5",
            value="5",  
            width=120,
            keyboard_type=ft.KeyboardType.NUMBER,
            border_radius=8,
            border_color="grey300"
        )
        self.status_text = ft.Text("", size=13, text_align=ft.TextAlign.CENTER, weight=ft.FontWeight.W_400)
        
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
                    "Mix Excel",
                    size=32,
                    weight=ft.FontWeight.W_700,
                    color="blue900"
                ),
                ft.Text(
                    "Transferencia de datos entre archivos",
                    size=13,
                    weight=ft.FontWeight.W_400,
                    color="grey700"
                ),
            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=5),
            alignment=ft.alignment.center,
            margin=ft.margin.only(bottom=30)
        )
        
        # Sección archivo origen
        source_section = ft.Container(
            content=ft.Column([
                ft.Text("Archivo Origen", size=14, weight=ft.FontWeight.W_500, color="grey800"),
                ft.ElevatedButton(
                    "Seleccionar archivo",
                    on_click=lambda _: self.pick_file("source"),
                    style=ft.ButtonStyle(
                        shape=ft.RoundedRectangleBorder(radius=8),
                        padding=15,
                    )
                ),
                self.source_file_text,
                ft.Divider(height=20, color="grey200"),
                self.source_ref_column_dropdown,
                self.source_column_dropdown,
            ], spacing=12),
            padding=25,
            bgcolor="white",
            border=ft.border.all(1, "grey200"),
            border_radius=12,
            margin=ft.margin.only(bottom=20)
        )
        
        # Sección archivo destino
        dest_section = ft.Container(
            content=ft.Column([
                ft.Text("Archivo Destino", size=14, weight=ft.FontWeight.W_500, color="grey800"),
                ft.ElevatedButton(
                    "Seleccionar archivo",
                    on_click=lambda _: self.pick_file("dest"),
                    style=ft.ButtonStyle(
                        shape=ft.RoundedRectangleBorder(radius=8),
                        padding=15,
                    )
                ),
                self.dest_file_text,
                ft.Divider(height=20, color="grey200"),
                self.dest_ref_column_dropdown,
                self.dest_adjacent_column_dropdown,
                self.dest_column_dropdown,
            ], spacing=12),
            padding=25,
            bgcolor="white",
            border=ft.border.all(1, "grey200"),
            border_radius=12,
            margin=ft.margin.only(bottom=20)
        )
        
        # Sección de tolerancia
        tolerance_section = ft.Container(
            content=ft.Column([
                ft.Text("Configuración", size=14, weight=ft.FontWeight.W_500, color="grey800"),
                ft.Row([
                    self.tolerance_field,
                    ft.Text("Tolerancia para comparación", size=12, color="grey600", weight=ft.FontWeight.W_400)
                ], alignment=ft.MainAxisAlignment.START, spacing=15),
            ], spacing=12),
            padding=25,
            bgcolor="white",
            border=ft.border.all(1, "grey200"),
            border_radius=12,
            margin=ft.margin.only(bottom=25)
        )
        
        # Botón de transferir
        transfer_button = ft.Container(
            content=ft.ElevatedButton(
                "Transferir Datos",
                on_click=self.transfer_data,
                style=ft.ButtonStyle(
                    color="white",
                    bgcolor="blue900",
                    shape=ft.RoundedRectangleBorder(radius=10),
                    padding=ft.padding.symmetric(horizontal=40, vertical=18),
                ),
            ),
            alignment=ft.alignment.center,
            margin=ft.margin.only(bottom=25)
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
                self.source_file_text.color = "blue900"
                self.load_columns("source")
            elif self.current_picker_type == "dest":
                self.dest_file = file_path
                self.dest_file_text.value = Path(file_path).name
                self.dest_file_text.color = "blue900"
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
                self.dest_adjacent_column_dropdown.options = [
                    ft.dropdown.Option(col) for col in columns
                ]
                self.dest_adjacent_column_dropdown.disabled = False
            
            self.page.update()
            
        except PermissionError:
            file_name = "origen" if file_type == "source" else "destino"
            self.show_status(
                f"Error: El archivo {file_name} está abierto. Ciérralo e intenta de nuevo.",
                "red700"
            )
        except Exception as e:
            self.show_status(f"Error al cargar columnas: {str(e)}", "red700")
    
    def transfer_data(self, e):
        """Transferir datos de una columna a otra"""
        try:
            # Validaciones
            if not self.source_file or not self.dest_file:
                self.show_status("Selecciona ambos archivos", "orange700")
                return
            
            if not self.source_column_dropdown.value or not self.source_ref_column_dropdown.value:
                self.show_status("Selecciona ambas columnas del archivo origen", "orange700")
                return
            
            if not self.dest_column_dropdown.value or not self.dest_ref_column_dropdown.value or not self.dest_adjacent_column_dropdown.value:
                self.show_status("Selecciona todas las columnas del archivo destino", "orange700")
                return
            
            # Leer datos
            source_df = pd.read_excel(self.source_file)
            dest_df = pd.read_excel(self.dest_file)
            
            source_col = self.source_column_dropdown.value
            source_ref_col = self.source_ref_column_dropdown.value
            dest_col = self.dest_column_dropdown.value
            dest_ref_col = self.dest_ref_column_dropdown.value
            dest_adjacent_col = self.dest_adjacent_column_dropdown.value
            
            # Obtener tolerancia
            try:
                tolerance = float(self.tolerance_field.value) / 100.0
            except:
                tolerance = 0.05  # 5% por defecto
            
            # Distribuir datos basándose en coincidencias
            matches_found = 0
            processed_refs = set()  # Para evitar procesar la misma referencia múltiples veces
            
            for idx, dest_ref_value in dest_df[dest_ref_col].items():
                if pd.isna(dest_ref_value) or dest_ref_value in processed_refs:
                    continue
                
                # Buscar TODAS las filas en destino con el mismo número de factura
                dest_matching_indices = dest_df[dest_df[dest_ref_col] == dest_ref_value].index.tolist()
                
                # Buscar TODAS las facturas con el mismo número en origen
                source_matching_rows = source_df[source_df[source_ref_col] == dest_ref_value]
                
                if len(source_matching_rows) > 0 and len(dest_matching_indices) > 0:
                    # Obtener todos los valores del origen
                    source_values = []
                    for src_idx, src_row in source_matching_rows.iterrows():
                        src_val = src_row[source_col]
                        if pd.notna(src_val):
                            try:
                                source_values.append(float(src_val))
                            except (ValueError, TypeError):
                                continue
                    
                    if len(source_values) > 0:
                        # Obtener los valores adyacentes del destino
                        dest_adjacent_values = []
                        for dest_idx in dest_matching_indices:
                            adj_val = dest_df.at[dest_idx, dest_adjacent_col]
                            if pd.notna(adj_val):
                                try:
                                    dest_adjacent_values.append((dest_idx, float(adj_val)))
                                except (ValueError, TypeError):
                                    pass
                        
                        if len(dest_adjacent_values) > 0:
                            # CASO 1: Si hay múltiples valores en origen, emparejar por similitud
                            if len(source_values) > 1:
                                used_source_values = set()
                                for dest_idx, dest_adj_val in dest_adjacent_values:
                                    best_match_value = None
                                    best_match_idx = None
                                    min_diff = float('inf')
                                    
                                    # Buscar el valor del origen más parecido que no se haya usado
                                    for i, src_val in enumerate(source_values):
                                        if i not in used_source_values:
                                            diff = abs(src_val - dest_adj_val) / dest_adj_val if dest_adj_val != 0 else float('inf')
                                            if diff < min_diff:
                                                min_diff = diff
                                                best_match_value = src_val
                                                best_match_idx = i
                                    
                                    # Asignar el mejor match
                                    if best_match_value is not None:
                                        if best_match_value == int(best_match_value):
                                            dest_df.at[dest_idx, dest_col] = int(best_match_value)
                                        else:
                                            dest_df.at[dest_idx, dest_col] = round(best_match_value, 2)
                                        used_source_values.add(best_match_idx)
                                        matches_found += 1
                            
                            # CASO 2: Si hay UN SOLO valor en origen, distribuir proporcionalmente
                            else:
                                total_to_distribute = source_values[0]
                                total_adjacent = sum([val for _, val in dest_adjacent_values])
                                
                                if total_adjacent > 0:
                                    # Distribuir proporcionalmente y ajustar el último valor para evitar errores de redondeo
                                    distributed_sum = 0
                                    for i, (dest_idx, adjacent_value) in enumerate(dest_adjacent_values):
                                        if i == len(dest_adjacent_values) - 1:
                                            # Último valor: asignar el restante para que sume exacto
                                            distributed_value = total_to_distribute - distributed_sum
                                        else:
                                            proportion = adjacent_value / total_adjacent
                                            distributed_value = total_to_distribute * proportion
                                            distributed_sum += distributed_value
                                        
                                        # Preservar tipo con 2 decimales
                                        dest_df.at[dest_idx, dest_col] = round(distributed_value, 2)
                                        matches_found += 1
                        
                        processed_refs.add(dest_ref_value)
            
            # Guardar
            try:
                dest_df.to_excel(self.dest_file, index=False)
                self.show_status(
                    f"Transferencia completada: {matches_found} coincidencias encontradas",
                    "green800"
                )
            except PermissionError:
                self.show_status(
                    "Error: El archivo de destino está abierto. Ciérralo e intenta de nuevo.",
                    "red700"
                )
            
        except PermissionError:
            self.show_status(
                "Error: Uno de los archivos está abierto. Ciérralos e intenta de nuevo.",
                "red700"
            )
        except Exception as e:
            error_msg = str(e)
            if "Permission denied" in error_msg or "Errno 13" in error_msg:
                self.show_status(
                    "Error: El archivo está abierto en otra aplicación. Ciérralo primero.",
                    "red700"
                )
            else:
                self.show_status(f"Error: {error_msg}", "red700")
    
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
