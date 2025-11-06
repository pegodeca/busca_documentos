"""
Aplicación de Búsqueda en Documentos con OCR
=============================================
Autor: Sistema de Desarrollo
Versión: 2.0.0
Descripción: Aplicación GUI para buscar texto en documentos Word, PDF, Excel, TXT, HTML, PHP
             Incluye soporte OCR para PDFs escaneados/imágenes

Estándares aplicados:
- PEP 8: Formato y convenciones de código Python
- Documentación clara con docstrings
- Manejo robusto de errores
- Separación de responsabilidades (SRP)
- Nombres descriptivos y significativos
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import threading
from typing import List, Dict
import os
import sys

# Librerías para lectura de documentos
try:
    from docx import Document  # python-docx
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    import PyPDF2
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

try:
    import openpyxl
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False

# Librerías para OCR
try:
    import pytesseract
    from pdf2image import convert_from_path
    from PIL import Image
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False


class DocumentSearcher:
    """
    Clase responsable de la búsqueda de texto en documentos.
    
    Principios aplicados:
    - Single Responsibility: Solo se encarga de buscar en documentos
    - Extensibilidad: Fácil añadir nuevos tipos de documentos
    """
    
    # Constante: extensiones soportadas
    SUPPORTED_EXTENSIONS = {'.txt', '.docx', '.pdf', '.xlsx', '.xls', '.html', '.htm', '.php'}
    
    def __init__(self):
        """Inicializa el buscador con configuración por defecto."""
        self.results: List[Dict[str, str]] = []
        self.search_cancelled = False
        self.use_ocr = False
        self.tesseract_path = None
        self.poppler_path = None
        self._configure_tesseract()
    
    def _configure_tesseract(self):
        """Configura las rutas de Tesseract y Poppler si están disponibles."""
        if not OCR_AVAILABLE:
            return
        
        # Rutas comunes de Tesseract en Windows
        common_tesseract_paths = [
            r'C:\Program Files\Tesseract-OCR\tesseract.exe',
            r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
            r'C:\Tesseract-OCR\tesseract.exe',
        ]
        
        # Intentar encontrar Tesseract
        for path in common_tesseract_paths:
            if os.path.exists(path):
                self.tesseract_path = path
                pytesseract.pytesseract.tesseract_cmd = path
                break
        
        # Rutas comunes de Poppler en Windows
        common_poppler_paths = [
            r'C:\Program Files\poppler\Library\bin',
            r'C:\Program Files (x86)\poppler\Library\bin',
            r'C:\poppler\Library\bin',
            r'C:\poppler-24.08.0\Library\bin',
        ]
        
        for path in common_poppler_paths:
            if os.path.exists(path):
                self.poppler_path = path
                break
    
    def set_tesseract_path(self, path: str):
        """
        Configura manualmente la ruta de Tesseract.
        
        Args:
            path: Ruta al ejecutable de Tesseract
        """
        if os.path.exists(path):
            self.tesseract_path = path
            pytesseract.pytesseract.tesseract_cmd = path
            return True
        return False
    
    def set_poppler_path(self, path: str):
        """
        Configura manualmente la ruta de Poppler.
        
        Args:
            path: Ruta al directorio bin de Poppler
        """
        if os.path.exists(path):
            self.poppler_path = path
            return True
        return False
    
    def search_in_directory(self, directory: str, search_term: str, 
                          case_sensitive: bool = False,
                          use_ocr: bool = False,
                          callback=None) -> List[Dict[str, str]]:
        """
        Busca un término en todos los documentos de un directorio.
        
        Args:
            directory: Ruta del directorio a buscar
            search_term: Texto a buscar
            case_sensitive: Si la búsqueda distingue mayúsculas/minúsculas
            use_ocr: Si se aplica OCR a PDFs e imágenes
            callback: Función callback para actualizar progreso
            
        Returns:
            Lista de diccionarios con resultados encontrados
            
        Raises:
            ValueError: Si el directorio no existe
        """
        if not os.path.exists(directory):
            raise ValueError(f"El directorio no existe: {directory}")
        
        if not search_term.strip():
            raise ValueError("El término de búsqueda no puede estar vacío")
        
        self.results = []
        self.search_cancelled = False
        self.use_ocr = use_ocr
        directory_path = Path(directory)
        
        # Normalizar término de búsqueda si no es case-sensitive
        normalized_term = search_term if case_sensitive else search_term.lower()
        
        # Obtener todos los archivos soportados
        files_to_search = self._get_supported_files(directory_path)
        total_files = len(files_to_search)
        
        if total_files == 0:
            return self.results
        
        # Buscar en cada archivo
        for index, file_path in enumerate(files_to_search, 1):
            if self.search_cancelled:
                break
                
            try:
                content = self._extract_text_from_file(file_path)
                
                if content:
                    normalized_content = content if case_sensitive else content.lower()
                    
                    if normalized_term in normalized_content:
                        # Contar ocurrencias
                        occurrences = normalized_content.count(normalized_term)
                        
                        self.results.append({
                            'file': str(file_path),
                            'filename': file_path.name,
                            'type': file_path.suffix,
                            'occurrences': occurrences
                        })
                
                # Actualizar progreso
                if callback:
                    progress = (index / total_files) * 100
                    callback(progress, file_path.name)
                    
            except Exception as e:
                # Log del error pero continúa buscando
                print(f"Error procesando {file_path}: {str(e)}")
                continue
        
        return self.results
    
    def cancel_search(self):
        """Cancela la búsqueda en progreso."""
        self.search_cancelled = True
    
    def _get_supported_files(self, directory: Path) -> List[Path]:
        """
        Obtiene recursivamente todos los archivos soportados.
        
        Args:
            directory: Directorio raíz
            
        Returns:
            Lista de rutas de archivos soportados
        """
        supported_files = []
        
        try:
            for file_path in directory.rglob('*'):
                if file_path.is_file() and file_path.suffix.lower() in self.SUPPORTED_EXTENSIONS:
                    supported_files.append(file_path)
        except PermissionError:
            pass  # Ignorar carpetas sin permisos
            
        return supported_files
    
    def _extract_text_from_file(self, file_path: Path) -> str:
        """
        Extrae texto del archivo según su tipo.
        
        Args:
            file_path: Ruta del archivo
            
        Returns:
            Texto extraído del archivo
        """
        extension = file_path.suffix.lower()
        
        try:
            if extension == '.txt':
                return self._read_txt(file_path)
            elif extension == '.docx' and DOCX_AVAILABLE:
                return self._read_docx(file_path)
            elif extension == '.pdf' and PDF_AVAILABLE:
                return self._read_pdf_with_ocr(file_path) if self.use_ocr else self._read_pdf(file_path)
            elif extension in {'.xlsx', '.xls'} and EXCEL_AVAILABLE:
                return self._read_excel(file_path)
            elif extension in {'.html', '.htm', '.php'}:
                return self._read_txt(file_path)  # Son archivos de texto plano
        except Exception as e:
            print(f"Error extrayendo texto de {file_path}: {str(e)}")
            
        return ""
    
    def _read_txt(self, file_path: Path) -> str:
        """Lee archivo de texto plano."""
        encodings = ['utf-8', 'latin-1', 'cp1252']
        
        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as file:
                    return file.read()
            except UnicodeDecodeError:
                continue
        
        return ""
    
    def _read_docx(self, file_path: Path) -> str:
        """Lee documento Word (.docx)."""
        if not DOCX_AVAILABLE:
            return ""
        
        doc = Document(file_path)
        return '\n'.join([paragraph.text for paragraph in doc.paragraphs])
    
    def _read_pdf(self, file_path: Path) -> str:
        """Lee documento PDF (solo texto nativo)."""
        if not PDF_AVAILABLE:
            return ""
        
        text = []
        with open(file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            for page in pdf_reader.pages:
                extracted = page.extract_text()
                if extracted:
                    text.append(extracted)
        
        return '\n'.join(text)
    
    def _read_pdf_with_ocr(self, file_path: Path) -> str:
        """
        Lee documento PDF aplicando OCR si es necesario.
        
        Args:
            file_path: Ruta al archivo PDF
            
        Returns:
            Texto extraído (combinando texto nativo + OCR)
        """
        if not OCR_AVAILABLE or not self.tesseract_path or not self.poppler_path:
            # Si OCR no está disponible, usar método normal
            return self._read_pdf(file_path)
        
        try:
            # Primero intentar extraer texto nativo
            native_text = self._read_pdf(file_path)
            
            # Si hay suficiente texto nativo, usarlo
            if len(native_text.strip()) > 100:
                return native_text
            
            # Si no hay texto o es muy poco, aplicar OCR
            return self._extract_text_with_ocr(file_path)
            
        except Exception as e:
            print(f"Error en OCR para {file_path}: {str(e)}")
            # Fallback al método normal
            return self._read_pdf(file_path)
    
    def _extract_text_with_ocr(self, file_path: Path) -> str:
        """
        Extrae texto de PDF usando OCR.
        
        Args:
            file_path: Ruta al archivo PDF
            
        Returns:
            Texto extraído mediante OCR
        """
        if not OCR_AVAILABLE:
            return ""
        
        try:
            # Convertir PDF a imágenes
            images = convert_from_path(
                str(file_path),
                poppler_path=self.poppler_path
            )
            
            # Aplicar OCR a cada página
            text_parts = []
            for image in images:
                text = pytesseract.image_to_string(image, lang='spa+eng')
                text_parts.append(text)
            
            return '\n'.join(text_parts)
            
        except Exception as e:
            print(f"Error aplicando OCR: {str(e)}")
            return ""
    
    def _read_excel(self, file_path: Path) -> str:
        """Lee archivo Excel (.xlsx, .xls)."""
        if not EXCEL_AVAILABLE:
            return ""
        
        workbook = openpyxl.load_workbook(file_path, data_only=True)
        text = []
        
        for sheet in workbook.worksheets:
            for row in sheet.iter_rows(values_only=True):
                text.extend([str(cell) for cell in row if cell is not None])
        
        return ' '.join(text)


class DocumentSearcherGUI:
    """
    Interfaz gráfica para la aplicación de búsqueda.
    
    Principios aplicados:
    - Separación Vista/Lógica: UI separada de lógica de búsqueda
    - Usabilidad: Interfaz intuitiva y retroalimentación clara
    """
    
    # Constantes de configuración visual
    COLOR_PRIMARY = "#2c3e50"
    COLOR_SECONDARY = "#3498db"
    COLOR_SUCCESS = "#27ae60"
    COLOR_WARNING = "#e74c3c"
    PADDING = 10
    
    def __init__(self, root: tk.Tk):
        """
        Inicializa la interfaz gráfica.
        
        Args:
            root: Ventana principal de Tkinter
        """
        self.root = root
        self.searcher = DocumentSearcher()
        self.search_thread = None
        
        self._setup_window()
        self._create_widgets()
        self._check_dependencies()
    
    def _setup_window(self):
        """Configura la ventana principal."""
        self.root.title("🔍 Buscador de Documentos v2.0 (con OCR)")
        self.root.geometry("950x750")
        self.root.minsize(850, 650)
        
        # Estilo moderno
        style = ttk.Style()
        style.theme_use('clam')
    
    def _create_widgets(self):
        """Crea todos los widgets de la interfaz."""
        # Frame principal con padding
        main_frame = ttk.Frame(self.root, padding=self.PADDING)
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurar expansión
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(9, weight=1)
        
        # --- Sección: Selección de carpeta ---
        ttk.Label(main_frame, text="📁 Carpeta a buscar:", 
                 font=('Helvetica', 10, 'bold')).grid(
            row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        self.directory_var = tk.StringVar()
        directory_entry = ttk.Entry(main_frame, textvariable=self.directory_var, 
                                    state='readonly')
        directory_entry.grid(row=1, column=0, columnspan=2, 
                           sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(main_frame, text="Seleccionar carpeta", 
                  command=self._select_directory).grid(
            row=1, column=2, padx=(5, 0), pady=(0, 10))
        
        # --- Sección: Término de búsqueda ---
        ttk.Label(main_frame, text="🔎 Texto a buscar:", 
                 font=('Helvetica', 10, 'bold')).grid(
            row=2, column=0, sticky=tk.W, pady=(0, 5))
        
        self.search_term_var = tk.StringVar()
        search_entry = ttk.Entry(main_frame, textvariable=self.search_term_var)
        search_entry.grid(row=3, column=0, columnspan=2, 
                         sticky=(tk.W, tk.E), pady=(0, 10))
        search_entry.bind('<Return>', lambda e: self._start_search())
        
        # --- Opciones de búsqueda ---
        options_frame = ttk.Frame(main_frame)
        options_frame.grid(row=4, column=0, columnspan=3, sticky=tk.W, pady=(0, 10))
        
        # Checkbox: Distinguir mayúsculas/minúsculas
        self.case_sensitive_var = tk.BooleanVar()
        ttk.Checkbutton(options_frame, text="Distinguir mayúsculas/minúsculas", 
                       variable=self.case_sensitive_var).pack(side=tk.LEFT, padx=(0, 15))
        
        # Checkbox: Aplicar OCR
        self.ocr_var = tk.BooleanVar()
        self.ocr_checkbox = ttk.Checkbutton(
            options_frame, 
            text="🔬 Aplicar OCR (PDFs escaneados/imágenes)", 
            variable=self.ocr_var,
            command=self._on_ocr_toggle
        )
        self.ocr_checkbox.pack(side=tk.LEFT)
        
        # Botón configurar OCR
        ttk.Button(options_frame, text="⚙️ Configurar OCR", 
                  command=self._configure_ocr).pack(side=tk.LEFT, padx=(5, 0))
        
        # --- Botones de acción ---
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=3, pady=(0, 10))
        
        self.search_button = ttk.Button(button_frame, text="🔍 Buscar", 
                                       command=self._start_search)
        self.search_button.pack(side=tk.LEFT, padx=5)
        
        self.cancel_button = ttk.Button(button_frame, text="✖ Cancelar", 
                                       command=self._cancel_search, 
                                       state='disabled')
        self.cancel_button.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="🗑 Limpiar resultados", 
                  command=self._clear_results).pack(side=tk.LEFT, padx=5)
        
        # --- Barra de progreso ---
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, 
                                           maximum=100, mode='determinate')
        self.progress_bar.grid(row=6, column=0, columnspan=3, 
                              sticky=(tk.W, tk.E), pady=(0, 5))
        
        self.status_var = tk.StringVar(value="Listo para buscar")
        status_label = ttk.Label(main_frame, textvariable=self.status_var, 
                                foreground=self.COLOR_PRIMARY)
        status_label.grid(row=7, column=0, columnspan=3, 
                         sticky=tk.W, pady=(0, 10))
        
        # --- Resultados ---
        ttk.Label(main_frame, text="📋 Resultados:", 
                 font=('Helvetica', 10, 'bold')).grid(
            row=8, column=0, sticky=tk.W, pady=(0, 5))
        
        # Frame para resultados con scrollbar
        results_frame = ttk.Frame(main_frame)
        results_frame.grid(row=9, column=0, columnspan=3, 
                          sticky=(tk.W, tk.E, tk.N, tk.S))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        
        # Treeview para mostrar resultados
        columns = ('filename', 'type', 'occurrences', 'path')
        self.results_tree = ttk.Treeview(results_frame, columns=columns, 
                                         show='headings', height=15)
        
        # Configurar columnas
        self.results_tree.heading('filename', text='Archivo')
        self.results_tree.heading('type', text='Tipo')
        self.results_tree.heading('occurrences', text='Coincidencias')
        self.results_tree.heading('path', text='Ruta completa')
        
        self.results_tree.column('filename', width=200)
        self.results_tree.column('type', width=80)
        self.results_tree.column('occurrences', width=120)
        self.results_tree.column('path', width=400)
        
        # Scrollbars
        vsb = ttk.Scrollbar(results_frame, orient="vertical", 
                           command=self.results_tree.yview)
        hsb = ttk.Scrollbar(results_frame, orient="horizontal", 
                           command=self.results_tree.xview)
        self.results_tree.configure(yscrollcommand=vsb.set, 
                                   xscrollcommand=hsb.set)
        
        self.results_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
        hsb.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Doble clic para abrir archivo
        self.results_tree.bind('<Double-1>', self._open_file)
        
        # Contador de resultados
        self.result_count_var = tk.StringVar(value="0 documentos encontrados")
        ttk.Label(main_frame, textvariable=self.result_count_var,
                 font=('Helvetica', 9)).grid(
            row=10, column=0, columnspan=3, sticky=tk.W, pady=(5, 0))
    
    def _check_dependencies(self):
        """Verifica y muestra advertencias sobre dependencias faltantes."""
        missing = []
        
        if not DOCX_AVAILABLE:
            missing.append("python-docx (archivos .docx)")
        if not PDF_AVAILABLE:
            missing.append("PyPDF2 (archivos .pdf)")
        if not EXCEL_AVAILABLE:
            missing.append("openpyxl (archivos .xlsx)")
        
        if missing:
            warning_msg = ("⚠️ Algunas librerías no están instaladas:\n\n" + 
                          "\n".join(f"• {lib}" for lib in missing) +
                          "\n\nInstálalas con: pip install python-docx PyPDF2 openpyxl")
            
            self.status_var.set("Advertencia: Algunas funcionalidades limitadas")
            messagebox.showwarning("Dependencias faltantes", warning_msg)
        
        # Verificar OCR
        if not OCR_AVAILABLE:
            self.ocr_checkbox.config(state='disabled')
            self.status_var.set("OCR no disponible - Instala: pip install pytesseract pdf2image pillow")
        elif not self.searcher.tesseract_path or not self.searcher.poppler_path:
            self.ocr_checkbox.config(state='disabled')
            missing_ocr = []
            if not self.searcher.tesseract_path:
                missing_ocr.append("Tesseract-OCR")
            if not self.searcher.poppler_path:
                missing_ocr.append("Poppler")
            self.status_var.set(f"OCR requiere: {', '.join(missing_ocr)} - Click en 'Configurar OCR'")
    
    def _on_ocr_toggle(self):
        """Maneja el evento de activar/desactivar OCR."""
        if self.ocr_var.get():
            if not OCR_AVAILABLE:
                messagebox.showwarning(
                    "OCR no disponible",
                    "Instala las librerías necesarias:\n\n"
                    "pip install pytesseract pdf2image pillow"
                )
                self.ocr_var.set(False)
                return
            
            if not self.searcher.tesseract_path or not self.searcher.poppler_path:
                messagebox.showwarning(
                    "Configuración OCR incompleta",
                    "Por favor configura las rutas de Tesseract y Poppler\n"
                    "usando el botón 'Configurar OCR'"
                )
                self.ocr_var.set(False)
                return
            
            # Advertir que OCR es más lento
            result = messagebox.askyesno(
                "OCR Activado",
                "⚠️ IMPORTANTE:\n\n"
                "El OCR hace la búsqueda MUCHO MÁS LENTA,\n"
                "especialmente con PDFs grandes.\n\n"
                "¿Continuar con OCR activado?",
                icon='warning'
            )
            
            if not result:
                self.ocr_var.set(False)
    
    def _configure_ocr(self):
        """Abre ventana de configuración de OCR."""
        config_window = tk.Toplevel(self.root)
        config_window.title("⚙️ Configuración OCR")
        config_window.geometry("600x400")
        config_window.transient(self.root)
        config_window.grab_set()
        
        frame = ttk.Frame(config_window, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Instrucciones
        ttk.Label(frame, text="Configuración de OCR", 
                 font=('Helvetica', 12, 'bold')).pack(pady=(0, 10))
        
        info_text = (
            "Para usar OCR necesitas instalar:\n\n"
            "1. Tesseract-OCR:\n"
            "   https://github.com/UB-Mannheim/tesseract/wiki\n\n"
            "2. Poppler para Windows:\n"
            "   https://github.com/oschwartz10612/poppler-windows/releases\n\n"
            "3. Librerías Python:\n"
            "   pip install pytesseract pdf2image pillow"
        )
        
        info_label = ttk.Label(frame, text=info_text, justify=tk.LEFT)
        info_label.pack(pady=(0, 20))
        
        # Tesseract Path
        ttk.Label(frame, text="Ruta de Tesseract:").pack(anchor=tk.W)
        tesseract_frame = ttk.Frame(frame)
        tesseract_frame.pack(fill=tk.X, pady=(5, 15))
        
        tesseract_var = tk.StringVar(value=self.searcher.tesseract_path or "")
        tesseract_entry = ttk.Entry(tesseract_frame, textvariable=tesseract_var)
        tesseract_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        def browse_tesseract():
            path = filedialog.askopenfilename(
                title="Seleccionar tesseract.exe",
                filetypes=[("Ejecutable", "*.exe"), ("Todos", "*.*")]
            )
            if path:
                tesseract_var.set(path)
        
        ttk.Button(tesseract_frame, text="Buscar...", 
                  command=browse_tesseract).pack(side=tk.LEFT)
        
        # Poppler Path
        ttk.Label(frame, text="Ruta del directorio bin de Poppler:").pack(anchor=tk.W)
        poppler_frame = ttk.Frame(frame)
        poppler_frame.pack(fill=tk.X, pady=(5, 20))
        
        poppler_var = tk.StringVar(value=self.searcher.poppler_path or "")
        poppler_entry = ttk.Entry(poppler_frame, textvariable=poppler_var)
        poppler_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        def browse_poppler():
            path = filedialog.askdirectory(title="Seleccionar carpeta bin de Poppler")
            if path:
                poppler_var.set(path)
        
        ttk.Button(poppler_frame, text="Buscar...", 
                  command=browse_poppler).pack(side=tk.LEFT)
        
        # Botones
        button_frame = ttk.Frame(frame)
        button_frame.pack(pady=(10, 0))
        
        def save_config():
            tesseract_ok = self.searcher.set_tesseract_path(tesseract_var.get())
            poppler_ok = self.searcher.set_poppler_path(poppler_var.get())
            
            if tesseract_ok and poppler_ok:
                messagebox.showinfo("Éxito", "Configuración guardada correctamente")
                self.ocr_checkbox.config(state='normal')
                self.status_var.set("OCR configurado y listo para usar")
                config_window.destroy()
            else:
                messagebox.showerror("Error", "Las rutas especificadas no son válidas")
        
        ttk.Button(button_frame, text="Guardar", 
                  command=save_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancelar", 
                  command=config_window.destroy).pack(side=tk.LEFT, padx=5)
    
    def _select_directory(self):
        """Abre diálogo para seleccionar directorio."""
        directory = filedialog.askdirectory(title="Seleccionar carpeta a buscar")
        if directory:
            self.directory_var.set(directory)
            self.status_var.set(f"Carpeta seleccionada: {Path(directory).name}")
    
    def _start_search(self):
        """Inicia la búsqueda en un hilo separado."""
        directory = self.directory_var.get()
        search_term = self.search_term_var.get()
        
        # Validaciones
        if not directory:
            messagebox.showwarning("Advertencia", 
                                 "Por favor selecciona una carpeta")
            return
        
        if not search_term.strip():
            messagebox.showwarning("Advertencia", 
                                 "Por favor ingresa un texto a buscar")
            return
        
        # Limpiar resultados anteriores
        self._clear_results()
        
        # Deshabilitar botón de búsqueda, habilitar cancelar
        self.search_button.config(state='disabled')
        self.cancel_button.config(state='normal')
        
        # Iniciar búsqueda en hilo separado (no bloquear UI)
        self.search_thread = threading.Thread(
            target=self._perform_search,
            args=(directory, search_term, self.case_sensitive_var.get(), self.ocr_var.get()),
            daemon=True
        )
        self.search_thread.start()
    
    def _perform_search(self, directory: str, search_term: str, 
                       case_sensitive: bool, use_ocr: bool):
        """
        Realiza la búsqueda (ejecutado en hilo separado).
        
        Args:
            directory: Directorio a buscar
            search_term: Término de búsqueda
            case_sensitive: Si distingue mayúsculas/minúsculas
            use_ocr: Si aplica OCR a PDFs
        """
        try:
            results = self.searcher.search_in_directory(
                directory, 
                search_term, 
                case_sensitive,
                use_ocr,
                callback=self._update_progress
            )
            
            # Actualizar UI en el hilo principal
            self.root.after(0, self._display_results, results)
            
        except Exception as e:
            self.root.after(0, messagebox.showerror, 
                          "Error", f"Error durante la búsqueda:\n{str(e)}")
        finally:
            self.root.after(0, self._search_completed)
    
    def _update_progress(self, progress: float, current_file: str):
        """
        Actualiza barra de progreso (thread-safe).
        
        Args:
            progress: Porcentaje completado (0-100)
            current_file: Archivo actual siendo procesado
        """
        self.root.after(0, self.progress_var.set, progress)
        status_text = f"Buscando... ({progress:.0f}%) - {current_file}"
        if self.ocr_var.get():
            status_text += " [OCR activado - puede ser lento]"
        self.root.after(0, self.status_var.set, status_text)
    
    def _display_results(self, results: List[Dict]):
        """
        Muestra resultados en la tabla.
        
        Args:
            results: Lista de resultados encontrados
        """
        for result in results:
            self.results_tree.insert('', tk.END, values=(
                result['filename'],
                result['type'],
                result['occurrences'],
                result['file']
            ))
        
        # Actualizar contador
        count = len(results)
        self.result_count_var.set(
            f"{count} documento{'s' if count != 1 else ''} encontrado{'s' if count != 1 else ''}"
        )
        
        if count == 0:
            self.status_var.set("Búsqueda completada - No se encontraron coincidencias")
        else:
            self.status_var.set(
                f"✓ Búsqueda completada - {count} documento(s) con coincidencias"
            )
    
    def _search_completed(self):
        """Restaura estado de la UI después de búsqueda."""
        self.search_button.config(state='normal')
        self.cancel_button.config(state='disabled')
        self.progress_var.set(0)
    
    def _cancel_search(self):
        """Cancela la búsqueda en progreso."""
        self.searcher.cancel_search()
        self.status_var.set("Búsqueda cancelada por el usuario")
        self._search_completed()
    
    def _clear_results(self):
        """Limpia los resultados mostrados."""
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        self.result_count_var.set("0 documentos encontrados")
        self.progress_var.set(0)
    
    def _open_file(self, event):
        """
        Abre el archivo seleccionado con la aplicación predeterminada.
        
        Args:
            event: Evento de doble clic
        """
        selection = self.results_tree.selection()
        if selection:
            item = self.results_tree.item(selection[0])
            file_path = item['values'][3]  # Ruta completa
            
            try:
                # Multiplataforma: abrir archivo
                import platform
                if platform.system() == 'Windows':
                    os.startfile(file_path)
                elif platform.system() == 'Darwin':  # macOS
                    os.system(f'open "{file_path}"')
                else:  # Linux
                    os.system(f'xdg-open "{file_path}"')
            except Exception as e:
                messagebox.showerror("Error", 
                                   f"No se pudo abrir el archivo:\n{str(e)}")


def main():
    """
    Función principal para ejecutar la aplicación.
    
    Punto de entrada del programa.
    """
    root = tk.Tk()
    app = DocumentSearcherGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
    
