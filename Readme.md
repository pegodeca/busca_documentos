# 🔍 Document Searcher (con OCR)

Aplicación GUI en Python para buscar texto en documentos Word, PDF, Excel, TXT, HTML y PHP. Incluye soporte OCR para PDFs escaneados e imágenes, con configuración persistente.

## 🚀 Características

- Búsqueda recursiva en directorios
- Soporte para `.txt`, `.docx`, `.pdf`, `.xlsx`, `.xls`, `.html`, `.htm`, `.php`
- OCR opcional con Tesseract y Poppler
- Interfaz gráfica con barra de progreso y tabla de resultados
- Configuración persistente en `~/.doc_searcher_config.json`
- Ventana de configuración OCR con validación visual
- Ventana de debug con log en tiempo real

## 🧱 Arquitectura

- `DocumentSearcher`: lógica de búsqueda y OCR
- `ConfigManager`: gestión de configuración persistente
- `DocumentSearcherGUI`: interfaz gráfica con `tkinter` y `ttk`

## 🧪 Requisitos

```bash
pip install python-docx PyPDF2 openpyxl pytesseract pdf2image pillow
