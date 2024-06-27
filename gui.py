import tkinter as tk
from tkinter import filedialog, messagebox, font
import threading
import json
import os
import webbrowser
from utils import process_xml_to_docx, load_config, save_config

CONFIG_FILE = 'styles_config.json'

def start_gui():
    def select_xml_file():
        file_path = filedialog.askopenfilename(filetypes=[("XML files", "*.xml")])
        xml_file_var.set(file_path)

    def select_output_folder():
        folder_path = filedialog.askdirectory()
        output_folder_var.set(folder_path)

    def start_processing():
        threading.Thread(target=process_file).start()

    def process_file():
        xml_file = xml_file_var.get()
        output_folder = output_folder_var.get()
        output_file_name = output_file_var.get()
        if not xml_file or not output_folder or not output_file_name:
            messagebox.showerror("Error", "Por favor, selecciona un archivo XML, una carpeta de salida y un nombre para el archivo de salida.")
            return

        if not output_file_name.endswith(".docx"):
            output_file_name += ".docx"

        status_var.set("Procesando...")
        log_message("Iniciando proceso...")

        try:
            process_xml_to_docx(xml_file, output_folder, output_file_name)
            status_var.set("Completado")
            log_message("Proceso completado exitosamente.")
        except Exception as e:
            status_var.set("Error")
            log_message(f"Error durante el proceso: {str(e)}")

    def update_config():
        config = {}
        for field, (type_var, style_var) in fields.items():
            config[field] = {'type': type_var.get(), 'style': style_var.get()}
        save_config(config)
        messagebox.showinfo("Información", "Configuración guardada correctamente.")

    def log_message(message):
        log_text.config(state=tk.NORMAL)
        log_text.insert(tk.END, message + "\n")
        log_text.see(tk.END)
        log_text.config(state=tk.DISABLED)

    def open_link(event):
        webbrowser.open_new("https://escandell.cat")

    root = tk.Tk()
    root.title("XML to DOCX Converter")

    xml_file_var = tk.StringVar()
    output_folder_var = tk.StringVar()
    output_file_var = tk.StringVar()
    status_var = tk.StringVar()

    config = load_config()

    fields = {}

    xml_structure = [
        'Evento-Principal-Titulo',
        'Evento-Principal-Dia',
        'Evento-Principal-Hora',
        'Evento-Principal-Lugar',
        'Evento-Principal-Descripcion',
        'Evento-Principal-info.extra',
        'Sub-evento-Titulo',
        'Sub-evento-Dia',
        'Sub-evento-Hora',
        'Sub-evento-Lugar',
        'Sub-evento-descripcion',
        'Sub-evento-info.extra',
        'actividad-titulo',
        'actividad-hora',
        'actividad-lugar',
        'actividad-descipcion',  # Actualitzat
        'actividad-info-extra'   # Actualitzat
    ]

    # Left column for XML fields and style settings
    left_frame = tk.Frame(root)
    left_frame.grid(row=0, column=0, padx=10, pady=10, sticky='n')

    tk.Label(left_frame, text="Conversión de campos XML a estilos de Word", font=("Helvetica", 16, "bold"), fg="#ffad67").grid(row=0, column=0, columnspan=4, padx=10, pady=10)
    tk.Label(left_frame, text="En este apartado, puedes asignar estilos de Word a los campos del XML.").grid(row=1, column=0, columnspan=4, padx=10, pady=10)

    for idx, field in enumerate(xml_structure):
        tk.Label(left_frame, text=field).grid(row=idx + 2, column=0, padx=10, pady=5)
        
        type_var = tk.StringVar(value=config.get(field, {}).get('type', 'caracter'))
        style_var = tk.StringVar(value=config.get(field, {}).get('style', field))  # Default style name to field name

        tk.Radiobutton(left_frame, text="Párrafo", variable=type_var, value='parrafo').grid(row=idx + 2, column=1, padx=5)
        tk.Radiobutton(left_frame, text="Carácter", variable=type_var, value='caracter').grid(row=idx + 2, column=2, padx=5)
        tk.Entry(left_frame, textvariable=style_var).grid(row=idx + 2, column=3, padx=10, pady=5)

        fields[field] = (type_var, style_var)

    tk.Button(left_frame, text="Guardar estilos", command=update_config).grid(row=len(xml_structure) + 2, column=0, columnspan=4, pady=20)

    # Right column for file selection and processing buttons
    right_frame = tk.Frame(root)
    right_frame.grid(row=0, column=1, padx=10, pady=10, sticky='n')

    tk.Label(right_frame, text="Configuración del archivo y ejecución", font=("Helvetica", 16, "bold"), fg="#ffad67").grid(row=0, column=0, columnspan=3, padx=10, pady=10)
    tk.Label(right_frame, text="Selecciona el archivo XML, la carpeta de salida y el nombre del archivo de salida. Luego inicia el proceso de conversión.").grid(row=1, column=0, columnspan=3, padx=10, pady=10)

    tk.Label(right_frame, text="Seleccionar archivo XML:").grid(row=2, column=0, padx=10, pady=10)
    tk.Entry(right_frame, textvariable=xml_file_var, width=50).grid(row=2, column=1, padx=10, pady=10)
    tk.Button(right_frame, text="Browse", command=select_xml_file).grid(row=2, column=2, padx=10, pady=10)

    tk.Label(right_frame, text="Seleccionar carpeta de salida:").grid(row=3, column=0, padx=10, pady=10)
    tk.Entry(right_frame, textvariable=output_folder_var, width=50).grid(row=3, column=1, padx=10, pady=10)
    tk.Button(right_frame, text="Browse", command=select_output_folder).grid(row=3, column=2, padx=10, pady=10)

    tk.Label(right_frame, text="Nombre del archivo de salida:").grid(row=4, column=0, padx=10, pady=10)
    tk.Entry(right_frame, textvariable=output_file_var, width=50).grid(row=4, column=1, padx=10, pady=10)

    tk.Button(right_frame, text="Iniciar Proceso", command=start_processing).grid(row=5, column=1, pady=10)

    tk.Label(right_frame, textvariable=status_var).grid(row=6, column=1, padx=10, pady=10)

    # Log Text Area
    log_frame = tk.Frame(right_frame)
    log_frame.grid(row=7, column=0, columnspan=3, padx=10, pady=10, sticky='ew')

    tk.Label(log_frame, text="Logs del Proceso:").grid(row=0, column=0, padx=10, pady=10, sticky='w')
    log_text = tk.Text(log_frame, height=10, width=60, state=tk.DISABLED)
    log_text.grid(row=1, column=0, padx=10, pady=10, sticky='ew')
    log_scroll = tk.Scrollbar(log_frame, command=log_text.yview)
    log_scroll.grid(row=1, column=1, sticky='ns')
    log_text['yscrollcommand'] = log_scroll.set

    # Footer
    footer_frame = tk.Frame(right_frame)
    footer_frame.grid(row=8, column=0, columnspan=3, padx=10, pady=10, sticky='ew')
    
    footer_label = tk.Label(footer_frame, text="Programado por Francesc Xavier Escandell ", cursor="hand2", fg="#ffad67")
    footer_label.grid(row=0, column=0, sticky='e')
    footer_label.bind("<Button-1>", open_link)
    
    footer_link = tk.Label(footer_frame, text="Escandell.cat", cursor="hand2", fg="#ffad67")
    footer_link.grid(row=0, column=1, sticky='w')
    footer_link.bind("<Button-1>", open_link)

    root.mainloop()

if __name__ == "__main__":
    start_gui()
