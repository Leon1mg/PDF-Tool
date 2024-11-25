import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import fitz  # PyMuPDF


# Hauptfenster erstellen
root = tk.Tk()
root.title("PDF-Tool")
root.state('zoomed')
root.state('zoomed')
root.iconbitmap("icon.ico")

# Globale Variablen
pdf_document = None
current_page = 0
zoom_level = 2  # Standardzoom für die Vorschau


# Funktion für die Startseite
def open_startpage():
    for widget in root.winfo_children():
        widget.destroy()

    title_label = tk.Label(root, text="Willkommen im PDF-Tool!", font=("Arial", 50, "bold"), pady=100)
    title_label.pack(pady=20)

    btn_merge = tk.Button(root, text="PDF zusammenfügen", command=open_pdf_merge, width=20, font=("Arial", 14, "bold"))
    btn_merge.pack(pady=10)

    btn_delete = tk.Button(root, text="PDF-Seiten löschen", command=open_pdf_delete_pages, width=20, font=("Arial", 14, "bold"))
    btn_delete.pack(pady=10)


def open_pdf_delete_pages():
    for widget in root.winfo_children():
        widget.destroy()


    global pdf_document, current_page
    pdf_document = None
    current_page = 0

    def load_pdf():
        global pdf_document, current_page
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if not file_path:
            return

        try:
            pdf_document = fitz.open(file_path)
            current_page = 0
            update_page_preview()
        except Exception as e:
            messagebox.showerror("Fehler", f"PDF konnte nicht geladen werden: {e}")

    def update_page_preview():
        if not pdf_document:
            return


        page = pdf_document.load_page(current_page)
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom_level, zoom_level))  # Zoom anwenden
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img_tk = ImageTk.PhotoImage(img)

        canvas.create_image(0, 0, anchor=tk.NW, image=img_tk)
        canvas.image = img_tk  # Verhindert Garbage Collection des Bildes

        canvas.config(scrollregion=canvas.bbox("all"))

    def delete_pages():
        global pdf_document
        if not pdf_document:
            messagebox.showwarning("Warnung", "Keine PDF geladen!")
            return

        try:
            pages_to_delete = list(map(int, pages_entry.get().split(',')))
            writer = fitz.open()

            for i, page in enumerate(pdf_document):
                if i + 1 not in pages_to_delete:
                    writer.insert_pdf(pdf_document, from_page=i, to_page=i)

            save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
            if save_path:
                writer.save(save_path)
                writer.close()
                messagebox.showinfo("Erfolg", f"PDF gespeichert unter: {save_path}")
        except Exception as e:
            messagebox.showerror("Fehler", f"Ein Fehler ist aufgetreten: {e}")

    def back_to_start():
        open_startpage()

    tk.Button(root, text="Zurück", command=back_to_start).pack(pady=10)

    tk.Button(root, text="PDF Laden", command=load_pdf).pack(pady=10)

    pages_label = tk.Label(root, text="Zu löschende Seiten (z. B. 1,3,5):")
    pages_label.pack(pady=5)
    pages_entry = tk.Entry(root, width=30)
    pages_entry.pack(pady=5)

    delete_btn = tk.Button(root, text="Seiten löschen und speichern", command=delete_pages)
    delete_btn.pack(pady=10)

    canvas_frame = tk.Frame(root)
    canvas_frame.pack(pady=10)

    canvas = tk.Canvas(canvas_frame, width=1500, height=700)
    canvas.pack(side=tk.LEFT)

    scrollbar = tk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=canvas.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    canvas.config(yscrollcommand=scrollbar.set)

    def next_page():
        global current_page
        if current_page < len(pdf_document) - 1:
            current_page += 1
            update_page_preview()

    def prev_page():
        global current_page
        if current_page > 0:
            current_page -= 1
            update_page_preview()

    nav_frame = tk.Frame(root)
    nav_frame.pack(pady=10)

    prev_button = tk.Button(nav_frame, text="Zurück", command=prev_page)
    prev_button.pack(side=tk.LEFT, padx=10)

    next_button = tk.Button(nav_frame, text="Weiter", command=next_page)
    next_button.pack(side=tk.LEFT, padx=10)

    # Menüleiste Start -----------------------------------------
    def toggle_fullscreen():
        is_fullscreen = root.attributes("-fullscreen")
        root.attributes("-fullscreen", not is_fullscreen)

    def toggle_window_mode():
        root.attributes("-fullscreen", False)

    def about():
        messagebox.showinfo("Help", "You can find detailed instructions on Github at: ")

    menubar = tk.Menu(root)

    file_menu = tk.Menu(menubar, tearoff=0)
    file_menu.add_command(label="Beenden", command=root.quit)
    menubar.add_cascade(label="File", menu=file_menu)

    view_menu = tk.Menu(menubar, tearoff=0)
    view_menu.add_command(label="Full screen", command=toggle_fullscreen)
    view_menu.add_command(label="Window mode", command=toggle_window_mode)
    menubar.add_cascade(label="View", menu=view_menu)

    help_menu = tk.Menu(menubar, tearoff=0)
    help_menu.add_command(label="About", command=about)
    menubar.add_cascade(label="Help", menu=help_menu)

    root.config(menu=menubar)
    # Menüleiste Ende -------------------------------------------


def open_pdf_merge():
    for widget in root.winfo_children():
        widget.destroy()

    pdf_files = []

    def add_files():
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if files:
            pdf_files.extend(files)
            update_file_list()

    def update_file_list():
        file_list.delete(0, tk.END)
        for file in pdf_files:
            file_list.insert(tk.END, file)

    def merge_pdfs():
        if len(pdf_files) < 2:
            messagebox.showwarning("Warning", "At least 2 PDF files are required.")
            return

        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not save_path:
            return

        try:
            merger = PdfMerger()
            for pdf in pdf_files:
                merger.append(pdf)
            merger.write(save_path)
            merger.close()
            messagebox.showinfo("Success", f"Successfully saved under: {save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def clear_files():
        pdf_files.clear()
        update_file_list()

    back_btn = tk.Button(root, text="Zurück zur Startseite", command=open_startpage)
    back_btn.pack(pady=10)

    frame = tk.Frame(root)
    frame.pack(pady=10)

    file_list = tk.Listbox(frame, width=50, height=10)
    file_list.pack(side=tk.LEFT, padx=5)

    scrollbar = tk.Scrollbar(frame, orient=tk.VERTICAL, command=file_list.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    file_list.config(yscrollcommand=scrollbar.set)

    btn_add = tk.Button(root, text="PDFs hinzufügen", command=add_files)
    btn_add.pack(pady=5)

    btn_merge = tk.Button(root, text="PDFs zusammenfügen", command=merge_pdfs)
    btn_merge.pack(pady=5)

    btn_clear = tk.Button(root, text="Liste leeren", command=clear_files)
    btn_clear.pack(pady=5)

    # Menüleiste Start -----------------------------------------
    def toggle_fullscreen():
        is_fullscreen = root.attributes("-fullscreen")
        root.attributes("-fullscreen", not is_fullscreen)

    def toggle_window_mode():
        root.attributes("-fullscreen", False)

    def about():
        messagebox.showinfo("Help", "You can find detailed instructions on Github at:")

    menubar = tk.Menu(root)

    file_menu = tk.Menu(menubar, tearoff=0)
    file_menu.add_command(label="Beenden", command=root.quit)
    menubar.add_cascade(label="File", menu=file_menu)

    view_menu = tk.Menu(menubar, tearoff=0)
    view_menu.add_command(label="Full screen", command=toggle_fullscreen)
    view_menu.add_command(label="Window mode", command=toggle_window_mode)
    menubar.add_cascade(label="View", menu=view_menu)

    help_menu = tk.Menu(menubar, tearoff=0)
    help_menu.add_command(label="About", command=about)
    menubar.add_cascade(label="Help", menu=help_menu)

    root.config(menu=menubar)
    # Menüleiste Ende -------------------------------------------




# Starte die Anwendung
open_startpage()
root.mainloop()
