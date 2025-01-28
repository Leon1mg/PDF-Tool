import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import fitz # PyMuPDF
import os
from win32com.client import Dispatch as CreateObject #pywin32
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2 import PdfMerger



root = tk.Tk()
root.title("PDF-Tool")
root.state('zoomed')
root.iconbitmap("Images/icon.ico")

pdf_document = None
current_page = 0
zoom_level = 2

def create_menuleiste(root):
    # Menüleiste Start -----------------------------------------
    def toggle_fullscreen():
        is_fullscreen = root.attributes("-fullscreen")
        root.attributes("-fullscreen", not is_fullscreen)

    def toggle_window_mode():
        root.attributes("-fullscreen", False)

    def about():
        messagebox.showinfo("Help", "You can find detailed instructions on Github at: \n Github.com/leon1mg/PDF-Tool")

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



def open_startpage():
    for widget in root.winfo_children():
        widget.destroy()

    title_label = tk.Label(root, text="Willkommen im PDF-Tool!", font=("Arial", 65, "bold"), pady=50)
    title_label.pack(pady=60)

    options_frame = tk.Frame(root)
    options_frame.pack(pady=150)

    img_merge = ImageTk.PhotoImage(Image.open("Images/kombinieren.png").resize((150, 220)))
    img_delete = ImageTk.PhotoImage(Image.open("Images/löschen.png").resize((150, 220)))
    img_crypt = ImageTk.PhotoImage(Image.open("Images/Verschlüsselung.png").resize((170, 220)))
    img_convert = ImageTk.PhotoImage(Image.open("Images/Converter.png").resize((185, 220)))
    img_extract = ImageTk.PhotoImage(Image.open("Images/extract.png").resize((150, 220)))
    img_watermark = ImageTk.PhotoImage(Image.open("Images/watermark.png").resize((150, 220)))


    merge_frame = tk.Frame(options_frame)
    merge_frame.pack(side=tk.LEFT, padx=20)

    merge_image_label = tk.Label(merge_frame, image=img_merge)
    merge_image_label.image = img_merge  # Verhindere Garbage Collection
    merge_image_label.pack()

    btn_merge = tk.Button(merge_frame, text="PDFs zusammenfügen", command=open_pdf_merge, width=20, font=("Arial", 14, "bold"))
    btn_merge.pack(pady=10)


    delete_frame = tk.Frame(options_frame)
    delete_frame.pack(side=tk.LEFT, padx=20)

    delete_image_label = tk.Label(delete_frame, image=img_delete)
    delete_image_label.image = img_delete  # Verhindere Garbage Collection
    delete_image_label.pack()

    btn_delete = tk.Button(delete_frame, text="PDF-Seiten löschen", command=open_pdf_delete_pages, width=20, font=("Arial", 14, "bold"))
    btn_delete.pack(pady=10)


    encrypt_frame = tk.Frame(options_frame)
    encrypt_frame.pack(side=tk.LEFT, padx=20)

    encrypt_image_label = tk.Label(encrypt_frame, image=img_crypt)
    encrypt_image_label.image = img_crypt  # Verhindere Garbage Collection
    encrypt_image_label.pack()

    btn_encrypt = tk.Button(encrypt_frame, text="PDF verschlüsseln", command=open_pdf_encrypt, width=20, font=("Arial", 14, "bold"))
    btn_encrypt.pack(pady=10)


    convert_frame = tk.Frame(options_frame)
    convert_frame.pack(side=tk.LEFT, padx=20)

    convert_image_label = tk.Label(convert_frame, image=img_convert)
    convert_image_label.image = img_convert  # Verhindere Garbage Collection
    convert_image_label.pack()

    btn_converter = tk.Button(convert_frame, text="Konverter", command=open_pdf_converter, width=20, font=("Arial", 14, "bold"))
    btn_converter.pack(pady=10)

    extract_frame = tk.Frame(options_frame)
    extract_frame.pack(side=tk.LEFT, padx=20)

    extract_image_label = tk.Label(extract_frame, image=img_extract)
    extract_image_label.image = img_extract  # Verhindere Garbage Collection
    extract_image_label.pack()

    btn_extract = tk.Button(extract_frame, text="Text extrahieren", command=open_pdf_extract_text, width=20, font=("Arial", 14, "bold"))
    btn_extract.pack(pady=10)

    create_menuleiste(root)

    watermark_frame = tk.Frame(options_frame)
    watermark_frame.pack(side=tk.LEFT, padx=20)

    watermark_image_label = tk.Label(watermark_frame, image=img_watermark)
    watermark_image_label.image = img_watermark  # Verhindere Garbage Collection
    watermark_image_label.pack()

    btn_watermark = tk.Button(watermark_frame, text="Wasserzeichen einfügen", command=open_pdf_watermark, width=20, font=("Arial", 14, "bold"))
    btn_watermark.pack(pady=10)






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

    create_menuleiste(root)
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
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom_level, zoom_level)) # Zoom anwenden
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img_tk = ImageTk.PhotoImage(img)

        canvas.create_image(0, 0, anchor=tk.NW, image=img_tk)
        canvas.image = img_tk # Verhindert Garbage Collection des Bildes

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

    create_menuleiste(root)


def open_pdf_encrypt():
    for widget in root.winfo_children():
        widget.destroy()

    def encrypt_pdf():
        input_file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if not input_file:
            return

        output_file = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if not output_file:
            return

        password = password_entry.get()
        if not password:
            messagebox.showwarning("Fehler", "Bitte ein Passwort eingeben!")
            return

        try:
            writer = PdfWriter()
            reader = PdfReader(input_file)

            # Inhalte in den Writer kopieren
            for page in reader.pages:
                writer.add_page(page)

            # Passwort setzen
            writer.encrypt(password)
            with open(output_file, "wb") as f:
                writer.write(f)

            messagebox.showinfo("Erfolg", f"PDF erfolgreich verschlüsselt und gespeichert unter: {output_file}")
        except Exception as e:
            messagebox.showerror("Fehler", f"Ein Fehler ist aufgetreten: {e}")

    def decrypt_pdf():
        input_file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if not input_file:
            return

        output_file = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if not output_file:
            return

        password = password_entry.get()
        if not password:
            messagebox.showwarning("Fehler", "Bitte ein Passwort eingeben!")
            return

        try:
            reader = PdfReader(input_file)
            if reader.is_encrypted:
                reader.decrypt(password)
            else:
                messagebox.showwarning("Warnung", "Diese PDF ist nicht passwortgeschützt.")
                return

            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)

            with open(output_file, "wb") as f:
                writer.write(f)

            messagebox.showinfo("Erfolg", f"PDF erfolgreich entschlüsselt und gespeichert unter: {output_file}")
        except Exception as e:
            messagebox.showerror("Fehler", f"Ein Fehler ist aufgetreten: {e}")

    # UI
    tk.Button(root, text="Zurück zur Startseite", command=open_startpage).pack(pady=10)

    instructions = tk.Label(root, text="Wählen Sie eine PDF aus, um sie zu verschlüsseln oder zu entschlüsseln.", font=("Arial", 14))
    instructions.pack(pady=20)

    password_label = tk.Label(root, text="Passwort:")
    password_label.pack(pady=5)
    password_entry = tk.Entry(root, width=30, show="*")
    password_entry.pack(pady=5)

    encrypt_btn = tk.Button(root, text="PDF verschlüsseln", command=encrypt_pdf)
    encrypt_btn.pack(pady=10)

    decrypt_btn = tk.Button(root, text="PDF entschlüsseln", command=decrypt_pdf)
    decrypt_btn.pack(pady=10)

    create_menuleiste(root)


def open_pdf_converter():
    for widget in root.winfo_children():
        widget.destroy()

    def convert_file():
        input_file = filedialog.askopenfilename(filetypes=[
            ("Alle unterstützten Dateien", "*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;*.png;*.jpeg;*.jpg"),
            ("Word Dateien", "*.doc;*.docx"),
            ("Excel Dateien", "*.xls;*.xlsx"),
            ("PowerPoint Dateien", "*.ppt;*.pptx"),
            ("Bilder", "*.png;*.jpeg;*.jpg")
        ])
        if not input_file:
            return

        output_file = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not output_file:
            return

        try:
            ext = os.path.splitext(input_file)[1].lower()
            if ext in ['.doc', '.docx']:
                word = CreateObject("Word.Application")
                doc = word.Documents.Open(input_file)
                doc.SaveAs(output_file, FileFormat=17)  # 17 = wdFormatPDF
                doc.Close()
                word.Quit()
            elif ext in ['.xls', '.xlsx']:
                excel = CreateObject("Excel.Application")
                wb = excel.Workbooks.Open(input_file)
                wb.ExportAsFixedFormat(0, output_file)  # 0 = xlTypePDF
                wb.Close()
                excel.Quit()
            elif ext in ['.ppt', '.pptx']:
                powerpoint = CreateObject("PowerPoint.Application")
                ppt = powerpoint.Presentations.Open(input_file)
                ppt.SaveAs(output_file, 32)  # 32 = ppSaveAsPDF
                ppt.Close()
                powerpoint.Quit()
            elif ext in ['.png', '.jpeg', '.jpg']:
                image = Image.open(input_file)
                if image.mode in ("RGBA", "LA"):
                    image = image.convert("RGB")
                image.save(output_file, "PDF", resolution=100.0)
            else:
                raise ValueError(f"Das Format {ext} wird nicht unterstützt.")
            messagebox.showinfo("Erfolg", f"Datei erfolgreich konvertiert und unter {output_file} gespeichert.")
        except Exception as e:
            messagebox.showerror("Fehler", f"Ein Fehler ist aufgetreten: {e}")

    back_btn = tk.Button(root, text="Zurück zur Startseite", command=open_startpage)
    back_btn.pack(pady=10)

    instructions = tk.Label(root, text="Wählen Sie eine Datei, um sie in PDF zu konvertieren.", font=("Arial", 14))
    instructions.pack(pady=20)
    instructions = tk.Label(root, text="Zur Auswahl stehen: doc, docx, xls, xlsx, ppt, pptx, png, jpg, jpeg --> PDF", font=("Arial", 14))
    instructions.pack(pady=20)


    convert_btn = tk.Button(root, text="Datei auswählen und konvertieren", command=convert_file, font=("Arial", 12))
    convert_btn.pack(pady=10)

    create_menuleiste(root)


def open_pdf_extract_text():
    for widget in root.winfo_children():
        widget.destroy()

    def extract_text():
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if not file_path:
            return

        try:
            pdf_document = fitz.open(file_path)
            extracted_text = ""
            for page in pdf_document:
                extracted_text += page.get_text()

            # Erstelle ein Textfeld mit einem Scrollbereich
            text_area.delete(1.0, tk.END)  # Vorherigen Text löschen
            text_area.insert(tk.END, extracted_text)

        except Exception as e:
            messagebox.showerror("Fehler", f"Ein Fehler ist aufgetreten: {e}")

    back_btn = tk.Button(root, text="Zurück zur Startseite", command=open_startpage)
    back_btn.pack(pady=10)

    instructions = tk.Label(root, text="Wählen Sie eine PDF aus, um den Text zu extrahieren.", font=("Arial", 14))
    instructions.pack(pady=20)

    extract_btn = tk.Button(root, text="Text extrahieren", command=extract_text)
    extract_btn.pack(pady=10)

    # Textfeld mit Scrollbereich
    text_area_frame = tk.Frame(root)
    text_area_frame.pack(pady=10)

    text_area = tk.Text(text_area_frame, width=100, height=20)
    text_area.pack(side=tk.LEFT)

    scrollbar = tk.Scrollbar(text_area_frame, orient=tk.VERTICAL, command=text_area.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    text_area.config(yscrollcommand=scrollbar.set)

    create_menuleiste(root)

def open_pdf_watermark():
    for widget in root.winfo_children():
        widget.destroy()

    def add_watermark():
        input_pdf = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if not input_pdf:
            return

        watermark_text = watermark_text_entry.get()
        if not watermark_text:
            messagebox.showwarning("Fehler", "Bitte ein Wasserzeichen eingeben!")
            return

        output_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if not output_pdf:
            return

        try:
            # Öffne das PDF-Dokument
            doc = fitz.open(input_pdf)

            # Iteriere durch alle Seiten und füge das Wasserzeichen hinzu
            for page in doc:
                # Position des Wasserzeichens (Zentrum der Seite)
                pos = fitz.Point(page.rect.width / 2, page.rect.height / 2)

                # Füge Text hinzu ohne Rotation
                page.insert_text(
                    pos,  # Position (Mitte der Seite)
                    watermark_text,  # Text des Wasserzeichens
                    fontsize=50,  # Schriftgröße
                    color=(0.7, 0.7, 0.7),  # Farbe (Hellgrau)
                    render_mode=2  # Render-Modus für Transparenz
                )

            # Speichern des neuen PDFs mit Wasserzeichen
            doc.save(output_pdf)
            messagebox.showinfo("Erfolg", f"Wasserzeichen hinzugefügt und PDF gespeichert unter: {output_pdf}")
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Hinzufügen des Wasserzeichens: {e}")

    back_btn = tk.Button(root, text="Zurück zur Startseite", command=open_startpage)
    back_btn.pack(pady=10)

    instructions = tk.Label(root, text="Geben Sie ein Wasserzeichen ein und fügen Sie es der PDF hinzu.", font=("Arial", 14))
    instructions.pack(pady=20)

    watermark_text_label = tk.Label(root, text="Wasserzeichen-Text:")
    watermark_text_label.pack(pady=5)

    watermark_text_entry = tk.Entry(root, width=40)
    watermark_text_entry.pack(pady=5)

    add_watermark_btn = tk.Button(root, text="Wasserzeichen hinzufügen", command=add_watermark)
    add_watermark_btn.pack(pady=10)

    create_menuleiste(root)


# Starte die Anwendung
open_startpage()
root.mainloop()

