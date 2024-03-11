import os
import sys
import PyPDF2 as pdf
import customtkinter as ctk
import fitz

def resource_path(relative_path):
    #Get absolute path to resource, works for dev and for PyInstaller
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


folder_semnate = "Extrase_semnate"
for file_name in os.listdir(resource_path(folder_semnate)):
    file_path = os.path.join(resource_path(folder_semnate), file_name)
    os.remove((resource_path(file_path)))


ctk.set_appearance_mode('dark')
ctk.set_default_color_theme('blue')

root = ctk.CTk()
root.title("Extras de cont")
root.geometry("300x220")

frame = ctk.CTkFrame(root,fg_color="#1F1F1F")
frame.pack(padx=5, pady=5)

#semnatura
path_furnizori = "Extrase\Extras de cont furnizori.PDF"
path_furnizori_valuta = "Extrase\Extras de cont furnizori valuta.PDF"
path_clienti = "Extrase\Extras de cont clienti.PDF"

rect_furnizori = fitz.Rect(468, 495, 548, 615)
rect_clienti = fitz.Rect(463, 490, 543, 610)
rect_furnizori_valuta = fitz.Rect(473, 500, 553, 580)

#semnatura

path_furnizori_semnat = "Extrase_semnate\Extras de cont furnizori.PDF"
path_furnizori_valuta_semnat = "Extrase_semnate\Extras de cont furnizori valuta.PDF"
path_clienti_semnat = "Extrase_semnate\Extras de cont clienti.PDF"
path_lista_parteneri = "Parteneri.txt"

def add_parteneri(fisier_parteneri, partener):
    with open(resource_path(fisier_parteneri), "a") as file:
        file.write(partener + '\n')
def delete_lista(fisier_parteneri):
    with open(resource_path(fisier_parteneri), 'w') as file:
        file.write("")

def furnizori_lei():
    delete_lista("Parteneri_Furnizori lei.txt")

    pdf_doc_furnizori = fitz.open(resource_path(path_furnizori))
    for page_furnizori in pdf_doc_furnizori:
        page_furnizori.insert_image(rect_furnizori, filename=resource_path("semnatura.png"))
        pdf_doc_furnizori.save(
            resource_path("Extrase_semnate\Extras de cont furnizori.PDF"))  # , incremental = True, encryption=0

    file_furnizori = open(resource_path(path_furnizori_semnat), "rb")
    pdf_reader_furnizori = pdf.PdfReader(file_furnizori)
    writer_furnizori = pdf.PdfWriter()

    for page in range(len(pdf_reader_furnizori.pages)):
        text_furnizori = pdf_reader_furnizori.pages[page].extract_text()
        selected_page_furnizori = pdf_reader_furnizori.pages[page]
        find_total_furnizori = text_furnizori.find("Total:")

        if find_total_furnizori > 0:
            writer_furnizori.add_page(selected_page_furnizori)
            start_index_furnizori = text_furnizori.index("creditoare: ")+len("creditoare: ")
            end_index_furnizori = text_furnizori.index(" Nr. de")
            nume_partener_furnizori = text_furnizori[start_index_furnizori:end_index_furnizori].rstrip()
            print(nume_partener_furnizori)


            file_name_furnizori = f"Split_PDF\{nume_partener_furnizori}.pdf"
            with open(resource_path(file_name_furnizori), "wb") as out:
                writer_furnizori.write(out)
            writer_furnizori = pdf.PdfWriter()

            add_parteneri("Parteneri_Furnizori lei.txt", nume_partener_furnizori)


        else:
            writer_furnizori.add_page(selected_page_furnizori)

def furnizori_valuta():
    delete_lista("Parteneri_Furnizori valuta.txt")

    pdf_doc_furnizori_valuta = fitz.open(resource_path(path_furnizori_valuta))
    for page_furnizori_valuta in pdf_doc_furnizori_valuta:
        page_furnizori_valuta.insert_image(rect_furnizori_valuta, filename=resource_path("semnatura.png"))
        pdf_doc_furnizori_valuta.save(
            resource_path("Extrase_semnate\Extras de cont furnizori valuta.PDF"))  # , incremental = True, encryption=0

    file_furnizori_valuta = open(resource_path(path_furnizori_valuta_semnat), "rb")
    pdf_reader_furnizori_valuta = pdf.PdfReader(file_furnizori_valuta)
    writer_furnizori_valuta = pdf.PdfWriter()

    for page_valuta in range(len(pdf_reader_furnizori_valuta.pages)):
        text_furnizori_valuta = pdf_reader_furnizori_valuta.pages[page_valuta].extract_text()
        selected_page_furnizori_valuta = pdf_reader_furnizori_valuta.pages[page_valuta]
        find_total_furnizori_valuta = text_furnizori_valuta.find("Total:")

        if find_total_furnizori_valuta > 0:
            writer_furnizori_valuta.add_page(selected_page_furnizori_valuta)
            start_index_furnizori_valuta = text_furnizori_valuta.index("To : ") + len("To : ")
            end_index_furnizori_valuta = text_furnizori_valuta.index("Address:")
            nume_partener_furnizori_valuta = text_furnizori_valuta[start_index_furnizori_valuta:end_index_furnizori_valuta].rstrip()

            if "/" in nume_partener_furnizori_valuta:
                nume_partener_furnizori_valuta = nume_partener_furnizori_valuta.replace("/","")
            print(nume_partener_furnizori_valuta)



            file_name_furnizori_valuta = f"Split_PDF\{nume_partener_furnizori_valuta}.pdf"
            with open(resource_path(file_name_furnizori_valuta), "wb") as out:
                writer_furnizori_valuta.write(out)
            writer_furnizori_valuta = pdf.PdfWriter()

            add_parteneri("Parteneri_Furnizori valuta.txt",nume_partener_furnizori_valuta)
        else:
            writer_furnizori_valuta.add_page(selected_page_furnizori_valuta)

def clienti_lei():
    delete_lista("Parteneri_Clienti lei.txt")

    pdf_doc_clienti = fitz.open(resource_path(path_clienti))
    for page_clienti in pdf_doc_clienti:
        page_clienti.insert_image(rect_clienti, filename=resource_path("semnatura.png"))
        pdf_doc_clienti.save(resource_path("Extrase_semnate\Extras de cont clienti.PDF"))  # , incremental = True, encryption=0

    file_clienti = open(resource_path(path_clienti_semnat), "rb")
    pdf_reader_clienti = pdf.PdfReader(file_clienti)
    writer_clienti = pdf.PdfWriter()

    for page_clienti in range(len(pdf_reader_clienti.pages)):
        text_clienti = pdf_reader_clienti.pages[page_clienti].extract_text()
        selected_page_clienti = pdf_reader_clienti.pages[page_clienti]
        find_total_clienti = text_clienti.find("Total:")

        if find_total_clienti > 0:
            writer_clienti.add_page(selected_page_clienti)
            start_index_clienti = text_clienti.index("Catre : ") + len("Catre : ")
            end_index_clienti = text_clienti.index(" / ")
            nume_partener_clienti = text_clienti[start_index_clienti:end_index_clienti].rstrip()
            print(nume_partener_clienti)



            file_name_clienti = f"Split_PDF\{nume_partener_clienti}.pdf"
            with open(resource_path(file_name_clienti), "wb") as out:
                writer_clienti.write(out)
            writer_clienti = pdf.PdfWriter()

            add_parteneri("Parteneri_Clienti lei.txt", nume_partener_clienti)
        else:
            writer_clienti.add_page(selected_page_clienti)

label = ctk.CTkLabel(frame, text="Alege extrasul de cont:", height=40, font=('Rockwell Bold', 20), text_color="#CDC6E1")
label.pack(padx=5, pady=5)

FurnizoriLei = ctk.CTkButton(frame, text = "Furnizori lei",height=35, font=('Rockwell Bold', 18), text_color="#CDC6E1", command=furnizori_lei )
FurnizoriLei.pack(fill = ctk.X, padx=5, pady=7)

ClientiLei = ctk.CTkButton(frame, text="Clienti lei",height=35, font=('Rockwell Bold', 18), text_color="#CDC6E1", command=clienti_lei)
ClientiLei.pack(fill = ctk.X, side=ctk.TOP, padx=5, pady=7)

FurnizoriValuta = ctk.CTkButton(frame, text="Furnizori valuta",height=35, font=('Rockwell Bold', 18), text_color="#CDC6E1", command=furnizori_valuta)
FurnizoriValuta.pack(fill = ctk.X, side=ctk.BOTTOM, padx=5, pady=7)

root.mainloop()







