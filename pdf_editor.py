import pikepdf
from pdf2docx import Converter
from docx2pdf import convert
from pdf2image import convert_from_path
from pikepdf import Pdf,Name,PdfImage,Page,Rectangle
from glob import glob

print("--------Please select 1-option for PDF operation")
print("1 :- PDF to WORD")
print("2 :- WORD to PDF")
print("3 :- PDF to IMAGE")
print("4 :- Extract Image from PDF")
print("5 :- Add watermarks/Overlays on PDF")
print("6 :- Merge PDF")
print("7 :- Reverse PDF")
print("8 :- Delete page from PDF")
print("9 :- Move Page in PDF")
print("10 :- Split PDF (Convert into single page pdf)")
print("11 :- Encrypt PDF (Set password)")
print("12 :- Rotate PDF file")

a = int(input("Select option :- "))

match a:
    case 1:
        obj = Converter(input("Enter .PDF path :-"))
        obj.convert(input("Enter name for .DOCX as you want :- "))
        obj.close()
        print("-convert Successfull !!!")

    case 2:
        convert(input("Enter .DOCX name :-"), input("Enter name for .PDF as you want :-"))
        print("-convert Successfull !!!")

    case 3:
        doc = input("Enter .PDF Path :- ")
        sv = input("Enter name for.JPEG image :- ")
        old_pdf = convert_from_path(doc, poppler_path =r"D:\project\python project\PDF Editor\poppler-23.07.0\Library\bin")  # download poppler file from python site and gie path here
        for i in range(len(old_pdf)):
            old_pdf[i].save(sv+str(i)+".jpg","JPEG")
        print("-convert Successfull !!!")

    case 4:
        old_pdf = Pdf.open(input("Enter .PDF name :- "))  #p1.pdf
        a = int(input("Enter page No.(that can be has a image) :- "))
        page_1 = old_pdf.pages[a]
        print("code of Images :- ",list(page_1.images.keys()))    # ['/X8']     <- image code
        raw_image = page_1.images['/X8']
        pdf_image = PdfImage(raw_image)
        pdf_image.extract_to(fileprefix="test1")
        print("-convert Successfull !!!")

    case 5:
        old_pdf1 = Pdf.open(input("Enter .PDF 1 name :- "))
        old_pdf2 = Pdf.open(input("Enter .PDF 2 name :- "))
        a = int(input("Enter PDF 1 page index :- "))
        b = int(input("Enter PDF 2 page index :- "))
        des_page = Page(old_pdf1.pages[a])
        sur_page = Page(old_pdf2.pages[b])
        d = int(input("Position of X :- "))
        e = int(input("Position of y :- "))
        f = int(input("height :- "))
        g = int(input("width  :- "))
        des_page .add_overlay(sur_page,Rectangle(d,e,f,g))      #position 0,0  height-width 500,500
        old_pdf1.save(input("Enter name .PDF SAVE AS :- "))
        print("-convert Successfull !!!")

    case 6:
        new_pdf = Pdf.new()
        for file in glob("*.pdf"):
            old_pdf = Pdf.open(file)
            new_pdf.pages.extend(old_pdf.pages)
        new_pdf.save(input("Enter name .PDF SAVE AS :- "))
        print("Successfull !!!")

    case 7:
        old_pdf = pikepdf.Pdf.open(input("Enter .PDF name :- "))
        old_pdf.pages.reverse()
        old_pdf.save(input("Enter name .PDF SAVE AS :- "))
        print("Successfull !!!")

    case 8:
        old_pdf = pikepdf.Pdf.open(input("Enter .PDF name :- "))
        a=int(input("Enter start page No. :- "))
        b=int(input("Enter end page No. :- "))
        del old_pdf.pages[a:b]
        old_pdf.save(input("Enter name .PDF SAVE AS :- "))
        print("Successfull !!!")

    case 9:
        old_pdf = pikepdf.Pdf.open(input("Enter .PDF name :- "))
        a = int(input("Page index (you want to move) :- "))
        b = int(input("Page index (where you want to move) :- "))
        old_pdf.pages[b] = old_pdf.pages[a]
        del old_pdf.pages[a]      #delete selected page
        old_pdf.save(input("Enter name .PDF SAVE AS :- "))
        print("Successfull !!!")

    case 10:
        old_pdf = pikepdf.Pdf.open(input("Enter .PDF name :- "))
        for n,pag_can in enumerate(old_pdf.pages):
            new_pdf = Pdf.new()
            new_pdf.pages.append(pag_can)
            new_pdf.save("new"+str(n)+".pdf")
        print("Successfull !!!")

    case 11:
        old_pdf = pikepdf.Pdf.open(input("Enter .PDF name :- "))
        no_ext = pikepdf.Permissions(extract=False)
        old_pdf.save(input("Enter name .PDF SAVE AS :- "),encryption=pikepdf.Encryption(user=input("Password :- "),owner=input("Owner name :- "),allow=no_ext))
        print("Successfull !!!")

    case 12:
        old_pdf = pikepdf.Pdf.open(input("Enter .PDF name :- "))
        a = int(input("Enter Rotate Degree :- "))
        st = int(input("Start No :- "))
        ed = int(input("End No :- "))
        b = input("Enter name .PDF SAVE AS :- ")
        for i in old_pdf.pages[st:ed]:
            i.Rotate = a
            old_pdf.save(b)
        print("Successfull !!!")

    case _:
        print("invalid choice !!! please try again")


