from docxtpl import DocxTemplate, InlineImage
from docx2pdf import convert
import pdfplumber
import flet as ft
import os
import re


def _data_(all_text):
    text = all_text.split("\n")
    index = []
    for i in range(len(text)):
        if ("pièce" or "pièces") in text[i]:
            index.append(i)
    line = " "
    all_line = []
    for i in range(len(index) - 2):
        for j in range(index[i], index[i+1]):
            line += text[j] + "\n"
        all_line.append(line.strip().split("\n"))
        line = " "
    quantite = []
    values = []
    numero = ""
    val = ''
    manu = ""
    ean = ""
    numSerie = ""
    tot2 = 0
    for article in all_line:
        cmp = article[0].split()
        for num in cmp:
            try:
                if int(num) and len(num) > 6:
                    numero = num
            except:
                pass
        quantite.append(cmp[0])
        mult = int(cmp[0].split('p')[0])

        for des in article:
            if 'Manufacturer' not in des:
                val += des + ' '
            else:
                break
        for i in range(len(cmp)):
            if cmp[i] == "0%":
                temp = i
                unit = cmp[i+1]
        for i in range(temp, len(cmp)):
            val = val.replace("0%", "").replace(
                str(numero), "").replace(cmp[i], "")
        if "pièce" in val:
            val = val.split('pièce')[1].strip()
        else:
            val = val.split('pièce')[2].strip()
        for des in article:
            if 'Manufacturer' in des:
                manu = des
            if 'EAN' in des:
                ean = des
            if 'Numéro de série' in des:
                numSerie = des
        try:
            val1 = val.split('    ')[0]
            val2 = val.split('    ')[1]
        except:
            val1 = val,
            val2 = ""
        mn = manu if manu != "" else ""
        ea = ean if ean != "" else ""
        nums = numSerie if numSerie != "" else ""

        values.append([cmp[0], val1, val2, mn, ea, nums, numero, "20%", str(round(float(
            unit.replace(",", '.'))/1.2, 2)).replace(".", ','), str(round(float(unit.replace(",", '.'))/1.2, 2)*mult).replace(".", ',')])
        val = ""
        tot2 += round(float(unit.replace(",", '.'))/1.2, 2)*mult
    return [values, tot2]


def getPdfData(path):
    with pdfplumber.open(path) as pdf:
        page = pdf.pages[0]
        text = page.extract_text()
    # Facture
    fac_re = re.compile("(?<=Facture).+")
    if fac_re.search(text):
        facture = fac_re.search(text).group(0).split()[0]
    else:
        facture = " "
    # Numero du document
    docName_re = re.compile("(?<=cNoummméruon aduet adiorecument).+")
    if docName_re.search(text):
        docName = docName_re.search(text).group(0).strip()
    else:
        docName = ' '
    # Numero
    No_re = re.compile("(?<=fNaoc.turation).+")
    if No_re.search(text):
        Numero = No_re.search(text).group(0).split()[0].strip()
    else:
        Numero = " "
    # Date de creation
    dt_re = re.compile("(?<=Date de création).+")
    if dt_re.search(text):
        Date = dt_re.search(text).group(0).split()[0].strip()
    else:
        Date = " "
    # Numero tva intra
    ntva_re = re.compile("(?<=dNo°c tuvma einntrta- ).+")
    if ntva_re.search(text):
        NumTva = ntva_re.search(text).group(0).split()[0].strip()
    else:
        NumTva = " "
    # Methode de livraison
    met_re = re.compile("(?<=Compte client de 113789 Méthode de livraison).+")
    if met_re.search(text):
        LivMet = met_re.search(text).group(0)
    else:
        LivMet = " "
    # Contenu de la livraison
    cont_re = re.compile("(?<=Contenu de la livraison\n).+")
    if cont_re.search(text):
        ContLiv = cont_re.search(text).group(0)
    else:
        ContLiv = " "
    # Manufacturer
    manu_re = re.compile("(?<=Manufacturer).+")
    if manu_re.findall(text):
        Manufacturer = manu_re.findall(text)
    else:
        Manufacturer = []
    # EAN
    ean_re = re.compile("(?<=EAN).+")
    if ean_re.findall(text):
        Ean = ean_re.findall(text)
    else:
        Ean = []
    # Num series
    serie_re = re.compile("(?<=Numéro de série ).+")
    if serie_re.findall(text):
        serie = serie_re.findall(text)
    else:
        serie = []
    # Qté
    qte_re = re.finditer("[0-9](pièce)?(.*)", "".join(text))
    Qte = []
    for line in qte_re:
        if line:
            if "pièce" in line.group(0):
                Qte.append(line.group(0).split()[0])

    art_re = re.compile("([0-9]pièc[a-z] )(.+\n)([0-9]pièc[a-z] ).+ ")
    art_ = []
    tot1 = 0
    art_data = []
    if art_re.search(text):
        article = art_re.search(text).group(0).split("\n")
        for art in article:
            art_.append(art.split())

        for art in art_:
            try:
                if len(art) > 2:
                    art_data.append([art[0], art[1] + " "+art[2], art[3], "20%", str(round(float(
                        art[5].replace(",", "."))/1.2, 2)).replace(".", ','), str(round(float(art[5].replace(",", "."))/1.2, 2)).replace(".", ',')])
                    tot1 += round(float(art[5].replace(",", "."))/1.2, 2)

                else:
                    art_data.append([art[0], art[1]])
            except:
                art_data = []
    else:
        art_data = []
    total = float(_data_(text)[1] + tot1)
    adresseData = getAdresse(text)
    data = {
        "header": {
            "facture": facture,
            "numero_document": docName,
            "numero": Numero,
            "date": Date,
            "numero_tva": NumTva,
            "methode_livraison": LivMet,
            "contenu_livraison": ContLiv
        },
        "container": {
            "Manufacturer": Manufacturer,
            "EAN": Ean,
            "numero_serie": serie,
            "data2": art_data,
            "total": total,
            "data": _data_(text)[0]
        },
        "pays": adresseData["pays"],
        "nom_client": adresseData["nom_client"],
        "tel": adresseData["tel"],
        "adresse1": adresseData["adresse1"],
        "adresse2": adresseData["adresse2"],
        "adresse3": adresseData["adresse3"]
    }

    return data


class PdfCreator():
    def __init__(self, data):
        super(PdfCreator, self).__init__()
        self.data = data

    def generate_pdf(self, output_name, filename, template, image):
        doc = DocxTemplate(template)
        for item in self.data["container"]["data"]:
            item.append(InlineImage(doc, image))

        total_ttc = str(round(self.data["container"]["total"] + self.data["container"]["total"]*0.2, 2)).replace('.', ",")
        doc.render({
            "facture": self.data["header"]["facture"],
            "numero_document": self.data["header"]["numero_document"],
            "numero": self.data["header"]["numero"],
            "date": self.data["header"]["date"],
            "total": str(round(self.data["container"]["total"],2)).replace(".", ","),
            "totaL_euro": str(round(self.data["container"]["total"]*0.2,2)).replace('.', ','),
            "ttc_euro": total_ttc,
            "numero_tva": self.data["header"]["numero_tva"],
            "methode_livraison": self.data["header"]["methode_livraison"],
            "contenu_livraison": self.data["header"]["contenu_livraison"],
            "data": self.data["container"]["data"],
            "data2": self.data["container"]["data2"],
            "nom_client": self.data["nom_client"],
            "pays": self.data["pays"],
            "tel": self.data["tel"],
            "adresse1": self.data["adresse1"],
            "adresse2": self.data["adresse2"],
            "adresse3": self.data["adresse3"]
        }, autoescape=True)
        doc_name = filename.replace(".pdf", "")+'.docx'
        doc.save(doc_name)
        convert(doc_name, output_name + '/' + filename)
        os.remove(doc_name)


def usage():
    use = """
    1- Copier les chemins des fichiers qui vous sont fournis et 
            renseignez les chemins dans chaque entrée
    
    2- Coller ce lien dans le champs réservé à cette fin
    
    3- Lancez l'opération
    
    4- Suivez l'évolution du traitement des fichiers
    
    5- A la fin, ouvrez votre explorateur de fichier et accédez au dossier de sortie
    
    NB: Assurez vous que tous vos fichiers sont conçus de la même façon et fermez MS
            Word si c'est ouvert, avant de lancer les opérations
    
    """
    return use


def getAdresse(text):
    data = text.split("\n")
    index = []
    adresse = []
    for i in range(len(data)):
        if "Adresse de livraison" in data[i]:
            index.append(i+1)
            break
    for i in range(len(data)):
        if "Envie de recevoir plus vite vos factures" in data[i]:
            index.append(i)
            break
    if len(index) > 1:
        for i in range(index[0], index[1]):
            adresse.append(data[i])
        if len(adresse) > 6:
            adresse3 = adresse[4].split("France")[0],
        else:
            adresse3 = ""
        adresseData = {
            "nom_client": adresse[0].split('ECOMLG')[0],
            "pays": adresse[len(adresse) - 1],
            "tel": adresse[1].split()[0],
            "adresse1": adresse[2].split("137 AVENUE DE LA REPUBLIQUE")[0],
            "adresse2": adresse[3].split("26270 LORIOL SUR DROME")[0],
            "adresse3": adresse3
        }
    else:
        adresseData = {
            "nom_client": "",
            "pays": "",
            "tel": "",
            "adresse1": "",
            "adresse2": "",
            "adresse3": ""
        }

    return adresseData


def main(page: ft.Page):
    page.title = "PDF TREAT"
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.horizontal_alignment = ft.MainAxisAlignment.CENTER
    page.theme_mode = "light"
    page.window_width = 800
    page.window_height = 600

    counter = 0
    chemin = ft.Text(value="", text_align='center', width=350)
    length = ft.Text(value="", text_align="center")
    treat = ft.Text(value=f"{counter}", text_align="center", visible=False)
    progress = ft.ProgressBar(width=200, visible=False)
    folder = ft.TextField(label="Le chemin vers votre dossier",
                          prefix_icon=ft.icons.FOLDER, border_radius=30, width=350)
    folder2 = ft.TextField(label="Le chemin vers le fichier template",
                          prefix_icon=ft.icons.FOLDER, border_radius=30, width=350)

    folder3 = ft.TextField(label="Le chemin vers le fichier image",
                          prefix_icon=ft.icons.FOLDER, border_radius=30, width=350)

    def close_dlg(e):
        dlg_modal.open = False
        page.update()

    dlg_modal = ft.AlertDialog(
        modal=True,
        title=ft.Text("Guide d'utilisation"),
        content=ft.Text(usage()),
        actions=[
            ft.TextButton("Compris", on_click=close_dlg)
        ],
        actions_alignment=ft.MainAxisAlignment.END
    )

    openFolder = ft.Container(content=ft.Text("\n\n\n\nTraitement terminé 👏! \n\nTous les fichiers ont\n été traités avec succès. \n\nParcourez le dossier de sortie pour voir vos \nnouveaux fichiers\n\n\n\n"
                             , color="white", text_align="center") ,bgcolor="blue" ,border_radius=10 ,padding=10 ,visible=False, width=300)

    def open_modal(e):
        dlg_modal.open = True
        page.update()

    def operation(e):
        if folder.value == "":
            chemin.value = "Le chemin vers le dossier est vide !"
        else:
            if os.path.isdir(folder.value):
                os.makedirs(folder.value + "/pdf_treat", exist_ok=True)
                chemin.value = "Le chemin de sortie est " + folder.value + "/pdf_treat"
                # faire tous les traitements ici
                pdf_files = [f for f in os.listdir(
                    folder.value) if f.endswith(".pdf")]
                length.value = str(len(pdf_files)) + ' fichiers pdf au total'
                progress.visible = True
                treat.visible = True

                for filePath in pdf_files:
                    treat.value = str(int(treat.value)+1)
                    data = getPdfData(folder.value+"/"+filePath)
                    PdfCreator(data).generate_pdf(
                        folder.value + "/pdf_treat", filePath, template=folder2.value, image = folder3.value)
                    page.update()

                if int(treat.value) == len(pdf_files):
                    openFolder.visible = True
                    treat.value = 0
                    progress.visible = False
                    treat.visible = False
                    chemin.value = ""
                    folder.value = ""
                    length.value = ""
                    page.update()
            else:
                chemin.value = "Le chemin vers le dossier est invalide ! Réessayez svp"
        page.update()

    page.add(
        ft.Row(
            [
                ft.Column([
                    ft.Container(
                        ft.Container(
                            ft.Column([
                                ft.Container(
                                    content=ft.Text(
                                        "Help", text_align="center", height=50, color="grey"),
                                    padding=0,
                                    on_click=open_modal,
                                ),
                                ft.Text("Bienvenue sur PDF TREAT", height=150,
                                        size=25, weight="bold", text_align="center"),
                                folder2,
                                folder3,
                                folder,
                                chemin,
                                ft.Container(
                                    content=ft.Text(
                                        "Lancer l'opération", color="white", text_align="center"),
                                    bgcolor="blue",
                                    border_radius=30,
                                    padding=10,
                                    on_click=operation,
                                    width=350,
                                    height=45
                                ),
                            ], alignment=ft.MainAxisAlignment.CENTER),
                        ),
                        padding=30,
                    )
                ]),
                ft.Column([
                    ft.Container(
                        ft.Container(
                            ft.Column([
                                length,
                                progress,
                                treat,
                                openFolder
                            ], alignment=ft.MainAxisAlignment.CENTER),
                        ),
                        padding=30,
                    )
                ]),
                dlg_modal
            ],
            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
        )
    )


ft.app(target=main)
