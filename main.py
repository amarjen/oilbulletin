# std libs
from pathlib import Path
import urllib.request
import filecmp
import shutil
import glob

# external libs
import xlrd
import PyPDF2
import xlsxwriter
import pandas as pd

class Boletines(object):
    """
    TODO: Refactorizar programa en una clase
    """
    def __init__(self):
        pass

    def descargar(self):
        pass

    def leer(self):
        pass
    
def descargar_boletines(forzar_descarga = False):
    """
    Descarga el listado PDF de boletines publicados de la CE
    y descarga los boletines xlsx o xlx que falten al directorio `raw_data`
    """
    # 1. descargar lista de boletines
    url = "https://ec.europa.eu/energy/observatory/reports/List-of-WOB.pdf"

    fichero = url.split('/')[-1].split('.')[0]
    urllib.request.urlretrieve(url, f"./{fichero}_tmp.pdf")

    # 2. comprobar si hay ya un listado antiguo, y si es el mismo 
    #    que el nuevo, si es así no procesar el listado.
    if Path(f"./{fichero}.pdf").is_file() and filecmp.cmp(f"./{fichero}.pdf", f"./{fichero}_tmp.pdf"):
        print("ficheros identicos")
        procesar = False

    else:
        print("ficheros distintos")
        shutil.copyfile(f"{fichero}_tmp.pdf", f"{fichero}.pdf")
        procesar = True

    Path(f"./{fichero}_tmp.pdf").unlink()
    
    if procesar or forzar_descarga:
        with open(f"{fichero}.pdf", "rb") as PDFFile:

            PDF = PyPDF2.PdfFileReader(PDFFile)
            pages = PDF.getNumPages()
            key = "/Annots"
            uri = "/URI"
            ank = "/A"

            for page in range(pages):
                print("Current Page: {}".format(page))
                pageSliced = PDF.getPage(page)
                pageObject = pageSliced.getObject()

                if key in pageObject.keys():
                    ann = pageObject[key]
                    for a in ann:
                        u = a.getObject()

                        if uri in u[ank].keys():
                            file_from = u[ank][uri]
                            file_to = Path(f"./raw_data/{u[ank][uri].split('/')[-1]}")

                            if "raw_data" in file_from and not file_to.is_file():
                                urllib.request.urlretrieve(file_from, file_to)
                                print(f"downloaded: {file_from}")

descargar_boletines()

## leer todos los boletines
forzar_lectura = False
if Path("./datos/boletines_agrupados.xlsx").is_file() or forzar_lectura:
    df = pd.read_excel("./datos/boletines_agrupados.xlsx")

else:
    ficheros = glob.glob("./raw_data/*")
    df = pd.concat([pd.read_excel(file) for file in ficheros], ignore_index=True)
    df.to_excel("./datos/boletines_agrupados.xlsx")

## generar tablas de precios semanal y mensual de gasoil de España.
df_sem = df[
    (df["Country Name"] == "Spain") & (df["Product Name"] == "Automotive gas oil")
][["Prices in force on", "Weekly price with taxes"]]
df_sem.columns = ["Fecha", "Precio"]
df_sem = df_sem.set_index("Fecha").sort_index(ascending=False)
df_sem["Precio"] = df_sem.apply(
    lambda x: float(str(x["Precio"]).replace(",", "")), axis=1
)

df_mes = (
    df_sem.groupby(pd.Grouper(freq="M")).mean().round(2).sort_index(ascending=False)
)

## Grabar resultado en excel
with pd.ExcelWriter(
    "evolucion_precio_gasoil.xlsx", engine="xlsxwriter", datetime_format="YYYY-MM-DD"
) as writer:
    df_sem.to_excel(writer, sheet_name="semanal")
    df_mes.to_excel(writer, sheet_name="mensual", startrow=3)

    # Add a header format.

    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    worksheet = writer.sheets["mensual"]
    worksheet.write(
        "A1",
        "Fuente: https://ec.europa.eu/energy/observatory/reports/Oil_Bulletin_Prices_History.xlsx",
    )
    # Add a header format.
    header_format = workbook.add_format(
        {
            "bold": True,
            "text_wrap": True,
            "valign": "top",
            "fg_color": "#D7E4BC",
            "border": 1,
        }
    )

    # Write the column headers with the defined format.
    for col_num, value in enumerate(df_mes.columns.values):
        worksheet.write(3, col_num + 1, value, header_format)
        
    writer.save()