"""Warsztat Single Customer View - skoroszyt z danymi, wskaźnikami i wykresami.

Uruchomienie:  python3 excel.py   (czyta dane.csv, zapisuje Warsztat_SCV.xlsx)
"""
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

KAT = Path(__file__).resolve().parent
df = pd.read_csv(KAT / "dane.csv")
N = len(df)
OST = N + 1

NAGLOWEK = PatternFill("solid", fgColor="D9E2F3")
CIENKA = Side(style="thin", color="000000")
RAMKA = Border(left=CIENKA, right=CIENKA, top=CIENKA, bottom=CIENKA)

wb = Workbook()
ws = wb.active
ws.title = "Wskaźniki"

naglowki = ["ID", "Kanał", "Format", "Temat", "Wyświetlenia", "Polubienia", "Komentarze", "Udostępnienia",
            "Kliknięcia w link", "Zakupy", "Interakcje razem", "Zaangażowanie", "Konwersja po kliknięciu", "Post"]
ws.append(naglowki)
for i, r in enumerate(df.itertuples(index=False), 2):
    ws.append([r.id, r.kanal, r.format, r.temat, r.wyswietlenia, r.polubienia, r.komentarze, r.udostepnienia,
               r.klikniecia, r.zakupy,
               f"=F{i}+G{i}+H{i}",
               f"=ROUND(K{i}/E{i}*100,1)/100",
               f"=ROUND(J{i}/I{i}*100,1)/100",
               f'=A{i}&" "&B{i}&", "&C{i}'])
    for c in "EFGHIJK":
        ws[f"{c}{i}"].number_format = "# ##0"
    ws[f"L{i}"].number_format = "0.0%"
    ws[f"M{i}"].number_format = "0.0%"

for r in ws.iter_rows(min_row=1, max_row=OST, max_col=len(naglowki)):
    for c in r:
        c.border = RAMKA
        c.alignment = Alignment(wrap_text=True, vertical="center")
for c in ws[1]:
    c.font = Font(bold=True)
    c.fill = NAGLOWEK
for kol, szer in zip("ABCDEFGHIJKLMN", (6, 11, 15, 20, 13, 11, 12, 14, 11, 8, 12, 15, 13, 28)):
    ws.column_dimensions[kol].width = szer
ws.row_dimensions[1].height = 32

r = OST + 2
ws.cell(row=r, column=1, value="Wzory").font = Font(bold=True)
ws.cell(row=r + 1, column=1, value="Zaangażowanie = (polubienia + komentarze + udostępnienia) ÷ wyświetlenia × 100%")
ws.cell(row=r + 2, column=1, value="Konwersja po kliknięciu = zakupy ÷ kliknięcia w link × 100%")

W = "Wskaźniki!"
wn = wb.create_sheet("Wnioski")
wn["A1"] = "Wnioski z analizy"
wn["A1"].font = Font(bold=True, size=12)
wnioski = [
    ("Najwyższe zaangażowanie", f"=INDEX({W}$N$2:$N${OST},MATCH(MAX({W}$L$2:$L${OST}),{W}$L$2:$L${OST},0))", f"=MAX({W}$L$2:$L${OST})"),
    ("Najwyższa konwersja po kliknięciu", f"=INDEX({W}$N$2:$N${OST},MATCH(MAX({W}$M$2:$M${OST}),{W}$M$2:$M${OST},0))", f"=MAX({W}$M$2:$M${OST})"),
    ("Dużo kliknięć, mało zakupów (najniższa konwersja)", f"=INDEX({W}$N$2:$N${OST},MATCH(MIN({W}$M$2:$M${OST}),{W}$M$2:$M${OST},0))", f"=MIN({W}$M$2:$M${OST})"),
]
wn.append([])
wn.append(["", "Post", "Wartość"])
for opis, post, wart in wnioski:
    wn.append([opis, post, wart])
    wn.cell(row=wn.max_row, column=3).number_format = "0.0%"
for rr in wn.iter_rows(min_row=3, max_row=wn.max_row, max_col=3):
    for c in rr:
        c.border = RAMKA
for c in wn[3][:3]:
    c.font = Font(bold=True)
    c.fill = NAGLOWEK
wn.column_dimensions["A"].width = 48
wn.column_dimensions["B"].width = 30
wn.column_dimensions["C"].width = 12
r = OST + 3


def wykres(kol, tytul, kolor, kotwica):
    ch = BarChart()
    ch.type = "bar"
    ch.style = 10
    ch.title = tytul
    ch.x_axis.scaling.orientation = "maxMin"
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.y_axis.number_format = "0%"
    ch.y_axis.majorGridlines = None
    ch.add_data(Reference(ws, min_col=kol, min_row=1, max_row=OST), titles_from_data=True)
    ch.series[0].graphicalProperties.solidFill = kolor
    ch.series[0].graphicalProperties.line.solidFill = kolor
    ch.set_categories(Reference(ws, min_col=14, min_row=2, max_row=OST))
    ch.dataLabels = DataLabelList(showVal=True, showCatName=False, showSerName=False, showLegendKey=False, showPercent=False)
    ch.dataLabels.numFmt = "0.0%"
    ch.legend = None
    ch.gapWidth = 60
    ch.width, ch.height = 13, 7
    ws.add_chart(ch, kotwica)


wykres(12, "Zaangażowanie", "4472C4", f"A{r + 3}")
wykres(13, "Konwersja po kliknięciu", "ED7D31", f"H{r + 3}")

wb.save(KAT / "Warsztat_SCV.xlsx")
print("Zapisano Warsztat_SCV.xlsx")
