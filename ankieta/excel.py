"""Ankieta „Opinia klienta” - skoroszyt Excel z tabelami i wykresami.

Uruchomienie:  python3 excel.py
Wczytuje odpowiedzi.csv i zapisuje Opinia_klienta_tabele.xlsx.
Wszystkie liczby w tabelach są liczone formułami z arkusza „Odpowiedzi”.
"""
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

KAT = Path(__file__).resolve().parent
XLSX = KAT / "Opinia_klienta_tabele.xlsx"

NAGLOWEK = PatternFill("solid", fgColor="D9E2F3")
CIENKA = Side(style="thin", color="000000")
RAMKA = Border(left=CIENKA, right=CIENKA, top=CIENKA, bottom=CIENKA)
NIEBIESKI, POMARANCZ = "4472C4", "ED7D31"

df = pd.read_csv(KAT / "odpowiedzi.csv")
df.columns = ["ts", "wiek", "zajecie", "czas", "sytuacje", "wazne", "cena", "styl", "marka", "gdzie", "brak"]
N = len(df)
OST = N + 1  # ostatni wiersz danych w arkuszu Odpowiedzi

# kolumny w arkuszu Odpowiedzi
KOL = {"nr": "A", "data": "B", "wiek": "C", "zajecie": "D", "grupa": "E", "czas": "F", "sytuacje": "G",
       "wazne": "H", "cena": "I", "styl": "J", "marka": "K", "gdzie": "L", "brak": "M"}


def zakres(kol):
    return f"Odpowiedzi!${KOL[kol]}$2:${KOL[kol]}${OST}"


def kategorie_uwagi(tekst):
    """Przypisuje odpowiedź otwartą do kategorii (najwyżej dwie)."""
    if not isinstance(tekst, str) or not tekst.strip():
        return []
    t = tekst.lower()
    if t.startswith("nic") or "niczego" in t:
        return ["Niczego nie brakuje"]
    kat = []
    if "wycisz" in t:
        kat.append("Lepsze wyciszenie otoczenia")
    if "bateri" in t:
        kat.append("Lepsza bateria")
    if "uciska" in t:
        kat.append("Słuchawki za bardzo uciskają głowę")
    if "dźwięk" in t or "dzwięk" in t:
        kat.append("Lepsza jakość dźwięku")
    if "złoż" in t:
        kat.append("Możliwość złożenia słuchawek")
    if "ładowa" in t:
        kat.append("Brak komunikatu głosowego o ładowaniu")
    return kat or ["Inne"]


def tylko_wartosci():
    return DataLabelList(showVal=True, showCatName=False, showSerName=False, showLegendKey=False,
                         showPercent=False, showBubbleSize=False, showLeaderLines=False)


def styl_tabeli(ws, r0, r1, c0, c1):
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            ws.cell(row=r, column=c).border = RAMKA
    for c in range(c0, c1 + 1):
        cell = ws.cell(row=r0, column=c)
        cell.font = Font(bold=True)
        cell.fill = NAGLOWEK
        cell.alignment = Alignment(wrap_text=True, vertical="center")


def wykres(ws, tytul, r_nagl, r_kon, kolumny, kotwica, wysokosc=None, os_y="Liczba osób", proc=False):
    """Wykres słupkowy poziomy z tabeli: kategorie w kolumnie A, serie w podanych kolumnach."""
    ch = BarChart()
    ch.type = "bar"
    ch.style = 10
    ch.title = tytul
    ch.y_axis.title = os_y
    ch.x_axis.scaling.orientation = "maxMin"  # kolejność od góry jak w tabeli
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.y_axis.majorGridlines = None
    if proc:
        ch.y_axis.number_format = "0%"
        ch.y_axis.scaling.min = 0
        ch.y_axis.scaling.max = 1
    kolory = [NIEBIESKI, POMARANCZ]
    for i, k in enumerate(kolumny):
        dane = Reference(ws, min_col=k, min_row=r_nagl, max_row=r_kon)
        ch.add_data(dane, titles_from_data=True)
        ch.series[-1].graphicalProperties.solidFill = kolory[i % 2]
        ch.series[-1].graphicalProperties.line.solidFill = kolory[i % 2]
    ch.set_categories(Reference(ws, min_col=1, min_row=r_nagl + 1, max_row=r_kon))
    ch.dataLabels = tylko_wartosci()
    if len(kolumny) == 1:
        ch.legend = None
    else:
        ch.legend.position = "b"
    ch.gapWidth = 60
    ch.width = 16
    ch.height = wysokosc or max(5.5, 1.1 * (r_kon - r_nagl) + 3)
    ws.add_chart(ch, kotwica)


wb = Workbook()

# --- Odpowiedzi -------------------------------------------------------------------
ws = wb.active
ws.title = "Odpowiedzi"
naglowki = ["Nr", "Data", "Wiek", "Zajęcie", "Grupa", "Czas dziennie w słuchawkach", "Sytuacje",
            "Najważniejsze cechy", "Maksymalny budżet", "Styl (1-5)", "Zaufanie do marki (1-5)",
            "Gdzie kupuje", "Czego brakuje"]
ws.append(naglowki)
for i, r in enumerate(df.itertuples(index=False), 2):
    data = datetime.strptime(r.ts.replace(" CET", ""), "%Y/%m/%d %I:%M:%S %p")
    ws.append([i - 1, data, r.wiek, r.zajecie, f'=IF(D{i}="Nauka lub studia","Studenci","Pracujący")', r.czas,
               r.sytuacje.replace(";", "; "), r.wazne.replace(";", "; "), r.cena, int(r.styl), int(r.marka),
               r.gdzie, r.brak if isinstance(r.brak, str) else ""])
    ws.cell(row=i, column=2).number_format = "dd.mm.yyyy hh:mm"
styl_tabeli(ws, 1, OST, 1, len(naglowki))
for wiersz in ws.iter_rows(min_row=2, max_row=OST):
    for c in wiersz:
        c.alignment = Alignment(wrap_text=True, vertical="top")
for kol, szer in zip("ABCDEFGHIJKLM", (5, 16, 11, 22, 11, 16, 30, 38, 14, 9, 11, 26, 45)):
    ws.column_dimensions[kol].width = szer
ws.row_dimensions[1].height = 32
ws.auto_filter.ref = f"A1:M{OST}"
ws.freeze_panes = "A2"

# --- Tabele -----------------------------------------------------------------------
wt = wb.create_sheet("Tabele")
wt["A1"] = "Wyniki ankiety „Opinia klienta” – tabele"
wt["A1"].font = Font(bold=True, size=14)
wt["A2"] = "Liczba respondentów"
wt["B2"] = f"=COUNTA(Odpowiedzi!$A$2:$A${OST})"
wt["A2"].font = Font(bold=True)
wt["B2"].font = Font(bold=True)
NR = "Tabele!$B$2"
wt.column_dimensions["A"].width = 34
wt.column_dimensions["B"].width = 13
wt.column_dimensions["C"].width = 16

PYTANIA = [
    ("1. Wiek respondentów", "wiek", ["Poniżej 18", "18–24", "25–34", "35–44", "45 i więcej"], False),
    ("2. Czym się zajmujesz?", "zajecie", ["Praca na etacie", "Nauka lub studia", "Freelance lub własna firma"], False),
    ("3. Ile czasu dziennie spędzasz w słuchawkach?", "czas", ["Do godziny", "1-3 godziny", "Powyżej 3 godzin"], False),
    ("4. W jakich sytuacjach najczęściej używasz słuchawek? (max 2)", "sytuacje",
     ["Podczas grania", "W drodze", "W pracy lub podczas nauki", "Na treningu", "W domu dla relaksu"], True),
    ("5. Co jest najważniejsze przy wyborze słuchawek? (3 odpowiedzi)", "wazne",
     ["Wygoda noszenia", "Jakość dźwięku i mocny bas", "Wyciszenie otoczenia (ANC)", "Mocna bateria",
      "Dobra cena i promocje", "Wygląd i design"], True),
    ("6. Ile maksymalnie zapłacisz za nowe słuchawki?", "cena", ["Do 200 zł", "201-400 zł", "401-700 zł", "Powyżej 700 zł"], False),
    ("9. Gdzie najchętniej kupujesz taki sprzęt?", "gdzie",
     ["W sklepach pośrednich online", "W sklepie stacjonarnym", "Na stronie producenta", "Na platformach typu Allegro czy Amazon"], False),
]
# usuń kategorie, których nikt nie wybrał w pytaniach jednokrotnego wyboru (np. 35-44)
PYTANIA = [(t_, k, [o for o in op if (wiele or (df[k] == o).any())], wiele) for t_, k, op, wiele in PYTANIA]

r = 4
for tytul, kol, opcje, wiele in PYTANIA:
    wt.cell(row=r, column=1, value=tytul).font = Font(bold=True, size=12)
    r += 1
    r_nagl = r
    wt.cell(row=r, column=1, value="Odpowiedź")
    wt.cell(row=r, column=2, value="Liczba osób")
    wt.cell(row=r, column=3, value="% respondentów")
    for o in opcje:
        r += 1
        wt.cell(row=r, column=1, value=o)
        kryt = f'"*"&A{r}&"*"' if wiele else f"A{r}"
        wt.cell(row=r, column=2, value=f"=COUNTIF({zakres(kol)},{kryt})")
        wt.cell(row=r, column=3, value=f"=B{r}/{NR}").number_format = "0%"
    r_kon = r
    if not wiele:
        r += 1
        wt.cell(row=r, column=1, value="Razem").font = Font(bold=True)
        wt.cell(row=r, column=2, value=f"=SUM(B{r_nagl + 1}:B{r_kon})").font = Font(bold=True)
        wt.cell(row=r, column=3, value=f"=SUM(C{r_nagl + 1}:C{r_kon})").number_format = "0%"
    styl_tabeli(wt, r_nagl, r, 1, 3)
    if wiele:
        wt.cell(row=r + 1, column=1, value="Można było zaznaczyć kilka odpowiedzi, więc procenty nie sumują się do 100%.").font = Font(italic=True, size=9, color="595959")
    wykres(wt, tytul.split(". ", 1)[1], r_nagl, r_kon, [2], f"E{r_nagl - 1}")
    r = max(r + 3, r_nagl - 1 + int(max(5.5, 1.1 * (r_kon - r_nagl) + 3) * 2) + 2)

# skale 1-5
wt.cell(row=r, column=1, value="7-8. Styl i zaufanie do marki (skala 1-5)").font = Font(bold=True, size=12)
r += 1
r_nagl = r
for c, v in enumerate(["Ocena", "Słuchawki to element mojego stylu", "Ufam sprawdzonym markom, np. JBL"], 1):
    wt.cell(row=r, column=c, value=v)
for ocena, opis in zip(range(1, 6), ["1 - zdecydowanie nie", "2", "3", "4", "5 - zdecydowanie tak"]):
    r += 1
    wt.cell(row=r, column=1, value=opis)
    wt.cell(row=r, column=2, value=f"=COUNTIF({zakres('styl')},{ocena})")
    wt.cell(row=r, column=3, value=f"=COUNTIF({zakres('marka')},{ocena})")
r_kon = r
r += 1
wt.cell(row=r, column=1, value="Średnia ocena").font = Font(bold=True)
wt.cell(row=r, column=2, value=f"=AVERAGE({zakres('styl')})").number_format = "0.0"
wt.cell(row=r, column=3, value=f"=AVERAGE({zakres('marka')})").number_format = "0.0"
styl_tabeli(wt, r_nagl, r, 1, 3)
wt.row_dimensions[r_nagl].height = 45
skale = BarChart()
skale.type = "col"
skale.style = 10
skale.title = "Styl i zaufanie do marki"
skale.y_axis.title = "Liczba osób"
skale.x_axis.title = "Ocena"
skale.x_axis.delete = False
skale.y_axis.delete = False
for i, k in enumerate((2, 3)):
    skale.add_data(Reference(wt, min_col=k, min_row=r_nagl, max_row=r_kon), titles_from_data=True)
    skale.series[-1].graphicalProperties.solidFill = (NIEBIESKI, POMARANCZ)[i]
    skale.series[-1].graphicalProperties.line.solidFill = (NIEBIESKI, POMARANCZ)[i]
skale.set_categories(Reference(wt, min_col=1, min_row=r_nagl + 1, max_row=r_kon))
skale.dataLabels = tylko_wartosci()
skale.legend.position = "b"
skale.width, skale.height = 16, 8
wt.add_chart(skale, f"E{r_nagl - 1}")
r += 20

# --- Porównanie grup ---------------------------------------------------------------
wg = wb.create_sheet("Porównanie grup")
wg["A1"] = "Porównanie: studenci a pracujący"
wg["A1"].font = Font(bold=True, size=14)
wg["A2"] = "Studenci = nauka lub studia. Pracujący = praca na etacie oraz freelance lub własna firma. Wartości to % osób w danej grupie."
wg["A2"].font = Font(italic=True, size=9, color="595959")
wg["A3"], wg["B3"], wg["C3"] = "Liczba osób", f'=COUNTIF({zakres("grupa")},"Studenci")', f'=COUNTIF({zakres("grupa")},"Pracujący")'
wg["B4"], wg["C4"] = "Studenci", "Pracujący"
for c in ("A3", "B3", "C3", "B4", "C4"):
    wg[c].font = Font(bold=True)
wg.column_dimensions["A"].width = 34
wg.column_dimensions["B"].width = 13
wg.column_dimensions["C"].width = 13

r = 6
GRUPOWE = [
    ("Najważniejsze cechy słuchawek", "wazne", PYTANIA[4][2], True),
    ("Sytuacje, w których używają słuchawek", "sytuacje", PYTANIA[3][2], True),
    ("Maksymalny budżet", "cena", PYTANIA[5][2], False),
    ("Gdzie kupują", "gdzie", PYTANIA[6][2], False),
]
for tytul, kol, opcje, wiele in GRUPOWE:
    wg.cell(row=r, column=1, value=tytul).font = Font(bold=True, size=12)
    r += 1
    r_nagl = r
    for c, v in enumerate(["Odpowiedź", "Studenci", "Pracujący"], 1):
        wg.cell(row=r, column=c, value=v)
    for o in opcje:
        r += 1
        wg.cell(row=r, column=1, value=o)
        kryt = f'"*"&$A{r}&"*"' if wiele else f"$A{r}"
        for c, grupa, kol_n in ((2, "Studenci", "B"), (3, "Pracujący", "C")):
            wg.cell(row=r, column=c, value=f'=COUNTIFS({zakres("grupa")},"{grupa}",{zakres(kol)},{kryt})/{kol_n}$3').number_format = "0%"
    r_kon = r
    styl_tabeli(wg, r_nagl, r_kon, 1, 3)
    wykres(wg, tytul, r_nagl, r_kon, [2, 3], f"E{r_nagl - 1}", proc=True, os_y="% osób w grupie")
    r = max(r + 3, r_nagl - 1 + int((1.1 * (r_kon - r_nagl) + 3) * 2) + 4)

wg.cell(row=r, column=1, value="Średnie oceny (skala 1-5)").font = Font(bold=True, size=12)
r += 1
r_nagl = r
for c, v in enumerate(["Pytanie", "Studenci", "Pracujący"], 1):
    wg.cell(row=r, column=c, value=v)
for kol, opis in (("styl", "Słuchawki to element mojego stylu"), ("marka", "Ufam sprawdzonym markom, np. JBL")):
    r += 1
    wg.cell(row=r, column=1, value=opis)
    for c, grupa in ((2, "Studenci"), (3, "Pracujący")):
        wg.cell(row=r, column=c, value=f'=AVERAGEIFS({zakres(kol)},{zakres("grupa")},"{grupa}")').number_format = "0.0"
styl_tabeli(wg, r_nagl, r, 1, 3)

# --- Uwagi ------------------------------------------------------------------------
wu = wb.create_sheet("Uwagi")
wu["A1"] = "10. Czego najbardziej brakuje Ci w słuchawkach, których używasz dzisiaj?"
wu["A1"].font = Font(bold=True, size=12)
wu.append([])
wu.append(["Nr", "Odpowiedź", "Kategoria 1", "Kategoria 2"])
r0 = 3
for i, tekst in enumerate(df.brak, 1):
    if isinstance(tekst, str) and tekst.strip():
        k = kategorie_uwagi(tekst) + ["", ""]
        wu.append([i, tekst.strip(), k[0], k[1]])
r1 = wu.max_row
styl_tabeli(wu, r0, r1, 1, 4)
for wiersz in wu.iter_rows(min_row=r0 + 1, max_row=r1):
    for c in wiersz:
        c.alignment = Alignment(wrap_text=True, vertical="top")
wu.column_dimensions["A"].width = 5
wu.column_dimensions["B"].width = 60
wu.column_dimensions["C"].width = 34
wu.column_dimensions["D"].width = 24

r = r1 + 3
wu.cell(row=r, column=2, value="Podsumowanie uwag").font = Font(bold=True, size=12)
r += 1
r_nagl = r
wu.cell(row=r, column=2, value="Kategoria")
wu.cell(row=r, column=3, value="Liczba osób")
kategorie = ["Niczego nie brakuje", "Lepsze wyciszenie otoczenia", "Lepsza bateria", "Słuchawki za bardzo uciskają głowę",
             "Lepsza jakość dźwięku", "Możliwość złożenia słuchawek", "Brak komunikatu głosowego o ładowaniu"]
for k in kategorie:
    r += 1
    wu.cell(row=r, column=2, value=k)
    wu.cell(row=r, column=3, value=f"=COUNTIF($C${r0 + 1}:$D${r1},B{r})")
r += 1
wu.cell(row=r, column=2, value="Brak odpowiedzi")
wu.cell(row=r, column=3, value=f"={NR}-COUNTA($B${r0 + 1}:$B${r1})")
styl_tabeli(wu, r_nagl, r, 2, 3)
wu.cell(row=r + 1, column=2, value="Jedna osoba mogła podać dwie uwagi, dlatego suma może być większa niż liczba respondentów.").font = Font(italic=True, size=9, color="595959")

wb.save(XLSX)
print("Zapisano", XLSX)
