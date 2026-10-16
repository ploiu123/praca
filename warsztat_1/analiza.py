"""Warsztat 1 - filtrowanie danych sklepu internetowego.

Wczytuje dane.csv, wypisuje zestawienia i buduje warsztat_1.xlsx
(arkusz Dane z filtrami + arkusz Wyniki z tabelami na formułach).

Uruchomienie:  python3 analiza.py
"""
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side

KATALOG = Path(__file__).resolve().parent
DANE = KATALOG / "dane.csv"
XLSX = KATALOG / "warsztat_1.xlsx"

NAGLOWKI_DANYCH = [
    "Użytkownik", "Źródło ruchu", "Urządzenie", "Kraj", "Nowy/Powracający",
    "Wskaźnik odrzuceń (%)", "Średni czas na stronie (min)", "Liczba przeglądanych stron",
    "Wartość koszyka (€)", "Czy dokonano zakupu?",
]
KOLUMNY = {"zrodlo_ruchu": "B", "urzadzenie": "C", "kraj": "D", "typ_uzytkownika": "E"}

CIENKA = Side(style="thin", color="000000")
RAMKA = Border(left=CIENKA, right=CIENKA, top=CIENKA, bottom=CIENKA)
EURO = '#,##0.00 "€"'


def wypisz(df):
    df = df.assign(kupil=(df.zakup == "Tak").astype(int))
    for kol in KOLUMNY:
        t = df.groupby(kol).agg(
            uzytkownicy=("kupil", "size"), zakupy=("kupil", "sum"),
            odrzucenia=("wskaznik_odrzucen_proc", "mean"), czas=("sredni_czas_min", "mean"))
        t["konwersja"] = t.zakupy / t.uzytkownicy * 100
        t["sredni_koszyk"] = df[df.kupil == 1].groupby(kol).wartosc_koszyka_eur.mean()
        print(t.round(2).to_string(), "\n")


def zakres(litera, n):
    return f"Dane!${litera}${2}:${litera}${n + 1}"


def arkusz_dane(wb, df):
    ws = wb.active
    ws.title = "Dane"
    ws.append(NAGLOWKI_DANYCH)
    for wiersz in df.itertuples(index=False):
        ws.append([v.item() if hasattr(v, "item") else v for v in wiersz])
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for wiersz in ws.iter_rows(min_row=1, max_row=len(df) + 1, max_col=len(NAGLOWKI_DANYCH)):
        for cell in wiersz:
            cell.border = RAMKA
    for wiersz in ws.iter_rows(min_row=2, min_col=9, max_col=9):
        wiersz[0].number_format = EURO
    for litera, szer in zip("ABCDEFGHIJ", (11, 15, 12, 16, 17, 12, 13, 13, 12, 11)):
        ws.column_dimensions[litera].width = szer
    ws.row_dimensions[1].height = 45
    ws.auto_filter.ref = f"A1:J{len(df) + 1}"
    ws.freeze_panes = "A2"


def tabela(ws, r, tytul, naglowki, wiersze):
    ws.cell(row=r, column=1, value=tytul).font = Font(bold=True, size=12)
    r += 1
    for c, n in enumerate(naglowki, 1):
        cell = ws.cell(row=r, column=c, value=n)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
        cell.border = RAMKA
    ws.row_dimensions[r].height = 32
    for wiersz in wiersze:
        r += 1
        for c, (v, fmt) in enumerate(wiersz(r), 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = RAMKA
            if fmt:
                cell.number_format = fmt
    return r + 3


def arkusz_wyniki(wb, n):
    ws = wb.create_sheet("Wyniki")
    zak = zakres("J", n)

    def podstawowe(kol, wartosc, z_czasem):
        k = zakres(kol, n)

        def wiersz(r):
            w = [
                (wartosc, None),
                (f"=COUNTIF({k},A{r})", None),
                (f'=COUNTIFS({k},A{r},{zak},"Tak")', None),
                (f"=C{r}/B{r}", "0.0%"),
                (f'=AVERAGEIFS({zakres("I", n)},{k},A{r},{zak},"Tak")', EURO),
                (f'=ROUND(AVERAGEIF({k},A{r},{zakres("F", n)}),1)', None),
            ]
            if z_czasem:
                w.append((f'=ROUND(AVERAGEIF({k},A{r},{zakres("G", n)}),1)', None))
            return w
        return wiersz

    glowne = ["Użytkownicy", "Zakupy", "Konwersja", "Średni koszyk", "Odrzucenia (%)"]
    r = 1
    r = tabela(ws, r, "1. Źródło ruchu", ["Źródło ruchu"] + glowne + ["Średni czas (min)"],
               [podstawowe("B", v, True) for v in ("Google Ads", "Social Media", "Organic Search")])
    r = tabela(ws, r, "2. Urządzenia", ["Urządzenie"] + glowne,
               [podstawowe("C", v, False) for v in ("Desktop", "Mobile", "Tablet")])
    r = tabela(ws, r, "3. Kraje", ["Kraj"] + glowne,
               [podstawowe("D", v, False) for v in ("Niemcy", "Polska", "Wielka Brytania")])

    def kraj_typ(kraj):
        d, e = zakres("D", n), zakres("E", n)

        def wiersz(r):
            return [
                (kraj, None),
                (f'=COUNTIFS({d},A{r},{e},"Nowy")', None),
                (f'=COUNTIFS({d},A{r},{e},"Nowy",{zak},"Tak")', None),
                (f'=COUNTIFS({d},A{r},{e},"Powracający")', None),
                (f'=COUNTIFS({d},A{r},{e},"Powracający",{zak},"Tak")', None),
                (f'=ROUND(AVERAGEIFS({zakres("F", n)},{d},A{r},{e},"Nowy"),1)', None),
                (f'=ROUND(AVERAGEIFS({zakres("F", n)},{d},A{r},{e},"Powracający"),1)', None),
            ]
        return wiersz

    r = tabela(ws, r, "3. Kraje – nowi i powracający",
               ["Kraj", "Nowi", "Zakupy nowych", "Powracający", "Zakupy powracających",
                "Odrzucenia nowych (%)", "Odrzucenia powracających (%)"],
               [kraj_typ(v) for v in ("Niemcy", "Polska", "Wielka Brytania")])

    def typ(wartosc):
        k = zakres("E", n)

        def wiersz(r):
            return [
                (wartosc, None),
                (f"=COUNTIF({k},A{r})", None),
                (f'=COUNTIFS({k},A{r},{zak},"Tak")', None),
                (f"=C{r}/B{r}", "0.0%"),
                (f'=ROUND(AVERAGEIF({k},A{r},{zakres("G", n)}),1)', None),
                (f'=ROUND(AVERAGEIF({k},A{r},{zakres("F", n)}),1)', None),
            ]
        return wiersz

    tabela(ws, r, "4. Nowi i powracający użytkownicy",
           ["Typ użytkownika", "Użytkownicy", "Zakupy", "Konwersja", "Średni czas (min)", "Odrzucenia (%)"],
           [typ(v) for v in ("Nowy", "Powracający")])

    ws.column_dimensions["A"].width = 18
    for litera in "BCDEFG":
        ws.column_dimensions[litera].width = 14


def zbuduj_xlsx(df):
    wb = Workbook()
    arkusz_dane(wb, df)
    arkusz_wyniki(wb, len(df))
    wb.save(XLSX)
    print(f"Zapisano {XLSX.name}")


if __name__ == "__main__":
    dane = pd.read_csv(DANE)
    wypisz(dane)
    zbuduj_xlsx(dane)
