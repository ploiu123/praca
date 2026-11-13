// Warsztat Single Customer View - odpowiedzi (Word).
// Uruchomienie: node raport.js   (wymaga pakietu docx i wykresu z wykres.py)
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType,
  LevelFormat, BorderStyle, ImageRun, ShadingType, Footer, PageNumber, VerticalAlign,
} = require("docx");

const KAT = __dirname;
const OUT = path.join(KAT, "Warsztat_SCV_odpowiedzi.docx");
const FONT = "Calibri";
const SZARY = "595959";

const t = (text, o = {}) => new TextRun({ text, font: FONT, ...o });
const p = (children) => new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 110, line: 276 }, children: typeof children === "string" ? [t(children)] : children });
const h1 = (text) => new Paragraph({ keepNext: true, spacing: { before: 280, after: 110 }, children: [t(text, { bold: true, size: 26 })] });
const li = (children) => new Paragraph({ numbering: { reference: "kropki", level: 0 }, spacing: { after: 60, line: 276 }, children: typeof children === "string" ? [t(children)] : children });

let nrWykresu = 0, nrTabeli = 0;
const podpis = (text) => new Paragraph({ keepNext: true, alignment: AlignmentType.CENTER, spacing: { before: 140, after: 60 }, children: [t(text, { bold: true, size: 20 })] });
const zrodlo = () => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 40, after: 180 }, children: [t("Źródło: opracowanie własne na podstawie danych z warsztatu.", { italics: true, size: 18, color: SZARY })] });

function tabela(tytul, naglowki, wiersze, widths, wyrownaj = []) {
  const linia = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
  const ramka = { top: linia, bottom: linia, left: linia, right: linia };
  const komorka = (text, w, nag, prawo) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    borders: ramka,
    shading: nag ? { type: ShadingType.CLEAR, fill: "D9E2F3", color: "auto" } : undefined,
    margins: { top: 50, bottom: 50, left: 80, right: 80 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ alignment: prawo ? AlignmentType.RIGHT : AlignmentType.LEFT, children: [t(text, { bold: nag, size: 20 })] })],
  });
  nrTabeli += 1;
  return [
    podpis(`Tabela ${nrTabeli}. ${tytul}`),
    new Table({
      width: { size: widths.reduce((s, w) => s + w, 0), type: WidthType.DXA },
      columnWidths: widths,
      rows: [
        new TableRow({ tableHeader: true, children: naglowki.map((n, i) => komorka(n, widths[i], true, false)) }),
        ...wiersze.map((w) => new TableRow({ cantSplit: true, children: w.map((v, i) => komorka(v, widths[i], false, wyrownaj[i])) })),
      ],
    }),
    zrodlo(),
  ];
}

function wykres(plik, tytul, skala) {
  const b = fs.readFileSync(path.join(KAT, plik));
  const w = b.readUInt32BE(16), h = b.readUInt32BE(20);
  nrWykresu += 1;
  return [
    podpis(`Wykres ${nrWykresu}. ${tytul}`),
    new Paragraph({ keepNext: true, alignment: AlignmentType.CENTER, children: [new ImageRun({ type: "png", data: b, transformation: { width: Math.round(w * skala), height: Math.round(h * skala) } })] }),
    zrodlo(),
  ];
}

function prostaTabela(naglowki, wiersze, widths, prawo) {
  const linia = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
  const ramka = { top: linia, bottom: linia, left: linia, right: linia };
  const komorka = (text, w, nag, pr) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    borders: ramka,
    shading: nag ? { type: ShadingType.CLEAR, fill: "D9E2F3", color: "auto" } : undefined,
    margins: { top: 50, bottom: 50, left: 80, right: 80 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ alignment: pr && !nag ? AlignmentType.RIGHT : AlignmentType.LEFT, children: [t(text, { bold: nag, size: 20 })] })],
  });
  return new Table({
    width: { size: widths.reduce((s, w) => s + w, 0), type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: naglowki.map((n, i) => komorka(n, widths[i], true, prawo[i])) }),
      ...wiersze.map((w) => new TableRow({ cantSplit: true, children: w.map((v, i) => komorka(v, widths[i], false, prawo[i])) })),
    ],
  });
}

function obraz(plik, skala) {
  const b = fs.readFileSync(path.join(KAT, plik));
  const w = b.readUInt32BE(16), h = b.readUInt32BE(20);
  return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 160, after: 80 }, children: [new ImageRun({ type: "png", data: b, transformation: { width: Math.round(w * skala), height: Math.round(h * skala) } })] });
}

const tresc = [
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60, after: 240 }, children: [t("Warsztat Single Customer View", { bold: true, size: 32 })] }),

  h1("2. Wskaźniki"),
  p("Zaangażowanie policzyliśmy jako sumę polubień, komentarzy i udostępnień podzieloną przez liczbę wyświetleń, a konwersję po kliknięciu jako liczbę zakupów podzieloną przez liczbę kliknięć w link. Wyniki są w procentach, zaokrąglone do jednego miejsca po przecinku."),
  prostaTabela(["Post", "Kanał i format", "Zaangażowanie", "Kliknięcia w link", "Zakupy", "Konwersja po kliknięciu"], [
    ["104", "Instagram, Stories", "6,4%", "600", "3", "0,5%"],
    ["108", "Instagram, Reels", "22,9%", "470", "9", "1,9%"],
    ["109", "Facebook, post promocyjny", "8,7%", "900", "20", "2,2%"],
    ["110", "Instagram, post z obrazem", "16,6%", "500", "16", "3,2%"],
  ], [700, 2676, 1550, 1450, 950, 1700], [false, false, true, true, true, true]),
  obraz("wykres_wskazniki.png", 0.36),
  p("Już na pierwszy rzut oka widać, że duże zaangażowanie nie oznacza dużej sprzedaży. Post, na który ludzie najchętniej reagują, wcale nie sprzedaje najlepiej."),

  h1("3. Wybór posta do dalszego testu"),
  p([t("Najwyższe zaangażowanie: ", { bold: true }), t("post 108, czyli Reels „Jak nosić elegancko” (22,9%). Krótki film z poradami dobrze się ogląda, dlatego zebrał dużo polubień i komentarzy. Z 470 kliknięć wyszło jednak tylko 9 zakupów (konwersja 1,9%). Taki materiał działa bardziej jako inspiracja i budowanie marki niż bezpośrednia zachęta do zakupu.")]),
  p([t("Najwyższa konwersja: ", { bold: true }), t("post 110, czyli post z obrazem „Inspiracje na zimę” (3,2%). Z 500 kliknięć 16 zakończyło się zakupem. Post pokazuje konkretne ubrania w stylizacjach, więc osoby, które klikają, wiedzą już, czego szukają w sklepie.")]),
  p([t("Dużo kliknięć, mało zakupów: ", { bold: true }), t("post 104, czyli Stories „Nowe trendy zimowe”. Dał 600 kliknięć, ale tylko 3 zakupy (0,5%). W Stories łatwo kliknąć link przy szybkim przeglądaniu, ale takie osoby często od razu wychodzą ze sklepu. Możliwe też, że strona, na którą prowadził link, nie pokazywała tego, co było w relacji.")]),
  p("Post 109 z wyprzedażą dał najwięcej zakupów (20), ale miał też największy zasięg (30 000 wyświetleń), a jego sprzedaż w dużej mierze napędzały rabaty, które obniżają marżę."),
  p([t("Do dalszego testu sprzedażowego wybieramy post 110. ", { bold: true }), t("Przemawiają za nim dwie liczby: najwyższa konwersja po kliknięciu (3,2%) i drugie najwyższe zaangażowanie (16,6%). Jako jedyny wypada dobrze w obu wskaźnikach, czyli i przyciąga uwagę, i sprzedaje. W teście sprawdzilibyśmy, czy przy większym budżecie utrzyma taką konwersję.")]),

  h1("4. Dane do profilu SCV"),
  p("Wyniki postów są zbiorcze: wiemy, ile osób kliknęło i kupiło, ale nie wiemy, kim one są i co robiły wcześniej. Wspólny profil klienta (SCV) łączy informacje o tej samej osobie z różnych miejsc. Wykorzystalibyśmy dwa źródła danych własnych marki:"),
  li([t("Sklep internetowy: ", { bold: true }), t("konto klienta, historia zamówień, kupione produkty i rozmiary, porzucone koszyki. Jeśli linki w postach mają oznaczenia UTM, sklep zapisuje też, z którego posta klient przyszedł.")]),
  li([t("Newsletter i program lojalnościowy (CRM): ", { bold: true }), t("zapis do newslettera, otwarcia i kliknięcia w maile, zgody marketingowe, punkty i nagrody.")]),
  p([t("Wspólny identyfikator: ", { bold: true }), t("adres e-mail. Klient podaje go przy rejestracji i zakupie w sklepie oraz przy zapisie do newslettera, więc po nim można połączyć oba źródła w jeden profil.")]),
  p([t("Przykład wykorzystania: ", { bold: true }), t("klient wszedł do sklepu z posta „Inspiracje na zimę”, dodał płaszcz do koszyka, ale go nie kupił. Z profilu wiemy, że jest zapisany do newslettera i zwykle kupuje rozmiar M. Następnego dnia wysyłamy mu maila z przypomnieniem o koszyku i dwiema podobnymi stylizacjami w jego rozmiarze. Wracamy do niego bez płacenia drugi raz za reklamę.")]),
];

const stopka = new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 20 })] })] });

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 22 } } } },
  numbering: { config: [{ reference: "kropki", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }] },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1418, bottom: 1418, left: 1418, right: 1418 } } }, footers: { default: stopka }, children: tresc }],
});

Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log("Zapisano", OUT); });
