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

  h1("Wskaźniki"),
  p("Zaangażowanie = (polubienia + komentarze + udostępnienia) / wyświetlenia × 100%, konwersja po kliknięciu = zakupy / kliknięcia w link × 100%."),
  prostaTabela(["Post", "Kanał i format", "Zaangażowanie", "Kliknięcia w link", "Zakupy", "Konwersja po kliknięciu"], [
    ["104", "Instagram, Stories", "6,4%", "600", "3", "0,5%"],
    ["108", "Instagram, Reels", "22,9%", "470", "9", "1,9%"],
    ["109", "Facebook, post promocyjny", "8,7%", "900", "20", "2,2%"],
    ["110", "Instagram, post z obrazem", "16,6%", "500", "16", "3,2%"],
  ], [700, 2676, 1550, 1450, 950, 1700], [false, false, true, true, true, true]),
  obraz("wykres_wskazniki.png", 0.36),

  h1("Wybór posta"),
  li("Najwyższe zaangażowanie: post 108 (Reels), 22,9%."),
  li("Najwyższa konwersja: post 110 (post z obrazem), 3,2%."),
  li("Dużo kliknięć, mało zakupów: post 104 (Stories), 600 kliknięć i tylko 3 zakupy (0,5%)."),
  p("Do dalszego testu wybieramy post 110 „Inspiracje na zimę”, bo ma najwyższą konwersję (3,2%) i przy tym wysokie zaangażowanie (16,6%). Post 108 jest najbardziej lubiany, ale słabiej sprzedaje (1,9%)."),

  h1("Dane do profilu SCV"),
  li("Źródło 1: sklep internetowy, czyli konta klientów i historia zamówień."),
  li("Źródło 2: newsletter i baza klientów (CRM), czyli zapisy i kliknięcia w maile."),
  li("Wspólny identyfikator: adres e-mail klienta."),
  p("Przykład: klient wszedł z posta 110, dodał kurtkę do koszyka, ale jej nie kupił. Jest zapisany do newslettera, więc następnego dnia dostaje maila z przypomnieniem o koszyku."),
];

const stopka = new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 20 })] })] });

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 22 } } } },
  numbering: { config: [{ reference: "kropki", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }] },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1418, bottom: 1418, left: 1418, right: 1418 } } }, footers: { default: stopka }, children: tresc }],
});

Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log("Zapisano", OUT); });
