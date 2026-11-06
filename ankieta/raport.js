// Raport z ankiety „Opinia klienta” (Word).
// Uruchomienie: node raport.js   (wymaga pakietu docx; wykresy z wykresy.py)
const fs = require("fs");
const path = require("path");
const sizeOf = (plik) => {
  const b = fs.readFileSync(plik);
  return { width: b.readUInt32BE(16), height: b.readUInt32BE(20) };
};
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  AlignmentType, LevelFormat, BorderStyle, ImageRun, ShadingType,
} = require("docx");

const KAT = __dirname;
const OUT = path.join(KAT, "Opinia_klienta_wyniki_ankiety.docx");
const FONT = "Calibri";
const AKCENT = "1C5CAB";
const SKALA = 0.42; // px wykresu (200 dpi) -> px w dokumencie, ta sama dla wszystkich wykresów

const t = (text, opts = {}) => new TextRun({ text, font: FONT, ...opts });
const p = (children, opts = {}) => new Paragraph({ spacing: { after: 120, line: 276 }, ...opts, children: typeof children === "string" ? [t(children)] : children });
const h1 = (text) => new Paragraph({ keepNext: true, spacing: { before: 320, after: 120 }, children: [t(text, { bold: true, size: 28, color: AKCENT })] });
const h2 = (text) => new Paragraph({ keepNext: true, spacing: { before: 200, after: 80 }, children: [t(text, { bold: true, size: 23 })] });
const li = (children) => new Paragraph({ numbering: { reference: "kropki", level: 0 }, spacing: { after: 60, line: 276 }, children: typeof children === "string" ? [t(children)] : children });
const cytat = (text) => new Paragraph({ indent: { left: 360 }, spacing: { after: 80, line: 276 }, children: [t(`„${text}”`, { italics: true, color: "404040" })] });

function wykres(plik) {
  const f = path.join(KAT, "wykresy", plik);
  const { width, height } = sizeOf(f);
  return new Paragraph({
    keepNext: false, alignment: AlignmentType.CENTER, spacing: { before: 60, after: 120 },
    children: [new ImageRun({ type: "png", data: fs.readFileSync(f), transformation: { width: Math.round(width * SKALA), height: Math.round(height * SKALA) } })],
  });
}

function tabela(naglowki, wiersze, widths) {
  const linia = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
  const ramka = { top: linia, bottom: linia, left: linia, right: linia };
  const komorka = (text, w, nag) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    borders: ramka,
    shading: nag ? { type: ShadingType.CLEAR, fill: "E8EEF7", color: "auto" } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ children: [t(text, { bold: nag, size: 20 })] })],
  });
  return new Table({
    width: { size: widths.reduce((s, w) => s + w, 0), type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: naglowki.map((n, i) => komorka(n, widths[i], true)) }),
      ...wiersze.map((w) => new TableRow({ cantSplit: true, children: w.map((v, i) => komorka(v, widths[i], false)) })),
    ],
  });
}
const odstep = () => new Paragraph({ spacing: { after: 80 }, children: [] });

const content = [
  new Paragraph({ spacing: { after: 60 }, children: [t("Opinia klienta – wyniki ankiety", { bold: true, size: 40, color: AKCENT })] }),
  p([t("Ankieta o słuchawkach bezprzewodowych", { size: 26, color: "595959" })], { spacing: { after: 240 } }),

  h1("O ankiecie"),
  p("Ankietę wypełniło 11 osób między 30 października a 6 listopada 2026 r. Ankieta miała 10 pytań: kim są respondenci, jak korzystają ze słuchawek, co jest dla nich ważne przy zakupie, ile chcą wydać, gdzie kupują i czego im brakuje w obecnych słuchawkach."),
  p("W pytaniach, w których można było zaznaczyć kilka odpowiedzi, procenty nie sumują się do 100%."),

  h1("1. Kim są respondenci"),
  wykres("1_profil.png"),
  p("Większość ankietowanych to młode osoby: 6 z 11 ma 18–24 lata, a 4 osoby mają 25–34 lata. Tylko jedna osoba ma 45 lat lub więcej. Pracujących na etacie (6 osób) i uczących się lub studiujących (5 osób) jest prawie po równo. Wśród osób w wieku 18–24 lata przeważają studenci (4 z 6), a wśród osób w wieku 25–34 lata pracujący (3 z 4)."),

  h1("2. Jak korzystają ze słuchawek"),
  h2("Ile czasu dziennie spędzają w słuchawkach"),
  wykres("2_czas.png"),
  p("Ponad połowa osób (6) używa słuchawek do godziny dziennie. Druga duża grupa (4 osoby) ma je na uszach ponad 3 godziny dziennie. Od 1 do 3 godzin korzysta z nich tylko 1 osoba."),
  h2("W jakich sytuacjach ich używają"),
  wykres("3_sytuacje.png"),
  p("Najczęściej słuchawek używa się podczas grania (5 osób). Zaraz potem są: w drodze, w pracy lub podczas nauki i w domu dla relaksu (po 4 osoby). Najrzadziej na treningu (3 osoby). Słuchawki towarzyszą więc respondentom w wielu różnych sytuacjach, a nie w jednej."),

  h1("3. Co jest najważniejsze przy wyborze słuchawek"),
  wykres("4_wazne.png"),
  p("Zdecydowanie najważniejsza jest wygoda noszenia, którą wybrało 10 z 11 osób. Dalej są jakość dźwięku z mocnym basem i mocna bateria (po 7 osób) oraz wyciszenie otoczenia ANC (6 osób). Cena i promocje liczą się tylko dla 3 osób, a wygląd i design tylko dla 1."),

  h1("4. Ile są gotowi zapłacić"),
  wykres("5_cena.png"),
  p("Najwięcej osób (4) wyda na nowe słuchawki bezprzewodowe 201–400 zł. Po 3 osoby wybrały do 200 zł i 401–700 zł, a tylko 1 osoba zapłaciłaby ponad 700 zł. Typowy budżet to średnia półka cenowa."),

  h1("5. Styl i zaufanie do marki"),
  wykres("6_skale.png"),
  p("Słuchawki raczej nie są dla respondentów elementem stylu. Średnia ocena to 2,9 na 5, a odpowiedzi rozkładają się dość równo. Pasuje to do tego, że wygląd i design wybrała tylko 1 osoba."),
  p("Opinie o zaufaniu do sprawdzonych marek, takich jak JBL, są podzielone. 4 osoby zdecydowanie się zgadzają, ale 4 raczej się nie zgadzają. Średnia to 3,5 na 5 i nikt nie wybrał odpowiedzi 1."),

  h1("6. Gdzie kupują"),
  wykres("7_gdzie.png"),
  p("Najwięcej osób (5) kupuje w internetowych sklepach pośredników. Po 2 osoby wybierają sklep stacjonarny, stronę producenta albo platformy typu Allegro czy Amazon. Łącznie aż 9 z 11 osób kupuje przez internet."),

  h1("7. Czego brakuje w obecnych słuchawkach"),
  p("Na to pytanie odpowiedziało 7 z 11 osób."),
  tabela(["Uwaga", "Liczba osób"], [
    ["Lepsze wyciszenie otoczenia", "2"],
    ["Większa wygoda (słuchawki za bardzo uciskają głowę)", "1"],
    ["Lepsza jakość dźwięku", "1"],
    ["Możliwość złożenia słuchawek", "1"],
    ["Bez komunikatu głosowego o ładowaniu w trakcie muzyki", "1"],
    ["Niczego nie brakuje (zadowolony użytkownik JBL)", "1"],
  ], [7226, 1800]),
  odstep(),
  p("Najczęstsza uwaga dotyczy wyciszenia otoczenia. Pozostałe dotyczą wygody, dźwięku, składania słuchawek i komunikatów głosowych. Wybrane wypowiedzi:"),
  cytat("Nic, bo używam JBL. Są wygodne, nie uciskają mnie i nie mam wrażenia, że moja głowa dźwiga coś ciężkiego. I fajnie wyciszają szumy z otoczenia jak potrzebuję się nad czymś mocno skupić."),
  cytat("za bardzo uciskają w głowe"),
  cytat("Usunęłabym głos kobiety w trakcie piosenek o potrzebie ładowania"),
  cytat("możliwości złożenia ich"),

  h1("8. Studenci a pracujący"),
  tabela(["", "Nauka lub studia (5 osób)", "Praca na etacie (6 osób)"], [
    ["Najczęstsze użycie", "podczas grania (3)", "w pracy lub nauce i w domu dla relaksu (po 3)"],
    ["Najważniejsze cechy", "jakość dźwięku i wygoda (po 4)", "wygoda (6) i mocna bateria (5)"],
    ["Budżet", "najczęściej 201–400 zł (2); jako jedyni także powyżej 700 zł (1)", "po 2 osoby: do 200 zł, 201–400 zł, 401–700 zł"],
    ["Słuchawki jako element stylu (średnio)", "2,8", "3,0"],
    ["Zaufanie do sprawdzonych marek (średnio)", "2,8", "4,0"],
  ], [2826, 3100, 3100]),
  odstep(),
  p("Pracujący bardziej ufają sprawdzonym markom (średnio 4,0 wobec 2,8) i częściej zwracają uwagę na baterię. Studenci częściej używają słuchawek do grania i bardziej cenią jakość dźwięku. Jedyna osoba gotowa wydać ponad 700 zł uczy się lub studiuje."),

  h1("Podsumowanie"),
  li("Typowy respondent to młoda osoba (18–34 lata), która pracuje albo studiuje i używa słuchawek do godziny dziennie lub ponad 3 godziny."),
  li("Przy wyborze najważniejsza jest wygoda (10 z 11 osób), potem jakość dźwięku, bateria i wyciszenie otoczenia. Wygląd prawie się nie liczy."),
  li("Najczęściej respondenci są gotowi wydać 201–400 zł."),
  li("Prawie wszyscy kupują przez internet (9 z 11 osób)."),
  li("Zaufanie do marek takich jak JBL jest podzielone, więc sama marka nie wystarczy. Trzeba pokazywać wygodę i jakość."),
  li("Najczęstsza uwaga to potrzeba lepszego wyciszenia otoczenia."),
];

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 22 } } } },
  numbering: {
    config: [{ reference: "kropki", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] }],
  },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1300, bottom: 1300, left: 1440, right: 1440 } } },
    children: content,
  }],
});

Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log("Zapisano", OUT); });
