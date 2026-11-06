// Raport z ankiety „Opinia klienta” (Word).
// Uruchomienie: node raport.js   (wymaga pakietu docx; wykresy z wykresy.py)
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType,
  LevelFormat, BorderStyle, ImageRun, ShadingType, Header, Footer, PageNumber, PageBreak,
  PageOrientation, VerticalAlign, TabStopType,
} = require("docx");

const KAT = __dirname;
const OUT = path.join(KAT, "Opinia_klienta_wyniki_ankiety.docx");
const FONT = "Calibri";
const AKCENT = "1C5CAB";
const JASNY = "EAF1FA";
const SZARY = "595959";
const SKALA = 0.42; // px wykresu (200 dpi) -> px w dokumencie, ta sama dla wszystkich wykresów
const SZER = 9026; // szerokość tekstu A4 pionowo (DXA)

// --- dane ----------------------------------------------------------------------
function czytajCsv(plik) {
  const txt = fs.readFileSync(plik, "utf8").replace(/^﻿/, "");
  const wiersze = [];
  let pole = "", wiersz = [], cudz = false;
  for (let i = 0; i < txt.length; i++) {
    const z = txt[i];
    if (cudz) {
      if (z === '"' && txt[i + 1] === '"') { pole += '"'; i++; }
      else if (z === '"') cudz = false;
      else pole += z;
    } else if (z === '"') cudz = true;
    else if (z === ",") { wiersz.push(pole); pole = ""; }
    else if (z === "\n" || z === "\r") {
      if (z === "\r" && txt[i + 1] === "\n") i++;
      wiersz.push(pole); pole = "";
      if (wiersz.some((v) => v !== "")) wiersze.push(wiersz);
      wiersz = [];
    } else pole += z;
  }
  if (pole !== "" || wiersz.length) { wiersz.push(pole); wiersze.push(wiersz); }
  return wiersze.slice(1);
}
const odp = czytajCsv(path.join(KAT, "odpowiedzi.csv"));

// --- klocki --------------------------------------------------------------------
const t = (text, o = {}) => new TextRun({ text, font: FONT, ...o });
const p = (children, o = {}) => new Paragraph({ spacing: { after: 120, line: 276 }, ...o, children: typeof children === "string" ? [t(children)] : children });
const h1 = (nr, text) => new Paragraph({
  keepNext: true, spacing: { before: 360, after: 140 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "C9D8EE", space: 4 } },
  children: [t(nr ? `${nr}  ` : "", { bold: true, size: 30, color: AKCENT }), t(text, { bold: true, size: 30, color: "1A1A1A" })],
});
const h2 = (text) => new Paragraph({ keepNext: true, spacing: { before: 220, after: 80 }, children: [t(text, { bold: true, size: 24, color: AKCENT })] });
const li = (children, ref = "kropki") => new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 80, line: 276 }, children: typeof children === "string" ? [t(children)] : children });
const cytat = (text) => new Paragraph({ indent: { left: 360 }, spacing: { after: 80, line: 276 }, children: [t(`„${text}”`, { italics: true, color: "404040" })] });
const nowaStrona = () => new Paragraph({ children: [new PageBreak()] });
const odstep = (po = 80) => new Paragraph({ spacing: { after: po }, children: [] });

function wniosek(text) {
  return new Paragraph({
    spacing: { before: 60, after: 200, line: 276 },
    indent: { left: 120, right: 120 },
    shading: { type: ShadingType.CLEAR, fill: JASNY, color: "auto" },
    border: { left: { style: BorderStyle.SINGLE, size: 24, color: AKCENT, space: 8 } },
    children: [t("Wniosek: ", { bold: true, color: AKCENT }), t(text)],
  });
}

function wykres(plik) {
  const f = path.join(KAT, "wykresy", plik);
  const b = fs.readFileSync(f);
  const w = b.readUInt32BE(16), h = b.readUInt32BE(20);
  return new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 80, after: 140 },
    children: [new ImageRun({ type: "png", data: b, transformation: { width: Math.round(w * SKALA), height: Math.round(h * SKALA) } })],
  });
}

const BRAK = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const BEZ_RAMEK = { top: BRAK, bottom: BRAK, left: BRAK, right: BRAK };

function tabela(naglowki, wiersze, widths, rozmiar = 20) {
  const linia = { style: BorderStyle.SINGLE, size: 4, color: "C9D0DA" };
  const ramka = { top: linia, bottom: linia, left: linia, right: linia };
  const komorka = (text, w, nag, pas) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    borders: ramka,
    shading: nag ? { type: ShadingType.CLEAR, fill: AKCENT, color: "auto" }
      : pas ? { type: ShadingType.CLEAR, fill: "F5F8FC", color: "auto" } : undefined,
    margins: { top: 60, bottom: 60, left: 90, right: 90 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ children: [t(text, { bold: nag, size: rozmiar, color: nag ? "FFFFFF" : undefined })] })],
  });
  return new Table({
    width: { size: widths.reduce((s, w) => s + w, 0), type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: naglowki.map((n, i) => komorka(n, widths[i], true)) }),
      ...wiersze.map((w, r) => new TableRow({ cantSplit: true, children: w.map((v, i) => komorka(v, widths[i], false, r % 2 === 1)) })),
    ],
  });
}

function kafelki(dane) {
  const w = Math.floor(SZER / dane.length);
  const widths = dane.map((_, i) => (i === dane.length - 1 ? SZER - w * (dane.length - 1) : w));
  const gruba = { style: BorderStyle.SINGLE, size: 24, color: "FFFFFF" };
  return new Table({
    width: { size: SZER, type: WidthType.DXA },
    columnWidths: widths,
    rows: [new TableRow({
      children: dane.map(([liczba, opis], i) => new TableCell({
        width: { size: widths[i], type: WidthType.DXA },
        borders: { top: gruba, bottom: gruba, left: gruba, right: gruba },
        shading: { type: ShadingType.CLEAR, fill: JASNY, color: "auto" },
        margins: { top: 200, bottom: 200, left: 160, right: 160 },
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [t(liczba, { bold: true, size: 40, color: AKCENT })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [t(opis, { size: 18, color: SZARY })] }),
        ],
      })),
    })],
  });
}

function baner(tytul, podtytul) {
  return new Table({
    width: { size: SZER, type: WidthType.DXA },
    columnWidths: [SZER],
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: SZER, type: WidthType.DXA },
        borders: BEZ_RAMEK,
        shading: { type: ShadingType.CLEAR, fill: AKCENT, color: "auto" },
        margins: { top: 600, bottom: 600, left: 400, right: 400 },
        children: [
          new Paragraph({ spacing: { after: 120 }, children: [t("RAPORT Z ANKIETY", { size: 20, bold: true, color: "BFD3F0" })] }),
          new Paragraph({ spacing: { after: 120 }, children: [t(tytul, { size: 56, bold: true, color: "FFFFFF" })] }),
          new Paragraph({ children: [t(podtytul, { size: 26, color: "E6EEF9" })] }),
        ],
      })],
    })],
  });
}

// --- treść ----------------------------------------------------------------------
const okladka = [
  odstep(600),
  baner("Opinia klienta", "Jak korzystamy ze słuchawek bezprzewodowych i czego od nich oczekujemy"),
  odstep(360),
  kafelki([
    ["11", "osób wypełniło ankietę"],
    ["91%", "wskazuje wygodę jako kluczową cechę"],
    ["82%", "kupuje sprzęt przez internet"],
    ["201–400 zł", "to najczęstszy budżet"],
  ]),
  odstep(360),
  p([t("O ankiecie", { bold: true, size: 24, color: AKCENT })], { spacing: { after: 80 } }),
  p("Ankietę wypełniło 11 osób między 30 października a 6 listopada 2026 r. Miała 10 pytań: kim są respondenci, jak korzystają ze słuchawek, co jest dla nich ważne przy zakupie, ile chcą wydać, gdzie kupują i czego im brakuje w słuchawkach, których używają dziś."),
  p("W pytaniach, w których można było zaznaczyć kilka odpowiedzi, procenty nie sumują się do 100%."),
  odstep(400),
  p([t("Listopad 2026", { color: SZARY })]),
];

const tresc = [
  h1("", "Najważniejsze wnioski"),
  li([t("Wygoda wygrywa. ", { bold: true }), t("Wygodę noszenia wskazało 10 z 11 osób. Wygląd i design wybrała tylko 1 osoba.")]),
  li([t("Słuchawki to sprzęt, nie ozdoba. ", { bold: true }), t("Zdanie „słuchawki to element mojego stylu” dostało średnio 2,9 na 5.")]),
  li([t("Średnia półka cenowa. ", { bold: true }), t("Najczęstszy budżet to 201–400 zł, a 7 z 11 osób nie wyda więcej niż 400 zł.")]),
  li([t("Zakupy w internecie. ", { bold: true }), t("9 z 11 osób kupuje słuchawki online, najczęściej w internetowych sklepach pośredników.")]),
  li([t("Marka dzieli. ", { bold: true }), t("4 osoby zdecydowanie ufają sprawdzonym markom, takim jak JBL, a 4 raczej nie. Pracujący ufają markom bardziej (średnio 4,0) niż studenci (2,8).")]),
  li([t("Czego brakuje? ", { bold: true }), t("Najczęstsza uwaga to lepsze wyciszenie otoczenia.")]),

  h1("1", "Kim są respondenci"),
  wykres("1_profil.png"),
  p("Większość ankietowanych to młode osoby: 6 z 11 ma 18–24 lata, a 4 osoby mają 25–34 lata. Tylko jedna osoba ma 45 lat lub więcej. Pracujących na etacie (6 osób) i uczących się lub studiujących (5 osób) jest prawie po równo. Wśród osób w wieku 18–24 lata przeważają studenci (4 z 6), a wśród osób w wieku 25–34 lata pracujący (3 z 4)."),
  wniosek("odbiorcy to głównie młodzi ludzie, którzy pracują albo studiują."),

  h1("2", "Jak korzystają ze słuchawek"),
  h2("Ile czasu dziennie spędzają w słuchawkach"),
  wykres("2_czas.png"),
  p("Ponad połowa osób (6) używa słuchawek do godziny dziennie. Druga duża grupa (4 osoby) ma je na uszach ponad 3 godziny dziennie. Od 1 do 3 godzin korzysta z nich tylko 1 osoba."),
  h2("W jakich sytuacjach ich używają"),
  wykres("3_sytuacje.png"),
  p("Najczęściej słuchawek używa się podczas grania (5 osób). Zaraz potem są: w drodze, w pracy lub podczas nauki i w domu dla relaksu (po 4 osoby). Najrzadziej na treningu (3 osoby)."),
  wniosek("są dwie grupy: krótkie, codzienne korzystanie (do godziny) i długie, kilkugodzinne sesje. Słuchawki towarzyszą w wielu sytuacjach, a nie w jednej."),

  h1("3", "Co jest najważniejsze przy wyborze słuchawek"),
  wykres("4_wazne.png"),
  p("Zdecydowanie najważniejsza jest wygoda noszenia, którą wybrało 10 z 11 osób. Dalej są jakość dźwięku z mocnym basem i mocna bateria (po 7 osób) oraz wyciszenie otoczenia ANC (6 osób). Cena i promocje liczą się tylko dla 3 osób, a wygląd i design tylko dla 1."),
  wniosek("liczy się to, co czuć na co dzień: wygoda, dźwięk i bateria. Cena i wygląd są na dalszym planie."),

  h1("4", "Ile są gotowi zapłacić"),
  wykres("5_cena.png"),
  p("Najwięcej osób (4) wyda na nowe słuchawki bezprzewodowe 201–400 zł. Po 3 osoby wybrały do 200 zł i 401–700 zł, a tylko 1 osoba zapłaciłaby ponad 700 zł."),
  wniosek("najważniejsza jest średnia półka cenowa, do 400 zł. Słuchawki premium powyżej 700 zł interesują pojedyncze osoby."),

  h1("5", "Styl i zaufanie do marki"),
  wykres("6_skale.png"),
  p("Słuchawki raczej nie są dla respondentów elementem stylu. Średnia ocena to 2,9 na 5, a odpowiedzi rozkładają się dość równo. Pasuje to do tego, że wygląd i design wybrała tylko 1 osoba."),
  p("Opinie o zaufaniu do sprawdzonych marek, takich jak JBL, są podzielone. 4 osoby zdecydowanie się zgadzają, ale 4 raczej się nie zgadzają. Średnia to 3,5 na 5 i nikt nie wybrał odpowiedzi 1."),
  wniosek("dla części klientów marka jest gwarancją jakości, ale dla drugiej części sama marka nie wystarczy."),

  h1("6", "Gdzie kupują"),
  wykres("7_gdzie.png"),
  p("Najwięcej osób (5) kupuje w internetowych sklepach pośredników. Po 2 osoby wybierają sklep stacjonarny, stronę producenta albo platformy typu Allegro czy Amazon."),
  wniosek("aż 9 z 11 osób kupuje przez internet, więc to tam trzeba być widocznym."),

  h1("7", "Czego brakuje w obecnych słuchawkach"),
  p("Na to pytanie odpowiedziało 7 z 11 osób. Odpowiedzi pogrupowaliśmy tematycznie."),
  tabela(["Uwaga", "Liczba osób"], [
    ["Lepsze wyciszenie otoczenia", "2"],
    ["Większa wygoda (słuchawki za bardzo uciskają głowę)", "1"],
    ["Lepsza jakość dźwięku", "1"],
    ["Możliwość złożenia słuchawek", "1"],
    ["Bez komunikatu głosowego o ładowaniu w trakcie muzyki", "1"],
    ["Niczego nie brakuje (zadowolony użytkownik JBL)", "1"],
  ], [7226, 1800]),
  odstep(),
  h2("Wybrane wypowiedzi"),
  cytat("Nic, bo używam JBL. Są wygodne, nie uciskają mnie i nie mam wrażenia, że moja głowa dźwiga coś ciężkiego. I fajnie wyciszają szumy z otoczenia jak potrzebuję się nad czymś mocno skupić."),
  cytat("za bardzo uciskają w głowe"),
  cytat("Usunęłabym głos kobiety w trakcie piosenek o potrzebie ładowania"),
  cytat("możliwości złożenia ich"),
  wniosek("najczęściej brakuje lepszego wyciszenia. Pozostałe uwagi dotyczą wygody, dźwięku, składania i komunikatów głosowych."),

  h1("8", "Studenci a pracujący"),
  p("Porównaliśmy dwie grupy: osoby, które uczą się lub studiują (5 osób), i osoby pracujące na etacie (6 osób). Wartości to odsetek osób w danej grupie."),
  h2("Co jest dla nich najważniejsze"),
  wykres("8_wazne_grupy.png"),
  p("Dla pracujących najważniejsza jest wygoda (wszyscy) i mocna bateria (5 z 6). Studenci równie mocno cenią wygodę i jakość dźwięku (po 4 z 5) i częściej zwracają uwagę na cenę i promocje."),
  h2("Kiedy używają słuchawek"),
  wykres("9_sytuacje_grupy.png"),
  p("Studenci najczęściej używają słuchawek podczas grania (3 z 5). Pracujący głównie w pracy lub podczas nauki i w domu dla relaksu (po 3 z 6)."),
  h2("Ile chcą wydać"),
  wykres("10_cena_grupy.png"),
  p("Budżet pracujących rozkłada się równo do 700 zł. Jedyna osoba gotowa wydać ponad 700 zł uczy się lub studiuje."),
  h2("Styl i zaufanie do marki"),
  wykres("11_skale_grupy.png"),
  p("W ocenie słuchawek jako elementu stylu obie grupy są podobne (2,8 i 3,0). Duża różnica jest w zaufaniu do sprawdzonych marek: pracujący średnio 4,0, studenci 2,8."),
  tabela(["", "Studenci i uczniowie (5 osób)", "Pracujący (6 osób)"], [
    ["Najczęstsze użycie", "podczas grania (3)", "w pracy lub nauce i w domu dla relaksu (po 3)"],
    ["Najważniejsze cechy", "wygoda i jakość dźwięku (po 4)", "wygoda (6) i mocna bateria (5)"],
    ["Budżet", "najczęściej 201–400 zł (2)", "po 2 osoby: do 200 zł, 201–400 zł, 401–700 zł"],
    ["Słuchawki jako element stylu (średnio)", "2,8", "3,0"],
    ["Zaufanie do sprawdzonych marek (średnio)", "2,8", "4,0"],
  ], [2826, 3100, 3100]),
  odstep(),
  wniosek("pracujący to klienci „na wygodę i baterię”, którzy ufają marce. Studenci to gracze, którzy szukają dobrego dźwięku i patrzą na cenę."),

  nowaStrona(),
  h1("", "Rekomendacje"),
  li([t("Komunikować wygodę. ", { bold: true }), t("To najważniejsza cecha dla 10 z 11 osób. Warto pokazywać, że słuchawki są lekkie, nie uciskają i można je nosić godzinami.")], "numery"),
  li([t("Wyciszenie w średniej półce. ", { bold: true }), t("ANC jest ważne dla ponad połowy osób i to najczęstsza uwaga. Dobrze, żeby było dostępne w modelach do 400 zł.")], "numery"),
  li([t("Skupić się na sprzedaży w internecie. ", { bold: true }), t("9 z 11 osób kupuje online, głównie w sklepach internetowych pośredników i na platformach typu Allegro czy Amazon.")], "numery"),
  li([t("Oferta dla studentów. ", { bold: true }), t("Modele do grania z dobrym dźwiękiem i promocje, bo studenci częściej zwracają uwagę na cenę.")], "numery"),
  li([t("Oferta dla pracujących. ", { bold: true }), t("Długa bateria i wygoda na cały dzień pracy. Ta grupa ufa marce, więc warto mocniej ją eksponować.")], "numery"),
  li([t("Drobne poprawki z uwag klientów. ", { bold: true }), t("Możliwość wyłączenia komunikatu głosowego o baterii i składana konstrukcja.")], "numery"),

  h1("", "Uwagi do badania"),
  li("Ankietę wypełniło 11 osób, więc wyniki pokazują kierunki, a nie dokładne udziały wśród wszystkich klientów."),
  li("W pytaniach 4 i 5 część osób zaznaczyła inną liczbę odpowiedzi niż w poleceniu (np. jedna osoba 6 cech zamiast 3). Policzyliśmy wszystkie zaznaczenia."),
  li("Pytanie 10 było otwarte. Odpowiedzi pogrupowaliśmy tematycznie, a cytaty podajemy w oryginalnej pisowni."),
];

const skroc = (s) => s
  .replace("Na platformach typu Allegro czy Amazon", "Allegro, Amazon")
  .replace("W sklepach pośrednich online", "Sklepy internetowe")
  .replace("W sklepie stacjonarnym", "Sklep stacjonarny")
  .replace("Na stronie producenta", "Strona producenta")
  .replace("W pracy lub podczas nauki", "Praca lub nauka")
  .replace("W domu dla relaksu", "Dom, relaks")
  .replace("Jakość dźwięku i mocny bas", "Dźwięk i bas")
  .replace("Wyciszenie otoczenia (ANC)", "ANC")
  .replace("Dobra cena i promocje", "Cena")
  .replace("Wygląd i design", "Wygląd")
  .replace("Wygoda noszenia", "Wygoda")
  .replace("Mocna bateria", "Bateria")
  .replace("Nauka lub studia", "Nauka / studia")
  .replace("Praca na etacie", "Praca na etacie");

const zalacznik = [
  h1("", "Załącznik: wszystkie odpowiedzi"),
  p([t("Każdy wiersz to jedna osoba. Styl i marka to oceny w skali 1–5 (pytania 7 i 8).", { color: SZARY, size: 20 })]),
  tabela(
    ["Nr", "Wiek", "Zajęcie", "Czas dziennie", "Sytuacje", "Najważniejsze cechy", "Budżet", "Styl", "Marka", "Gdzie kupuje", "Czego brakuje"],
    odp.map((r, i) => [
      String(i + 1), r[1], skroc(r[2]), r[3].replace("1-3", "1–3"),
      r[4].split(";").map((x) => skroc(x.trim())).join(", "),
      r[5].split(";").map((x) => skroc(x.trim())).join(", "),
      r[6].replace("201-400", "201–400").replace("401-700", "401–700"), r[7], r[8], skroc(r[9]), r[10] || "–",
    ]),
    [480, 820, 1200, 1250, 2100, 2650, 1150, 600, 650, 1500, 2438], 16),
];

const zrobNaglowek = (szer) => new Header({
  children: [new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: szer }],
    children: [t("Opinia klienta – wyniki ankiety", { size: 18, color: SZARY }), t("\tListopad 2026", { size: 18, color: SZARY })],
  })],
});
const naglowek = zrobNaglowek(SZER);
const naglowekPoziomy = zrobNaglowek(16838 - 2000);
const stopka = new Footer({
  children: [new Paragraph({
    alignment: AlignmentType.RIGHT,
    children: [t("Strona ", { size: 18, color: SZARY }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: SZARY }),
      t(" z ", { size: 18, color: SZARY }), new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 18, color: SZARY })],
  })],
});
const marg = { top: 1300, bottom: 1200, left: 1440, right: 1440 };

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 22 } } } },
  numbering: {
    config: [
      { reference: "kropki", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } }, run: { color: AKCENT } } }] },
      { reference: "numery", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 360 } }, run: { bold: true, color: AKCENT } } }] },
    ],
  },
  sections: [
    { properties: { page: { size: { width: 11906, height: 16838 }, margin: marg } }, children: okladka },
    { properties: { page: { size: { width: 11906, height: 16838 }, margin: marg } }, headers: { default: naglowek }, footers: { default: stopka }, children: tresc },
    {
      properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 1000, bottom: 1000, left: 1000, right: 1000 } } },
      headers: { default: naglowekPoziomy }, footers: { default: stopka }, children: zalacznik,
    },
  ],
});

Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log("Zapisano", OUT, "| odpowiedzi:", odp.length); });
