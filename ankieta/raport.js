// Raport z ankiety „Opinia klienta” (Word).
// Uruchomienie: node raport.js   (wymaga pakietu docx; wykresy z wykresy.py)
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType,
  LevelFormat, BorderStyle, ImageRun, ShadingType, Footer, PageNumber, PageOrientation, VerticalAlign,
} = require("docx");

const KAT = __dirname;
const OUT = path.join(KAT, "Opinia_klienta_wyniki_ankiety.docx");
const FONT = "Calibri";
const SZARY = "595959";
const SKALA = 0.42; // px wykresu (200 dpi) -> px w dokumencie, ta sama dla wszystkich wykresów

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
const p = (text) => new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 120, line: 288 }, children: [t(text)] });
const h1 = (text) => new Paragraph({ keepNext: true, spacing: { before: 320, after: 120 }, children: [t(text, { bold: true, size: 26 })] });
const li = (text) => new Paragraph({ numbering: { reference: "kropki", level: 0 }, spacing: { after: 60, line: 288 }, children: [t(text)] });

let nrWykresu = 0, nrTabeli = 0;
const podpis = (text) => new Paragraph({ keepNext: true, alignment: AlignmentType.CENTER, spacing: { before: 160, after: 60 }, children: [t(text, { bold: true, size: 20 })] });
const zrodlo = () => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 40, after: 200 }, children: [t("Źródło: opracowanie własne na podstawie wyników ankiety.", { italics: true, size: 18, color: SZARY })] });

function wykres(plik, tytul) {
  const b = fs.readFileSync(path.join(KAT, "wykresy", plik));
  const w = b.readUInt32BE(16), h = b.readUInt32BE(20);
  nrWykresu += 1;
  return [
    podpis(`Wykres ${nrWykresu}. ${tytul}`),
    new Paragraph({ keepNext: true, alignment: AlignmentType.CENTER, spacing: { after: 0 },
      children: [new ImageRun({ type: "png", data: b, transformation: { width: Math.round(w * SKALA), height: Math.round(h * SKALA) } })] }),
    zrodlo(),
  ];
}

function tabela(tytul, naglowki, wiersze, widths, rozmiar = 20) {
  const linia = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
  const ramka = { top: linia, bottom: linia, left: linia, right: linia };
  const komorka = (text, w, nag) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    borders: ramka,
    shading: nag ? { type: ShadingType.CLEAR, fill: "D9E2F3", color: "auto" } : undefined,
    margins: { top: 50, bottom: 50, left: 80, right: 80 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ children: [t(text, { bold: nag, size: rozmiar })] })],
  });
  nrTabeli += 1;
  return [
    podpis(`Tabela ${nrTabeli}. ${tytul}`),
    new Table({
      width: { size: widths.reduce((s, w) => s + w, 0), type: WidthType.DXA },
      columnWidths: widths,
      rows: [
        new TableRow({ tableHeader: true, children: naglowki.map((n, i) => komorka(n, widths[i], true)) }),
        ...wiersze.map((w) => new TableRow({ cantSplit: true, children: w.map((v, i) => komorka(v, widths[i], false)) })),
      ],
    }),
    zrodlo(),
  ];
}

// --- treść ----------------------------------------------------------------------
const tresc = [
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 80 }, children: [t("Raport z badania ankietowego „Opinia klienta”", { bold: true, size: 34 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [t("Ankieta dotycząca korzystania ze słuchawek bezprzewodowych", { size: 24 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [t("Listopad 2026", { size: 22, color: SZARY })] }),

  h1("Wstęp"),
  p("Celem ankiety było sprawdzenie, jak respondenci korzystają ze słuchawek bezprzewodowych, na co zwracają uwagę przy ich wyborze, ile są w stanie na nie wydać i gdzie je kupują. Zapytaliśmy też, czego brakuje im w słuchawkach, których używają obecnie."),
  p("Ankieta była prowadzona przez internet (Formularze Google) od 30 października do 6 listopada 2026 r. Wypełniło ją 11 osób. Składała się z 10 pytań, w tym jednego otwartego. Ze względu na małą liczbę odpowiedzi wyniki trzeba traktować orientacyjnie. W pytaniach, w których można było zaznaczyć kilka odpowiedzi, procenty nie sumują się do 100%."),

  h1("1. Charakterystyka respondentów"),
  p("W ankiecie wzięły udział głównie młode osoby. Ponad połowa (6 osób) ma 18-24 lata, 4 osoby mają 25-34 lata, a tylko jedna osoba ma 45 lat lub więcej. Jeśli chodzi o zajęcie, 6 osób pracuje na etacie, a 5 uczy się lub studiuje. Wśród najmłodszych respondentów przeważają studenci (4 z 6), natomiast w grupie 25-34 lata więcej jest osób pracujących (3 z 4)."),
  ...wykres("1_profil.png", "Wiek i zajęcie respondentów"),

  h1("2. Korzystanie ze słuchawek"),
  p("Najwięcej osób, bo 6, używa słuchawek do godziny dziennie. Spora grupa (4 osoby) korzysta z nich ponad 3 godziny dziennie, a od 1 do 3 godzin tylko jedna osoba. Widać więc podział na osoby, które zakładają słuchawki na chwilę, i takie, które mają je na uszach przez dużą część dnia."),
  ...wykres("2_czas.png", "Czas spędzany dziennie w słuchawkach"),
  p("W kolejnym pytaniu można było wybrać maksymalnie dwie sytuacje, w których respondenci najczęściej używają słuchawek. Najwięcej wskazań miało granie (5 osób). Po 4 osoby wybrały korzystanie w drodze, w pracy lub podczas nauki oraz w domu dla relaksu, a 3 osoby trening. Odpowiedzi rozkładają się dość równo, więc słuchawki są używane w bardzo różnych sytuacjach."),
  ...wykres("3_sytuacje.png", "Sytuacje, w których respondenci używają słuchawek"),

  h1("3. Najważniejsze cechy słuchawek"),
  p("W tym pytaniu należało wybrać trzy najważniejsze cechy. Zdecydowanie najczęściej wybierana była wygoda noszenia, którą wskazało 10 z 11 osób. Na kolejnych miejscach znalazły się jakość dźwięku i mocny bas oraz mocna bateria (po 7 osób), a także wyciszenie otoczenia ANC (6 osób). Cena i promocje były ważne tylko dla 3 osób, a wygląd i design dla jednej."),
  p("Można z tego wnioskować, że przy wyborze słuchawek liczą się przede wszystkim cechy, które odczuwa się podczas codziennego używania, a wygląd ma niewielkie znaczenie."),
  ...wykres("4_wazne.png", "Najważniejsze cechy przy wyborze słuchawek"),

  h1("4. Cena"),
  p("Najwięcej osób (4) jest gotowych wydać na nowe słuchawki bezprzewodowe od 201 do 400 zł. Po 3 osoby wybrały kwotę do 200 zł oraz 401-700 zł, a tylko jedna osoba zapłaciłaby ponad 700 zł. Łącznie 7 z 11 respondentów nie wydałoby więcej niż 400 zł, czyli najbardziej interesuje ich średnia półka cenowa."),
  ...wykres("5_cena.png", "Maksymalna kwota, jaką respondenci wydadzą na nowe słuchawki"),

  h1("5. Styl i zaufanie do marki"),
  p("Dwa pytania dotyczyły stosunku do słuchawek i do marek. Odpowiadano w skali od 1 (zdecydowanie nie) do 5 (zdecydowanie tak)."),
  p("Ze zdaniem, że słuchawki są ważnym elementem stylu, respondenci raczej się nie zgadzali, bo średnia ocena wyniosła 2,9. Odpowiedzi były dość rozproszone, najczęściej wybierano 3 i 4 (po 3 osoby). Zgadza się to z wcześniejszym wynikiem, w którym wygląd był ważny tylko dla jednej osoby."),
  p("W pytaniu o zaufanie do sprawdzonych marek, takich jak JBL, zdania były podzielone. 4 osoby zaznaczyły 5, czyli zdecydowanie ufają takim markom, ale tyle samo osób wybrało 2. Średnia wyniosła 3,5, a nikt nie zaznaczył odpowiedzi 1."),
  ...wykres("6_skale.png", "Odpowiedzi na pytania o styl i zaufanie do marki"),

  h1("6. Miejsce zakupu"),
  p("Najwięcej respondentów (5 osób) kupuje słuchawki w internetowych sklepach pośredników. Pozostałe odpowiedzi rozłożyły się po równo: po 2 osoby kupują w sklepie stacjonarnym, na stronie producenta i na platformach typu Allegro czy Amazon. Oznacza to, że 9 z 11 osób robi takie zakupy przez internet."),
  ...wykres("7_gdzie.png", "Miejsce zakupu słuchawek"),

  h1("7. Uwagi respondentów"),
  p("Ostatnie pytanie było otwarte i dotyczyło tego, czego brakuje respondentom w słuchawkach, których obecnie używają. Odpowiedziało na nie 7 osób. Dwie z nich chciałyby lepszego wyciszenia otoczenia. Pozostałe uwagi pojawiły się po jednym razie: słuchawki za bardzo uciskają głowę, brakuje dobrej jakości dźwięku i możliwości złożenia słuchawek, a jedna osoba chciałaby, żeby w trakcie muzyki nie pojawiał się komunikat głosowy o potrzebie ładowania. Jedna osoba odpisała, że niczego jej nie brakuje, bo używa słuchawek JBL, które są wygodne i dobrze wyciszają otoczenie."),
  ...tabela("Czego brakuje respondentom w obecnych słuchawkach", ["Uwaga", "Liczba osób"], [
    ["Lepsze wyciszenie otoczenia", "2"],
    ["Słuchawki za bardzo uciskają głowę", "1"],
    ["Lepsza jakość dźwięku", "1"],
    ["Możliwość złożenia słuchawek", "1"],
    ["Brak komunikatu głosowego o ładowaniu w trakcie muzyki", "1"],
    ["Niczego nie brakuje", "1"],
    ["Brak odpowiedzi", "4"],
  ], [7226, 1800]),

  h1("8. Porównanie studentów i osób pracujących"),
  p("Na koniec porównaliśmy odpowiedzi osób uczących się lub studiujących (5 osób) i pracujących na etacie (6 osób). Na wykresach podano odsetek osób w danej grupie."),
  p("Dla osób pracujących najważniejsza jest wygoda (wszystkie 6 osób) i mocna bateria (5 osób). Studenci równie często wybierali wygodę i jakość dźwięku (po 4 z 5) i częściej niż pracujący zwracali uwagę na cenę i promocje."),
  ...wykres("8_wazne_grupy.png", "Najważniejsze cechy słuchawek w obu grupach"),
  p("Różnice widać też w sytuacjach, w których używane są słuchawki. Studenci najczęściej korzystają z nich podczas grania (3 z 5), a pracujący w pracy lub podczas nauki oraz w domu dla relaksu (po 3 z 6)."),
  ...wykres("9_sytuacje_grupy.png", "Sytuacje, w których obie grupy używają słuchawek"),
  p("Budżet osób pracujących rozkłada się równo w przedziałach do 700 zł. Jedyna osoba, która wydałaby ponad 700 zł, uczy się lub studiuje."),
  ...wykres("10_cena_grupy.png", "Maksymalna kwota na słuchawki w obu grupach"),
  p("Największa różnica dotyczy zaufania do sprawdzonych marek. Osoby pracujące oceniły je średnio na 4,0, a studenci na 2,8. W ocenie słuchawek jako elementu stylu obie grupy wypadły podobnie (3,0 i 2,8)."),
  ...tabela("Porównanie studentów i osób pracujących", ["", "Studenci (5 osób)", "Pracujący (6 osób)"], [
    ["Najczęstsze użycie", "podczas grania (3)", "w pracy lub nauce i w domu dla relaksu (po 3)"],
    ["Najważniejsze cechy", "wygoda i jakość dźwięku (po 4)", "wygoda (6) i mocna bateria (5)"],
    ["Najczęstszy budżet", "201-400 zł (2)", "po 2 osoby: do 200 zł, 201-400 zł, 401-700 zł"],
    ["Słuchawki jako element stylu (średnia)", "2,8", "3,0"],
    ["Zaufanie do sprawdzonych marek (średnia)", "2,8", "4,0"],
  ], [2826, 3100, 3100]),

  h1("9. Podsumowanie i wnioski"),
  p("Ankieta pokazuje, że dla respondentów słuchawki są przede wszystkim sprzętem do codziennego użytku, a nie elementem stylu. Przy wyborze najważniejsza jest wygoda, a w dalszej kolejności jakość dźwięku, bateria i wyciszenie otoczenia. Większość osób szuka słuchawek w średniej cenie, do 400 zł, i kupuje je przez internet."),
  p("Zaufanie do znanych marek, takich jak JBL, jest podzielone. Bardziej ufają im osoby pracujące, a studenci częściej zwracają uwagę na cenę i używają słuchawek głównie do grania."),
  p("Na podstawie wyników firma mogłaby:"),
  li("w reklamach mocniej podkreślać wygodę noszenia,"),
  li("oferować wyciszenie otoczenia (ANC) w modelach do 400 zł,"),
  li("skupić sprzedaż i promocje na sklepach internetowych,"),
  li("przygotować osobną ofertę dla studentów i graczy, np. z promocjami,"),
  li("dodać możliwość wyłączenia komunikatów głosowych i pomyśleć o składanej konstrukcji."),
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
  .replace("Nauka lub studia", "Nauka / studia");

const zalacznik = [
  h1("Załącznik. Zestawienie wszystkich odpowiedzi"),
  p("Każdy wiersz to odpowiedzi jednej osoby. W kolumnach „Styl” i „Marka” podano oceny w skali od 1 do 5 (pytania 7 i 8)."),
  ...tabela("Wszystkie odpowiedzi z ankiety",
    ["Nr", "Wiek", "Zajęcie", "Czas dziennie", "Sytuacje", "Najważniejsze cechy", "Budżet", "Styl", "Marka", "Gdzie kupuje", "Czego brakuje"],
    odp.map((r, i) => [
      String(i + 1), r[1].replace("–", "-"), skroc(r[2]), r[3],
      r[4].split(";").map((x) => skroc(x.trim())).join(", "),
      r[5].split(";").map((x) => skroc(x.trim())).join(", "),
      r[6], r[7], r[8], skroc(r[9]), r[10] || "-",
    ]),
    [480, 820, 1200, 1250, 2100, 2650, 1150, 600, 650, 1500, 2438], 16),
];

const stopka = new Footer({
  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 20 })] })],
});

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 22 } } } },
  numbering: {
    config: [{ reference: "kropki", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }],
  },
  sections: [
    { properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1418, bottom: 1418, left: 1418, right: 1418 } } }, footers: { default: stopka }, children: tresc },
    {
      properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 1000, bottom: 1000, left: 1000, right: 1000 } } },
      footers: { default: stopka }, children: zalacznik,
    },
  ],
});

Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log("Zapisano", OUT, "| odpowiedzi:", odp.length); });
