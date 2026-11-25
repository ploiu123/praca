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
const SKALA = 0.40; // px wykresu (200 dpi) -> px w dokumencie, ta sama dla wszystkich wykresów

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
const surowe = czytajCsv(path.join(KAT, "odpowiedzi.csv"));
const wszystkie = surowe.length;
// te same odpowiedzi wysłane kilka razy pod rząd liczymy jako jedną
const widziane = new Set();
const odp = surowe.filter((r) => { const k = r.slice(1).join("|"); if (widziane.has(k)) return false; widziane.add(k); return true; });
const powtorzone = wszystkie - odp.length;

// --- klocki --------------------------------------------------------------------
const t = (text, o = {}) => new TextRun({ text, font: FONT, ...o });
const p = (text) => new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 110, line: 276 }, children: [t(text)] });
const h1 = (text) => new Paragraph({ keepNext: true, spacing: { before: 320, after: 120 }, children: [t(text, { bold: true, size: 26 })] });
const li = (text) => new Paragraph({ numbering: { reference: "kropki", level: 0 }, spacing: { after: 50, line: 276 }, children: [t(text)] });

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
  p(`Ankieta była prowadzona przez internet (Formularze Google) od 30 października do 9 listopada 2026 r. Formularz wysłano ${wszystkie} razy, ale ${powtorzone + 1} zgłoszeń z 9 listopada było identycznych i przyszło w ciągu niecałych 30 sekund. Uznaliśmy, że to jedna osoba, która wysłała formularz kilka razy, dlatego tę odpowiedź policzyliśmy tylko raz. Do analizy wzięliśmy ${odp.length} odpowiedzi.`),
  p("Ankieta składała się z 10 pytań, w tym jednego otwartego. Próba nie jest duża, więc wyniki trzeba traktować orientacyjnie. W pytaniach, w których można było zaznaczyć kilka odpowiedzi, procenty nie sumują się do 100%. Część osób zaznaczyła więcej lub mniej odpowiedzi, niż było w poleceniu, ale uwzględniliśmy wszystkie wskazania."),

  h1("1. Charakterystyka respondentów"),
  p("Większość respondentów to młode osoby. Najwięcej (10 osób) ma 18-24 lata, 8 osób ma 25-34 lata, a 5 osób ma mniej niż 18 lat. Starszych respondentów jest niewielu: po 2 osoby w wieku 35-44 lata oraz 45 lat i więcej. Pracujących na etacie i uczących się lub studiujących jest tyle samo, po 11 osób, a 5 osób prowadzi własną firmę lub pracuje jako freelancer. Wszystkie osoby poniżej 18 lat się uczą, w grupie 18-24 lata studentów i pracujących jest po równo (po 5), a wśród osób od 25 lat w górę przeważają pracujący (11 z 12)."),
  ...wykres("1_profil.png", "Wiek i zajęcie respondentów"),

  h1("2. Korzystanie ze słuchawek"),
  p("Prawie połowa respondentów (13 osób) używa słuchawek ponad 3 godziny dziennie. 8 osób zakłada je najwyżej na godzinę, a 6 osób używa ich od 1 do 3 godzin. Dla dużej części badanych słuchawki są więc sprzętem na cały dzień, a nie tylko na chwilę."),
  ...wykres("2_czas.png", "Czas spędzany dziennie w słuchawkach"),
  p("W kolejnym pytaniu można było wybrać maksymalnie dwie sytuacje, w których respondenci najczęściej używają słuchawek. Najwięcej wskazań miała praca lub nauka (14 osób, czyli ponad połowa), a na drugim miejscu jest granie (12 osób). Po 9 osób używa słuchawek na treningu i w domu dla relaksu, a 8 osób w drodze. Słuchawki są więc używane w bardzo różnych sytuacjach, ale najczęściej wtedy, gdy trzeba się skupić, i podczas grania."),
  ...wykres("3_sytuacje.png", "Sytuacje, w których respondenci używają słuchawek"),

  h1("3. Najważniejsze cechy słuchawek"),
  p("W tym pytaniu należało wybrać trzy najważniejsze cechy. Najczęściej wybierana była wygoda noszenia, którą wskazało 20 z 27 osób (74%). Na drugim miejscu jest jakość dźwięku i mocny bas (18 osób). Wyciszenie otoczenia ANC wybrało 13 osób, a mocną baterię 12, czyli mniej więcej połowa badanych. Cena i promocje były ważne dla 10 osób, a wygląd i design tylko dla 2."),
  p("Przy wyborze słuchawek liczą się więc przede wszystkim cechy, które odczuwa się podczas codziennego używania, a wygląd ma niewielkie znaczenie."),
  ...wykres("4_wazne.png", "Najważniejsze cechy przy wyborze słuchawek"),

  h1("4. Cena"),
  p("Najwięcej osób (11) jest gotowych wydać na nowe słuchawki od 401 do 700 zł. 6 osób wyda od 201 do 400 zł, a po 5 osób do 200 zł i ponad 700 zł. Ponad połowa respondentów (16 z 27) zapłaciłaby więc za słuchawki więcej niż 400 zł."),
  ...wykres("5_cena.png", "Maksymalna kwota, jaką respondenci wydadzą na nowe słuchawki"),

  h1("5. Styl i zaufanie do marki"),
  p("Dwa pytania dotyczyły stosunku do słuchawek i do marek. Odpowiadano w skali od 1 (zdecydowanie nie) do 5 (zdecydowanie tak)."),
  p("Ze zdaniem, że słuchawki są ważnym elementem stylu, respondenci zgadzali się umiarkowanie. Średnia ocena wyniosła 3,3, a najczęściej wybierano odpowiedź środkową, czyli 3 (9 osób). 12 osób zaznaczyło 4 lub 5, a 6 osób 1 lub 2. Mimo to wygląd i design znalazł się wśród najważniejszych cech tylko u 2 osób. Styl ma więc dla części osób znaczenie, ale przy samym wyborze słuchawek przegrywa z wygodą i dźwiękiem."),
  p("Zaufanie do sprawdzonych marek, takich jak JBL, jest wyższe. Średnia wyniosła 3,7, a 15 osób zaznaczyło 4 lub 5. Najczęściej wybierano odpowiedzi 3 i 5 (po 8 osób). Nikt nie zaznaczył odpowiedzi 1, a odpowiedź 2 wybrały 4 osoby."),
  ...wykres("6_skale.png", "Odpowiedzi na pytania o styl i zaufanie do marki"),

  h1("6. Miejsce zakupu"),
  p("Najwięcej respondentów (12 osób) kupuje słuchawki w sklepie stacjonarnym. 7 osób kupuje w internetowych sklepach pośredników, a po 4 osoby na stronie producenta i na platformach typu Allegro czy Amazon. Łącznie przez internet kupuje 15 z 27 osób, ale to sklep stacjonarny jest najczęściej wybieranym pojedynczym miejscem. Możliwe, że wiele osób chce przed zakupem przymierzyć słuchawki i ich posłuchać."),
  ...wykres("7_gdzie.png", "Miejsce zakupu słuchawek"),

  h1("7. Uwagi respondentów"),
  p("Ostatnie pytanie było otwarte i dotyczyło tego, czego brakuje respondentom w słuchawkach, których obecnie używają. Odpowiedziała na nie 21 osób. 4 osoby napisały, że niczego im nie brakuje, a jedna z nich dodała, że używa słuchawek JBL, które są wygodne i dobrze wyciszają otoczenie."),
  p("Najczęściej powtarzały się dwie uwagi: lepsze wyciszenie otoczenia (4 osoby) oraz wygoda i dopasowanie (4 osoby). W tej drugiej grupie respondenci pisali, że słuchawki za mocno uciskają głowę, że nauszniki z czasem się rozciągają i przydałaby się regulacja docisku albo że słuchawki źle pasują do kształtu ucha. Po 2 osoby chciałyby lepszej baterii, lepszego dźwięku i ładniejszego wyglądu. Pojedyncze osoby wspomniały o mikrofonie, łatwym przełączaniu między urządzeniami, możliwości złożenia słuchawek i o komunikacie głosowym o ładowaniu, który przerywa muzykę. Jedna osoba podała dwie uwagi, dlatego liczby w tabeli nie sumują się do 27."),
  ...tabela("Czego brakuje respondentom w obecnych słuchawkach", ["Uwaga", "Liczba osób"], [
    ["Niczego nie brakuje", "4"],
    ["Lepsze wyciszenie otoczenia", "4"],
    ["Wygoda i dopasowanie (ucisk, docisk, kształt ucha)", "4"],
    ["Lepsza bateria", "2"],
    ["Lepsza jakość dźwięku", "2"],
    ["Lepszy wygląd", "2"],
    ["Lepszy mikrofon", "1"],
    ["Łatwe przełączanie między urządzeniami", "1"],
    ["Możliwość złożenia słuchawek", "1"],
    ["Brak komunikatu głosowego o ładowaniu w trakcie muzyki", "1"],
    ["Brak odpowiedzi", "6"],
  ], [7226, 1800]),

  h1("8. Porównanie studentów i osób pracujących"),
  p("Porównaliśmy odpowiedzi osób uczących się lub studiujących (11 osób) i osób pracujących (16 osób: 11 pracuje na etacie, a 5 prowadzi własną firmę lub pracuje jako freelancer). Na wykresach podano odsetek osób w danej grupie."),
  p("Studentom najbardziej zależy na jakości dźwięku (8 z 11), a pracującym na wygodzie (13 z 16). Studenci częściej zwracali uwagę na cenę i promocje (5 z 11, a wśród pracujących 5 z 16) oraz na wyciszenie otoczenia (6 z 11 wobec 7 z 16). Wygląd wybrały tylko 2 osoby pracujące i żaden student."),
  ...wykres("8_wazne_grupy.png", "Najważniejsze cechy słuchawek w obu grupach"),
  p("Studenci zdecydowanie najczęściej korzystają ze słuchawek podczas grania (8 z 11). Pracujący używają ich głównie w pracy (10 z 16) i w domu dla relaksu (7 z 16)."),
  ...wykres("9_sytuacje_grupy.png", "Sytuacje, w których obie grupy używają słuchawek"),
  p("Budżety studentów rozkładają się prawie równo: po 3 osoby wybrały do 200 zł, 401-700 zł i ponad 700 zł. Wśród pracujących wyraźnie przeważa przedział 401-700 zł, który wybrała połowa z nich (8 z 16). Ciekawe jest to, że ponad 700 zł częściej deklarowali studenci (3 z 11) niż pracujący (2 z 16)."),
  ...wykres("10_cena_grupy.png", "Maksymalna kwota na słuchawki w obu grupach"),
  p("Pracujący bardziej ufają sprawdzonym markom: ich średnia ocena wyniosła 3,9, a studentów 3,5. W ocenie słuchawek jako elementu stylu wyniki są podobne (3,5 u studentów i 3,3 u pracujących)."),
  ...tabela("Porównanie studentów i osób pracujących", ["", "Studenci (11 osób)", "Pracujący (16 osób)"], [
    ["Ponad 3 godziny dziennie w słuchawkach", "6 osób", "7 osób"],
    ["Najczęstsze użycie", "podczas grania (8)", "w pracy lub nauce (10)"],
    ["Najważniejsze cechy", "jakość dźwięku (8), wygoda (7)", "wygoda (13), jakość dźwięku (10)"],
    ["Najczęstszy budżet", "do 200 zł, 401-700 zł i ponad 700 zł (po 3)", "401-700 zł (8)"],
    ["Gdzie najczęściej kupują", "sklep stacjonarny (6)", "sklep stacjonarny (6), sklepy internetowe (5)"],
    ["Słuchawki jako element stylu (średnia)", "3,5", "3,3"],
    ["Zaufanie do sprawdzonych marek (średnia)", "3,5", "3,9"],
  ], [2826, 3100, 3100]),

  h1("9. Podsumowanie i wnioski"),
  p("Ankieta pokazuje, że przy wyborze słuchawek najważniejsza jest dla respondentów wygoda, a zaraz po niej jakość dźwięku. Wyciszenie otoczenia i bateria są ważne dla mniej więcej połowy osób, a wygląd prawie się nie liczy, chociaż część osób traktuje słuchawki jako element swojego stylu. Prawie połowa badanych używa słuchawek ponad 3 godziny dziennie, a ponad połowa jest gotowa zapłacić za nie więcej niż 400 zł. Najczęściej wybieranym miejscem zakupu jest sklep stacjonarny, chociaż w sumie więcej osób kupuje przez internet."),
  p("Sprawdzonym markom, takim jak JBL, ufa ponad połowa respondentów, a bardziej ufają im osoby pracujące. Studenci częściej używają słuchawek do grania i bardziej zwracają uwagę na dźwięk i cenę, a pracujący szukają przede wszystkim wygody na wiele godzin pracy."),
  p("Na podstawie wyników firma mogłaby:"),
  li("w reklamach mocniej podkreślać wygodę noszenia i jakość dźwięku,"),
  li("pokazywać, że słuchawki są wygodne przez wiele godzin, np. dzięki regulacji docisku i dobremu dopasowaniu do ucha,"),
  li("dać klientom możliwość przetestowania słuchawek w sklepach stacjonarnych, a jednocześnie rozwijać sprzedaż przez internet,"),
  li("przygotować ofertę dla studentów i graczy, np. modele z dobrym dźwiękiem i promocje,"),
  li("dodać możliwość wyłączenia komunikatów głosowych i łatwe przełączanie między urządzeniami."),
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
  .replace("Freelance lub własna firma", "Freelance / firma");

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
