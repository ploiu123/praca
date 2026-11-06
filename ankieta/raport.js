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
const odp = czytajCsv(path.join(KAT, "odpowiedzi.csv"));

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
  p("Ankieta była prowadzona przez internet (Formularze Google) od 30 października do 6 listopada 2026 r. Wypełniło ją 15 osób. Składała się z 10 pytań, w tym jednego otwartego. Ze względu na małą liczbę odpowiedzi wyniki trzeba traktować orientacyjnie. W pytaniach, w których można było zaznaczyć kilka odpowiedzi, procenty nie sumują się do 100%. W kilku przypadkach respondenci zaznaczyli więcej lub mniej odpowiedzi, niż było w poleceniu, ale uwzględniliśmy wszystkie wskazania."),

  h1("1. Charakterystyka respondentów"),
  p("W ankiecie wzięły udział głównie młode osoby. Ponad połowa (8 osób) ma 18-24 lata, 5 osób ma 25-34 lata, a po jednej osobie ma mniej niż 18 lat oraz 45 lat lub więcej. Pracujących na etacie i uczących się lub studiujących jest tyle samo, po 7 osób, a jedna osoba prowadzi własną firmę lub pracuje jako freelancer. Wśród osób w wieku 18-24 lata przeważają studenci (5 z 8), a w grupie 25-34 lata osoby pracujące (4 z 5)."),
  ...wykres("1_profil.png", "Wiek i zajęcie respondentów"),

  h1("2. Korzystanie ze słuchawek"),
  p("Po 6 osób używa słuchawek do godziny dziennie albo ponad 3 godziny dziennie, a 3 osoby korzystają z nich od 1 do 3 godzin. Widać więc podział na osoby, które zakładają słuchawki na chwilę, i takie, które mają je na uszach przez dużą część dnia."),
  ...wykres("2_czas.png", "Czas spędzany dziennie w słuchawkach"),
  p("W kolejnym pytaniu można było wybrać maksymalnie dwie sytuacje, w których respondenci najczęściej używają słuchawek. Najwięcej wskazań miało granie (8 osób, czyli ponad połowa). Na drugim miejscu jest korzystanie w drodze (6 osób), a po 5 osób wybrało pracę lub naukę, trening i relaks w domu. Słuchawki są więc używane w bardzo różnych sytuacjach, ale granie wyraźnie się wyróżnia."),
  ...wykres("3_sytuacje.png", "Sytuacje, w których respondenci używają słuchawek"),

  h1("3. Najważniejsze cechy słuchawek"),
  p("W tym pytaniu należało wybrać trzy najważniejsze cechy. Zdecydowanie najczęściej wybierana była wygoda noszenia, którą wskazało 13 z 15 osób. Na drugim miejscu jest jakość dźwięku i mocny bas (11 osób), a dalej wyciszenie otoczenia ANC (8 osób) i mocna bateria (7 osób). Cena i promocje były ważne dla 4 osób, a wygląd i design tylko dla jednej."),
  p("Można z tego wnioskować, że przy wyborze słuchawek liczą się przede wszystkim cechy, które odczuwa się podczas codziennego używania, a wygląd ma niewielkie znaczenie."),
  ...wykres("4_wazne.png", "Najważniejsze cechy przy wyborze słuchawek"),

  h1("4. Cena"),
  p("Budżety respondentów są dość zróżnicowane. Najwięcej osób (5) jest gotowych wydać na nowe słuchawki od 401 do 700 zł, 4 osoby od 201 do 400 zł, a po 3 osoby do 200 zł i ponad 700 zł. Ponad połowa respondentów (8 z 15) zapłaciłaby więc za słuchawki więcej niż 400 zł."),
  ...wykres("5_cena.png", "Maksymalna kwota, jaką respondenci wydadzą na nowe słuchawki"),

  h1("5. Styl i zaufanie do marki"),
  p("Dwa pytania dotyczyły stosunku do słuchawek i do marek. Odpowiadano w skali od 1 (zdecydowanie nie) do 5 (zdecydowanie tak)."),
  p("Ze zdaniem, że słuchawki są ważnym elementem stylu, respondenci zgadzali się umiarkowanie. Średnia ocena wyniosła 3,3, a najczęściej wybierano odpowiedzi 3 i 4 (po 4 osoby). 7 osób zaznaczyło 4 lub 5, a 4 osoby 1 lub 2. Mimo to wygląd i design znalazł się wśród najważniejszych cech tylko u jednej osoby. Styl ma więc dla części osób znaczenie, ale przy samym wyborze słuchawek przegrywa z wygodą i dźwiękiem."),
  p("W pytaniu o zaufanie do sprawdzonych marek, takich jak JBL, najwięcej osób (6) zaznaczyło 5, czyli zdecydowanie ufa takim markom. Z drugiej strony 4 osoby wybrały 2, a kolejne 4 odpowiedź środkową. Średnia wyniosła 3,6, a nikt nie zaznaczył odpowiedzi 1."),
  ...wykres("6_skale.png", "Odpowiedzi na pytania o styl i zaufanie do marki"),

  h1("6. Miejsce zakupu"),
  p("Najwięcej respondentów (5 osób) kupuje słuchawki w internetowych sklepach pośredników. Po 4 osoby kupują w sklepie stacjonarnym i na stronie producenta, a 2 osoby na platformach typu Allegro czy Amazon. Łącznie 11 z 15 osób robi takie zakupy przez internet, ale sklep stacjonarny też ma swoich zwolenników."),
  ...wykres("7_gdzie.png", "Miejsce zakupu słuchawek"),

  h1("7. Uwagi respondentów"),
  p("Ostatnie pytanie było otwarte i dotyczyło tego, czego brakuje respondentom w słuchawkach, których obecnie używają. Odpowiedziało na nie 11 osób. Najczęściej pisano, że niczego nie brakuje (4 osoby), a jedna z tych osób dodała, że używa słuchawek JBL, które są wygodne i dobrze wyciszają otoczenie. Wśród uwag najczęściej powtarzało się lepsze wyciszenie otoczenia (3 osoby). Pozostałe uwagi pojawiły się po jednym razie: lepsza bateria, słuchawki za bardzo uciskają głowę, brakuje dobrej jakości dźwięku i możliwości złożenia słuchawek, a jedna osoba chciałaby, żeby w trakcie muzyki nie pojawiał się komunikat głosowy o potrzebie ładowania. Jedna osoba podała dwie uwagi, dlatego liczby w tabeli nie sumują się do 15."),
  ...tabela("Czego brakuje respondentom w obecnych słuchawkach", ["Uwaga", "Liczba osób"], [
    ["Niczego nie brakuje", "4"],
    ["Lepsze wyciszenie otoczenia", "3"],
    ["Lepsza bateria", "1"],
    ["Słuchawki za bardzo uciskają głowę", "1"],
    ["Lepsza jakość dźwięku", "1"],
    ["Możliwość złożenia słuchawek", "1"],
    ["Brak komunikatu głosowego o ładowaniu w trakcie muzyki", "1"],
    ["Brak odpowiedzi", "4"],
  ], [7226, 1800]),

  h1("8. Porównanie studentów i osób pracujących"),
  p("Porównaliśmy odpowiedzi osób uczących się lub studiujących (7 osób) i osób pracujących (8 osób, w tym jedna osoba prowadząca własną firmę lub pracująca jako freelancer). Na wykresach podano odsetek osób w danej grupie."),
  p("Dla obu grup najważniejsza jest wygoda (6 z 7 studentów i 7 z 8 pracujących). Studenci równie często wybierali jakość dźwięku (6 z 7) i częściej zwracali uwagę na cenę i promocje (3 z 7, a wśród pracujących 1 z 8). Pracujący częściej wskazywali mocną baterię (5 z 8, a wśród studentów 2 z 7)."),
  ...wykres("8_wazne_grupy.png", "Najważniejsze cechy słuchawek w obu grupach"),
  p("Studenci zdecydowanie najczęściej korzystają ze słuchawek podczas grania (5 z 7). Pracujący używają ich głównie w pracy lub podczas nauki i w domu dla relaksu (po 4 z 8)."),
  ...wykres("9_sytuacje_grupy.png", "Sytuacje, w których obie grupy używają słuchawek"),
  p("Ciekawie wygląda porównanie budżetów. Wszystkie 3 osoby, które wydałyby na słuchawki ponad 700 zł, uczą się lub studiują. Wśród pracujących nikt nie wybrał tej odpowiedzi, za to połowa z nich (4 z 8) wydałaby od 401 do 700 zł."),
  ...wykres("10_cena_grupy.png", "Maksymalna kwota na słuchawki w obu grupach"),
  p("Największa różnica dotyczy zaufania do sprawdzonych marek. Pracujący ocenili je średnio na 4,0, a studenci na 3,1. W ocenie słuchawek jako elementu stylu obie grupy wypadły tak samo (średnio 3,3)."),
  ...tabela("Porównanie studentów i osób pracujących", ["", "Studenci (7 osób)", "Pracujący (8 osób)"], [
    ["Najczęstsze użycie", "podczas grania (5)", "w pracy lub nauce i w domu dla relaksu (po 4)"],
    ["Najważniejsze cechy", "wygoda i jakość dźwięku (po 6)", "wygoda (7), jakość dźwięku i bateria (po 5)"],
    ["Najczęstszy budżet", "powyżej 700 zł (3)", "401-700 zł (4)"],
    ["Gdzie najczęściej kupują", "sklep stacjonarny (3)", "sklepy internetowe i strona producenta (po 3)"],
    ["Słuchawki jako element stylu (średnia)", "3,3", "3,3"],
    ["Zaufanie do sprawdzonych marek (średnia)", "3,1", "4,0"],
  ], [2826, 3100, 3100]),

  h1("9. Podsumowanie i wnioski"),
  p("Ankieta pokazuje, że przy wyborze słuchawek najważniejsza jest dla respondentów wygoda, a zaraz po niej jakość dźwięku. Wyciszenie otoczenia i bateria są ważne dla mniej więcej połowy osób, a wygląd prawie się nie liczy, chociaż część osób traktuje słuchawki jako element swojego stylu. Budżety są zróżnicowane, ale ponad połowa respondentów jest gotowa zapłacić więcej niż 400 zł. Większość osób (11 z 15) kupuje słuchawki przez internet."),
  p("Sprawdzonym markom, takim jak JBL, ufa mniej więcej połowa respondentów, a bardziej ufają im osoby pracujące. Studenci częściej używają słuchawek do grania i bardziej zwracają uwagę na dźwięk i cenę, a jednocześnie to wśród nich są osoby gotowe wydać na słuchawki najwięcej."),
  p("Na podstawie wyników firma mogłaby:"),
  li("w reklamach mocniej podkreślać wygodę noszenia i jakość dźwięku,"),
  li("oferować wyciszenie otoczenia (ANC) zarówno w tańszych, jak i w droższych modelach,"),
  li("prowadzić sprzedaż głównie przez internet, ale nie rezygnować ze sklepów stacjonarnych,"),
  li("przygotować ofertę dla studentów i graczy, np. modele z dobrym dźwiękiem i promocje,"),
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
