// Projekt 3: Segmentacja rynku – słuchawki bezprzewodowe JBL (Word).
// Uruchomienie: node segmentacja.js   (wymaga pakietu docx; wykres 11_segmenty.png z wykresy.py)
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType,
  LevelFormat, BorderStyle, ImageRun, ShadingType, Footer, PageNumber, VerticalAlign,
} = require("docx");

const KAT = __dirname;
const OUT = path.join(KAT, "Segmentacja_rynku_JBL.docx");
const FONT = "Calibri";
const SZARY = "595959";

const t = (text, o = {}) => new TextRun({ text, font: FONT, ...o });
const p = (children) => new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 110, line: 276 }, children: typeof children === "string" ? [t(children)] : children });
const h1 = (text) => new Paragraph({ keepNext: true, spacing: { before: 300, after: 120 }, children: [t(text, { bold: true, size: 26 })] });
const h2 = (text) => new Paragraph({ keepNext: true, spacing: { before: 200, after: 80 }, children: [t(text, { bold: true, size: 23 })] });
const li = (children) => new Paragraph({ numbering: { reference: "kropki", level: 0 }, spacing: { after: 60, line: 276 }, children: typeof children === "string" ? [t(children)] : children });

let nrWykresu = 0, nrTabeli = 0;
const podpis = (text) => new Paragraph({ keepNext: true, alignment: AlignmentType.CENTER, spacing: { before: 160, after: 60 }, children: [t(text, { bold: true, size: 20 })] });
const zrodlo = () => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 40, after: 200 }, children: [t("Źródło: opracowanie własne na podstawie wyników ankiety.", { italics: true, size: 18, color: SZARY })] });

function wykres(plik, tytul, skala = 0.40) {
  const b = fs.readFileSync(path.join(KAT, "wykresy", plik));
  const w = b.readUInt32BE(16), h = b.readUInt32BE(20);
  nrWykresu += 1;
  return [
    podpis(`Wykres ${nrWykresu}. ${tytul}`),
    new Paragraph({ keepNext: true, alignment: AlignmentType.CENTER, children: [new ImageRun({ type: "png", data: b, transformation: { width: Math.round(w * skala), height: Math.round(h * skala) } })] }),
    zrodlo(),
  ];
}

// komórka może być tekstem albo obiektem { text, bold, fill, center }
function tabela(tytul, naglowki, wiersze, widths, o = {}) {
  const linia = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
  const ramka = { top: linia, bottom: linia, left: linia, right: linia };
  const rozmiar = o.rozmiar || 20;
  const komorka = (v, w, nag, pierwsza) => {
    const c = typeof v === "string" ? { text: v } : v;
    return new TableCell({
      width: { size: w, type: WidthType.DXA },
      borders: ramka,
      shading: nag || c.fill ? { type: ShadingType.CLEAR, fill: c.fill || "D9E2F3", color: "auto" } : undefined,
      margins: { top: 50, bottom: 50, left: 80, right: 80 },
      verticalAlign: VerticalAlign.CENTER,
      children: c.text.split("\n").map((linijka, i) => new Paragraph({
        alignment: c.center ? AlignmentType.CENTER : AlignmentType.LEFT,
        children: [t(linijka, { bold: nag || c.bold || (o.pierwszaPogrubiona && pierwsza) || (c.center && i === 0), size: rozmiar })],
      })),
    });
  };
  nrTabeli += 1;
  return [
    podpis(`Tabela ${nrTabeli}. ${tytul}`),
    new Table({
      width: { size: widths.reduce((s, w) => s + w, 0), type: WidthType.DXA },
      columnWidths: widths,
      rows: [
        new TableRow({ tableHeader: true, children: naglowki.map((n, i) => komorka(n, widths[i], true, i === 0)) }),
        ...wiersze.map((w) => new TableRow({ cantSplit: true, children: w.map((v, i) => komorka(v, widths[i], false, i === 0)) })),
      ],
    }),
    zrodlo(),
  ];
}

const tresc = [
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: 80 }, children: [t("Segmentacja rynku słuchawek bezprzewodowych JBL", { bold: true, size: 34 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [t("Projekt 3: Segmentacja rynku", { size: 24 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [t("Listopad 2026", { size: 22, color: SZARY })] }),

  h1("1. Wybór produktu"),
  p("Wybraliśmy słuchawki bezprzewodowe marki JBL. To produkt, który zna prawie każdy, a słuchawek używają zarówno uczniowie i studenci, jak i osoby pracujące. Na rynku jest dużo modeli w różnych cenach, od tanich słuchawek dousznych po drogie modele z wyciszaniem otoczenia, więc klienci mogą się bardzo różnić tym, czego szukają i ile chcą zapłacić."),

  h1("2. Cel badania"),
  p("Celem badania było pomóc marce JBL lepiej zrozumieć swoich klientów. Chcieliśmy sprawdzić, jak respondenci korzystają ze słuchawek, na co zwracają uwagę przy ich wyborze, ile są w stanie na nie wydać i gdzie je kupują. Na tej podstawie podzieliliśmy respondentów na segmenty i wskazaliśmy, do którego z nich firma powinna kierować ofertę w pierwszej kolejności."),

  h1("3. Kryteria segmentacji"),
  p("Wybraliśmy pięć kategorii segmentacyjnych. Każdej z nich odpowiadają konkretne pytania w ankiecie."),
  ...tabela("Kryteria segmentacji i pytania w ankiecie", ["Kryterium", "Co sprawdzamy", "Pytania"], [
    ["Demograficzne", "wiek, czym się zajmuje", "1, 2"],
    ["Behawioralne", "jak długo i kiedy używa słuchawek, gdzie je kupuje", "3, 4, 9"],
    ["Preferencje (poszukiwane korzyści)", "najważniejsze cechy słuchawek, czego brakuje w obecnych", "5, 10"],
    ["Cenowe", "ile maksymalnie zapłaci za nowe słuchawki", "6"],
    ["Psychograficzne", "słuchawki jako element stylu, zaufanie do marki", "7, 8"],
  ], [2700, 5126, 1200], { pierwszaPogrubiona: true }),

  h1("4. Ankieta i przebieg badania"),
  p("Ankietę przygotowaliśmy w Formularzach Google. Ma 10 pytań różnego typu: pytania zamknięte jednokrotnego i wielokrotnego wyboru, dwa pytania w skali Likerta od 1 (zdecydowanie nie) do 5 (zdecydowanie tak) oraz jedno pytanie otwarte."),
  ...tabela("Pytania w ankiecie", ["Nr", "Pytanie", "Typ pytania", "Kryterium"], [
    ["1", "Ile masz lat?", "jednokrotny wybór", "demograficzne"],
    ["2", "Czym się obecnie zajmujesz?", "jednokrotny wybór", "demograficzne"],
    ["3", "Ile czasu dziennie spędzasz w słuchawkach?", "jednokrotny wybór", "behawioralne"],
    ["4", "W jakich sytuacjach najczęściej ich używasz? (max 2 odpowiedzi)", "wielokrotny wybór", "behawioralne"],
    ["5", "Co jest dla Ciebie najważniejsze, gdy wybierasz słuchawki? (zaznacz 3)", "wielokrotny wybór", "preferencje"],
    ["6", "Ile maksymalnie zapłacisz za nowe słuchawki bezprzewodowe?", "jednokrotny wybór", "cenowe"],
    ["7", "„Słuchawki to dla mnie nie tylko sprzęt, ale też ważny element mojego stylu”", "skala Likerta 1-5", "psychograficzne"],
    ["8", "„Kupując elektronikę, ufam tylko sprawdzonym markom, takim jak JBL”", "skala Likerta 1-5", "psychograficzne"],
    ["9", "Gdzie najchętniej kupujesz taki sprzęt?", "jednokrotny wybór", "behawioralne"],
    ["10", "Czego najbardziej brakuje Ci w słuchawkach, których używasz dzisiaj?", "otwarte", "preferencje"],
  ], [500, 5026, 1800, 1700]),
  p("Odpowiedzi zbieraliśmy od 30 października do 9 listopada 2026 r. Formularz wysłano 32 razy, ale 6 zgłoszeń z 9 listopada było identycznych i przyszło w ciągu niecałych 30 sekund, więc policzyliśmy je jako jedną odpowiedź. Do analizy wzięliśmy 27 osób. Wśród respondentów są dwie wyraźne grupy: osoby uczące się lub studiujące (11 osób) i osoby pracujące (16 osób). Wyniki obu grup porównaliśmy w punkcie 6."),

  h1("5. Sposób analizy wyników"),
  p("Respondentów podzieliliśmy według dwóch kryteriów, które najbardziej różnicowały odpowiedzi:"),
  li([t("intensywność używania ", { bold: true }), t("(kryterium behawioralne, pytanie 3): do godziny dziennie albo dłużej,")]),
  li([t("budżet ", { bold: true }), t("(kryterium cenowe, pytanie 6): do 400 zł albo ponad 400 zł.")]),
  p("Z połączenia tych dwóch podziałów powstały cztery segmenty, podobnie jak w przykładzie z wykładu, gdzie klientów dzielono według tego, czy kupują drogie czy tanie produkty i czy robią to często czy rzadko. Pozostałe pytania (wiek, zajęcie, ważne cechy, miejsce zakupu, styl i zaufanie do marki) posłużyły do opisania, czym segmenty różnią się od siebie."),
  ...tabela("Podział respondentów na segmenty", ["", "Do godziny dziennie", "Ponad godzinę dziennie"], [
    [{ text: "Budżet ponad 400 zł", bold: true }, { text: "C Wierni marce\n2 osoby (7%)", center: true }, { text: "A Entuzjaści dźwięku\n14 osób (52%)", center: true, fill: "FBE4D5" }],
    [{ text: "Budżet do 400 zł", bold: true }, { text: "D Okazjonalni oszczędni\n6 osób (22%)", center: true }, { text: "B Praktyczni pracujący\n5 osób (19%)", center: true }],
  ], [2226, 3400, 3400]),
  ...wykres("11_segmenty.png", "Liczebność segmentów"),

  h1("6. Prezentacja wyników"),
  h2("Opis segmentów"),
  p([t("Segment A: Entuzjaści dźwięku (14 osób, 52%). ", { bold: true }), t("Młode osoby (9 z 14 ma mniej niż 25 lat), uczące się albo pracujące, które mają słuchawki na uszach ponad godzinę dziennie (10 z 14 ponad 3 godziny) i są gotowe wydać na nie ponad 400 zł, a 5 z nich nawet ponad 700 zł. Najbardziej zależy im na jakości dźwięku i wygodzie (po 10 z 14 osób), a połowa wskazała wyciszenie otoczenia. Najczęściej kupują w sklepie stacjonarnym (9 z 14). Dla tej grupy słuchawki są też najbardziej elementem stylu (średnio 3,8).")]),
  p([t("Segment B: Praktyczni pracujący (5 osób, 19%). ", { bold: true }), t("Głównie osoby pracujące na etacie (4 z 5), najczęściej w wieku 25-34 lata. Używają słuchawek długo, przede wszystkim w pracy (4 z 5), ale nie chcą wydawać więcej niż 400 zł, a 3 z 5 wyda najwyżej 200 zł. Wszyscy wskazali wygodę noszenia, a 3 z 5 baterię i cenę. Kupują głównie w sklepach internetowych (4 z 5).")]),
  p([t("Segment C: Wierni marce (2 osoby, 7%). ", { bold: true }), t("Osoby pracujące, które używają słuchawek krótko, ale wydadzą na nie od 401 do 700 zł. Kupują przez internet (Allegro, strona producenta). Obie osoby zaznaczyły, że ufają tylko sprawdzonym markom (ocena 5), a styl nie ma dla nich większego znaczenia (średnio 2,5). W tym segmencie jest osoba, która napisała, że niczego jej nie brakuje, bo używa wygodnych słuchawek JBL.")]),
  p([t("Segment D: Okazjonalni oszczędni (6 osób, 22%). ", { bold: true }), t("Głównie uczniowie i studenci (4 z 6), którzy używają słuchawek najwyżej godzinę dziennie, najczęściej do grania, i mają budżet do 400 zł. Zwracają uwagę na dźwięk i baterię (po 4 z 6), a połowa na cenę i promocje. Styl jest dla nich mało ważny (średnio 2,3). To najmniej zadowolona grupa: 5 z 6 osób napisało, czego brakuje im w obecnych słuchawkach, np. wygody, dłuższej baterii, lepszego dźwięku i wyciszenia.")]),
  ...tabela("Profile segmentów", ["", "A Entuzjaści dźwięku", "B Praktyczni pracujący", "C Wierni marce", "D Okazjonalni oszczędni"], [
    ["Liczebność", "14 osób (52%)", "5 osób (19%)", "2 osoby (7%)", "6 osób (22%)"],
    ["Kto", "uczniowie, studenci i pracujący, 9 z 14 poniżej 25 lat", "głównie pracujący na etacie, 25-34 lata", "pracujący na etacie", "głównie uczniowie i studenci"],
    ["Czas dziennie", "ponad godzinę, 10 z 14 ponad 3 h", "ponad godzinę", "do godziny", "do godziny"],
    ["Budżet", "ponad 400 zł, 5 osób ponad 700 zł", "do 400 zł", "401-700 zł", "do 400 zł"],
    ["Co ważne", "dźwięk i wygoda (po 10), ANC (7)", "wygoda (5), dźwięk (4)", "wygoda (2)", "dźwięk i bateria (po 4)"],
    ["Gdzie kupują", "sklep stacjonarny (9)", "sklepy internetowe (4)", "Allegro, strona producenta", "różnie"],
    ["Styl (średnia 1-5)", "3,8", "3,6", "2,5", "2,3"],
    ["Zaufanie do marki (średnia 1-5)", "3,6", "3,4", "5,0", "3,7"],
  ], [1626, 2000, 1800, 1700, 1900], { rozmiar: 18, pierwszaPogrubiona: true }),

  h2("Porównanie dwóch grup respondentów"),
  p("Osobno porównaliśmy osoby uczące się lub studiujące i osoby pracujące. Studenci najczęściej używają słuchawek do grania i bardziej zwracają uwagę na dźwięk, a pracujący używają ich głównie w pracy i najbardziej cenią wygodę. Pracujący mają wyższe zaufanie do sprawdzonych marek, a budżet ponad 400 zł deklaruje podobna część obu grup."),
  ...tabela("Studenci i pracujący", ["", "Studenci (11 osób)", "Pracujący (16 osób)"], [
    ["Ponad 3 godziny dziennie w słuchawkach", "6 osób", "7 osób"],
    ["Najczęstsze użycie", "podczas grania (8)", "w pracy lub nauce (10)"],
    ["Najważniejsza cecha", "jakość dźwięku (8)", "wygoda (13)"],
    ["Budżet ponad 400 zł", "6 osób", "10 osób"],
    ["Zakup w sklepie stacjonarnym", "6 osób", "6 osób"],
    ["Zaufanie do marki (średnia 1-5)", "3,5", "3,9"],
  ], [3426, 2800, 2800]),

  h2("Segment docelowy"),
  p([t("Jako segment docelowy wybieramy segment A, czyli Entuzjastów dźwięku. ", { bold: true }), t("Przemawiają za tym trzy liczby: to ponad połowa respondentów (14 z 27), 10 z 14 osób używa słuchawek ponad 3 godziny dziennie, a wszyscy są gotowi wydać na nie ponad 400 zł. To klienci, którzy kupują droższe modele i używają ich codziennie, więc najbardziej opłaca się do nich dotrzeć.")]),
  p([t("Pozycjonowanie: ", { bold: true }), t("słuchawki JBL z mocnym, czystym dźwiękiem, które są wygodne nawet po kilku godzinach grania, nauki czy pracy.")]),

  h1("Podsumowanie i wnioski"),
  p("Ankieta pokazuje, że klienci różnią się przede wszystkim tym, jak długo używają słuchawek i ile chcą na nie wydać. Największa grupa to osoby, które mają słuchawki na uszach przez dużą część dnia i są gotowe zapłacić ponad 400 zł. Dla większości (20 z 27) ważna jest wygoda, a zaraz po niej jakość dźwięku (18 z 27). Wygląd prawie się nie liczy. Mimo popularności zakupów przez internet aż 12 z 27 osób kupuje słuchawki w sklepie stacjonarnym, głównie w segmencie A."),
  p("Działania, które firma mogłaby podjąć wobec poszczególnych segmentów:"),
  ...tabela("Działania marketingowe dla segmentów", ["Segment", "Proponowane działania"], [
    ["A Entuzjaści dźwięku", "stanowiska do testowania dźwięku i przymierzania słuchawek w sklepach stacjonarnych, modele z ANC, reklamy pokazujące kilka godzin grania lub nauki w słuchawkach, współpraca z twórcami gamingowymi"],
    ["B Praktyczni pracujący", "modele do 400 zł z długą baterią i wygodnymi nausznikami, promocje w sklepach internetowych, łatwe przełączanie między telefonem a laptopem"],
    ["C Wierni marce", "program lojalnościowy, newsletter o nowościach, wcześniejszy dostęp do nowych modeli w sklepie producenta"],
    ["D Okazjonalni oszczędni", "rabaty i zniżki studenckie, promocje np. w Black Friday, proste modele z dobrym dźwiękiem i baterią, lepsze dopasowanie do ucha"],
  ], [2600, 6426], { pierwszaPogrubiona: true }),
  p("Trzeba pamiętać, że 27 odpowiedzi to mała próba, a segmenty C i D liczą tylko kilka osób. Wyniki pokazują kierunek, który warto sprawdzić na większej grupie klientów."),
];

const stopka = new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 20 })] })] });

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 22 } } } },
  numbering: { config: [{ reference: "kropki", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }] },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1418, bottom: 1418, left: 1418, right: 1418 } } }, footers: { default: stopka }, children: tresc }],
});

Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log("Zapisano", OUT); });
