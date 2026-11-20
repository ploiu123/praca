// Prezentacja z rozwiązań warsztatów z analityki (PowerPoint).
// Uruchomienie: NODE_PATH=<katalog z pptxgenjs> node build.js
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.APPLY_THEME);

const KAT = __dirname;
const REPO = path.resolve(KAT, "..");
const OUT = path.join(KAT, "Analityka_warsztaty.pptx");

const THEME = {
  name: "Analityka warsztaty",
  headFontFace: "Calibri",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1E2B37", lt1: "FFFFFF", dk2: "4A5866", lt2: "EDF3F2",
    accent1: "1A7F7A", accent2: "E07A2E", accent3: "7FB3AE", accent4: "C9DEDB",
    accent5: "B9472E", accent6: "6B7680", hlink: "1A7F7A", folHlink: "6B7680",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625 in
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Analityka – rozwiązania warsztatów";
const C = pres.SchemeColor;

// --- układy slajdów ---------------------------------------------------------------
pres.defineSlideMaster({
  title: "TYTUL",
  background: { color: C.text1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.5, w: 8.8, h: 1.4, fontSize: 38, bold: true, color: C.background1, valign: "bottom", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 3.05, w: 8.8, h: 1.3, fontSize: 16, color: C.accent4, valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "SEKCJA",
  background: { color: C.text1 },
  objects: [
    { placeholder: { options: { name: "nr", type: "body", x: 0.6, y: 1.25, w: 3, h: 0.9, fontSize: 44, bold: true, color: C.accent3, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 2.15, w: 8.8, h: 1.1, fontSize: 34, bold: true, color: C.background1, valign: "top", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 3.35, w: 8.8, h: 0.9, fontSize: 16, color: C.accent4, valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "TRESC",
  background: { color: C.background1 },
  slideNumber: { x: 9.15, y: 5.2, w: 0.5, h: 0.3, fontSize: 10, color: C.accent6, align: "right" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 9, h: 0.7, fontSize: 26, bold: true, color: C.text1, valign: "middle", align: "left", margin: 0 }, text: "" } },
  ],
});

// --- pomocnicze -------------------------------------------------------------------
const OS = { catAxisLabelColor: HEX.dk2, valAxisLabelColor: HEX.dk2, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
  catAxisLabelFontSize: 11, valAxisLabelFontSize: 10, dataLabelFontFace: "+mn-lt", dataLabelFontSize: 11, dataLabelColor: HEX.dk1,
  titleFontFace: "+mn-lt", titleFontSize: 13, titleColor: HEX.dk1, legendFontFace: "+mn-lt", legendFontSize: 11, legendColor: HEX.dk2 };

function slajd(tytul, sekcja) {
  const s = pres.addSlide({ masterName: "TRESC", sectionTitle: sekcja });
  s.addText(tytul, { placeholder: "title" });
  return s;
}

function sekcja(nr, tytul, opis) {
  pres.addSection({ title: tytul });
  const s = pres.addSlide({ masterName: "SEKCJA", sectionTitle: tytul });
  s.addText(nr, { placeholder: "nr" });
  s.addText(tytul, { placeholder: "title" });
  s.addText(opis, { placeholder: "body" });
  return s;
}

function tekst(s, kawalki, x, y, w, h, o = {}) {
  s.addText(kawalki, { x, y, w, h, isTextBox: true, fontSize: 14, color: C.text1, valign: "top", margin: 0, paraSpaceAfter: 6, ...o });
}

function punkty(lista) {
  return lista.map((p, i) => ({ text: p, options: { bullet: { indent: 14 }, breakLine: i < lista.length - 1 } }));
}

function karta(s, x, y, w, h, nazwa) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: C.background2 }, line: { type: "none" }, rectRadius: 0.08, objectName: nazwa });
}

function liczba(s, x, y, w, duza, opis, kolor = C.accent1) {
  s.addText(duza, { x, y, w, h: 0.6, isTextBox: true, fontSize: 30, bold: true, color: kolor, margin: 0 });
  s.addText(opis, { x, y: y + 0.62, w, h: 0.55, isTextBox: true, fontSize: 12, color: C.text2, margin: 0, valign: "top" });
}

function slupki(s, etykiety, wartosci, o) {
  s.addChart(pres.charts.BAR, [{ name: o.seria || "Wartość", labels: etykiety, values: wartosci }], {
    x: o.x, y: o.y, w: o.w, h: o.h, barDir: o.pion ? "col" : "bar", catAxisOrientation: o.pion ? "minMax" : "maxMin",
    chartColors: [o.kolor || HEX.accent1], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: o.format || "0",
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, valAxisMinVal: 0, valAxisMaxVal: o.max,
    showLegend: false, showTitle: !!o.tytul, title: o.tytul || "", barGapWidthPct: 60, catAxisLineShow: true,
    catAxisLineColor: HEX.accent4, ...OS, objectName: o.nazwa,
  });
}

function grupowe(s, etykiety, serie, o) {
  s.addChart(pres.charts.BAR, serie.map(([name, values]) => ({ name, labels: etykiety, values })), {
    x: o.x, y: o.y, w: o.w, h: o.h, barDir: o.pion ? "col" : "bar", barGrouping: "clustered",
    catAxisOrientation: o.pion ? "minMax" : "maxMin",
    chartColors: [HEX.accent1, HEX.accent2], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: o.format || "0",
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, valAxisMinVal: 0, valAxisMaxVal: o.max,
    showLegend: true, legendPos: "t", showTitle: !!o.tytul, title: o.tytul || "", barGapWidthPct: 50,
    catAxisLineShow: true, catAxisLineColor: HEX.accent4, ...OS, objectName: o.nazwa,
  });
}

function tabela(s, naglowki, wiersze, o) {
  const nag = naglowki.map((t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.accent1 }, fontSize: o.font || 12 } }));
  const ciało = wiersze.map((w) => w.map((t, i) => ({ text: t, options: { fontSize: o.font || 12, color: C.text1, align: o.prawo && o.prawo[i] ? "right" : "left" } })));
  s.addTable([nag, ...ciało], { x: o.x, y: o.y, w: o.w, colW: o.colW, border: { type: "solid", pt: 0.75, color: HEX.accent4 },
    valign: "middle", margin: [0.04, 0.08, 0.04, 0.08], rowH: o.rowH || 0.32, objectName: o.nazwa });
}

// =============================================================================
// 1. Strona tytułowa i plan
// =============================================================================
pres.addSection({ title: "Wstęp" });
let s = pres.addSlide({ masterName: "TYTUL", sectionTitle: "Wstęp" });
s.addText("Analityka – rozwiązania warsztatów", { placeholder: "title" });
s.addText("Filtrowanie danych, korelacja i DAG, ankieta „Opinia klienta”, Single Customer View\nListopad 2026", { placeholder: "body" });
s.addNotes("Prezentacja zbiera rozwiązania czterech zadań z zajęć z analityki.");

s = slajd("Co zrobiliśmy", "Wstęp");
const plan = [
  ["01", "Filtrowanie danych", "Sklep internetowy: które grupy klientów wymagają uwagi"],
  ["02", "Korelacja i DAG", "Czy reklama naprawdę zwiększa sprzedaż?"],
  ["03", "Ankieta „Opinia klienta”", "Jak korzystamy ze słuchawek i czego od nich oczekujemy"],
  ["04", "Single Customer View", "Który post sprzedaje i jak zbudować profil klienta"],
];
plan.forEach(([nr, tyt, op], i) => {
  const x = 0.5 + (i % 2) * 4.6, y = 1.25 + Math.floor(i / 2) * 1.95;
  karta(s, x, y, 4.4, 1.75, `plan_${nr}`);
  s.addShape(pres.shapes.OVAL, { x: x + 0.3, y: y + 0.3, w: 0.7, h: 0.7, fill: { color: C.accent1 }, line: { type: "none" }, objectName: `kolo_${nr}` });
  s.addText(nr, { x: x + 0.3, y: y + 0.3, w: 0.7, h: 0.7, isTextBox: true, fontSize: 18, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0 });
  s.addText(tyt, { x: x + 1.2, y: y + 0.28, w: 3.0, h: 0.45, isTextBox: true, fontSize: 17, bold: true, color: C.text1, margin: 0, valign: "middle" });
  s.addText(op, { x: x + 1.2, y: y + 0.78, w: 3.0, h: 0.8, isTextBox: true, fontSize: 13, color: C.text2, margin: 0, valign: "top" });
});

// =============================================================================
// 2. Warsztat 1 – filtrowanie danych
// =============================================================================
const S1 = "Filtrowanie danych";
sekcja("01", S1, "Sklep internetowy: 15 użytkowników, 8 zakupów, 935 € przychodu");

s = slajd("Trzy wskaźniki, które wybraliśmy", S1);
const kpi = [
  ["53%", "Współczynnik konwersji", "Mierzy cel zarządu, czyli liczbę zakupów"],
  ["117 €", "Średnia wartość koszyka", "Razem z konwersją decyduje o przychodzie"],
  ["52%", "Wskaźnik odrzuceń", "Najlepiej pokazuje, dlaczego ludzie nie kupują"],
];
kpi.forEach(([liczbaTxt, nazwa, op], i) => {
  const x = 0.5 + i * 3.07;
  karta(s, x, 1.3, 2.87, 2.6, `kpi_${i}`);
  s.addText(liczbaTxt, { x: x + 0.3, y: 1.55, w: 2.3, h: 0.8, isTextBox: true, fontSize: 36, bold: true, color: C.accent1, margin: 0 });
  s.addText(nazwa, { x: x + 0.3, y: 2.4, w: 2.3, h: 0.45, isTextBox: true, fontSize: 15, bold: true, color: C.text1, margin: 0 });
  s.addText(op, { x: x + 0.3, y: 2.85, w: 2.3, h: 0.9, isTextBox: true, fontSize: 13, color: C.text2, margin: 0, valign: "top" });
});
tekst(s, "Liczby to wynik dla całego sklepu. Czasu na stronie i liczby stron nie wybraliśmy, bo pokazują to samo co odrzucenia.", 0.5, 4.2, 9, 0.6, { fontSize: 13, color: C.text2 });

s = slajd("Źródło ruchu: Google Ads sprzedaje najlepiej", S1);
slupki(s, ["Google Ads", "Social Media", "Organic Search"], [80, 60, 20], { x: 0.4, y: 1.15, w: 5.2, h: 3.4, tytul: "Konwersja (%)", max: 100, nazwa: "wykres_zrodla" });
tekst(s, punkty([
  "Google Ads: 4 z 8 transakcji, koszyk 125 €.",
  "Organic Search: najwięcej odrzuceń (66%) i tylko 1 zakup.",
  "Uwaga: z Google Ads przyszli głównie powracający klienci, a z wyszukiwarki nowi. To oni robią różnicę, nie sam kanał.",
]), 5.9, 1.4, 3.6, 3.3);

s = slajd("Urządzenia: mobile wymaga poprawy", S1);
slupki(s, ["Desktop", "Mobile", "Tablet"], [29, 67, 57], { x: 0.4, y: 1.15, w: 5.2, h: 3.4, tytul: "Wskaźnik odrzuceń (%)", max: 100, kolor: HEX.accent2, nazwa: "wykres_urzadzenia" });
tekst(s, punkty([
  "Desktop: najwięcej zakupów (4) i najwyższy koszyk (162,50 €).",
  "Mobile: największy ruch (7 z 15 osób), ale 67% odrzuceń.",
  "Najgorzej wypadają nowi użytkownicy na telefonach.",
]), 5.9, 1.4, 3.6, 3.3);

s = slajd("Kraje: najwyższa konwersja w Niemczech", S1);
tabela(s, ["Kraj", "Konwersja", "Średni koszyk", "Odrzucenia"], [
  ["Niemcy", "80%", "87,50 €", "50,0%"],
  ["Polska", "50%", "128,33 €", "50,8%"],
  ["Wielka Brytania", "25%", "200,00 €", "57,5%"],
], { x: 0.5, y: 1.35, w: 5.4, colW: [1.8, 1.1, 1.4, 1.1], prawo: [false, true, true, true], font: 13, rowH: 0.42, nazwa: "tabela_kraje" });
tekst(s, punkty([
  "Najniższe odrzucenia mają Niemcy i Polska (ok. 50%).",
  "W każdym kraju powracający kupują zawsze, a nowi prawie wcale.",
  "W Wielkiej Brytanii najwyższy koszyk, ale żaden nowy klient nic nie kupił.",
]), 6.2, 1.4, 3.3, 3.3);

s = slajd("Nowi a powracający: największa różnica", S1);
grupowe(s, ["Konwersja (%)", "Odrzucenia (%)"], [["Nowi", [12.5, 72]], ["Powracający", [100, 30]]], { x: 0.4, y: 1.15, w: 5.2, h: 3.5, max: 115, pion: true, format: "General", nazwa: "wykres_nowi" });
tekst(s, [{ text: "Jak przekonać nowych:", options: { bold: true, breakLine: true } }, ...punkty([
  "rabat lub darmowa dostawa na pierwsze zakupy,",
  "przypomnienie mailem o koszyku,",
  "reklamy do osób, które już były na stronie,",
  "prostszy zakup na telefonie.",
])], 5.9, 1.4, 3.6, 3.3);

s = slajd("Wnioski z filtrowania", S1);
karta(s, 0.5, 1.25, 9, 2.9, "wnioski_1");
tekst(s, punkty([
  "Sklep dobrze sprzedaje klientom, którzy już go znają: powracający generują 97% przychodu.",
  "Nowi użytkownicy prawie nie kupują, szczególnie na telefonach i z wyszukiwarki.",
  "Warto poprawić wersję mobilną strony i strony wejścia z Google.",
  "Nowych trzeba zachęcić do pierwszego zakupu i do powrotu: oferta powitalna, maile, remarketing.",
]), 0.85, 1.55, 8.3, 3.0, { fontSize: 16, paraSpaceAfter: 12 });

// =============================================================================
// 3. Warsztat 2 – korelacja i DAG
// =============================================================================
const S2 = "Korelacja i DAG";
sekcja("02", S2, "Czy reklama naprawdę zwiększa sprzedaż? Dyrektor chce dać 30% więcej na reklamę");

s = slajd("Reklama i sprzedaż rosną razem", S2);
const reklama = [20, 22, 25, 27, 30, 33, 35, 34, 38, 42, 65, 70];
const sprzedaz = [210, 220, 245, 250, 275, 290, 310, 305, 340, 365, 520, 570];
s.addChart(pres.charts.SCATTER, [{ name: "Reklama", values: reklama }, { name: "Sprzedaż", values: sprzedaz }], {
  x: 0.4, y: 1.1, w: 5.4, h: 3.6, lineSize: 0, lineDataSymbol: "circle", lineDataSymbolSize: 8, chartColors: [HEX.accent1],
  showTitle: true, title: "Reklama a sprzedaż (tys. zł)", showLegend: false, valAxisMinVal: 150, catAxisMinVal: 15, catAxisMaxVal: 75,
  valGridLine: { color: "E3E8EA", size: 0.5 }, catGridLine: { style: "none" }, showCatAxisTitle: true, catAxisTitle: "Reklama",
  showValAxisTitle: true, valAxisTitle: "Sprzedaż", catAxisTitleFontSize: 11, valAxisTitleFontSize: 11, catAxisTitleColor: HEX.dk2, valAxisTitleColor: HEX.dk2,
  ...OS, objectName: "wykres_rozrzutu",
});
tabela(s, ["Korelacja", "Cały rok", "Bez XI i XII"], [
  ["Reklama – sprzedaż", "0,999", "0,996"],
  ["Popyt – sprzedaż", "0,999", "0,996"],
  ["Popyt – reklama", "0,998", "0,998"],
], { x: 6.0, y: 1.3, w: 3.5, colW: [1.7, 0.9, 0.9], prawo: [false, true, true], font: 12, rowH: 0.36, nazwa: "tabela_korelacje" });
tekst(s, "Popyt rośnie razem z reklamą, więc z samej korelacji nie wiemy, co naprawdę napędza sprzedaż.", 6.0, 3.0, 3.5, 1.4, { fontSize: 13 });

s = slajd("Pytania do dyskusji", S2);
const pyt = [
  ["Czy silna korelacja oznacza, że reklama zwiększa sprzedaż?", "Nie. Firma więcej się reklamuje wtedy, gdy i tak jest większe zainteresowanie, np. przed świętami."],
  ["Czy wyłączenie listopada i grudnia zmienia wynik?", "Nie bardzo. Korelacja spada tylko z 0,999 do 0,996, czyli problem jest przez cały rok."],
  ["Czy takie wyłączenie jest uzasadnione?", "Nie do końca. Tracimy 2 z 12 miesięcy. Lepiej zostawić wszystkie dane i uwzględnić sezon w modelu."],
];
pyt.forEach(([q, a], i) => {
  const y = 1.2 + i * 1.25;
  karta(s, 0.5, y, 9, 1.1, `pytanie_${i}`);
  s.addText(q, { x: 0.8, y: y + 0.12, w: 8.4, h: 0.4, isTextBox: true, fontSize: 15, bold: true, color: C.text1, margin: 0, valign: "middle" });
  s.addText(a, { x: 0.8, y: y + 0.52, w: 8.4, h: 0.5, isTextBox: true, fontSize: 13, color: C.text2, margin: 0, valign: "top" });
});

s = slajd("Model przyczynowy (DAG)", S2);
s.addImage({ path: path.join(REPO, "warsztat_2", "diagram_DAG.png"), x: 0.4, y: 1.05, w: 5.0, h: 5.0 * 1033 / 1280, altText: "Diagram DAG: sezonowość, popyt, rabaty i konkurencja wpływają na wydatki reklamowe i sprzedaż; reklama działa przez ruch na stronie", objectName: "diagram_dag" });
tekst(s, [
  { text: "Zmienne zakłócające", options: { bold: true, breakLine: true } },
  { text: "sezonowość, popyt rynkowy, rabaty, działania konkurencji", options: { breakLine: true } },
  { text: " ", options: { breakLine: true, fontSize: 6 } },
  { text: "Mediator", options: { bold: true, breakLine: true } },
  { text: "ruch na stronie", options: { breakLine: true } },
  { text: " ", options: { breakLine: true, fontSize: 6 } },
  { text: "Model regresji", options: { bold: true, breakLine: true } },
  { text: "sprzedaż ~ reklama + sezonowość + popyt + rabaty + konkurencja. Bez ruchu na stronie, bo to przez niego działa reklama." },
], 5.7, 1.35, 3.8, 3.6);

s = slajd("Czy dać 30% więcej na reklamę?", S2);
s.addText("Sama korelacja tego nie uzasadnia", { x: 0.5, y: 1.15, w: 9, h: 0.5, isTextBox: true, fontSize: 20, bold: true, color: C.accent5, margin: 0 });
karta(s, 0.5, 1.85, 4.4, 2.3, "endogenicznosc");
karta(s, 5.1, 1.85, 4.4, 2.3, "jak_sprawdzic");
tekst(s, [{ text: "Skąd bierze się problem", options: { bold: true, breakLine: true } }, ...punkty([
  "popyt i sezon wpływają na reklamę i sprzedaż naraz,",
  "budżet ustala się pod spodziewaną sprzedaż,",
  "rabaty i konkurencja zmieniają się razem z reklamą.",
])], 0.8, 2.1, 3.9, 2.5, { fontSize: 13 });
tekst(s, [{ text: "Jak to sprawdzić", options: { bold: true, breakLine: true } }, ...punkty([
  "eksperyment: więcej reklamy tylko w części regionów i porównanie sprzedaży,",
  "regresja z popytem, sezonem, rabatami i konkurencją na dłuższych danych.",
])], 5.4, 2.1, 3.9, 2.5, { fontSize: 13 });

// =============================================================================
// 4. Ankieta „Opinia klienta”
// =============================================================================
const S3 = "Ankieta „Opinia klienta”";
sekcja("03", S3, "Słuchawki bezprzewodowe: 15 odpowiedzi, 30 października – 6 listopada 2026");

s = slajd("Kto odpowiadał", S3);
slupki(s, ["Poniżej 18", "18-24", "25-34", "45 i więcej"], [1, 8, 5, 1], { x: 0.4, y: 1.15, w: 4.5, h: 2.9, tytul: "Wiek (liczba osób)", max: 10, nazwa: "wykres_wiek" });
slupki(s, ["Praca na etacie", "Nauka lub studia", "Freelance, firma"], [7, 7, 1], { x: 5.1, y: 1.15, w: 4.5, h: 2.9, tytul: "Zajęcie (liczba osób)", max: 9, nazwa: "wykres_zajecie" });
tekst(s, "Głównie młode osoby: ponad połowa ma 18-24 lata. Pracujących i studentów jest po równo.", 0.5, 4.3, 9, 0.5, { fontSize: 14, color: C.text2 });

s = slajd("Jak korzystają ze słuchawek", S3);
slupki(s, ["Podczas grania", "W drodze", "W pracy lub nauce", "Na treningu", "W domu dla relaksu"], [8, 6, 5, 5, 5], { x: 0.4, y: 1.15, w: 5.4, h: 3.5, tytul: "Sytuacje (liczba wskazań)", max: 10, nazwa: "wykres_sytuacje" });
karta(s, 6.1, 1.3, 3.4, 3.2, "czas_dziennie");
s.addText("Czas dziennie", { x: 6.4, y: 1.45, w: 2.9, h: 0.4, isTextBox: true, fontSize: 14, bold: true, color: C.text1, margin: 0 });
liczba(s, 6.4, 1.9, 2.9, "6 osób", "do godziny dziennie");
liczba(s, 6.4, 3.1, 2.9, "6 osób", "ponad 3 godziny dziennie", C.accent2);

s = slajd("Wygoda jest najważniejsza", S3);
slupki(s, ["Wygoda noszenia", "Jakość dźwięku i bas", "Wyciszenie (ANC)", "Mocna bateria", "Cena i promocje", "Wygląd i design"], [13, 11, 8, 7, 4, 1], { x: 0.4, y: 1.15, w: 5.6, h: 3.6, tytul: "Najważniejsze cechy (liczba wskazań)", max: 15, nazwa: "wykres_cechy" });
liczba(s, 6.3, 1.5, 3.2, "13 z 15", "osób wybrało wygodę noszenia");
liczba(s, 6.3, 2.75, 3.2, "1 osoba", "wskazała wygląd i design", C.accent2);
tekst(s, "Liczy się to, co czuć na co dzień, a nie wygląd.", 6.3, 4.0, 3.2, 0.6, { fontSize: 13, color: C.text2 });

s = slajd("Ile wydadzą i gdzie kupują", S3);
slupki(s, ["Do 200 zł", "201-400 zł", "401-700 zł", "Powyżej 700 zł"], [3, 4, 5, 3], { x: 0.4, y: 1.15, w: 4.5, h: 3.0, tytul: "Maksymalny budżet (liczba osób)", max: 7, nazwa: "wykres_budzet" });
slupki(s, ["Sklepy internetowe", "Sklep stacjonarny", "Strona producenta", "Allegro, Amazon"], [5, 4, 4, 2], { x: 5.1, y: 1.15, w: 4.5, h: 3.0, tytul: "Miejsce zakupu (liczba osób)", max: 7, kolor: HEX.accent2, nazwa: "wykres_gdzie" });
tekst(s, "8 z 15 osób wyda ponad 400 zł, a 11 z 15 kupuje przez internet.", 0.5, 4.4, 9, 0.5, { fontSize: 14, color: C.text2 });

s = slajd("Styl i zaufanie do marki", S3);
grupowe(s, ["1", "2", "3", "4", "5"], [["Słuchawki to element mojego stylu", [2, 2, 4, 4, 3]], ["Ufam sprawdzonym markom, np. JBL", [0, 4, 4, 1, 6]]], { x: 0.4, y: 1.1, w: 5.8, h: 3.7, max: 7, pion: true, nazwa: "wykres_skale" });
liczba(s, 6.5, 1.4, 3.0, "3,3", "średnio: styl (skala 1-5)");
liczba(s, 6.5, 2.6, 3.0, "3,6", "średnio: zaufanie do marki", C.accent2);
tekst(s, "Zaufanie do marki jest podzielone: 6 osób dało 5, ale 4 osoby tylko 2.", 6.5, 3.8, 3.0, 0.9, { fontSize: 13, color: C.text2 });

s = slajd("Studenci a pracujący", S3);
tabela(s, ["", "Studenci (7)", "Pracujący (8)"], [
  ["Najczęstsze użycie", "granie (5)", "praca lub nauka, dom (po 4)"],
  ["Najważniejsze cechy", "wygoda i dźwięk (po 6)", "wygoda (7), dźwięk i bateria (po 5)"],
  ["Najczęstszy budżet", "powyżej 700 zł (3)", "401-700 zł (4)"],
  ["Zaufanie do marki", "3,1", "4,0"],
], { x: 0.5, y: 1.3, w: 9, colW: [2.4, 3.0, 3.6], font: 13, rowH: 0.45, nazwa: "tabela_grupy" });
tekst(s, "Pracujący bardziej ufają marce. Studenci częściej grają, patrzą na dźwięk i cenę, ale to wśród nich są osoby gotowe wydać najwięcej.", 0.5, 3.85, 9, 0.8, { fontSize: 14, color: C.text2 });

s = slajd("Co z tego wynika dla firmy", S3);
karta(s, 0.5, 1.25, 9, 2.9, "wnioski_ankieta");
tekst(s, punkty([
  "W reklamach podkreślać wygodę i jakość dźwięku.",
  "Oferować wyciszenie otoczenia (ANC) w tańszych i droższych modelach.",
  "Sprzedawać głównie przez internet, ale nie rezygnować ze sklepów stacjonarnych.",
  "Przygotować ofertę dla studentów i graczy.",
  "Dodać możliwość wyłączenia komunikatów głosowych o baterii.",
]), 0.85, 1.55, 8.3, 3.0, { fontSize: 16, paraSpaceAfter: 10 });

// =============================================================================
// 5. Single Customer View
// =============================================================================
const S4 = "Single Customer View";
sekcja("04", S4, "Marka odzieżowa: cztery posty i wspólny profil klienta");

s = slajd("Zaangażowanie nie oznacza sprzedaży", S4);
const posty = ["104 Stories", "108 Reels", "109 Post promocyjny", "110 Post z obrazem"];
slupki(s, posty, [6.4, 22.9, 8.7, 16.6], { x: 0.4, y: 1.15, w: 4.6, h: 3.3, tytul: "Zaangażowanie (%)", max: 28, format: "0.0", nazwa: "wykres_zaang" });
slupki(s, posty, [0.5, 1.9, 2.2, 3.2], { x: 5.0, y: 1.15, w: 4.6, h: 3.3, tytul: "Konwersja po kliknięciu (%)", max: 4, format: "0.0", kolor: HEX.accent2, nazwa: "wykres_konw" });
tekst(s, "Reels ma najwięcej reakcji, ale najlepiej sprzedaje post z obrazem.", 0.5, 4.55, 9, 0.45, { fontSize: 14, color: C.text2 });

s = slajd("Który post wybieramy do testu", S4);
const posty3 = [
  ["108 Reels", "22,9%", "najwyższe zaangażowanie, ale z 470 kliknięć tylko 9 zakupów"],
  ["110 Post z obrazem", "3,2%", "najwyższa konwersja: 16 zakupów z 500 kliknięć"],
  ["104 Stories", "0,5%", "600 kliknięć i tylko 3 zakupy"],
];
posty3.forEach(([nazwa, wartosc, op], i) => {
  const x = 0.5 + i * 3.07;
  karta(s, x, 1.2, 2.87, 2.0, `post_${i}`);
  s.addText(nazwa, { x: x + 0.25, y: 1.35, w: 2.4, h: 0.4, isTextBox: true, fontSize: 14, bold: true, color: C.text1, margin: 0 });
  s.addText(wartosc, { x: x + 0.25, y: 1.75, w: 2.4, h: 0.6, isTextBox: true, fontSize: 28, bold: true, color: i === 2 ? C.accent5 : C.accent1, margin: 0 });
  s.addText(op, { x: x + 0.25, y: 2.4, w: 2.4, h: 0.75, isTextBox: true, fontSize: 12, color: C.text2, margin: 0, valign: "top" });
});
tekst(s, [
  { text: "Do testu wybieramy post 110 „Inspiracje na zimę”. ", options: { bold: true } },
  { text: "Ma najwyższą konwersję (3,2%) i drugie najwyższe zaangażowanie (16,6%), więc i przyciąga uwagę, i sprzedaje. Post 109 dał najwięcej zakupów (20), ale przy największym zasięgu i rabatach." },
], 0.5, 3.5, 9, 1.3, { fontSize: 14 });

s = slajd("Jak zbudować wspólny profil klienta", S4);
karta(s, 0.5, 1.2, 4.4, 1.3, "zrodlo_sklep");
karta(s, 5.1, 1.2, 4.4, 1.3, "zrodlo_crm");
tekst(s, [{ text: "Sklep internetowy", options: { bold: true, breakLine: true } }, { text: "konto, historia zamówień, rozmiary, porzucone koszyki, wejścia z linków UTM" }], 0.8, 1.4, 3.9, 1.4, { fontSize: 13 });
tekst(s, [{ text: "Newsletter i program lojalnościowy", options: { bold: true, breakLine: true } }, { text: "zapisy, otwarcia i kliknięcia w maile, zgody, punkty" }], 5.4, 1.4, 3.9, 1.4, { fontSize: 13 });
s.addText([{ text: "Wspólny identyfikator: ", options: { bold: true } }, { text: "adres e-mail klienta" }], { x: 0.5, y: 2.85, w: 9, h: 0.45, isTextBox: true, fontSize: 16, color: C.accent1, margin: 0 });
tekst(s, [{ text: "Przykład: ", options: { bold: true } }, { text: "klient wszedł z posta 110, dodał płaszcz do koszyka i nie kupił. Jest zapisany do newslettera, więc następnego dnia dostaje maila z przypomnieniem i podobnymi stylizacjami w swoim rozmiarze." }], 0.5, 3.45, 9, 1.1, { fontSize: 14 });

// =============================================================================
// Zakończenie
// =============================================================================
pres.addSection({ title: "Zakończenie" });
s = pres.addSlide({ masterName: "TYTUL", sectionTitle: "Zakończenie" });
s.addText("Dziękujemy", { placeholder: "title" });
s.addText("Pytania?", { placeholder: "body" });

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("Zapisano", OUT);
})();
