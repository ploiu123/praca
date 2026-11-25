// Prezentacja zaliczeniowa: Podstawy analityki internetowej (4 projekty).
// Uruchomienie: NODE_PATH=<katalog z pptxgenjs> APPLY_THEME=<apply_theme.js> node build_zaliczenie.js
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.APPLY_THEME);

const KAT = __dirname;
const REPO = path.resolve(KAT, "..");
const OUT = path.join(KAT, "Podstawy_analityki_projekty.pptx");

const THEME = {
  name: "Podstawy analityki",
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
pres.layout = "LAYOUT_16x9";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Podstawy analityki internetowej – projekty";
const C = pres.SchemeColor;

// --- układy -----------------------------------------------------------------------
pres.defineSlideMaster({
  title: "TYTUL",
  background: { color: C.text1 },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: 0.6, y: 1.15, w: 8.8, h: 0.4, fontSize: 14, bold: true, color: C.accent3, align: "left", margin: 0, charSpacing: 1 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.6, w: 8.8, h: 1.3, fontSize: 38, bold: true, color: C.background1, valign: "top", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 3.05, w: 8.8, h: 1.4, fontSize: 16, color: C.accent4, valign: "top", align: "left", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "SEKCJA",
  background: { color: C.text1 },
  objects: [
    { placeholder: { options: { name: "nr", type: "body", x: 0.6, y: 1.0, w: 4, h: 0.9, fontSize: 48, bold: true, color: C.accent3, align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.95, w: 8.8, h: 1.4, fontSize: 32, bold: true, color: C.background1, valign: "top", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 3.5, w: 8.8, h: 1.0, fontSize: 16, color: C.accent4, valign: "top", align: "left", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "TRESC",
  background: { color: C.background1 },
  slideNumber: { x: 9.15, y: 5.22, w: 0.5, h: 0.25, fontSize: 10, color: C.accent6, align: "right" },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: 0.5, y: 0.25, w: 9, h: 0.3, fontSize: 11, bold: true, color: C.accent1, align: "left", margin: 0, charSpacing: 1 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.52, w: 9, h: 0.62, fontSize: 24, bold: true, color: C.text1, valign: "middle", align: "left", margin: 0 }, text: "" } },
  ],
});

// --- pomocnicze -------------------------------------------------------------------
const OS = { catAxisLabelColor: HEX.dk2, valAxisLabelColor: HEX.dk2, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
  catAxisLabelFontSize: 11, valAxisLabelFontSize: 10, dataLabelFontFace: "+mn-lt", dataLabelFontSize: 11, dataLabelColor: HEX.dk1,
  titleFontFace: "+mn-lt", titleFontSize: 12, titleColor: HEX.dk1, legendFontFace: "+mn-lt", legendFontSize: 11, legendColor: HEX.dk2 };

let sekcjaAkt = "Wstęp";
let kickerAkt = "";

function slajd(tytul, notatki) {
  const s = pres.addSlide({ masterName: "TRESC", sectionTitle: sekcjaAkt });
  if (kickerAkt) s.addText(kickerAkt, { placeholder: "kicker" });
  s.addText(tytul, { placeholder: "title" });
  if (notatki) s.addNotes(notatki);
  return s;
}

function sekcja(nr, tytul, opis, kicker, notatki) {
  sekcjaAkt = tytul;
  kickerAkt = kicker;
  pres.addSection({ title: tytul });
  const s = pres.addSlide({ masterName: "SEKCJA", sectionTitle: tytul });
  s.addText(nr, { placeholder: "nr" });
  s.addText(tytul, { placeholder: "title" });
  s.addText(opis, { placeholder: "body" });
  if (notatki) s.addNotes(notatki);
  return s;
}

function tekst(s, kawalki, x, y, w, h, o = {}) {
  s.addText(kawalki, { x, y, w, h, isTextBox: true, fontSize: 14, color: C.text1, valign: "top", margin: 0, paraSpaceAfter: 6, ...o });
}
const B = (t, br = false) => ({ text: t, options: { bold: true, breakLine: br } });
const T = (t, br = false) => ({ text: t, options: { breakLine: br } });
function punkty(lista) {
  return lista.map((p, i) => ({ text: p, options: { bullet: { indent: 14 }, breakLine: i < lista.length - 1 } }));
}
function karta(s, x, y, w, h, nazwa, kolor) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: kolor || C.background2 }, line: { type: "none" }, rectRadius: 0.08, objectName: nazwa });
}
function liczba(s, x, y, w, duza, opis, kolor = C.accent1, rozmiar = 28) {
  s.addText(duza, { x, y, w, h: 0.55, isTextBox: true, fontSize: rozmiar, bold: true, color: kolor, margin: 0, valign: "bottom" });
  s.addText(opis, { x, y: y + 0.58, w, h: 0.5, isTextBox: true, fontSize: 12, color: C.text2, margin: 0, valign: "top" });
}
function kolko(s, x, y, d, napis, kolor = C.accent1, rozmiar = 14) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: kolor }, line: { type: "none" }, objectName: `kolko_${napis}_${x}_${y}` });
  s.addText(napis, { x, y, w: d, h: d, isTextBox: true, fontSize: rozmiar, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0 });
}
function strzalka(s, x1, y1, x2, y2) {
  s.addShape(pres.shapes.LINE, { x: x1, y: y1, w: x2 - x1, h: y2 - y1, line: { color: HEX.dk2, width: 1.5, endArrowType: "triangle" } });
}

function slupki(s, etykiety, wartosci, o) {
  s.addChart(pres.charts.BAR, [{ name: o.seria || "Wartość", labels: etykiety, values: wartosci }], {
    x: o.x, y: o.y, w: o.w, h: o.h, barDir: o.pion ? "col" : "bar", catAxisOrientation: o.pion ? "minMax" : "maxMin",
    chartColors: [o.kolor || HEX.accent1], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: o.format || "General",
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, valAxisMinVal: 0, valAxisMaxVal: o.max,
    showLegend: false, showTitle: !!o.tytul, title: o.tytul || "", barGapWidthPct: 55, catAxisLineShow: true,
    catAxisLineColor: HEX.accent4, ...OS, objectName: o.nazwa,
  });
}
function grupowe(s, etykiety, serie, o) {
  s.addChart(pres.charts.BAR, serie.map(([name, values]) => ({ name, labels: etykiety, values })), {
    x: o.x, y: o.y, w: o.w, h: o.h, barDir: o.pion ? "col" : "bar", barGrouping: "clustered",
    catAxisOrientation: o.pion ? "minMax" : "maxMin",
    chartColors: [HEX.accent1, HEX.accent2], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: o.format || "General",
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, valAxisMinVal: 0, valAxisMaxVal: o.max,
    showLegend: true, legendPos: "t", showTitle: !!o.tytul, title: o.tytul || "", barGapWidthPct: 45,
    catAxisLineShow: true, catAxisLineColor: HEX.accent4, ...OS, objectName: o.nazwa,
  });
}
function tabela(s, naglowki, wiersze, o) {
  const f = o.font || 12;
  const nag = naglowki.map((t, i) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.accent1 }, fontSize: f, align: o.prawo && o.prawo[i] ? "right" : "left" } }));
  const cialo = wiersze.map((w, r) => w.map((t, i) => ({ text: t, options: {
    fontSize: f, color: C.text1, bold: o.pierwszaPogrubiona && i === 0, align: o.prawo && o.prawo[i] ? "right" : "left",
    fill: o.wyroznij && o.wyroznij.includes(r) ? { color: C.background2 } : undefined,
  } })));
  s.addTable([nag, ...cialo], { x: o.x, y: o.y, w: o.w, colW: o.colW, border: { type: "solid", pt: 0.75, color: HEX.accent4 },
    valign: "middle", margin: [0.03, 0.07, 0.03, 0.07], rowH: o.rowH || 0.3, objectName: o.nazwa });
}

// =============================================================================
// Wstęp
// =============================================================================
pres.addSection({ title: "Wstęp" });
let s = pres.addSlide({ masterName: "TYTUL", sectionTitle: "Wstęp" });
s.addText("PROJEKTY ZALICZENIOWE", { placeholder: "kicker" });
s.addText("Podstawy analityki internetowej", { placeholder: "title" });
s.addText("Filtrowanie danych, korelacja i model przyczynowy, segmentacja rynku, analiza SCV\nListopad 2026", { placeholder: "body" });
s.addNotes("Prezentujemy cztery projekty warsztatowe z analizy danych z różnych źródeł internetowych, zgodnie z zasadami zaliczenia przedmiotu.");

s = slajd("Cztery projekty", "Każdy projekt to osobna analiza na innych danych: dane ze sklepu internetowego, dane miesięczne o reklamie i sprzedaży, nasza ankieta o słuchawkach oraz wyniki postów w social mediach.");
const plan = [
  ["1", "Filtrowanie danych", "Które grupy klientów sklepu internetowego wymagają uwagi?"],
  ["2", "Korelacja i model przyczynowy", "Czy reklama naprawdę zwiększa sprzedaż?"],
  ["3", "Segmentacja rynku", "Ankieta o słuchawkach: jakie segmenty klientów widać w danych?"],
  ["4", "Analiza SCV", "Który post sprzedaje i jak zbudować wspólny profil klienta?"],
];
plan.forEach(([nr, tyt, op], i) => {
  const x = 0.5 + (i % 2) * 4.6, y = 1.35 + Math.floor(i / 2) * 1.85;
  karta(s, x, y, 4.4, 1.65, `plan_${nr}`);
  kolko(s, x + 0.3, y + 0.3, 0.65, nr, C.accent1, 20);
  s.addText(tyt, { x: x + 1.15, y: y + 0.25, w: 3.05, h: 0.45, isTextBox: true, fontSize: 16, bold: true, color: C.text1, margin: 0, valign: "middle" });
  s.addText(op, { x: x + 1.15, y: y + 0.72, w: 3.05, h: 0.8, isTextBox: true, fontSize: 13, color: C.text2, margin: 0, valign: "top" });
});

// =============================================================================
// PROJEKT 1 – filtrowanie danych
// =============================================================================
sekcja("01", "Filtrowanie danych statystycznych z różnych źródeł ruchu", "Sklep internetowy: 15 użytkowników z Google Ads, Social Media i wyszukiwarki",
  "PROJEKT 1 · FILTROWANIE DANYCH", "Pierwszy projekt to analiza danych ze sklepu internetowego. Filtrowaliśmy dane według źródła ruchu, urządzenia, kraju i typu użytkownika.");

s = slajd("Zadanie i dane", "Zarząd sklepu chce zwiększyć liczbę zakupów i poprawić efektywność marketingu. Mieliśmy dane o 15 użytkownikach: skąd przyszli, z jakiego urządzenia, z jakiego kraju, czy są nowi, jak się zachowywali na stronie i czy kupili.");
tekst(s, [B("Sytuacja", true), T("Sklep internetowy inwestuje w Google Ads i social media. Zarząd chce więcej zakupów i lepszej efektywności marketingu.", true),
  T(" ", true), B("Nasze zadanie", true), T("Wskazać grupy klientów, które wymagają uwagi, i działania warte przetestowania.")], 0.5, 1.4, 4.2, 3.5);
[["15", "użytkowników"], ["8", "zakupów"], ["53%", "konwersja"], ["935 €", "przychodu"]].forEach(([l, o], i) => {
  const x = 5.0 + (i % 2) * 2.3, y = 1.4 + Math.floor(i / 2) * 1.55;
  karta(s, x, y, 2.1, 1.35, `fakt_${i}`);
  liczba(s, x + 0.25, y + 0.15, 1.7, l, o);
});

s = slajd("Pytania biznesowe i wskaźniki", "Zanim zaczęliśmy liczyć, ustaliliśmy pytania biznesowe i wskaźniki. Wskaźnik odrzuceń traktujemy jako metrykę pomocniczą, bo sam nie jest celem, ale wyjaśnia, dlaczego konwersja jest niska.");
karta(s, 0.5, 1.35, 4.4, 2.95, "pytania_biznesowe");
tekst(s, [B("Pytania przed analizą", true), ...punkty([
  "Które źródło ruchu przynosi najwięcej zakupów?",
  "Na jakim urządzeniu klienci kupują najczęściej?",
  "Z których krajów kupuje się najwięcej?",
  "Czy nowi klienci kupują tak samo jak powracający?",
  "Ile kosztuje pozyskanie jednego klienta?",
])], 0.75, 1.55, 3.95, 3.3, { fontSize: 13 });
tabela(s, ["Wskaźnik", "Dlaczego"], [
  ["Współczynnik konwersji", "mierzy cel zarządu: liczbę zakupów"],
  ["Średnia wartość koszyka", "razem z konwersją daje przychód"],
  ["Wskaźnik odrzuceń", "pokazuje, gdzie tracimy klientów"],
  ["Koszt pozyskania (CAC)", "ocena efektywności, brak danych o kosztach"],
], { x: 5.1, y: 1.35, w: 4.4, colW: [1.75, 2.65], font: 12, rowH: 0.55, pierwszaPogrubiona: true, nazwa: "tabela_kpi" });

s = slajd("Źródło ruchu: Google Ads sprzedaje najlepiej", "Google Ads ma najwyższą konwersję i najwyższy koszyk. Organic Search wypada najsłabiej: 66% odrzuceń i tylko jeden zakup.");
tabela(s, ["Źródło", "Użytk.", "Zakupy", "Konwersja", "Koszyk", "Odrzucenia"], [
  ["Google Ads", "5", "4", "80%", "125 €", "39%"],
  ["Social Media", "5", "3", "60%", "120 €", "52%"],
  ["Organic Search", "5", "1", "20%", "75 €", "66%"],
], { x: 0.5, y: 1.4, w: 5.1, colW: [1.3, 0.6, 0.65, 0.9, 0.7, 0.95], prawo: [false, true, true, true, true, true], font: 11, rowH: 0.38, pierwszaPogrubiona: true, nazwa: "tabela_zrodla" });
slupki(s, ["Google Ads", "Social Media", "Organic Search"], [80, 60, 20], { x: 5.7, y: 1.25, w: 3.9, h: 2.4, tytul: "Konwersja (%)", max: 100, nazwa: "wykres_zrodla" });
tekst(s, [B("Co widać: "), T("Google Ads daje 4 z 8 transakcji i 53% przychodu. Użytkownicy z Google Ads są na stronie prawie 2 razy dłużej (3,5 min) niż z wyszukiwarki (2,0 min).")], 0.5, 3.3, 5.1, 1.6, { fontSize: 13 });

s = slajd("Dlaczego Google Ads wygrywa z wyszukiwarką", "To ważna pułapka interpretacji: średnia może ukrywać segmenty. Różnica między kanałami wynika głównie z tego, że Google Ads przyprowadził powracających klientów, a wyszukiwarka nowych. W tych samych grupach kanały wypadają tak samo.");
tabela(s, ["", "Nowi: osoby", "Nowi: zakupy", "Powracający: osoby", "Powracający: zakupy"], [
  ["Google Ads", "1", "0", "4", "4"],
  ["Social Media", "3", "1", "2", "2"],
  ["Organic Search", "4", "0", "1", "1"],
], { x: 0.5, y: 1.4, w: 5.6, colW: [1.4, 1.0, 1.0, 1.1, 1.1], prawo: [false, true, true, true, true], font: 12, rowH: 0.4, pierwszaPogrubiona: true, nazwa: "tabela_mix" });
karta(s, 6.4, 1.4, 3.1, 3.4, "pulapka");
tekst(s, [B("Średnia może ukrywać segmenty", true), T("Z Google Ads przyszli głównie powracający klienci (4 z 5), a z wyszukiwarki nowi (4 z 5).", true), T(" ", true),
  T("Powracający kupili wszyscy, a nowi prawie nikt. Różnicę robi typ klienta, a nie sam kanał.")], 6.65, 1.6, 2.65, 3.1, { fontSize: 13 });
tekst(s, "Wniosek: przed oceną kanału trzeba sprawdzić, kogo on przyprowadza.", 0.5, 3.35, 5.6, 0.8, { fontSize: 14, bold: true, color: C.accent1 });

s = slajd("Urządzenia: mobile wymaga poprawy", "Mobile ma największy ruch, ale najwięcej odrzuceń i najniższy koszyk. Najgorzej wypadają nowi użytkownicy na telefonach: 79% odrzuceń i 1 zakup na 5 osób.");
tabela(s, ["Urządzenie", "Użytk.", "Zakupy", "Konwersja", "Koszyk", "Odrzucenia"], [
  ["Desktop", "5", "4", "80%", "162,50 €", "29%"],
  ["Mobile", "7", "3", "43%", "66,67 €", "67%"],
  ["Tablet", "3", "1", "33%", "85,00 €", "57%"],
], { x: 0.5, y: 1.4, w: 5.1, colW: [1.1, 0.6, 0.65, 0.9, 0.9, 0.95], prawo: [false, true, true, true, true, true], font: 11, rowH: 0.38, pierwszaPogrubiona: true, nazwa: "tabela_urzadzenia" });
slupki(s, ["Desktop", "Mobile", "Tablet"], [29, 67, 57], { x: 5.7, y: 1.25, w: 3.9, h: 2.4, tytul: "Wskaźnik odrzuceń (%)", max: 100, kolor: HEX.accent2, nazwa: "wykres_urzadzenia" });
tekst(s, [B("Co widać: "), T("desktop daje połowę zakupów i 70% przychodu. Mobile to prawie połowa ruchu, ale 67% osób od razu wychodzi ze strony.")], 0.5, 3.3, 5.1, 1.6, { fontSize: 13 });

s = slajd("Kraje: wszędzie ten sam problem z nowymi", "W każdym kraju powracający kupują zawsze, a nowi prawie nigdy. Wielka Brytania ma najwyższy koszyk, ale żaden nowy klient tam nie kupił, więc warto sprawdzić dopasowanie oferty: waluta, dostawa, cła.");
tabela(s, ["Kraj", "Konwersja", "Koszyk", "Odrzucenia"], [
  ["Niemcy", "80%", "87,50 €", "50,0%"],
  ["Polska", "50%", "128,33 €", "50,8%"],
  ["Wielka Brytania", "25%", "200,00 €", "57,5%"],
], { x: 0.5, y: 1.4, w: 4.4, colW: [1.45, 0.95, 0.95, 1.05], prawo: [false, true, true, true], font: 12, rowH: 0.38, pierwszaPogrubiona: true, nazwa: "tabela_kraje" });
tabela(s, ["Kraj", "Nowi kupili", "Powracający kupili"], [
  ["Niemcy", "1 z 2", "3 z 3"],
  ["Polska", "0 z 3", "3 z 3"],
  ["Wielka Brytania", "0 z 3", "1 z 1"],
], { x: 5.1, y: 1.4, w: 4.4, colW: [1.6, 1.2, 1.6], prawo: [false, true, true], font: 12, rowH: 0.38, pierwszaPogrubiona: true, nazwa: "tabela_kraje_nowi" });
tekst(s, punkty([
  "Najwyższa konwersja w Niemczech (80%), najwyższy koszyk w Wielkiej Brytanii (200 €).",
  "Najniższe odrzucenia mają Niemcy i Polska, ok. 50%.",
  "W Wielkiej Brytanii nowi klienci nic nie kupili, więc warto sprawdzić ofertę na ten rynek.",
]), 0.5, 3.15, 9, 1.7, { fontSize: 13 });

s = slajd("Nowi a powracający: tu jest największa różnica", "To najważniejszy wynik projektu. Powracający generują prawie cały przychód. Dlatego działania warto skierować na pierwszy zakup nowych klientów i na to, żeby wrócili.");
grupowe(s, ["Konwersja (%)", "Odrzucenia (%)"], [["Nowi (8 osób)", [12.5, 72]], ["Powracający (7 osób)", [100, 30]]], { x: 0.4, y: 1.25, w: 5.0, h: 3.6, max: 115, pion: true, nazwa: "wykres_nowi" });
karta(s, 5.7, 1.35, 3.8, 1.2, "przychod_powr");
liczba(s, 5.95, 1.4, 3.4, "97%", "przychodu dają powracający klienci (905 z 935 €)");
tekst(s, [B("Jak przekonać nowych", true), ...punkty([
  "rabat lub darmowa dostawa na pierwsze zakupy,",
  "mail z przypomnieniem o koszyku,",
  "reklamy do osób, które były na stronie,",
  "prostszy zakup na telefonie.",
])], 5.7, 2.8, 3.8, 2.1, { fontSize: 13 });

s = slajd("Wnioski i rekomendacje", "Podsumowując projekt 1: sklep sprzedaje dobrze tym, którzy go znają. Największa szansa jest w nowych klientach, szczególnie na telefonach i z wyszukiwarki.");
const rek1 = [
  ["Poprawić wersję mobilną", "prostszy koszyk, szybsze płatności (BLIK, Google Pay)"],
  ["Zadbać o pierwszy zakup", "oferta powitalna i maile do nowych"],
  ["Poprawić strony z wyszukiwarki", "Organic Search: 66% odrzuceń, 1 zakup"],
  ["Policzyć koszty kampanii", "bez nich nie ocenimy efektywności (CAC, ROAS)"],
];
rek1.forEach(([t1, t2], i) => {
  const x = 0.5 + (i % 2) * 4.6, y = 1.4 + Math.floor(i / 2) * 1.75;
  karta(s, x, y, 4.4, 1.55, `rek1_${i}`);
  kolko(s, x + 0.25, y + 0.25, 0.5, String(i + 1), C.accent2, 14);
  s.addText(t1, { x: x + 0.95, y: y + 0.22, w: 3.25, h: 0.45, isTextBox: true, fontSize: 15, bold: true, color: C.text1, margin: 0, valign: "middle" });
  s.addText(t2, { x: x + 0.95, y: y + 0.7, w: 3.25, h: 0.7, isTextBox: true, fontSize: 13, color: C.text2, margin: 0, valign: "top" });
});

// =============================================================================
// PROJEKT 2 – korelacja i DAG
// =============================================================================
sekcja("02", "Analiza korelacji i model przyczynowy skuteczności reklamy", "Dyrektor marketingu chce zwiększyć budżet reklamowy o 30%. Czy dane to uzasadniają?",
  "PROJEKT 2 · KORELACJA I MODEL PRZYCZYNOWY", "Drugi projekt: dyrektor zauważył, że w miesiącach z większą reklamą była większa sprzedaż, i proponuje dać 30% więcej na reklamę. Sprawdzamy, czy to uzasadnione.");

s = slajd("Dane: 12 miesięcy sprzedaży i reklamy", "Dane są syntetyczne, w tysiącach złotych. Indeks popytu to umowny wskaźnik zainteresowania kategorią, niezależny od reklam firmy, np. z trendów wyszukiwania. Widać duży skok w listopadzie i grudniu.");
const mies = ["Sty", "Lut", "Mar", "Kwi", "Maj", "Cze", "Lip", "Sie", "Wrz", "Paź", "Lis", "Gru"];
const reklama = [20, 22, 25, 27, 30, 33, 35, 34, 38, 42, 65, 70];
const sprzedaz = [210, 220, 245, 250, 275, 290, 310, 305, 340, 365, 520, 570];
const popyt = [85, 88, 93, 95, 102, 106, 110, 108, 115, 120, 155, 170];
const rabat = [0, 0, 5, 0, 5, 0, 10, 5, 5, 10, 20, 15];
tabela(s, ["", ...mies], [
  ["Reklama (tys. zł)", ...reklama.map(String)],
  ["Sprzedaż (tys. zł)", ...sprzedaz.map(String)],
  ["Indeks popytu", ...popyt.map(String)],
  ["Rabat (%)", ...rabat.map(String)],
], { x: 0.5, y: 1.4, w: 9, colW: [1.56, ...Array(12).fill(0.62)], prawo: [false, ...Array(12).fill(true)], font: 11, rowH: 0.36, pierwszaPogrubiona: true, nazwa: "tabela_dane2" });
tekst(s, [B("Co sprawdzamy: "), T("korelację Pearsona między reklamą, sprzedażą i popytem, dla całego roku i bez listopada i grudnia, a potem budujemy model przyczynowy (DAG).")], 0.5, 3.45, 9, 1.0, { fontSize: 14 });

s = slajd("Korelacja: wszystko rośnie razem", "Współczynnik Pearsona bliski 1 oznacza prawie idealną zależność liniową. Problem w tym, że popyt jest tak samo mocno związany ze sprzedażą jak reklama, a z reklamą prawie idealnie.");
tabela(s, ["Para zmiennych", "Cały rok", "Bez XI i XII"], [
  ["Reklama – sprzedaż", "0,999", "0,996"],
  ["Popyt – sprzedaż", "0,999", "0,996"],
  ["Popyt – reklama", "0,998", "0,998"],
], { x: 0.5, y: 1.4, w: 4.6, colW: [2.2, 1.2, 1.2], prawo: [false, true, true], font: 13, rowH: 0.45, pierwszaPogrubiona: true, nazwa: "tabela_korelacje" });
karta(s, 5.4, 1.4, 4.1, 3.4, "co_znaczy_r");
tekst(s, [B("Jak to czytać", true), ...punkty([
  "r = 0,999 to prawie idealny związek, ale tylko związek, nie przyczyna.",
  "Reklama i popyt rosną razem (0,998), więc nie wiemy, co napędza sprzedaż.",
  "Bez listopada i grudnia korelacje są prawie takie same, czyli problem jest przez cały rok.",
])], 5.65, 1.6, 3.65, 3.1, { fontSize: 13 });
tekst(s, "Korelacja ≠ przyczynowość", 0.5, 3.3, 4.6, 0.5, { fontSize: 20, bold: true, color: C.accent5 });

s = slajd("Wykres rozrzutu i linia regresji", "Linia regresji: sprzedaż = 7,16 × reklama + 61,8, R² = 0,998. Formalnie 1 tys. zł reklamy więcej to 7,16 tys. zł sprzedaży więcej, ale ta liczba zawiera też wpływ popytu i sezonu, więc nie jest prawdziwym efektem reklamy.");
s.addChart(pres.charts.SCATTER, [{ name: "Reklama", values: reklama }, { name: "Sprzedaż", values: sprzedaz }], {
  x: 0.4, y: 1.25, w: 5.6, h: 3.7, lineSize: 0, lineDataSymbol: "circle", lineDataSymbolSize: 8, chartColors: [HEX.accent1],
  showTitle: true, title: "Reklama a sprzedaż (tys. zł)", showLegend: false, valAxisMinVal: 150, catAxisMinVal: 15, catAxisMaxVal: 75,
  valGridLine: { color: "E3E8EA", size: 0.5 }, catGridLine: { style: "none" }, showCatAxisTitle: true, catAxisTitle: "Reklama",
  showValAxisTitle: true, valAxisTitle: "Sprzedaż", catAxisTitleFontSize: 11, valAxisTitleFontSize: 11, catAxisTitleColor: HEX.dk2, valAxisTitleColor: HEX.dk2,
  ...OS, objectName: "wykres_rozrzutu",
});
tabela(s, ["Linia regresji", "Cały rok", "Bez XI i XII"], [
  ["Nachylenie", "7,16", "7,12"],
  ["Wyraz wolny", "61,8", "63,3"],
  ["R²", "0,998", "0,992"],
], { x: 6.2, y: 1.4, w: 3.3, colW: [1.3, 1.0, 1.0], prawo: [false, true, true], font: 12, rowH: 0.38, pierwszaPogrubiona: true, nazwa: "tabela_regresja" });
tekst(s, "Listopad i grudzień leżą na tej samej linii co reszta roku, tylko dalej w prawo.", 6.2, 3.1, 3.3, 1.2, { fontSize: 13, color: C.text2 });

s = slajd("Pytania do dyskusji", "Odpowiedzi na pytania z karty zadania. Najważniejsze: wyłączenie miesięcy świątecznych nie rozwiązuje problemu, bo popyt i reklama rosną razem przez cały rok.");
[
  ["Czy silna korelacja oznacza, że reklama zwiększa sprzedaż?", "Nie. Firma reklamuje się więcej wtedy, gdy i tak jest większe zainteresowanie, np. przed świętami. Sama korelacja nie oddzieli wpływu reklamy od popytu."],
  ["Czy wyłączenie listopada i grudnia zmienia interpretację?", "Nie bardzo. Korelacja spada tylko z 0,999 do 0,996, więc wysoka korelacja nie wynika tylko ze świąt."],
  ["Czy takie wyłączenie jest uzasadnione metodologicznie?", "Nie do końca. Tracimy 2 z 12 miesięcy, akurat te z największą reklamą. Lepiej zostawić wszystkie dane i uwzględnić sezon w modelu."],
].forEach(([q, a], i) => {
  const y = 1.3 + i * 1.25;
  karta(s, 0.5, y, 9, 1.12, `pytanie_${i}`);
  kolko(s, 0.7, y + 0.28, 0.55, String(i + 1), C.accent1, 14);
  s.addText(q, { x: 1.45, y: y + 0.1, w: 7.85, h: 0.38, isTextBox: true, fontSize: 14, bold: true, color: C.text1, margin: 0, valign: "middle" });
  s.addText(a, { x: 1.45, y: y + 0.48, w: 7.85, h: 0.6, isTextBox: true, fontSize: 12, color: C.text2, margin: 0, valign: "top" });
});

s = slajd("Model przyczynowy (DAG)", "Strzałka A → B oznacza, że A wpływa na B. Zielone strzałki to ścieżka, którą działa reklama. Czerwone to zmienne zakłócające wpływające na wydatki reklamowe, a niebieskie to ich wpływ na sprzedaż.");
s.addImage({ path: path.join(REPO, "warsztat_2", "diagram_DAG.png"), x: 0.4, y: 1.2, w: 4.95, h: 4.95 * 1033 / 1280, altText: "Diagram DAG z siedmioma zmiennymi", objectName: "diagram_dag" });
tabela(s, ["Zmienna", "Rola", "W modelu"], [
  ["Wydatki reklamowe", "przyczyna", "tak"],
  ["Sezonowość", "zakłócająca", "tak"],
  ["Popyt rynkowy", "zakłócająca", "tak"],
  ["Rabaty", "zakłócająca", "tak"],
  ["Działania konkurencji", "zakłócająca", "tak"],
  ["Ruch na stronie", "mediator", "nie"],
  ["Sprzedaż", "skutek", "–"],
], { x: 5.6, y: 1.3, w: 3.9, colW: [1.7, 1.3, 0.9], font: 11, rowH: 0.34, pierwszaPogrubiona: true, wyroznij: [5], nazwa: "tabela_role" });
tekst(s, "Ruchu na stronie nie dodajemy, bo to przez niego działa reklama. Zaniżyłby jej efekt.", 5.6, 4.35, 3.9, 0.7, { fontSize: 12, color: C.text2 });

s = slajd("Dlaczego takie strzałki", "Uzasadnienie najważniejszych strzałek w modelu. Zmienne zakłócające mają strzałkę zarówno do reklamy, jak i do sprzedaży, dlatego tworzą fałszywy związek między nimi.");
[
  ["Popyt → reklama", "Firma zwiększa budżet, gdy widzi większe zainteresowanie. W danych korelacja wynosi 0,998."],
  ["Popyt → sprzedaż", "Przy większym zainteresowaniu ludzie kupują więcej, nawet bez reklamy firmy."],
  ["Reklama → ruch → sprzedaż", "Reklama sprowadza ludzi na stronę, a część z nich kupuje. To efekt, który chcemy zmierzyć."],
  ["Sezonowość → popyt, rabaty, reklama", "W listopadzie i grudniu rośnie popyt, są większe rabaty i więcej reklam."],
].forEach(([t1, t2], i) => {
  const x = 0.5 + (i % 2) * 4.6, y = 1.35 + Math.floor(i / 2) * 1.8;
  karta(s, x, y, 4.4, 1.6, `strzalka_${i}`);
  s.addText(t1, { x: x + 0.3, y: y + 0.2, w: 3.8, h: 0.4, isTextBox: true, fontSize: 15, bold: true, color: C.accent1, margin: 0 });
  s.addText(t2, { x: x + 0.3, y: y + 0.65, w: 3.8, h: 0.85, isTextBox: true, fontSize: 13, color: C.text1, margin: 0, valign: "top" });
});

s = slajd("Czy dać 30% więcej na reklamę?", "Odpowiedź dla dyrektora: korelacja nie wystarczy. Endogeniczność oznacza, że wydatki na reklamę nie są niezależne, tylko same zależą od popytu, sezonu i spodziewanej sprzedaży. Najlepszym sposobem jest eksperyment z grupą kontrolną.");
s.addText("Nie na podstawie samej korelacji", { x: 0.5, y: 1.3, w: 9, h: 0.5, isTextBox: true, fontSize: 20, bold: true, color: C.accent5, margin: 0 });
karta(s, 0.5, 1.95, 4.4, 2.3, "endogenicznosc");
karta(s, 5.1, 1.95, 4.4, 2.3, "jak_sprawdzic");
tekst(s, [B("Źródła endogeniczności", true), ...punkty([
  "popyt i sezon wpływają na reklamę i sprzedaż naraz,",
  "budżet ustala się pod spodziewaną sprzedaż,",
  "rabaty i konkurencja zmieniają się razem z reklamą.",
])], 0.75, 2.15, 3.95, 2.5, { fontSize: 13 });
tekst(s, [B("Jak wiarygodnie zmierzyć efekt", true), ...punkty([
  "eksperyment: więcej reklamy tylko w części regionów i porównanie z resztą,",
  "regresja z popytem, sezonem, rabatami i konkurencją na danych z dłuższego okresu.",
])], 5.35, 2.15, 3.95, 2.5, { fontSize: 13 });

// =============================================================================
// PROJEKT 3 – segmentacja rynku
// =============================================================================
sekcja("03", "Segmentacja rynku", "Ankieta „Opinia klienta”: słuchawki bezprzewodowe, 15 respondentów",
  "PROJEKT 3 · SEGMENTACJA RYNKU", "Trzeci projekt to nasza ankieta o słuchawkach bezprzewodowych. Na jej podstawie podzieliliśmy respondentów na segmenty i wybraliśmy segment docelowy.");

s = slajd("Badanie i kryteria segmentacji", "Ankieta była prowadzona w Formularzach Google od 30 października do 6 listopada 2026. Pytania dobraliśmy tak, żeby pokrywały różne kryteria segmentacji: demograficzne, behawioralne, korzyści, cenowe i psychograficzne.");
tabela(s, ["Kryterium", "Pytania w ankiecie"], [
  ["Demograficzne", "wiek, czym się zajmujesz"],
  ["Behawioralne", "czas w słuchawkach, sytuacje, miejsce zakupu"],
  ["Poszukiwane korzyści", "najważniejsze cechy słuchawek"],
  ["Cenowe", "ile maksymalnie zapłacisz"],
  ["Psychograficzne", "słuchawki jako styl, zaufanie do marki"],
], { x: 0.5, y: 1.4, w: 5.4, colW: [1.9, 3.5], font: 12, rowH: 0.48, pierwszaPogrubiona: true, nazwa: "tabela_kryteria" });
[["15", "respondentów"], ["10", "pytań, w tym 1 otwarte"], ["8 dni", "zbierania odpowiedzi"]].forEach(([l, o], i) => {
  karta(s, 6.2, 1.4 + i * 1.2, 3.3, 1.05, `badanie_${i}`);
  liczba(s, 6.45, 1.42 + i * 1.2, 2.9, l, o, C.accent1, 24);
});

s = slajd("Kim są respondenci", "Większość to młode osoby: ponad połowa ma 18-24 lata. Pracujących na etacie i studentów jest tyle samo, po 7 osób, i jedna osoba prowadzi własną firmę.");
slupki(s, ["Poniżej 18", "18-24", "25-34", "45 i więcej"], [1, 8, 5, 1], { x: 0.4, y: 1.25, w: 4.5, h: 3.0, tytul: "Wiek (liczba osób)", max: 10, nazwa: "wykres_wiek" });
slupki(s, ["Praca na etacie", "Nauka lub studia", "Freelance, firma"], [7, 7, 1], { x: 5.1, y: 1.25, w: 4.5, h: 3.0, tytul: "Zajęcie (liczba osób)", max: 9, nazwa: "wykres_zajecie" });
tekst(s, "Ponad połowa (8 z 15) ma 18-24 lata. Studenci i pracujący to dwie równe grupy.", 0.5, 4.45, 9, 0.5, { fontSize: 13, color: C.text2 });

s = slajd("Jak korzystają ze słuchawek", "Widać podział na osoby, które używają słuchawek krótko, i takie, które mają je na uszach przez dużą część dnia. Najczęstsza sytuacja to granie.");
slupki(s, ["Podczas grania", "W drodze", "W pracy lub nauce", "Na treningu", "W domu dla relaksu"], [8, 6, 5, 5, 5], { x: 0.4, y: 1.25, w: 5.4, h: 3.6, tytul: "Sytuacje (liczba wskazań)", max: 10, nazwa: "wykres_sytuacje" });
karta(s, 6.1, 1.35, 3.4, 3.45, "czas_dziennie");
s.addText("Czas dziennie", { x: 6.35, y: 1.5, w: 2.9, h: 0.35, isTextBox: true, fontSize: 14, bold: true, color: C.text1, margin: 0 });
liczba(s, 6.35, 1.85, 2.9, "6 osób", "do godziny dziennie", C.accent1, 24);
liczba(s, 6.35, 2.75, 2.9, "3 osoby", "od 1 do 3 godzin", C.accent3, 24);
liczba(s, 6.35, 3.65, 2.9, "6 osób", "ponad 3 godziny dziennie", C.accent2, 24);

s = slajd("Jakich korzyści szukają", "Wygoda wygrywa zdecydowanie: wybrało ją 13 z 15 osób. Zaraz potem jest jakość dźwięku. Wygląd wskazała tylko jedna osoba, więc to cecha drugorzędna.");
slupki(s, ["Wygoda noszenia", "Jakość dźwięku i bas", "Wyciszenie (ANC)", "Mocna bateria", "Cena i promocje", "Wygląd i design"], [13, 11, 8, 7, 4, 1], { x: 0.4, y: 1.25, w: 5.6, h: 3.6, tytul: "Najważniejsze cechy (liczba wskazań)", max: 15, nazwa: "wykres_cechy" });
karta(s, 6.3, 1.35, 3.2, 1.45, "kpi_wygoda");
liczba(s, 6.55, 1.4, 2.8, "87%", "wybrało wygodę noszenia", C.accent1, 26);
karta(s, 6.3, 3.0, 3.2, 1.45, "kpi_wyglad");
liczba(s, 6.55, 3.05, 2.8, "7%", "wskazało wygląd i design", C.accent2, 26);

s = slajd("Budżet i kanały zakupu", "Budżety są zróżnicowane, ale ponad połowa wyda więcej niż 400 zł. Większość kupuje przez internet, ale sklep stacjonarny wybrały aż 4 osoby.");
slupki(s, ["Do 200 zł", "201-400 zł", "401-700 zł", "Powyżej 700 zł"], [3, 4, 5, 3], { x: 0.4, y: 1.25, w: 4.5, h: 3.1, tytul: "Maksymalny budżet (liczba osób)", max: 7, nazwa: "wykres_budzet" });
slupki(s, ["Sklepy internetowe", "Sklep stacjonarny", "Strona producenta", "Allegro, Amazon"], [5, 4, 4, 2], { x: 5.1, y: 1.25, w: 4.5, h: 3.1, tytul: "Miejsce zakupu (liczba osób)", max: 7, kolor: HEX.accent2, nazwa: "wykres_gdzie" });
tekst(s, "8 z 15 osób wyda ponad 400 zł, a 11 z 15 kupuje przez internet.", 0.5, 4.5, 9, 0.45, { fontSize: 13, color: C.text2 });

s = slajd("Styl i zaufanie do marki", "Pytania w skali 1-5. Styl ma umiarkowane znaczenie (3,3). Zaufanie do marki jest podzielone: 6 osób dało 5, ale 4 osoby tylko 2.");
grupowe(s, ["1", "2", "3", "4", "5"], [["Słuchawki to element mojego stylu", [2, 2, 4, 4, 3]], ["Ufam sprawdzonym markom, np. JBL", [0, 4, 4, 1, 6]]], { x: 0.4, y: 1.2, w: 5.8, h: 3.75, max: 7, pion: true, nazwa: "wykres_skale" });
karta(s, 6.45, 1.35, 3.05, 1.35, "sr_styl");
liczba(s, 6.7, 1.4, 2.6, "3,3", "średnio: styl (1-5)", C.accent1, 26);
karta(s, 6.45, 2.9, 3.05, 1.35, "sr_marka");
liczba(s, 6.7, 2.95, 2.6, "3,6", "średnio: zaufanie do marki", C.accent2, 26);

// Macierz segmentów
s = slajd("Segmentacja: cztery grupy klientów", "Segmenty wyznaczyliśmy według dwóch kryteriów: ile czasu dziennie ktoś używa słuchawek (do godziny albo dłużej) i ile jest gotów wydać (do 400 zł albo więcej). To proste kryteria behawioralne i cenowe, które najlepiej różnicują odpowiedzi.");
const MX = 1.55, MY = 1.3, MW = 3.85, MH = 1.75, GAP = 0.12;
const kw = [
  [MX, MY, "C", "Wierni marce", "2 osoby · 13%", "krótko używają, ale wydadzą 401-700 zł", C.background2, C.text1],
  [MX + MW + GAP, MY, "A", "Entuzjaści dźwięku", "6 osób · 40%", "długo używają i wydadzą ponad 400 zł", C.accent1, C.background1],
  [MX, MY + MH + GAP, "D", "Okazjonalni oszczędni", "4 osoby · 27%", "do godziny dziennie, budżet do 400 zł", C.background2, C.text1],
  [MX + MW + GAP, MY + MH + GAP, "B", "Praktyczni pracujący", "3 osoby · 20%", "długo używają, budżet do 400 zł", C.background2, C.text1],
];
kw.forEach(([x, y, lit, nazwa, ile, op, fill, kol]) => {
  karta(s, x, y, MW, MH, `segment_${lit}`, fill);
  s.addText(`${lit}  ${nazwa}`, { x: x + 0.25, y: y + 0.18, w: MW - 0.5, h: 0.42, isTextBox: true, fontSize: 16, bold: true, color: kol, margin: 0 });
  s.addText(ile, { x: x + 0.25, y: y + 0.62, w: MW - 0.5, h: 0.45, isTextBox: true, fontSize: 20, bold: true, color: kol === C.background1 ? C.background1 : C.accent1, margin: 0 });
  s.addText(op, { x: x + 0.25, y: y + 1.1, w: MW - 0.5, h: 0.5, isTextBox: true, fontSize: 12, color: kol === C.background1 ? C.background2 : C.text2, margin: 0, valign: "top" });
});
s.addText("Budżet\nponad 400 zł", { x: 0.2, y: MY, w: 1.25, h: MH, isTextBox: true, fontSize: 11, bold: true, color: C.text2, align: "right", valign: "middle", margin: 0 });
s.addText("Budżet\ndo 400 zł", { x: 0.2, y: MY + MH + GAP, w: 1.25, h: MH, isTextBox: true, fontSize: 11, bold: true, color: C.text2, align: "right", valign: "middle", margin: 0 });
s.addText("Do godziny dziennie", { x: MX, y: MY + 2 * MH + GAP + 0.05, w: MW, h: 0.3, isTextBox: true, fontSize: 11, bold: true, color: C.text2, align: "center", margin: 0 });
s.addText("Ponad godzinę dziennie", { x: MX + MW + GAP, y: MY + 2 * MH + GAP + 0.05, w: MW, h: 0.3, isTextBox: true, fontSize: 11, bold: true, color: C.text2, align: "center", margin: 0 });

s = slajd("Profile segmentów", "Profil każdego segmentu z odpowiedzi w ankiecie. Entuzjaści to głównie młodzi studenci, którzy cenią dźwięk i kupują w sklepach stacjonarnych lub u producenta. Praktyczni pracujący potrzebują wygody i baterii na cały dzień i kupują online.");
tabela(s, ["", "A Entuzjaści dźwięku", "B Praktyczni pracujący", "C Wierni marce", "D Okazjonalni oszczędni"], [
  ["Liczebność", "6 osób (40%)", "3 osoby (20%)", "2 osoby (13%)", "4 osoby (27%)"],
  ["Kto", "głównie studenci, 18-24 lata", "wszyscy pracują na etacie", "pracujący", "głównie studenci"],
  ["Jak używają", "granie, w drodze, trening", "w pracy i w domu", "praca, dom", "granie, w drodze"],
  ["Co ważne", "dźwięk (6 z 6), wygoda", "bateria i wygoda (3 z 3)", "wygoda", "dźwięk, wygoda, cena"],
  ["Budżet", "3 osoby ponad 700 zł", "do 400 zł", "401-700 zł", "201-400 zł"],
  ["Gdzie kupują", "sklep stacjonarny, producent", "sklepy internetowe", "Allegro, producent", "różnie"],
  ["Zaufanie do marki", "3,3", "3,3", "5,0", "3,5"],
], { x: 0.5, y: 1.3, w: 9, colW: [1.45, 2.05, 1.9, 1.6, 2.0], font: 11, rowH: 0.47, pierwszaPogrubiona: true, nazwa: "tabela_segmenty" });

s = slajd("Segment docelowy i działania", "Jako segment docelowy wybieramy Entuzjastów dźwięku: to największa grupa (40%), najdłużej używa słuchawek i ma najwyższy budżet. Dla pozostałych segmentów proponujemy osobne, tańsze działania.");
karta(s, 0.5, 1.3, 4.4, 3.55, "cel_A", C.accent1);
tekst(s, [{ text: "Segment docelowy: A Entuzjaści dźwięku", options: { bold: true, color: C.background1, breakLine: true, fontSize: 16 } },
  { text: " ", options: { breakLine: true, fontSize: 6 } },
  { text: "40% respondentów, najdłużej używają słuchawek, 3 z 6 wydadzą ponad 700 zł.", options: { color: C.background1, breakLine: true } },
  { text: " ", options: { breakLine: true, fontSize: 6 } },
  { text: "Pozycjonowanie: słuchawki z mocnym dźwiękiem, wygodne na wiele godzin grania i w drodze.", options: { color: C.background2, italic: true } }],
  0.8, 1.55, 3.85, 3.1, { fontSize: 14 });
tabela(s, ["Segment", "Co proponujemy"], [
  ["A Entuzjaści", "testy dźwięku w sklepach, treści dla graczy"],
  ["B Praktyczni", "długa bateria i wygoda w modelach do 400 zł, sprzedaż online"],
  ["C Wierni marce", "program lojalnościowy, nowości dla stałych klientów"],
  ["D Okazjonalni", "promocje i rabaty na pierwsze słuchawki"],
], { x: 5.1, y: 1.3, w: 4.4, colW: [1.35, 3.05], font: 11, rowH: 0.55, pierwszaPogrubiona: true, wyroznij: [0], nazwa: "tabela_dzialania" });
tekst(s, "15 odpowiedzi to mała próba, więc segmenty pokazują kierunki do sprawdzenia na większej grupie.", 5.1, 4.45, 4.4, 0.5, { fontSize: 11, color: C.text2 });

// =============================================================================
// PROJEKT 4 – SCV
// =============================================================================
sekcja("04", "Analiza SCV (Single Customer View)", "Marka odzieżowa: cztery posty i wspólny profil klienta",
  "PROJEKT 4 · ANALIZA SCV", "Czwarty projekt: marka odzieżowa planuje kampanię sprzedażową. Porównujemy cztery posty i proponujemy, jakie dane połączyć we wspólny profil klienta.");

s = slajd("Dane i wzory", "Zakładamy, że zakupy przypisano do kliknięć w link danego posta. Liczymy dwa wskaźniki: zaangażowanie i konwersję po kliknięciu, w procentach z jednym miejscem po przecinku.");
tabela(s, ["Post", "Kanał i format", "Wyświetl.", "Polub.", "Koment.", "Udost.", "Kliknięcia", "Zakupy"], [
  ["104", "Instagram, Stories", "15 000", "900", "50", "5", "600", "3"],
  ["108", "Instagram, Reels", "14 000", "2 800", "350", "60", "470", "9"],
  ["109", "Facebook, post promocyjny", "30 000", "2 100", "400", "100", "900", "20"],
  ["110", "Instagram, post z obrazem", "21 000", "3 000", "420", "75", "500", "16"],
], { x: 0.5, y: 1.35, w: 9, colW: [0.6, 2.4, 1.0, 0.9, 0.95, 0.85, 1.25, 1.05], prawo: [false, false, true, true, true, true, true, true], font: 12, rowH: 0.38, pierwszaPogrubiona: true, nazwa: "tabela_posty" });
karta(s, 0.5, 3.45, 4.4, 1.4, "wzor_zaang");
karta(s, 5.1, 3.45, 4.4, 1.4, "wzor_konw");
tekst(s, [B("Zaangażowanie", true), T("(polubienia + komentarze + udostępnienia) ÷ wyświetlenia × 100%")], 0.75, 3.62, 3.95, 1.1, { fontSize: 13 });
tekst(s, [B("Konwersja po kliknięciu", true), T("zakupy ÷ kliknięcia w link × 100%")], 5.35, 3.62, 3.95, 1.1, { fontSize: 13 });

s = slajd("Wyniki: zaangażowanie nie oznacza sprzedaży", "Przykład dla posta 104: (900 + 50 + 5) ÷ 15 000 × 100% = 6,4% zaangażowania, a 3 ÷ 600 × 100% = 0,5% konwersji. Reels ma najwięcej reakcji, ale najlepiej sprzedaje post z obrazem.");
tabela(s, ["Post", "Zaangażowanie", "Konwersja"], [
  ["104 Stories", "6,4%", "0,5%"],
  ["108 Reels", "22,9%", "1,9%"],
  ["109 Post promocyjny", "8,7%", "2,2%"],
  ["110 Post z obrazem", "16,6%", "3,2%"],
], { x: 0.5, y: 1.35, w: 3.75, colW: [1.55, 1.3, 0.9], prawo: [false, true, true], font: 11, rowH: 0.4, pierwszaPogrubiona: true, wyroznij: [3], nazwa: "tabela_wskazniki" });
const posty = ["104 Stories", "108 Reels", "109 Promocyjny", "110 Z obrazem"];
slupki(s, posty, [6.4, 22.9, 8.7, 16.6], { x: 4.4, y: 1.2, w: 2.65, h: 3.0, tytul: "Zaangażowanie (%)", max: 30, format: "0.0", nazwa: "wykres_zaang" });
slupki(s, posty, [0.5, 1.9, 2.2, 3.2], { x: 7.05, y: 1.2, w: 2.65, h: 3.0, tytul: "Konwersja (%)", max: 4.2, format: "0.0", kolor: HEX.accent2, nazwa: "wykres_konw" });
tekst(s, "Post, na który ludzie najchętniej reagują, nie jest tym, który najlepiej sprzedaje.", 0.5, 4.4, 9, 0.5, { fontSize: 13, color: C.text2 });

s = slajd("Który post wybieramy do testu", "Do dalszego testu wybieramy post 110, bo jako jedyny jest wysoko w obu wskaźnikach. Post 109 dał najwięcej zakupów, ale przy największym zasięgu i dzięki rabatom.");
[
  ["108 Reels", "22,9%", "najwyższe zaangażowanie, ale z 470 kliknięć tylko 9 zakupów", C.accent1],
  ["110 Post z obrazem", "3,2%", "najwyższa konwersja: 16 zakupów z 500 kliknięć", C.accent1],
  ["104 Stories", "0,5%", "dużo kliknięć (600), a tylko 3 zakupy", C.accent5],
].forEach(([n, w, o, kol], i) => {
  const x = 0.5 + i * 3.07;
  karta(s, x, 1.3, 2.87, 1.95, `post_${i}`);
  s.addText(n, { x: x + 0.25, y: 1.45, w: 2.4, h: 0.35, isTextBox: true, fontSize: 14, bold: true, color: C.text1, margin: 0 });
  s.addText(w, { x: x + 0.25, y: 1.8, w: 2.4, h: 0.55, isTextBox: true, fontSize: 28, bold: true, color: kol, margin: 0 });
  s.addText(o, { x: x + 0.25, y: 2.4, w: 2.4, h: 0.75, isTextBox: true, fontSize: 12, color: C.text2, margin: 0, valign: "top" });
});
karta(s, 0.5, 3.45, 9, 1.4, "decyzja", C.accent1);
tekst(s, [{ text: "Do testu sprzedażowego: post 110 „Inspiracje na zimę”", options: { bold: true, color: C.background1, breakLine: true, fontSize: 16 } },
  { text: "Konwersja 3,2% (najwyższa) i zaangażowanie 16,6% (drugie najwyższe): przyciąga uwagę i sprzedaje.", options: { color: C.background1 } }],
  0.8, 3.62, 8.4, 1.1, { fontSize: 14 });

s = slajd("Jak zbudować wspólny profil klienta (SCV)", "Dane o postach są zbiorcze, więc nie wiemy, kto konkretnie kliknął i kupił. Łączymy dwa źródła danych własnych marki po adresie e-mail. Dzięki temu możemy np. wysłać przypomnienie o porzuconym koszyku bez płacenia drugi raz za reklamę.");
karta(s, 0.5, 1.3, 4.4, 1.35, "zrodlo_sklep");
karta(s, 5.1, 1.3, 4.4, 1.35, "zrodlo_crm");
tekst(s, [B("Źródło 1: sklep internetowy", true), T("konto, historia zamówień, rozmiary, porzucone koszyki, wejścia z linków UTM")], 0.75, 1.45, 3.95, 1.15, { fontSize: 13 });
tekst(s, [B("Źródło 2: newsletter i program lojalnościowy", true), T("zapisy, otwarcia i kliknięcia w maile, zgody, punkty")], 5.35, 1.45, 3.95, 1.15, { fontSize: 13 });
s.addText([B("Wspólny identyfikator: "), { text: "adres e-mail klienta", options: {} }], { x: 0.5, y: 2.8, w: 9, h: 0.4, isTextBox: true, fontSize: 16, color: C.accent1, margin: 0 });
const kroki = ["Klik w post 110 (link z UTM)", "Płaszcz w koszyku, brak zakupu", "Profil po e-mailu: newsletter, rozmiar M", "Mail z przypomnieniem i stylizacjami"];
kroki.forEach((k, i) => {
  const x = 0.5 + i * 2.33;
  karta(s, x, 3.45, 1.95, 1.3, `krok_${i}`, i === 3 ? C.accent2 : C.background2);
  s.addText(k, { x: x + 0.12, y: 3.5, w: 1.71, h: 1.2, isTextBox: true, fontSize: 12, bold: i === 3, color: i === 3 ? C.background1 : C.text1, align: "center", valign: "middle", margin: 0 });
  if (i < 3) strzalka(s, x + 1.98, 4.1, x + 2.3, 4.1);
});

// =============================================================================
// Podsumowanie
// =============================================================================
sekcjaAkt = "Podsumowanie";
kickerAkt = "PODSUMOWANIE";
pres.addSection({ title: "Podsumowanie" });
s = slajd("Najważniejsze wnioski", "Krótko o każdym projekcie: w jednym zdaniu, co z niego wynika.");
[
  ["1", "Filtrowanie danych", "Powracający dają 97% przychodu. Trzeba zadbać o pierwszy zakup nowych, zwłaszcza na telefonach."],
  ["2", "Korelacja i DAG", "Korelacja 0,999 nie dowodzi efektu reklamy. Decyzję o +30% warto oprzeć na eksperymencie."],
  ["3", "Segmentacja", "Segment docelowy to Entuzjaści dźwięku: 40% osób, długie użytkowanie i najwyższy budżet."],
  ["4", "Analiza SCV", "Post 110 łączy zaangażowanie z konwersją. E-mail łączy dane ze sklepu i newslettera."],
].forEach(([nr, t1, t2], i) => {
  const x = 0.5 + (i % 2) * 4.6, y = 1.35 + Math.floor(i / 2) * 1.85;
  karta(s, x, y, 4.4, 1.65, `wniosek_${nr}`);
  kolko(s, x + 0.25, y + 0.25, 0.55, nr, C.accent1, 16);
  s.addText(t1, { x: x + 1.0, y: y + 0.22, w: 3.2, h: 0.42, isTextBox: true, fontSize: 15, bold: true, color: C.text1, margin: 0, valign: "middle" });
  s.addText(t2, { x: x + 1.0, y: y + 0.68, w: 3.2, h: 0.9, isTextBox: true, fontSize: 12, color: C.text2, margin: 0, valign: "top" });
});

s = pres.addSlide({ masterName: "TYTUL", sectionTitle: "Podsumowanie" });
s.addText("PODSTAWY ANALITYKI INTERNETOWEJ", { placeholder: "kicker" });
s.addText("Dziękujemy", { placeholder: "title" });
s.addText("Pytania?", { placeholder: "body" });

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("Zapisano", OUT);
})();
