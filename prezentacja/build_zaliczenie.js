// Prezentacja zaliczeniowa: Podstawy analityki internetowej, projekty 2, 3 i 4.
// Uruchomienie: NODE_PATH=<katalog z pptxgenjs, react, react-dom, react-icons, sharp> APPLY_THEME=<apply_theme.js> node build_zaliczenie.js
// Potem: python3 trendline.py Podstawy_analityki_projekty.pptx (linia trendu na wykresie rozrzutu)
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa6");
const { applyTheme } = require(process.env.APPLY_THEME);

const KAT = __dirname;
const REPO = path.resolve(KAT, "..");
const OUT = path.join(KAT, "Podstawy_analityki_projekty.pptx");

const THEME = {
  name: "Podstawy analityki",
  headFontFace: "Calibri",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "16232F", lt1: "FFFFFF", dk2: "4A5866", lt2: "EDF4F3",
    accent1: "0F7B74", accent2: "E8742C", accent3: "7FC3BC", accent4: "CFE3E0",
    accent5: "C0392B", accent6: "6B7680", hlink: "0F7B74", folHlink: "6B7680",
  },
};
const HEX = THEME.colors;

// --- ikony (react-icons -> PNG) -------------------------------------------------------
const IKONY = {
  dag: "FaDiagramProject", sluchawki: "FaHeadphones", profil: "FaIdCard", zarowka: "FaLightbulb",
  wykres: "FaChartLine", ostrzezenie: "FaTriangleExclamation", eksperyment: "FaFlask", cel: "FaBullseye",
  produkt: "FaBox", ankieta: "FaClipboardList", sklep: "FaStore", mail: "FaEnvelope", koszyk: "FaCartShopping",
  link: "FaLink", klik: "FaArrowPointer", wyslij: "FaPaperPlane", cytat: "FaQuoteLeft", student: "FaGraduationCap",
  praca: "FaBriefcase", pytanie: "FaCircleQuestion", ksiazka: "FaBook", osoby: "FaUserGroup", obrot: "FaArrowsRotate",
};
const IK = {};
async function przygotujIkony() {
  for (const [klucz, nazwa] of Object.entries(IKONY)) {
    const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(fa[nazwa], { color: "#FFFFFF", size: "256" }));
    const png = await sharp(Buffer.from(svg)).resize(256, 256, { fit: "contain", background: { r: 0, g: 0, b: 0, alpha: 0 } }).png().toBuffer();
    IK[klucz] = "image/png;base64," + png.toString("base64");
  }
}

function zbuduj() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.title = "Podstawy analityki internetowej – projekty 2, 3 i 4";
  const C = pres.SchemeColor;

  // --- układy -----------------------------------------------------------------------
  pres.defineSlideMaster({
    title: "TYTUL",
    background: { color: C.text1 },
    objects: [
      { placeholder: { options: { name: "kicker", type: "body", x: 0.6, y: 0.8, w: 8.8, h: 0.4, fontSize: 14, bold: true, color: C.accent3, align: "left", margin: 0, charSpacing: 2 }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.2, w: 8.8, h: 1.05, fontSize: 40, bold: true, color: C.background1, valign: "top", align: "left", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 2.3, w: 8.8, h: 1.1, fontSize: 16, color: C.accent4, valign: "top", align: "left", margin: 0 }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "SEKCJA",
    background: { color: C.text1 },
    objects: [
      { placeholder: { options: { name: "nr", type: "body", x: 0.6, y: 0.85, w: 4, h: 0.85, fontSize: 48, bold: true, color: C.accent3, align: "left", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.75, w: 6.2, h: 1.7, fontSize: 30, bold: true, color: C.background1, valign: "top", align: "left", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 3.55, w: 6.2, h: 1.2, fontSize: 15, color: C.accent4, valign: "top", align: "left", margin: 0 }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "TRESC",
    background: { color: C.background1 },
    slideNumber: { x: 9.1, y: 5.22, w: 0.5, h: 0.25, fontSize: 10, color: C.accent6, align: "right" },
    objects: [
      { placeholder: { options: { name: "kicker", type: "body", x: 0.5, y: 0.28, w: 9, h: 0.28, fontSize: 11, bold: true, color: C.accent1, align: "left", margin: 0, charSpacing: 1 }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.56, w: 9, h: 0.6, fontSize: 26, bold: true, color: C.text1, valign: "middle", align: "left", margin: 0 }, text: "" } },
    ],
  });

  // --- pomocnicze -------------------------------------------------------------------
  const OS = { catAxisLabelColor: HEX.dk2, valAxisLabelColor: HEX.dk2, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
    catAxisLabelFontSize: 11, valAxisLabelFontSize: 10, dataLabelFontFace: "+mn-lt", dataLabelFontSize: 11, dataLabelColor: HEX.dk1,
    titleFontFace: "+mn-lt", titleFontSize: 12, titleColor: HEX.dk1, legendFontFace: "+mn-lt", legendFontSize: 11, legendColor: HEX.dk2 };

  let sekcjaAkt = "Wstęp";
  let kickerAkt = "";
  let licznik = 0;
  const id = (p) => `${p}_${++licznik}`;

  function slajd(tytul, notatki) {
    const s = pres.addSlide({ masterName: "TRESC", sectionTitle: sekcjaAkt });
    if (kickerAkt) s.addText(kickerAkt, { placeholder: "kicker" });
    s.addText(tytul, { placeholder: "title" });
    if (notatki) s.addNotes(notatki);
    return s;
  }
  function karta(s, x, y, w, h, nazwa, kolor) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: kolor || C.background2 }, line: { type: "none" }, rectRadius: 0.08, objectName: id(nazwa) });
  }
  function ikonaKolo(s, x, y, d, klucz, kolor = C.accent1) {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: kolor }, line: { type: "none" }, objectName: id(`kolo_${klucz}`) });
    const m = d * 0.25;
    s.addImage({ data: IK[klucz], x: x + m, y: y + m, w: d - 2 * m, h: d - 2 * m, altText: `ikona: ${klucz}`, objectName: id(`ikona_${klucz}`) });
  }
  function sekcja(nr, tytul, opis, kicker, ikona, notatki) {
    sekcjaAkt = tytul;
    kickerAkt = kicker;
    pres.addSection({ title: tytul });
    const s = pres.addSlide({ masterName: "SEKCJA", sectionTitle: tytul });
    s.addText(nr, { placeholder: "nr" });
    s.addText(tytul, { placeholder: "title" });
    s.addText(opis, { placeholder: "body" });
    ikonaKolo(s, 7.15, 1.55, 2.1, ikona, C.accent1);
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
  function kolko(s, x, y, d, napis, kolor = C.accent1, rozmiar = 14) {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: kolor }, line: { type: "none" }, objectName: id(`kolko_${napis}`) });
    s.addText(napis, { x, y, w: d, h: d, isTextBox: true, fontSize: rozmiar, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0 });
  }
  function liczba(s, x, y, w, duza, opis, kolor = C.accent1, rozmiar = 28) {
    s.addText(duza, { x, y, w, h: 0.55, isTextBox: true, fontSize: rozmiar, bold: true, color: kolor, margin: 0, valign: "bottom" });
    s.addText(opis, { x, y: y + 0.58, w, h: 0.5, isTextBox: true, fontSize: 12, color: C.text2, margin: 0, valign: "top" });
  }
  function strzalka(s, x1, y1, x2, y2) {
    s.addShape(pres.shapes.LINE, { x: x1, y: y1, w: x2 - x1, h: y2 - y1, line: { color: HEX.dk2, width: 1.5, endArrowType: "triangle" }, objectName: id("strzalka") });
  }
  // ramka z wnioskiem na dole slajdu (motyw całej prezentacji)
  function wniosek(s, txt) {
    karta(s, 0.5, 4.6, 8.45, 0.6, "wniosek", C.background2);
    ikonaKolo(s, 0.63, 4.71, 0.38, "zarowka", C.accent2);
    s.addText([{ text: "Wniosek: ", options: { bold: true, color: C.accent2 } }, { text: txt, options: { bold: true } }],
      { x: 1.15, y: 4.6, w: 7.7, h: 0.6, isTextBox: true, fontSize: 13, color: C.text1, valign: "middle", margin: 0 });
  }
  function slupki(s, etykiety, wartosci, o) {
    s.addChart(pres.charts.BAR, [{ name: o.seria || "Wartość", labels: etykiety, values: wartosci }], {
      x: o.x, y: o.y, w: o.w, h: o.h, barDir: o.pion ? "col" : "bar", catAxisOrientation: o.pion ? "minMax" : "maxMin",
      chartColors: [o.kolor || HEX.accent1], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: o.format || "General",
      valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, valAxisMinVal: 0, valAxisMaxVal: o.max,
      showLegend: false, showTitle: !!o.tytul, title: o.tytul || "", barGapWidthPct: 55, catAxisLineShow: true,
      catAxisLineColor: HEX.accent4, ...OS, objectName: id(o.nazwa),
    });
  }
  function grupowe(s, etykiety, serie, o) {
    s.addChart(pres.charts.BAR, serie.map(([name, values]) => ({ name, labels: etykiety, values })), {
      x: o.x, y: o.y, w: o.w, h: o.h, barDir: o.pion ? "col" : "bar", barGrouping: "clustered",
      catAxisOrientation: o.pion ? "minMax" : "maxMin",
      chartColors: [HEX.accent1, HEX.accent2], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: o.format || "General",
      valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, valAxisMinVal: 0, valAxisMaxVal: o.max,
      showLegend: true, legendPos: "t", showTitle: !!o.tytul, title: o.tytul || "", barGapWidthPct: 45,
      catAxisLineShow: true, catAxisLineColor: HEX.accent4, ...OS, objectName: id(o.nazwa),
    });
  }
  function tabela(s, naglowki, wiersze, o) {
    const f = o.font || 12;
    const nag = naglowki.map((t, i) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.accent1 }, fontSize: f, align: o.prawo && o.prawo[i] ? "right" : "left" } }));
    const cialo = wiersze.map((w, r) => w.map((t, i) => ({ text: t, options: {
      fontSize: f, color: C.text1, bold: (o.pierwszaPogrubiona && i === 0) || (o.wyroznij && o.wyroznij.includes(r) && o.pogrubWyroznione),
      align: o.prawo && o.prawo[i] ? "right" : "left",
      fill: o.wyroznij && o.wyroznij.includes(r) ? { color: C.background2 } : undefined,
    } })));
    s.addTable([nag, ...cialo], { x: o.x, y: o.y, w: o.w, colW: o.colW, border: { type: "solid", pt: 0.75, color: HEX.accent4 },
      valign: "middle", margin: [0.03, 0.07, 0.03, 0.07], rowH: o.rowH || 0.3, objectName: id(o.nazwa) });
  }

  // ===================================================================================
  // Wstęp
  // ===================================================================================
  pres.addSection({ title: "Wstęp" });
  let s = pres.addSlide({ masterName: "TYTUL", sectionTitle: "Wstęp" });
  s.addText("PROJEKTY ZALICZENIOWE 2, 3 I 4", { placeholder: "kicker" });
  s.addText("Podstawy analityki internetowej", { placeholder: "title" });
  s.addText("WSIZ · prowadząca: dr n.t. M. Marecka · listopad 2026", { placeholder: "body" });
  [["dag", "Projekt 2", "Analiza korelacji i model przyczynowy skuteczności reklamy"],
    ["sluchawki", "Projekt 3", "Segmentacja rynku"],
    ["profil", "Projekt 4", "Analiza SCV"]].forEach(([ik, nr, nazwa], i) => {
    const x = 0.6 + i * 3.0;
    ikonaKolo(s, x, 3.75, 0.62, ik, C.accent1);
    s.addText([{ text: nr, options: { bold: true, color: C.accent3, breakLine: true } }, { text: nazwa, options: { color: C.background1 } }],
      { x: x + 0.78, y: 3.65, w: 2.1, h: 0.9, isTextBox: true, fontSize: 12, valign: "middle", margin: 0 });
  });
  s.addNotes("Dzień dobry. Przedstawiamy trzy projekty z zajęć z Podstaw analityki internetowej: projekt 2, czyli analizę korelacji i model przyczynowy skuteczności reklamy, projekt 3, czyli segmentację rynku na podstawie naszej ankiety o słuchawkach JBL, oraz projekt 4, czyli analizę SCV, Single Customer View.");

  s = slajd("Trzy projekty, trzy pytania", "Każdy projekt to inne dane i inna metoda. W projekcie 2 sprawdzamy, czy reklama naprawdę zwiększa sprzedaż. W projekcie 3 dzielimy klientów na segmenty na podstawie naszej ankiety. W projekcie 4 porównujemy posty w social mediach i proponujemy wspólny profil klienta. We wszystkich chodzi o to samo: żeby decyzje marketingowe opierać na danych, a nie na intuicji.");
  [["dag", "Projekt 2", "Korelacja i model przyczynowy", "Czy reklama naprawdę zwiększa sprzedaż?", "12 miesięcy: reklama, sprzedaż, popyt, rabaty"],
    ["sluchawki", "Projekt 3", "Segmentacja rynku", "Jakie grupy klientów słuchawek JBL widać w ankiecie?", "nasza ankieta: 10 pytań, 27 osób"],
    ["profil", "Projekt 4", "Analiza SCV", "Który post sprzedaje i jak połączyć dane o kliencie?", "4 posty marki odzieżowej"]].forEach(([ik, nr, nazwa, pyt, dane], i) => {
    const x = 0.5 + i * 3.07;
    karta(s, x, 1.3, 2.87, 3.1, `plan_${i}`);
    ikonaKolo(s, x + 0.25, 1.5, 0.7, ik, C.accent1);
    s.addText(nr.toUpperCase(), { x: x + 1.1, y: 1.55, w: 1.6, h: 0.3, isTextBox: true, fontSize: 11, bold: true, color: C.accent1, margin: 0, charSpacing: 1 });
    s.addText(nazwa, { x: x + 1.1, y: 1.83, w: 1.65, h: 0.4, isTextBox: true, fontSize: 13, bold: true, color: C.text1, margin: 0, valign: "top" });
    s.addText(pyt, { x: x + 0.25, y: 2.5, w: 2.4, h: 1.0, isTextBox: true, fontSize: 15, bold: true, color: C.text1, margin: 0, valign: "top" });
    s.addText(dane, { x: x + 0.25, y: 3.6, w: 2.4, h: 0.65, isTextBox: true, fontSize: 12, color: C.text2, margin: 0, valign: "top" });
  });
  wniosek(s, "każdy projekt to inne dane, ale cel jest ten sam: decyzje oparte na danych, a nie na intuicji.");

  // ===================================================================================
  // PROJEKT 2 – korelacja i DAG
  // ===================================================================================
  sekcja("02", "Analiza korelacji i model przyczynowy skuteczności reklamy", "Dyrektor marketingu chce zwiększyć budżet reklamowy o 30%. Czy dane to uzasadniają?",
    "PROJEKT 2 · KORELACJA I MODEL PRZYCZYNOWY", "dag",
    "Projekt 2. Firma e-commerce analizuje skuteczność płatnych kampanii. Dyrektor marketingu zauważył, że w miesiącach z wyższymi wydatkami na reklamę sprzedaż też była wyższa, i proponuje zwiększyć budżet o 30%. Naszym zadaniem było ocenić, czy ta korelacja uzasadnia decyzję, wskazać źródła endogeniczności i zaproponować wiarygodny sposób oszacowania efektu reklamy.");

  const mies = ["Sty", "Lut", "Mar", "Kwi", "Maj", "Cze", "Lip", "Sie", "Wrz", "Paź", "Lis", "Gru"];
  const reklama = [20, 22, 25, 27, 30, 33, 35, 34, 38, 42, 65, 70];
  const sprzedaz = [210, 220, 245, 250, 275, 290, 310, 305, 340, 365, 520, 570];
  const popyt = [85, 88, 93, 95, 102, 106, 110, 108, 115, 120, 155, 170];
  const rabat = [0, 0, 5, 0, 5, 0, 10, 5, 5, 10, 20, 15];

  s = slajd("Sytuacja i dane: 12 miesięcy", "Dane są syntetyczne, z pliku Excela z zajęć: 12 miesięcy, kwoty w tysiącach złotych, rabat w punktach procentowych. Indeks popytu to umowny wskaźnik zainteresowania kategorią, niezależny od reklam firmy. Widać, że w listopadzie i grudniu wszystko skacze naraz: reklama z 42 do 65 i 70 tysięcy, sprzedaż z 365 do 520 i 570 tysięcy, popyt i rabaty też. To pierwsza wskazówka, że za wzrostem sprzedaży może stać nie tylko reklama.");
  tekst(s, [B("Teza dyrektora: "), T("w miesiącach z większą reklamą była większa sprzedaż, więc trzeba dać 30% więcej na reklamę. "), B("Nasze zadanie: "), T("sprawdzić, czy dane to uzasadniają.")], 0.5, 1.3, 9, 0.6, { fontSize: 14 });
  tabela(s, ["", ...mies], [
    ["Reklama (tys. zł)", ...reklama.map(String)],
    ["Sprzedaż (tys. zł)", ...sprzedaz.map(String)],
    ["Indeks popytu", ...popyt.map(String)],
    ["Rabat (%)", ...rabat.map(String)],
  ], { x: 0.5, y: 2.05, w: 9, colW: [1.56, ...Array(12).fill(0.62)], prawo: [false, ...Array(12).fill(true)], font: 11, rowH: 0.36, pierwszaPogrubiona: true, nazwa: "tabela_dane2" });
  tekst(s, "Dane syntetyczne z zajęć. Indeks popytu: umowny wskaźnik zainteresowania kategorią, niezależny od reklam firmy.", 0.5, 3.95, 9, 0.4, { fontSize: 11, color: C.text2 });
  wniosek(s, "w listopadzie i grudniu rosną naraz reklama, sprzedaż, popyt i rabaty.");

  s = slajd("Etap I: korelacja Pearsona", "Policzyliśmy współczynnik korelacji Pearsona funkcją PEARSON w Excelu, dla całego roku i bez listopada i grudnia. Reklama ze sprzedażą: 0,999, czyli prawie idealny związek liniowy. Ale popyt ze sprzedażą ma dokładnie tak samo wysoką korelację, 0,999, a popyt z reklamą 0,998. Czyli reklama i popyt rosną razem i z samej korelacji nie da się powiedzieć, co naprawdę napędza sprzedaż. Bez listopada i grudnia korelacje spadają tylko do 0,996, więc problem jest przez cały rok, nie tylko w święta.\n\nJeśli padnie pytanie, co znaczy r: to siła i kierunek związku liniowego, od minus 1 do 1. Wartość bliska 1 to silny związek dodatni, ale nie dowód przyczyny.");
  tabela(s, ["Para zmiennych", "Cały rok", "Bez XI i XII"], [
    ["Reklama – sprzedaż", "0,999", "0,996"],
    ["Popyt – sprzedaż", "0,999", "0,996"],
    ["Popyt – reklama", "0,998", "0,998"],
  ], { x: 0.5, y: 1.35, w: 4.6, colW: [2.2, 1.2, 1.2], prawo: [false, true, true], font: 13, rowH: 0.45, pierwszaPogrubiona: true, nazwa: "tabela_korelacje" });
  ikonaKolo(s, 0.5, 3.35, 0.6, "ostrzezenie", C.accent5);
  s.addText([{ text: "Korelacja ≠ przyczynowość", options: { bold: true, color: C.accent5, fontSize: 20, breakLine: true } },
    { text: "r bliskie 1 to silny związek, ale nie dowód, że reklama go powoduje", options: { color: C.text2, fontSize: 12 } }],
  { x: 1.25, y: 3.25, w: 3.85, h: 0.85, isTextBox: true, margin: 0, valign: "middle" });
  karta(s, 5.4, 1.35, 4.1, 3.0, "jak_czytac");
  tekst(s, [B("Jak to czytać", true), ...punkty([
    "Reklama i sprzedaż: r = 0,999, prawie idealny związek liniowy.",
    "Popyt ze sprzedażą: też 0,999, a popyt z reklamą 0,998. Wszystko rośnie razem.",
    "Bez listopada i grudnia: dalej 0,996, więc to nie tylko efekt świąt.",
  ])], 5.65, 1.55, 3.65, 2.7, { fontSize: 13 });
  wniosek(s, "popyt jest związany ze sprzedażą tak samo mocno jak reklama, więc korelacja nie wskazuje przyczyny.");

  s = slajd("Wykres rozrzutu i linia regresji", "Każdy punkt to jeden miesiąc: na osi poziomej reklama, na pionowej sprzedaż. Linia regresji z Excela: sprzedaż = 7,16 × reklama + 61,8, R² = 0,998, czyli prosta wyjaśnia 99,8% zmienności sprzedaży. Bez listopada i grudnia: 7,12 × reklama + 63,3, R² = 0,992, prawie to samo. Formalnie 1 tysiąc złotych reklamy więcej to 7,16 tysiąca sprzedaży więcej, ale ta liczba zawiera też wpływ popytu, sezonu i rabatów, więc nie jest prawdziwym efektem reklamy. Listopad i grudzień leżą na tej samej linii, tylko dalej w prawo.");
  s.addChart(pres.charts.SCATTER, [{ name: "Reklama", values: reklama }, { name: "Sprzedaż", values: sprzedaz }], {
    x: 0.4, y: 1.2, w: 5.6, h: 3.3, lineSize: 0, lineDataSymbol: "circle", lineDataSymbolSize: 8, chartColors: [HEX.accent1],
    showTitle: true, title: "Reklama a sprzedaż (tys. zł)", showLegend: false, valAxisMinVal: 150, catAxisMinVal: 15, catAxisMaxVal: 75,
    valGridLine: { color: "E3E8EA", size: 0.5 }, catGridLine: { style: "none" }, showCatAxisTitle: true, catAxisTitle: "Reklama",
    showValAxisTitle: true, valAxisTitle: "Sprzedaż", catAxisTitleFontSize: 11, valAxisTitleFontSize: 11, catAxisTitleColor: HEX.dk2, valAxisTitleColor: HEX.dk2,
    ...OS, objectName: id("wykres_rozrzutu"),
  });
  tabela(s, ["Linia regresji", "Cały rok", "Bez XI i XII"], [
    ["Nachylenie", "7,16", "7,12"],
    ["Wyraz wolny", "61,8", "63,3"],
    ["R²", "0,998", "0,992"],
  ], { x: 6.2, y: 1.35, w: 3.3, colW: [1.3, 1.0, 1.0], prawo: [false, true, true], font: 12, rowH: 0.38, pierwszaPogrubiona: true, nazwa: "tabela_regresja" });
  karta(s, 6.2, 3.0, 3.3, 1.35, "rownanie");
  tekst(s, [B("sprzedaż = 7,16 × reklama + 61,8", true), T("Listopad i grudzień leżą na tej samej linii, tylko dalej w prawo.")], 6.4, 3.12, 2.95, 1.15, { fontSize: 12 });
  wniosek(s, "7,16 tys. zł sprzedaży na 1 tys. zł reklamy to opis związku, a nie zmierzony efekt reklamy.");

  s = slajd("Pytania do dyskusji z karty zadania", "To odpowiedzi na trzy pytania do dyskusji z karty zadania. Pierwsze: silna korelacja nie oznacza, że reklama odpowiada za wzrost przychodów, bo firma reklamuje się więcej wtedy, kiedy i tak jest większe zainteresowanie. Drugie: wyłączenie listopada i grudnia prawie nic nie zmienia, korelacja spada z 0,999 do 0,996. Trzecie: takie wyłączenie nie jest dobrą metodą, bo usuwamy 2 z 12 miesięcy, akurat te z największą reklamą i sprzedażą, a wpływ popytu dalej zostaje w danych. Lepiej zostawić wszystkie miesiące i uwzględnić sezon i popyt w modelu.");
  [
    ["Czy silna korelacja oznacza, że reklama odpowiada za wzrost przychodów?", "Nie. Firma reklamuje się więcej wtedy, gdy i tak jest większe zainteresowanie, np. przed świętami. Sama korelacja nie oddzieli wpływu reklamy od popytu."],
    ["Czy wyłączenie miesięcy świątecznych zmienia interpretację wyników?", "Nie bardzo. Korelacja spada tylko z 0,999 do 0,996, więc wysoka korelacja nie wynika tylko ze świąt."],
    ["Czy takie wyłączenie jest samo w sobie metodologicznie uzasadnione?", "Nie do końca. Tracimy 2 z 12 miesięcy, akurat te z największą reklamą. Lepiej zostawić wszystkie dane i uwzględnić sezon w modelu."],
  ].forEach(([q, a], i) => {
    const y = 1.3 + i * 1.07;
    karta(s, 0.5, y, 9, 0.97, `pytanie_${i}`);
    kolko(s, 0.7, y + 0.23, 0.5, String(i + 1), C.accent1, 14);
    s.addText(q, { x: 1.4, y: y + 0.08, w: 7.9, h: 0.34, isTextBox: true, fontSize: 14, bold: true, color: C.text1, margin: 0, valign: "middle" });
    s.addText(a, { x: 1.4, y: y + 0.44, w: 7.9, h: 0.48, isTextBox: true, fontSize: 12, color: C.text2, margin: 0, valign: "top" });
  });
  wniosek(s, "wyłączenie świąt nie rozwiązuje problemu, bo popyt i reklama rosną razem przez cały rok.");

  s = slajd("Etap II: model przyczynowy (DAG)", "W drugim etapie zbudowaliśmy DAG, czyli graf skierowany bez cykli z siedmioma zmiennymi z zadania. Strzałka zawsze prowadzi od przyczyny do skutku. Kolory: zielone strzałki to ścieżka przyczynowa, przez którą działa reklama: wydatki reklamowe, ruch na stronie, sprzedaż. Czerwone to wpływ zmiennych zakłócających na wydatki reklamowe, a niebieskie ich wpływ na sprzedaż i na inne zmienne. Zmienne zakłócające to sezonowość, popyt rynkowy, rabaty i działania konkurencji, bo wpływają i na reklamę, i na sprzedaż. Ruch na stronie to zmienna pośrednicząca, czyli mediator. Sprzedaż to zmienna wynikowa.\n\nDo regresji, żeby oszacować całkowity efekt reklamy, bierzemy: wydatki reklamowe, sezonowość, popyt, rabaty i działania konkurencji. Ruchu na stronie nie dodajemy, bo to przez niego reklama działa, więc zaniżylibyśmy jej efekt.");
  s.addImage({ path: path.join(REPO, "warsztat_2", "diagram_DAG.png"), x: 0.45, y: 1.22, w: 4.0, h: 4.0 * 1033 / 1280, altText: "Diagram DAG z siedmioma zmiennymi: wydatki reklamowe, ruch na stronie, sprzedaż, sezonowość, popyt rynkowy, rabaty, działania konkurencji", objectName: id("diagram_dag") });
  tabela(s, ["Zmienna", "Rola w modelu", "W regresji"], [
    ["Wydatki reklamowe", "przyczyna (badany wpływ)", "tak"],
    ["Sezonowość", "zmienna zakłócająca", "tak"],
    ["Popyt rynkowy", "zmienna zakłócająca", "tak"],
    ["Rabaty", "zmienna zakłócająca", "tak"],
    ["Działania konkurencji", "zmienna zakłócająca", "tak"],
    ["Ruch na stronie", "zmienna pośrednicząca (mediator)", "nie"],
    ["Sprzedaż", "zmienna wynikowa", "–"],
  ], { x: 4.75, y: 1.3, w: 4.75, colW: [1.7, 2.25, 0.8], font: 11, rowH: 0.36, pierwszaPogrubiona: true, wyroznij: [5], nazwa: "tabela_role" });
  wniosek(s, "sprzedaż = reklama + sezonowość + popyt + rabaty + konkurencja. Mediatora (ruchu) nie dodajemy.");

  s = slajd("Uzasadnienie najważniejszych strzałek", "W zadaniu trzeba było uzasadnić co najmniej trzy strzałki, mamy cztery. Popyt na reklamę: firma zwiększa budżet, kiedy widzi większe zainteresowanie, w danych popyt z reklamą ma korelację 0,998. Popyt na sprzedaż: kiedy ludzie bardziej interesują się kategorią, kupują więcej nawet bez reklamy firmy. Reklama, ruch, sprzedaż: reklama sprowadza ludzi na stronę, a część z nich kupuje, i to jest efekt, który chcemy zmierzyć. Sezonowość: w listopadzie i grudniu rośnie popyt, są większe rabaty i firma daje więcej reklam.");
  [
    ["Popyt → wydatki reklamowe", "Firma zwiększa budżet, gdy widzi większe zainteresowanie. Popyt z reklamą: r = 0,998."],
    ["Popyt → sprzedaż", "Przy większym zainteresowaniu kategorią ludzie kupują więcej, nawet bez reklamy firmy."],
    ["Reklama → ruch na stronie → sprzedaż", "Reklama sprowadza ludzi na stronę, a część z nich kupuje. To efekt, który chcemy zmierzyć."],
    ["Sezonowość → popyt, rabaty, reklama", "W listopadzie i grudniu rośnie popyt, są większe rabaty (20% i 15%) i więcej reklam."],
  ].forEach(([t1, t2], i) => {
    const x = 0.5 + (i % 2) * 4.6, y = 1.3 + Math.floor(i / 2) * 1.6;
    karta(s, x, y, 4.4, 1.45, `strzalka_${i}`);
    kolko(s, x + 0.22, y + 0.22, 0.45, String(i + 1), i === 2 ? C.accent1 : C.accent2, 13);
    s.addText(t1, { x: x + 0.85, y: y + 0.18, w: 3.4, h: 0.5, isTextBox: true, fontSize: 14, bold: true, color: C.accent1, margin: 0, valign: "middle" });
    s.addText(t2, { x: x + 0.85, y: y + 0.7, w: 3.4, h: 0.7, isTextBox: true, fontSize: 12, color: C.text1, margin: 0, valign: "top" });
  });
  wniosek(s, "zmienne zakłócające wpływają i na reklamę, i na sprzedaż, więc tworzą pozorny związek między nimi.");

  s = slajd("Czy dać 30% więcej na reklamę?", "Nasza odpowiedź dla dyrektora: korelacja nie uzasadnia zwiększenia budżetu o 30%. Problemem jest endogeniczność, czyli to, że wydatki na reklamę nie są niezależne od czynników, które wpływają na sprzedaż. Na wykładzie były trzy mechanizmy i wszystkie tu występują. Pominięta zmienna: popyt, sezon, rabaty i konkurencja wpływają naraz na reklamę i sprzedaż. Odwrotna przyczynowość: firma ustala budżet pod spodziewaną sprzedaż. Błąd pomiaru: znamy wydatki, a nie to, ile osób faktycznie zobaczyło reklamę.\n\nJak to sprawdzić: najlepiej geo-eksperymentem, czyli zwiększyć budżet o 30% tylko w losowo wybranych regionach i porównać sprzedaż z regionami kontrolnymi. Można też zrobić test A/B albo regresję z popytem, sezonem, rabatami i konkurencją na danych z dłuższego okresu, ale sama regresja nie gwarantuje usunięcia endogeniczności.");
  s.addText("Nie na podstawie samej korelacji", { x: 0.5, y: 1.25, w: 9, h: 0.4, isTextBox: true, fontSize: 18, bold: true, color: C.accent5, margin: 0 });
  karta(s, 0.5, 1.75, 4.4, 2.65, "endogenicznosc");
  karta(s, 5.1, 1.75, 4.4, 2.65, "jak_sprawdzic", C.accent1);
  tekst(s, [B("Endogeniczność: 3 mechanizmy", true),
    B("Pominięta zmienna: "), T("popyt, sezon, rabaty i konkurencja wpływają naraz na reklamę i sprzedaż.", true),
    B("Odwrotna przyczynowość: "), T("budżet ustala się pod spodziewaną sprzedaż.", true),
    B("Błąd pomiaru: "), T("znamy wydatki, a nie rzeczywisty zasięg reklamy.")], 0.75, 1.92, 3.95, 2.4, { fontSize: 12, paraSpaceAfter: 5 });
  tekst(s, [{ text: "Jak wiarygodnie oszacować efekt", options: { bold: true, breakLine: true } },
    { text: "Geo-eksperyment: ", options: { bold: true } }, { text: "+30% budżetu tylko w losowo wybranych regionach i porównanie z regionami kontrolnymi.", options: { breakLine: true } },
    { text: "Test A/B: ", options: { bold: true } }, { text: "grupa z reklamą i grupa kontrolna.", options: { breakLine: true } },
    { text: "Regresja (MMM): ", options: { bold: true } }, { text: "z popytem, sezonem, rabatami i konkurencją, na dłuższych danych." }],
  5.35, 1.92, 3.95, 2.4, { fontSize: 12, color: C.background1, paraSpaceAfter: 5 });
  wniosek(s, "korelacja nie uzasadnia +30% budżetu. Efekt reklamy trzeba najpierw zmierzyć eksperymentem.");

  // ===================================================================================
  // PROJEKT 3 – segmentacja rynku
  // ===================================================================================
  sekcja("03", "Segmentacja rynku", "Nasza ankieta „Opinia klienta”: słuchawki bezprzewodowe JBL, 27 respondentów",
    "PROJEKT 3 · SEGMENTACJA RYNKU", "sluchawki",
    "Projekt 3. Według karty zadania mieliśmy wybrać produkt, przygotować ankietę z 10 pytań, zebrać odpowiedzi, podzielić respondentów na segmenty i zaproponować działania marketingowe. Wybraliśmy słuchawki bezprzewodowe JBL.");

  s = slajd("Produkt, cel i kryteria segmentacji", "Produkt to słuchawki bezprzewodowe JBL: znana marka, słuchawek używają i uczniowie, i osoby pracujące, a modele są w bardzo różnych cenach. Cel badania: pomóc firmie zrozumieć klientów, podzielić ich na segmenty i wybrać segment docelowy. W zadaniu trzeba było wybrać co najmniej 3 kryteria segmentacji, my mamy pięć. Nazwy są takie jak na wykładzie: demograficzna, behawioralna, wartości klienta i psychograficzna, a do tego preferencje, czyli korzyści, których klient szuka. Geograficznej nie użyliśmy, bo nie pytaliśmy o miejsce zamieszkania.");
  karta(s, 0.5, 1.3, 3.6, 1.45, "produkt");
  ikonaKolo(s, 0.7, 1.5, 0.55, "produkt", C.accent1);
  tekst(s, [B("Produkt", true), T("słuchawki bezprzewodowe JBL")], 1.45, 1.47, 2.5, 1.2, { fontSize: 13 });
  karta(s, 0.5, 2.9, 3.6, 1.5, "cel");
  ikonaKolo(s, 0.7, 3.1, 0.55, "cel", C.accent2);
  tekst(s, [B("Cel badania", true), T("poznać grupy klientów i wybrać segment docelowy")], 1.45, 3.07, 2.5, 1.25, { fontSize: 13 });
  tabela(s, ["Rodzaj segmentacji (wykład)", "Pytania w ankiecie"], [
    ["Demograficzna", "wiek, czym się zajmujesz"],
    ["Behawioralna", "czas w słuchawkach, sytuacje, miejsce zakupu"],
    ["Wartości klienta", "ile maksymalnie zapłacisz"],
    ["Psychograficzna", "słuchawki jako styl, zaufanie do marki"],
    ["Preferencje (korzyści)", "najważniejsze cechy, czego brakuje"],
  ], { x: 4.35, y: 1.3, w: 5.15, colW: [1.95, 3.2], font: 12, rowH: 0.5, pierwszaPogrubiona: true, nazwa: "tabela_kryteria" });
  wniosek(s, "mamy 5 kryteriów segmentacji, a w karcie zadania minimum to 3.");

  s = slajd("Ankieta i przebieg badania", "Ankietę zrobiliśmy w Formularzach Google. Ma 10 pytań różnych typów, tak jak wymagała karta zadania: 5 pytań jednokrotnego wyboru, 2 wielokrotnego wyboru, 2 pytania w skali Likerta od 1 do 5 i jedno pytanie otwarte. Odpowiedzi zbieraliśmy od 30 października do 9 listopada 2026. Formularz wysłano 32 razy, ale 6 zgłoszeń z 9 listopada było identycznych i przyszło w ciągu 30 sekund, więc to jedna osoba, która kliknęła wyślij kilka razy. Liczymy ją raz, czyli usuwamy 5 powtórzeń i zostaje 27 odpowiedzi. To dbanie o jakość danych, bo inaczej jedna osoba ważyłaby sześć razy więcej.");
  [["5", "pytań jednokrotnego wyboru"], ["2", "pytania wielokrotnego wyboru"], ["2", "pytania w skali Likerta"], ["1", "pytanie otwarte"]].forEach(([l, o], i) => {
    const x = 0.5 + (i % 2) * 2.2, y = 1.3 + Math.floor(i / 2) * 1.55;
    karta(s, x, y, 2.05, 1.4, `typ_${i}`);
    liczba(s, x + 0.22, y + 0.12, 1.7, l, o, i === 3 ? C.accent2 : C.accent1, 30);
  });
  karta(s, 4.95, 1.3, 4.55, 2.95, "przebieg");
  ikonaKolo(s, 5.15, 1.48, 0.5, "ankieta", C.accent1);
  tekst(s, [B("Przebieg badania", true), T("Formularze Google, 30.10–9.11.2026")], 5.8, 1.47, 3.5, 0.6, { fontSize: 13, paraSpaceAfter: 2 });
  [["32", "zgłoszenia", C.text2], ["−5", "powtórzeń", C.accent5], ["27", "do analizy", C.accent1]].forEach(([l, o, kol], i) => {
    const x = 5.2 + i * 1.45;
    s.addText(l, { x, y: 2.3, w: 1.25, h: 0.6, isTextBox: true, fontSize: 30, bold: true, color: kol, margin: 0, align: "center" });
    s.addText(o, { x, y: 2.9, w: 1.25, h: 0.3, isTextBox: true, fontSize: 12, color: C.text2, margin: 0, align: "center" });
  });
  tekst(s, "6 identycznych odpowiedzi z 9.11 wysłanych w 30 sekund to jedna osoba.", 5.2, 3.4, 4.1, 0.7, { fontSize: 12, color: C.text2 });
  wniosek(s, "10 pytań różnych typów, 27 odpowiedzi po usunięciu powtórzeń, czyli dbamy o jakość danych.");

  s = slajd("Kim są respondenci", "Większość respondentów to młode osoby: 23 z 27 ma mniej niż 35 lat, a najwięcej, 10 osób, ma 18-24 lata. Studentów i pracujących na etacie jest po 11, a 5 osób prowadzi własną firmę albo pracuje jako freelancer. Dzięki temu mogliśmy porównać dwie grupy: uczących się i pracujących.");
  slupki(s, ["Poniżej 18", "18-24", "25-34", "35-44", "45 i więcej"], [5, 10, 8, 2, 2], { x: 0.4, y: 1.2, w: 4.5, h: 3.25, tytul: "Wiek (liczba osób)", max: 12, nazwa: "wykres_wiek" });
  slupki(s, ["Praca na etacie", "Nauka lub studia", "Freelance, firma"], [11, 11, 5], { x: 5.1, y: 1.2, w: 4.5, h: 3.25, tytul: "Zajęcie (liczba osób)", max: 13, nazwa: "wykres_zajecie" });
  wniosek(s, "23 z 27 osób ma mniej niż 35 lat, a studentów i pracujących na etacie jest po 11.");

  s = slajd("Jak korzystają ze słuchawek", "Prawie połowa respondentów, 13 z 27, używa słuchawek ponad 3 godziny dziennie. 8 osób najwyżej godzinę, a 6 od 1 do 3 godzin. Najczęstsza sytuacja to praca lub nauka, 14 wskazań, zaraz potem granie, 12. W tym pytaniu można było zaznaczyć dwie odpowiedzi, dlatego to liczba wskazań, a nie osób.");
  slupki(s, ["W pracy lub nauce", "Podczas grania", "Na treningu", "W domu dla relaksu", "W drodze"], [14, 12, 9, 9, 8], { x: 0.4, y: 1.2, w: 5.4, h: 3.25, tytul: "Sytuacje (liczba wskazań)", max: 17, nazwa: "wykres_sytuacje" });
  karta(s, 6.1, 1.3, 3.4, 3.1, "czas_dziennie");
  s.addText("Czas dziennie w słuchawkach", { x: 6.35, y: 1.42, w: 3.0, h: 0.35, isTextBox: true, fontSize: 13, bold: true, color: C.text1, margin: 0 });
  liczba(s, 6.35, 1.7, 2.9, "8 osób", "do godziny dziennie", C.accent3, 22);
  liczba(s, 6.35, 2.55, 2.9, "6 osób", "od 1 do 3 godzin", C.accent1, 22);
  liczba(s, 6.35, 3.4, 2.9, "13 osób", "ponad 3 godziny dziennie", C.accent2, 22);
  wniosek(s, "prawie połowa (13 z 27) ma słuchawki na uszach ponad 3 godziny dziennie.");

  s = slajd("Jakich korzyści szukają", "W pytaniu o najważniejsze cechy trzeba było wybrać trzy. Wygoda noszenia wygrywa: wskazało ją 20 z 27 osób, czyli 74%. Zaraz potem jest jakość dźwięku i mocny bas, 18 osób. Wyciszenie ANC i bateria są ważne dla mniej więcej połowy. Wygląd wskazały tylko 2 osoby, czyli 7%, więc to cecha drugorzędna.");
  slupki(s, ["Wygoda noszenia", "Jakość dźwięku i bas", "Wyciszenie (ANC)", "Mocna bateria", "Cena i promocje", "Wygląd i design"], [20, 18, 13, 12, 10, 2], { x: 0.4, y: 1.2, w: 5.6, h: 3.25, tytul: "Najważniejsze cechy (liczba wskazań)", max: 24, nazwa: "wykres_cechy" });
  karta(s, 6.3, 1.3, 3.2, 1.45, "kpi_wygoda");
  liczba(s, 6.55, 1.38, 2.8, "74%", "wybrało wygodę noszenia (20 z 27)", C.accent1, 28);
  karta(s, 6.3, 2.95, 3.2, 1.45, "kpi_wyglad");
  liczba(s, 6.55, 3.03, 2.8, "7%", "wskazało wygląd i design (2 z 27)", C.accent2, 28);
  wniosek(s, "liczą się wygoda i dźwięk, a nie wygląd.");

  s = slajd("Budżet i kanały zakupu", "Najwięcej osób, 11, wyda na słuchawki od 401 do 700 złotych. Razem ponad 400 złotych wyda 16 z 27 osób, czyli ponad połowa. Najczęściej wybierane miejsce zakupu to sklep stacjonarny, 12 osób, ale jeśli zsumujemy sklepy internetowe, stronę producenta i Allegro czy Amazon, to przez internet kupuje 15 osób.");
  slupki(s, ["Do 200 zł", "201-400 zł", "401-700 zł", "Powyżej 700 zł"], [5, 6, 11, 5], { x: 0.4, y: 1.2, w: 4.5, h: 3.25, tytul: "Maksymalny budżet (liczba osób)", max: 13, nazwa: "wykres_budzet" });
  slupki(s, ["Sklep stacjonarny", "Sklepy internetowe", "Strona producenta", "Allegro, Amazon"], [12, 7, 4, 4], { x: 5.1, y: 1.2, w: 4.5, h: 3.25, tytul: "Miejsce zakupu (liczba osób)", max: 14, kolor: HEX.accent2, nazwa: "wykres_gdzie" });
  wniosek(s, "16 z 27 wyda ponad 400 zł. Sklep stacjonarny wybrało 12 osób, internet łącznie 15.");

  s = slajd("Styl i zaufanie do marki", "Dwa pytania w skali Likerta od 1 do 5. Pierwsze: czy słuchawki to element stylu. Średnia 3,3, najczęściej odpowiedź środkowa, 3. Drugie: czy kupując elektronikę, ufam tylko sprawdzonym markom, takim jak JBL. Tu średnia jest wyższa, 3,7: 15 osób dało 4 lub 5, a nikt nie dał 1. Dla marki to dobra wiadomość: zaufanie do sprawdzonych marek jest wyraźnie wyższe niż znaczenie stylu.");
  grupowe(s, ["1", "2", "3", "4", "5"], [["Słuchawki to element mojego stylu", [3, 3, 9, 6, 6]], ["Ufam sprawdzonym markom, np. JBL", [0, 4, 8, 7, 8]]], { x: 0.4, y: 1.2, w: 5.8, h: 3.3, max: 11, pion: true, nazwa: "wykres_skale" });
  karta(s, 6.45, 1.3, 3.05, 1.45, "sr_styl");
  liczba(s, 6.7, 1.38, 2.6, "3,3", "średnio: słuchawki jako styl (1-5)", C.accent1, 28);
  karta(s, 6.45, 2.95, 3.05, 1.45, "sr_marka");
  liczba(s, 6.7, 3.03, 2.6, "3,7", "średnio: zaufanie do marki (1-5)", C.accent2, 28);
  wniosek(s, "zaufanie do marki (3,7) jest wyższe niż znaczenie stylu (3,3). 15 osób dało marce 4 lub 5.");

  s = slajd("Czego brakuje klientom: pytanie otwarte", "Ostatnie pytanie było otwarte: czego brakuje w słuchawkach, których używasz dzisiaj. Odpowiedziało 21 z 27 osób. Pogrupowaliśmy odpowiedzi w kategorie. Najczęściej: lepsze wyciszenie otoczenia, 4 osoby, oraz wygoda i dopasowanie, też 4: słuchawki uciskają głowę, nauszniki się rozciągają albo źle pasują do ucha. 4 osoby napisały, że niczego im nie brakuje, w tym jedna, która używa JBL. Pojedyncze uwagi: mikrofon, przełączanie między urządzeniami, składanie słuchawek i komunikat o ładowaniu w trakcie muzyki.");
  slupki(s, ["Lepsze wyciszenie", "Wygoda i dopasowanie", "Niczego nie brakuje", "Dłuższa bateria", "Lepszy dźwięk", "Lepszy wygląd", "Inne (po 1 osobie)"], [4, 4, 4, 2, 2, 2, 4], { x: 0.4, y: 1.2, w: 5.0, h: 3.25, tytul: "Uwagi (liczba osób, odpowiedziało 21 z 27)", max: 5, nazwa: "wykres_uwagi" });
  karta(s, 5.6, 1.3, 3.9, 3.1, "cytaty");
  ikonaKolo(s, 5.8, 1.45, 0.45, "cytat", C.accent2);
  tekst(s, [
    { text: "„Nic, bo używam JBL. Są wygodne, nie uciskają mnie (…). I fajnie wyciszają szumy z otoczenia”", options: { italic: true, breakLine: true } },
    { text: "„Lepszego dopasowania do kształtu ucha”", options: { italic: true, breakLine: true } },
    { text: "„Usunęłabym głos kobiety w trakcie piosenek o potrzebie ładowania”", options: { italic: true } },
  ], 5.85, 2.05, 3.45, 2.3, { fontSize: 14, paraSpaceAfter: 12 });
  wniosek(s, "najczęstsze braki to wyciszenie oraz wygoda i dopasowanie, czyli wskazówki dla produktu.");

  s = slajd("Dwie grupy: studenci i pracujący", "Karta zadania mówi o badaniu w dwóch grupach, więc porównaliśmy osoby uczące się lub studiujące, 11 osób, i pracujące, 16 osób, w tym 5 z własną firmą. Studenci najczęściej używają słuchawek do grania, 8 z 11, i najbardziej cenią dźwięk. Pracujący używają ich głównie w pracy, 10 z 16, i najbardziej cenią wygodę, 81%. Pracujący mają też wyższe zaufanie do marki: 3,9 wobec 3,5.");
  grupowe(s, ["Wygoda", "Dźwięk i bas", "Wyciszenie (ANC)", "Bateria", "Cena i promocje"], [["Studenci (11)", [0.64, 0.73, 0.55, 0.45, 0.45]], ["Pracujący (16)", [0.81, 0.62, 0.44, 0.44, 0.31]]],
    { x: 0.4, y: 1.2, w: 5.2, h: 3.3, max: 1, format: "0%", tytul: "Co ważne przy wyborze (% osób w grupie)", nazwa: "wykres_grupy" });
  tabela(s, ["", "Studenci", "Pracujący"], [
    ["Najczęstsze użycie", "granie (8 z 11)", "praca (10 z 16)"],
    ["Ponad 3 h dziennie", "6 z 11", "7 z 16"],
    ["Budżet ponad 400 zł", "6 z 11", "10 z 16"],
    ["Zaufanie do marki", "3,5", "3,9"],
  ], { x: 5.85, y: 1.35, w: 3.65, colW: [1.45, 1.1, 1.1], font: 11, rowH: 0.5, pierwszaPogrubiona: true, nazwa: "tabela_grupy" });
  wniosek(s, "studenci to granie i dźwięk, pracujący to praca i wygoda oraz wyższe zaufanie do marki.");

  // Macierz segmentów
  s = slajd("Segmentacja: cztery grupy klientów", "Segmenty wyznaczyliśmy według dwóch kryteriów, które najbardziej różnicowały odpowiedzi. Pierwsze jest behawioralne: ile czasu dziennie ktoś używa słuchawek, do godziny albo dłużej. Drugie to wartość klienta: ile jest gotów wydać, do 400 złotych albo więcej. Tak jak w przykładzie z wykładu, gdzie klientów dzielono na kupujących drogo lub tanio i często lub rzadko. Wyszły cztery segmenty. A, Entuzjaści dźwięku: 14 osób, 52%. B, Praktyczni pracujący: 5 osób. C, Wierni marce: 2 osoby. D, Okazjonalni oszczędni: 6 osób. Wierni marce przypominają lojalsów z wykładu, a Okazjonalni oszczędni klientów okazjonalnych i łowców promocji.");
  const MX = 1.55, MY = 1.28, MW = 3.85, MH = 1.38, GAP = 0.12;
  const kw = [
    [MX, MY, "C", "Wierni marce", "2 osoby · 7%", "krótko używają, ale wydadzą 401-700 zł", C.background2, C.text1],
    [MX + MW + GAP, MY, "A", "Entuzjaści dźwięku", "14 osób · 52%", "długo używają i wydadzą ponad 400 zł", C.accent1, C.background1],
    [MX, MY + MH + GAP, "D", "Okazjonalni oszczędni", "6 osób · 22%", "do godziny dziennie, budżet do 400 zł", C.background2, C.text1],
    [MX + MW + GAP, MY + MH + GAP, "B", "Praktyczni pracujący", "5 osób · 19%", "długo używają, budżet do 400 zł", C.background2, C.text1],
  ];
  kw.forEach(([x, y, lit, nazwa, ile, op, fill, kol]) => {
    karta(s, x, y, MW, MH, `segment_${lit}`, fill);
    s.addText(`${lit}  ${nazwa}`, { x: x + 0.25, y: y + 0.13, w: MW - 0.5, h: 0.38, isTextBox: true, fontSize: 15, bold: true, color: kol, margin: 0 });
    s.addText(ile, { x: x + 0.25, y: y + 0.5, w: MW - 0.5, h: 0.42, isTextBox: true, fontSize: 19, bold: true, color: kol === C.background1 ? C.background1 : C.accent1, margin: 0 });
    s.addText(op, { x: x + 0.25, y: y + 0.94, w: MW - 0.5, h: 0.35, isTextBox: true, fontSize: 12, color: kol === C.background1 ? C.background2 : C.text2, margin: 0, valign: "top" });
  });
  s.addText("Budżet\nponad 400 zł", { x: 0.2, y: MY, w: 1.25, h: MH, isTextBox: true, fontSize: 11, bold: true, color: C.text2, align: "right", valign: "middle", margin: 0 });
  s.addText("Budżet\ndo 400 zł", { x: 0.2, y: MY + MH + GAP, w: 1.25, h: MH, isTextBox: true, fontSize: 11, bold: true, color: C.text2, align: "right", valign: "middle", margin: 0 });
  s.addText("Do godziny dziennie", { x: MX, y: MY + 2 * MH + GAP + 0.04, w: MW, h: 0.26, isTextBox: true, fontSize: 11, bold: true, color: C.text2, align: "center", margin: 0 });
  s.addText("Ponad godzinę dziennie", { x: MX + MW + GAP, y: MY + 2 * MH + GAP + 0.04, w: MW, h: 0.26, isTextBox: true, fontSize: 11, bold: true, color: C.text2, align: "center", margin: 0 });
  wniosek(s, "dzielimy według zachowania (czas) i wartości klienta (budżet), jak klientów A–D z wykładu.");

  s = slajd("Profile segmentów", "Profil każdego segmentu z odpowiedzi w ankiecie. A, Entuzjaści dźwięku: głównie młode osoby, 9 z 14 ma mniej niż 25 lat, 10 z 14 używa słuchawek ponad 3 godziny dziennie, cenią dźwięk i wygodę i kupują w sklepach stacjonarnych. B, Praktyczni pracujący: głównie pracujący na etacie, używają słuchawek w pracy, chcą wygody i kupują w sklepach internetowych. C, Wierni marce: obie osoby oceniają zaufanie do marki na 5. D, Okazjonalni oszczędni: głównie uczniowie i studenci z budżetem do 400 złotych. To najmniej zadowolona grupa: 5 z 6 napisało, czego brakuje im w obecnych słuchawkach.");
  tabela(s, ["", "A Entuzjaści dźwięku", "B Praktyczni pracujący", "C Wierni marce", "D Okazjonalni oszczędni"], [
    ["Liczebność", "14 osób (52%)", "5 osób (19%)", "2 osoby (7%)", "6 osób (22%)"],
    ["Kto", "9 z 14 poniżej 25 lat", "pracujący na etacie, 25-34", "pracujący na etacie", "głównie uczniowie i studenci"],
    ["Jak używają", "10 z 14 ponad 3 h dziennie", "głównie w pracy (4 z 5)", "do godziny dziennie", "do godziny, często granie"],
    ["Co ważne", "dźwięk i wygoda (10 z 14)", "wygoda (5 z 5)", "wygoda (2 z 2)", "dźwięk, bateria (4 z 6)"],
    ["Budżet", "5 osób ponad 700 zł", "3 z 5 do 200 zł", "401-700 zł", "do 400 zł"],
    ["Gdzie kupują", "sklep stacjonarny (9 z 14)", "sklepy internetowe (4 z 5)", "Allegro, strona producenta", "różnie"],
    ["Zaufanie do marki", "3,6", "3,4", "5,0", "3,7"],
  ], { x: 0.5, y: 1.3, w: 9, colW: [1.4, 1.9, 1.9, 1.9, 1.9], font: 11, rowH: 0.4, pierwszaPogrubiona: true, nazwa: "tabela_segmenty" });
  wniosek(s, "Okazjonalni oszczędni są najmniej zadowoleni: 5 z 6 napisało, czego im brakuje.");

  s = slajd("Segment docelowy i działania", "Jako segment docelowy wybieramy A, Entuzjastów dźwięku. Przemawiają za tym trzy liczby: to ponad połowa respondentów, 14 z 27, 10 z 14 używa słuchawek ponad 3 godziny dziennie, a wszyscy wydadzą ponad 400 złotych, w tym 5 ponad 700. Pozycjonowanie: słuchawki JBL z mocnym dźwiękiem, wygodne przez wiele godzin grania, nauki i pracy. Dla pozostałych segmentów proponujemy osobne, tańsze działania: dla Praktycznych modele do 400 złotych z długą baterią i promocje online, dla Wiernych marce program lojalnościowy, a dla Okazjonalnych zniżki studenckie i promocje, np. w Black Friday.\n\nOgraniczenie: 27 odpowiedzi to mała próba, segmenty C i D są małe, więc wyniki pokazują kierunek do sprawdzenia na większej grupie.");
  karta(s, 0.5, 1.3, 4.4, 3.1, "cel_A", C.accent1);
  ikonaKolo(s, 0.75, 1.5, 0.55, "cel", C.accent2);
  tekst(s, [{ text: "Segment docelowy", options: { color: C.accent4, breakLine: true, fontSize: 12, bold: true } },
    { text: "A Entuzjaści dźwięku", options: { bold: true, color: C.background1, fontSize: 18 } }], 1.45, 1.47, 3.3, 0.7, { paraSpaceAfter: 0 });
  tekst(s, [{ text: "14 z 27 osób (52%), 10 z 14 ponad 3 h dziennie, wszyscy wydadzą ponad 400 zł.", options: { color: C.background1, breakLine: true } },
    { text: " ", options: { breakLine: true, fontSize: 6 } },
    { text: "Pozycjonowanie: słuchawki JBL z mocnym dźwiękiem, wygodne przez wiele godzin grania, nauki i pracy.", options: { color: C.background2, italic: true } }],
  0.8, 2.35, 3.85, 1.95, { fontSize: 13 });
  tabela(s, ["Segment", "Co proponujemy"], [
    ["A Entuzjaści", "testy dźwięku w sklepach stacjonarnych, modele z ANC, treści dla graczy"],
    ["B Praktyczni", "modele do 400 zł z długą baterią, promocje w sklepach internetowych"],
    ["C Wierni marce", "program lojalnościowy, nowości dla stałych klientów"],
    ["D Okazjonalni", "zniżki studenckie, promocje np. w Black Friday"],
  ], { x: 5.1, y: 1.3, w: 4.4, colW: [1.35, 3.05], font: 11, rowH: 0.6, pierwszaPogrubiona: true, wyroznij: [0], nazwa: "tabela_dzialania" });
  wniosek(s, "główna kampania dla segmentu A, a dla B, C i D tańsze działania dopasowane do ich potrzeb.");

  // ===================================================================================
  // PROJEKT 4 – SCV
  // ===================================================================================
  sekcja("04", "Analiza SCV (Single Customer View)", "Marka odzieżowa: cztery posty w social mediach i wspólny profil klienta",
    "PROJEKT 4 · ANALIZA SCV", "profil",
    "Projekt 4. Marka odzieżowa planuje kampanię, której celem jest sprzedaż w sklepie internetowym. Mieliśmy porównać wyniki czterech postów i zaproponować, jakie dane pomogłyby zbudować wspólny profil klienta, czyli SCV.");

  s = slajd("Dane i wzory", "Dane z karty zadania: cztery posty, trzy na Instagramie i jeden na Facebooku. Przyjmujemy, że zakupy przypisano do kliknięć w link danego posta. Liczymy dwa wskaźniki. Zaangażowanie to suma polubień, komentarzy i udostępnień podzielona przez wyświetlenia. Konwersja po kliknięciu to zakupy podzielone przez kliknięcia w link. Wyniki w procentach z jednym miejscem po przecinku.");
  tabela(s, ["Post", "Kanał i format", "Wyświetl.", "Polub.", "Koment.", "Udost.", "Kliknięcia", "Zakupy"], [
    ["104", "Instagram, Stories", "15 000", "900", "50", "5", "600", "3"],
    ["108", "Instagram, Reels", "14 000", "2 800", "350", "60", "470", "9"],
    ["109", "Facebook, post promocyjny", "30 000", "2 100", "400", "100", "900", "20"],
    ["110", "Instagram, post z obrazem", "21 000", "3 000", "420", "75", "500", "16"],
  ], { x: 0.5, y: 1.3, w: 9, colW: [0.6, 2.4, 1.0, 0.9, 0.95, 0.85, 1.25, 1.05], prawo: [false, false, true, true, true, true, true, true], font: 12, rowH: 0.38, pierwszaPogrubiona: true, nazwa: "tabela_posty" });
  karta(s, 0.5, 3.35, 4.4, 1.05, "wzor_zaang");
  karta(s, 5.1, 3.35, 4.4, 1.05, "wzor_konw");
  tekst(s, [B("Zaangażowanie", true), T("(polubienia + komentarze + udostępnienia) ÷ wyświetlenia × 100%")], 0.75, 3.47, 3.95, 0.85, { fontSize: 13, paraSpaceAfter: 2 });
  tekst(s, [B("Konwersja po kliknięciu", true), T("zakupy ÷ kliknięcia w link × 100%")], 5.35, 3.47, 3.95, 0.85, { fontSize: 13, paraSpaceAfter: 2 });
  wniosek(s, "zakupy są przypisane do kliknięć w link danego posta, więc możemy policzyć konwersję.");

  s = slajd("Wyniki: zaangażowanie nie oznacza sprzedaży", "Wyniki w tabeli. Przykład dla posta 110: 3000 plus 420 plus 75 to 3495 reakcji, podzielone przez 21 000 wyświetleń daje 16,6% zaangażowania, a 16 zakupów na 500 kliknięć to 3,2% konwersji. Najwyższe zaangażowanie ma Reels, post 108: 22,9%. Najwyższą konwersję ma post z obrazem, 110: 3,2%. Najsłabiej wypada Stories, post 104: 0,5% konwersji.");
  tabela(s, ["Post", "Zaangażowanie", "Konwersja"], [
    ["104 Stories", "6,4%", "0,5%"],
    ["108 Reels", "22,9%", "1,9%"],
    ["109 Post promocyjny", "8,7%", "2,2%"],
    ["110 Post z obrazem", "16,6%", "3,2%"],
  ], { x: 0.5, y: 1.3, w: 3.75, colW: [1.55, 1.3, 0.9], prawo: [false, true, true], font: 11, rowH: 0.4, pierwszaPogrubiona: true, wyroznij: [3], pogrubWyroznione: true, nazwa: "tabela_wskazniki" });
  tekst(s, [B("Przykład (post 110)", true), T("(3 000 + 420 + 75) ÷ 21 000 = 16,6%", true), T("16 ÷ 500 = 3,2%")], 0.5, 3.4, 3.75, 1.0, { fontSize: 12, color: C.text2, paraSpaceAfter: 2 });
  const posty = ["104 Stories", "108 Reels", "109 Promocyjny", "110 Z obrazem"];
  slupki(s, posty, [6.4, 22.9, 8.7, 16.6], { x: 4.4, y: 1.2, w: 2.65, h: 3.25, tytul: "Zaangażowanie (%)", max: 30, format: "0.0", nazwa: "wykres_zaang" });
  slupki(s, posty, [0.5, 1.9, 2.2, 3.2], { x: 7.05, y: 1.2, w: 2.65, h: 3.25, tytul: "Konwersja (%)", max: 4.2, format: "0.0", kolor: HEX.accent2, nazwa: "wykres_konw" });
  wniosek(s, "post, na który ludzie najchętniej reagują (108), nie sprzedaje najlepiej (110).");

  s = slajd("Który post wybieramy do testu", "Odpowiedzi na punkt 3. Najwyższe zaangażowanie: Reels 108, 22,9%, ale z 470 kliknięć tylko 9 zakupów, więc to raczej treść budująca markę. Najwyższa konwersja: post z obrazem 110, 3,2%, 16 zakupów z 500 kliknięć. Dużo kliknięć i mało zakupów: Stories 104, 600 kliknięć, a tylko 3 zakupy. Do testu sprzedażowego wybieramy post 110 i uzasadniamy dwiema liczbami: najwyższa konwersja 3,2% i drugie najwyższe zaangażowanie 16,6%. Jako jedyny wypada dobrze w obu wskaźnikach.\n\nJeśli zapyta o post 109: dał najwięcej zakupów, 20, ale miał największy zasięg, 30 000 wyświetleń, i sprzedaż napędzały rabaty, które obniżają marżę. Jego konwersja to 2,2%, niższa niż 3,2% posta 110.");
  [
    ["108 Reels", "22,9%", "najwyższe zaangażowanie, ale z 470 kliknięć tylko 9 zakupów", C.accent1],
    ["110 Post z obrazem", "3,2%", "najwyższa konwersja: 16 zakupów z 500 kliknięć", C.accent1],
    ["104 Stories", "0,5%", "dużo kliknięć (600), a tylko 3 zakupy", C.accent5],
  ].forEach(([n, w, o, kol], i) => {
    const x = 0.5 + i * 3.07;
    karta(s, x, 1.3, 2.87, 1.8, `post_${i}`);
    s.addText(n, { x: x + 0.25, y: 1.42, w: 2.4, h: 0.32, isTextBox: true, fontSize: 14, bold: true, color: C.text1, margin: 0 });
    s.addText(w, { x: x + 0.25, y: 1.74, w: 2.4, h: 0.55, isTextBox: true, fontSize: 28, bold: true, color: kol, margin: 0 });
    s.addText(o, { x: x + 0.25, y: 2.33, w: 2.4, h: 0.7, isTextBox: true, fontSize: 12, color: C.text2, margin: 0, valign: "top" });
  });
  karta(s, 0.5, 3.25, 9, 1.15, "decyzja", C.accent1);
  ikonaKolo(s, 0.72, 3.5, 0.65, "cel", C.accent2);
  tekst(s, [{ text: "Do testu sprzedażowego: post 110 „Inspiracje na zimę”", options: { bold: true, color: C.background1, breakLine: true, fontSize: 16 } },
    { text: "Konwersja 3,2% (najwyższa) i zaangażowanie 16,6% (drugie najwyższe): przyciąga uwagę i sprzedaje.", options: { color: C.background1 } }],
  1.6, 3.37, 7.7, 0.95, { fontSize: 13, valign: "middle", paraSpaceAfter: 2 });
  wniosek(s, "post 109 ma najwięcej zakupów (20), ale przy największym zasięgu (30 000) i dzięki rabatom.");

  s = slajd("Jak zbudować wspólny profil klienta (SCV)", "Punkt 4. Wyniki postów są zbiorcze: wiemy, ile osób kliknęło i kupiło, ale nie wiemy, kim one są. SCV, Single Customer View, to jeden profil klienta, który łączy dane z różnych źródeł. Na wykładzie nazywało się to integracją danych, np. połączenie danych z CRM z danymi transakcyjnymi. Bierzemy dwa źródła danych własnych marki: sklep internetowy, czyli konto, zamówienia, rozmiary, porzucone koszyki i wejścia z linków UTM, oraz newsletter i program lojalnościowy. Wspólny identyfikator to adres e-mail, bo klient podaje go i w sklepie, i przy zapisie do newslettera. Przykład: klient klika post 110, dodaje płaszcz do koszyka, ale nie kupuje. Z profilu wiemy, że jest w newsletterze i nosi rozmiar M, więc następnego dnia wysyłamy mu mail z przypomnieniem i stylizacjami w jego rozmiarze, bez płacenia drugi raz za reklamę. Ważne: na takie maile potrzebna jest zgoda marketingowa, zgodnie z RODO.");
  karta(s, 0.5, 1.3, 4.4, 1.2, "zrodlo_sklep");
  karta(s, 5.1, 1.3, 4.4, 1.2, "zrodlo_crm");
  ikonaKolo(s, 0.68, 1.48, 0.5, "sklep", C.accent1);
  ikonaKolo(s, 5.28, 1.48, 0.5, "mail", C.accent1);
  tekst(s, [B("Źródło 1: sklep internetowy", true), T("konto, zamówienia, rozmiary, porzucone koszyki, wejścia z linków UTM")], 1.35, 1.42, 3.4, 1.0, { fontSize: 12, paraSpaceAfter: 2 });
  tekst(s, [B("Źródło 2: newsletter i program lojalnościowy", true), T("zapisy, otwarcia i kliknięcia w maile, zgody, punkty")], 5.95, 1.42, 3.4, 1.0, { fontSize: 12, paraSpaceAfter: 2 });
  ikonaKolo(s, 0.5, 2.63, 0.42, "link", C.accent2);
  s.addText([B("Wspólny identyfikator: "), { text: "adres e-mail klienta (za zgodą marketingową)", options: {} }], { x: 1.05, y: 2.62, w: 8.4, h: 0.44, isTextBox: true, fontSize: 15, color: C.accent1, margin: 0, valign: "middle" });
  const kroki = [["klik", "Klik w post 110 (link z UTM)"], ["koszyk", "Płaszcz w koszyku, brak zakupu"], ["profil", "Profil po e-mailu: newsletter, rozmiar M"], ["wyslij", "Mail z przypomnieniem i stylizacjami"]];
  kroki.forEach(([ik, k], i) => {
    const x = 0.5 + i * 2.33;
    const ostatni = i === 3;
    karta(s, x, 3.2, 1.95, 1.2, `krok_${i}`, ostatni ? C.accent2 : C.background2);
    ikonaKolo(s, x + 0.12, 3.32, 0.4, ik, ostatni ? C.text1 : C.accent1);
    s.addText(k, { x: x + 0.12, y: 3.76, w: 1.72, h: 0.58, isTextBox: true, fontSize: 11, bold: true, color: ostatni ? C.background1 : C.text1, valign: "top", margin: 0 });
    if (i < 3) strzalka(s, x + 1.98, 3.8, x + 2.3, 3.8);
  });
  wniosek(s, "SCV łączy dane ze sklepu i newslettera w jeden profil klienta, a wspólnym kluczem jest e-mail.");

  // ===================================================================================
  // Podsumowanie
  // ===================================================================================
  sekcjaAkt = "Podsumowanie";
  kickerAkt = "PODSUMOWANIE";
  pres.addSection({ title: "Podsumowanie" });
  s = slajd("Najważniejsze wnioski", "Na koniec w skrócie. Projekt 2: korelacja 0,999 nie dowodzi, że reklama zwiększa sprzedaż, bo razem z reklamą rosną popyt i sezon. Decyzję o 30% większym budżecie trzeba oprzeć na geo-eksperymencie. Projekt 3: segment docelowy to Entuzjaści dźwięku, 52% respondentów, długo używają słuchawek i mają najwyższy budżet. Projekt 4: do testu wybieramy post 110, a wspólny profil klienta budujemy z danych sklepu i newslettera połączonych adresem e-mail. Wspólna lekcja ze wszystkich projektów: średnie i korelacje mogą mylić, dlatego patrzymy na segmenty i przyczyny.");
  [["dag", "Projekt 2", "Korelacja i DAG", "Korelacja 0,999 nie dowodzi efektu reklamy. O\u00A0+30% budżetu powinien zdecydować eksperyment z\u00A0grupą kontrolną."],
    ["sluchawki", "Projekt 3", "Segmentacja", "Segment docelowy: Entuzjaści dźwięku. 52% osób, 10 z 14 ponad 3\u00A0h dziennie, budżet ponad 400\u00A0zł."],
    ["profil", "Projekt 4", "Analiza SCV", "Do testu post 110: konwersja 3,2%, zaangażowanie 16,6%. SCV łączy sklep i newsletter po adresie e\u2011mail."]].forEach(([ik, nr, nazwa, txt], i) => {
    const x = 0.5 + i * 3.07;
    karta(s, x, 1.3, 2.87, 3.1, `wniosek_${i}`);
    ikonaKolo(s, x + 0.25, 1.5, 0.7, ik, C.accent1);
    s.addText(nr.toUpperCase(), { x: x + 1.1, y: 1.55, w: 1.6, h: 0.3, isTextBox: true, fontSize: 11, bold: true, color: C.accent1, margin: 0, charSpacing: 1 });
    s.addText(nazwa, { x: x + 1.1, y: 1.83, w: 1.65, h: 0.4, isTextBox: true, fontSize: 14, bold: true, color: C.text1, margin: 0, valign: "top" });
    s.addText(txt, { x: x + 0.25, y: 2.45, w: 2.4, h: 1.85, isTextBox: true, fontSize: 14, color: C.text1, margin: 0, valign: "top" });
  });
  wniosek(s, "średnie i korelacje mogą mylić, dlatego zawsze patrzymy na segmenty i przyczyny.");

  s = pres.addSlide({ masterName: "TYTUL", sectionTitle: "Podsumowanie" });
  s.addText("PODSTAWY ANALITYKI INTERNETOWEJ", { placeholder: "kicker" });
  s.addText("Dziękujemy", { placeholder: "title" });
  s.addText("Chętnie odpowiemy na pytania", { placeholder: "body" });
  ikonaKolo(s, 0.6, 3.75, 0.62, "pytanie", C.accent2);
  s.addText("W dodatku: pojęcia z zajęć i odpowiedzi na najczęstsze pytania do projektów 2, 3 i 4.", { x: 1.38, y: 3.75, w: 7.5, h: 0.62, isTextBox: true, fontSize: 13, color: C.accent4, valign: "middle", margin: 0 });
  s.addNotes("Dziękujemy za uwagę. Jeśli będą pytania, odpowiedzi na najczęstsze są w dodatku na kolejnych slajdach.");

  // ===================================================================================
  // Dodatek
  // ===================================================================================
  sekcjaAkt = "Dodatek";
  kickerAkt = "DODATEK · POJĘCIA Z ZAJĘĆ";
  pres.addSection({ title: "Dodatek" });
  s = slajd("Pojęcia: korelacja i przyczynowość (projekt 2)", "Ściąga z pojęć do projektu 2, zgodnie z wykładem. Gdy padnie pytanie o definicję, wystarczy przeczytać środkową kolumnę i podać przykład z prawej.");
  tabela(s, ["Pojęcie", "Co znaczy", "U nas"], [
    ["Korelacja Pearsona (r)", "siła i kierunek związku liniowego, od −1 do 1; nie mówi, co jest przyczyną", "reklama–sprzedaż: 0,999"],
    ["R²", "jaką część zmienności wyjaśnia linia regresji", "0,998, czyli 99,8%"],
    ["DAG", "graf skierowany bez cykli: model naszych założeń o tym, co na co wpływa, a nie dowód", "7 zmiennych, strzałki od przyczyny do skutku"],
    ["Zmienna zakłócająca", "wpływa i na badaną przyczynę, i na wynik (confounder)", "popyt, sezonowość, rabaty, konkurencja"],
    ["Zmienna pośrednicząca", "leży na ścieżce od przyczyny do wyniku i wyjaśnia mechanizm (mediator)", "ruch na stronie"],
    ["Zmienna wynikowa", "to, co chcemy wyjaśnić (outcome)", "sprzedaż"],
    ["Endogeniczność", "badana zmienna nie jest niezależna od czynników, które wpływają na wynik", "budżet reklamy zależy od popytu i sezonu"],
  ], { x: 0.5, y: 1.3, w: 9, colW: [2.0, 4.3, 2.7], font: 12, rowH: 0.48, pierwszaPogrubiona: true, nazwa: "tabela_pojecia1" });

  s = slajd("Pojęcia: segmentacja i SCV (projekty 3 i 4)", "Ściąga z pojęć do projektów 3 i 4. Segmentacja i jej rodzaje są z wykładu o segmentacji klientów, a integracja danych z wykładu o zarządzaniu danymi.");
  tabela(s, ["Pojęcie", "Co znaczy", "U nas"], [
    ["Segmentacja", "podział klientów na grupy o podobnych cechach, potrzebach lub zachowaniach", "4 segmenty z ankiety"],
    ["Rodzaje segmentacji", "demograficzna, geograficzna, behawioralna, wartości klienta, psychograficzna", "4 z nich + preferencje"],
    ["Segment docelowy", "grupa, do której firma kieruje ofertę w pierwszej kolejności", "A Entuzjaści dźwięku"],
    ["Pozycjonowanie", "jak chcemy, żeby klienci postrzegali produkt", "mocny dźwięk i wygoda na wiele godzin"],
    ["SCV (Single Customer View)", "jeden profil klienta, który łączy dane z różnych źródeł (integracja danych)", "sklep + newsletter, klucz: e-mail"],
    ["Dane własne", "dane, które marka zbiera sama, za zgodą klienta (first-party)", "konto, zamówienia, newsletter"],
    ["UTM", "znaczniki w linku, które pokazują, z jakiego posta lub kampanii ktoś przyszedł", "link w poście 110"],
  ], { x: 0.5, y: 1.3, w: 9, colW: [2.2, 4.3, 2.5], font: 12, rowH: 0.48, pierwszaPogrubiona: true, nazwa: "tabela_pojecia2" });

  kickerAkt = "DODATEK · PYTANIA DO OBRONY";
  function pytania(tytul, notatka, lista) {
    const sl = slajd(tytul, notatka);
    lista.forEach(([q, a], i) => {
      const x = 0.5 + (i % 2) * 4.6, y = 1.3 + Math.floor(i / 2) * 1.95;
      karta(sl, x, y, 4.4, 1.8, `qa_${i}`);
      ikonaKolo(sl, x + 0.2, y + 0.2, 0.42, "pytanie", C.accent2);
      sl.addText(q, { x: x + 0.78, y: y + 0.14, w: 3.45, h: 0.55, isTextBox: true, fontSize: 13, bold: true, color: C.text1, margin: 0, valign: "middle" });
      sl.addText(a, { x: x + 0.25, y: y + 0.78, w: 3.95, h: 0.95, isTextBox: true, fontSize: 12, color: C.text2, margin: 0, valign: "top" });
    });
  }
  pytania("Pytania do projektu 2", "Najczęstsze pytania do projektu 2 i krótkie odpowiedzi do przeczytania.", [
    ["Czy r = 0,999 oznacza, że reklama zwiększa sprzedaż?", "Nie. To tylko siła związku. Popyt ze sprzedażą też ma 0,999, a z reklamą 0,998, więc nie wiemy, co naprawdę napędza sprzedaż."],
    ["Dlaczego nie dodajecie ruchu na stronie do regresji?", "Bo to mediator: przez niego reklama działa na sprzedaż. Po dodaniu go zmierzylibyśmy tylko część efektu, a chcemy efekt całkowity."],
    ["Co to jest endogeniczność w tym zadaniu?", "Budżet reklamy nie jest niezależny: firma daje więcej reklam, gdy spodziewa się większej sprzedaży. Do tego pominięte zmienne i błąd pomiaru."],
    ["Jak w praktyce zmierzyć efekt reklamy?", "Geo-eksperyment: +30% budżetu tylko w losowo wybranych regionach i porównanie sprzedaży z regionami kontrolnymi."],
  ]);
  pytania("Pytania do projektu 3", "Najczęstsze pytania do projektu 3 i krótkie odpowiedzi do przeczytania.", [
    ["Dlaczego dzielicie według czasu i budżetu?", "Te dwa kryteria najbardziej różnicowały odpowiedzi. Łączą zachowanie z wartością klienta, jak przykład klientów A–D z wykładu."],
    ["Dlaczego segment A jest docelowy?", "To 14 z 27 osób (52%), 10 z 14 używa słuchawek ponad 3 godziny dziennie, a wszyscy wydadzą ponad 400 zł."],
    ["Dlaczego z 32 odpowiedzi zostało 27?", "6 identycznych odpowiedzi przyszło 9.11 w ciągu 30 sekund. To jedna osoba, więc liczymy ją raz i usuwamy 5 powtórzeń."],
    ["Jakie są ograniczenia badania?", "Mała próba (27 osób), brak pytań o płeć i miejsce zamieszkania, a segmenty C i D są małe. Wyniki trzeba sprawdzić na większej grupie."],
  ]);
  pytania("Pytania do projektu 4", "Najczęstsze pytania do projektu 4 i krótkie odpowiedzi do przeczytania.", [
    ["Dlaczego post 110, a nie 109 z największą liczbą zakupów?", "109 miał największy zasięg (30 000) i rabaty. Post 110 ma wyższą konwersję (3,2% wobec 2,2%) i prawie 2 razy większe zaangażowanie."],
    ["Dlaczego Stories ma tak niską konwersję?", "Link łatwo kliknąć przy szybkim przeglądaniu (600 kliknięć), ale kupiły tylko 3 osoby. Strona po kliknięciu mogła nie pokazywać tego, co w relacji."],
    ["Czym jest SCV i po co je budować?", "To jeden profil klienta z danych z różnych źródeł. Wiemy, kim jest klient i co robił, więc możemy do niego wrócić bez płacenia znowu za reklamę."],
    ["Dlaczego e-mail jako identyfikator?", "Klient podaje go i w sklepie, i przy zapisie do newslettera. Jest unikalny i się nie zmienia. Potrzebna jest zgoda marketingowa (RODO)."],
  ]);

  return pres;
}

(async () => {
  await przygotujIkony();
  const pres = zbuduj();
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("Zapisano", OUT);
})();
