# Ćwiczenie I: filtrowanie danych, rozwiązanie

**Sytuacja:** sklep internetowy pozyskuje ruch z Google Ads, Social Media i wyszukiwarki (Organic Search). Zarząd chce więcej zakupów i lepszej efektywności marketingu. Cel: wskazać grupy klientów, które wymagają uwagi, i działania do przetestowania.

**Dane:** 15 użytkowników, 8 zakupów, przychód 935 €. Konwersja ogółem 53,3%, średnia wartość koszyka kupujących 116,88 €, średni wskaźnik odrzuceń 52,3%.

> **Definicje użyte w obliczeniach**
> - *Konwersja* = liczba zakupów / liczba użytkowników w grupie.
> - *Średnia wartość koszyka (AOV)* = przychód / liczba zakupów, czyli liczona tylko dla kupujących.
> - *Przychód na użytkownika (RPU)* = przychód / wszyscy użytkownicy w grupie. Wychodzi to samo co średnia „wartość koszyka" z zerami włącznie, a RPU = konwersja × AOV.

> **Ważne zastrzeżenia (warto je powiedzieć na forum)**
> 1. Próba jest bardzo mała: w grupie 5 osób jeden użytkownik zmienia konwersję o 20 p.p. Wyniki to **hipotezy do weryfikacji** na pełnych danych (np. GA4), a nie pewne wnioski.
> 2. Najsilniej różnicuje dane typ użytkownika: **wszyscy 7 powracających kupiło, z 8 nowych kupił tylko 1**. Kanały i urządzenia mają różny udział nowych i powracających, więc część różnic między nimi wynika z tego **miksu**, a nie z samego kanału czy urządzenia. Dlatego niżej są też filtry krzyżowe.

---

## 1. Pytania biznesowe przed rozpoczęciem analizy

**Cel i definicje**
- Co jest celem nadrzędnym: liczba transakcji, przychód czy marża? W jakim horyzoncie czasowym i z jakim celem liczbowym (np. konwersja +X p.p.)?
- Co liczymy jako konwersję: tylko zakup, czy też mikrokonwersje (dodanie do koszyka, rejestracja, zapis do newslettera)?

**Koszty i efektywność**
- Ile wydajemy na Google Ads i Social Media? Bez kosztów nie policzymy ROAS ani kosztu pozyskania klienta (CAC), a to one mierzą *efektywność* marketingu.
- Jakie kampanie działają w Google Ads (brand, remarketing, prospecting)? Czy płacimy za użytkowników, którzy i tak by wrócili?

**Klienci i lejek**
- Jak przypisujemy zakup do kanału (model atrybucji)? Powracający użytkownik mógł pierwszy raz trafić na stronę z innego źródła.
- Na którym etapie lejka odpadają użytkownicy: strona wejścia, karta produktu, koszyk czy płatność?
- Jaka jest wartość klienta w czasie (LTV) i jak często nowi klienci wracają?

**Kontekst i jakość danych**
- Czy oferta różni się między krajami (ceny, waluta, koszty i czas dostawy, wersja językowa)?
- Czy strona jest zoptymalizowana pod mobile (czas ładowania, płatności mobilne)?
- Z jakiego okresu są dane, czy są reprezentatywne (sezonowość, promocje) i jak zdefiniowano wskaźnik odrzuceń?

## 2. Wybór KPI

Spośród dostępnych wskaźników wybieramy trzy:

| # | KPI | Dlaczego |
|---|-----|----------|
| 1 | **Współczynnik konwersji** | Bezpośrednio mierzy cel zarządu („zwiększyć liczbę zakupów"). Pozwala porównać jakość ruchu z różnych kanałów, urządzeń i krajów. |
| 2 | **Średnia wartość koszyka (AOV)** | Pokazuje, *ile* warty jest zakup. Przychód = użytkownicy × konwersja × AOV, więc razem z konwersją wyjaśnia przychód. Dla porównań między kanałami używamy też RPU (konwersja × AOV w jednej liczbie). |
| 3 | **Wskaźnik odrzuceń** | Wczesny sygnał jakości ruchu i dopasowania strony docelowej, mierzony już przy pierwszej wizycie. Najmocniej z danych zachowania wiąże się z zakupem (korelacja r = −0,81). Daje się bezpośrednio poprawiać przez optymalizację stron wejścia. |

**Dlaczego nie czas na stronie ani liczba stron?** W tych danych są niemal tym samym sygnałem co wskaźnik odrzuceń: korelacja z odrzuceniami wynosi −0,99 dla czasu i −0,97 dla liczby stron. Dodanie ich nic nie wnosi, więc wystarczy jeden reprezentant zaangażowania.

**Docelowo** do oceny *efektywności* warto dodać **ROAS / CAC** (wymaga danych o kosztach, których w zestawie nie ma).

---

## 3. Filtr: źródło ruchu

| Źródło | Użytk. | Zakupy | Konwersja | Odrzucenia | Czas (min) | Strony | Przychód | AOV | RPU |
|---|---|---|---|---|---|---|---|---|---|
| Google Ads | 5 | **4** | **80%** | **39%** | **3,48** | **7,6** | **500 € (53%)** | **125 €** | **100 €** |
| Social Media | 5 | 3 | 60% | 52% | 2,60 | 6,0 | 360 € (39%) | 120 € | 72 € |
| Organic Search | 5 | 1 | 20% | 66% | 1,96 | 4,0 | 75 € (8%) | 75 € | 15 € |

- **Porównanie konwersji i wartości koszyka:** Google Ads ma najwyższą konwersję (80%) i najwyższy AOV (125 €). Social Media jest blisko pod względem koszyka (120 €), ale konwertuje słabiej (60%). Organic Search wypada najsłabiej (20% konwersji, 75 €).
- **Najwięcej transakcji:** **Google Ads**, 4 z 8 zakupów i 53% przychodu.
- **Google Ads a Organic Search:** tak, różnica jest wyraźna. Użytkownicy z Google Ads rzadziej odrzucają stronę (39% wobec 66%), są na niej prawie 2 razy dłużej (3,5 wobec 2,0 min), oglądają 2 razy więcej stron (7,6 wobec 4,0) i kupują 4 razy częściej.
  **Ale:** z Google Ads przyszło 4 powracających i 1 nowy, a z Organic odwrotnie, 4 nowych i 1 powracający. W tych samych grupach wyniki są zbliżone: nowi mają 0% konwersji w obu kanałach (odrzucenia 75% i 71%), powracający 100% w obu. Różnica wynika więc głównie z tego, **kogo** kanał przyprowadza. Google Ads w dużej mierze „łapie" powracających (prawdopodobnie kampanie brandowe lub remarketing). Organic jest głównym źródłem nowych użytkowników, którzy nie kupują.

## 4. Filtr: urządzenie

| Urządzenie | Użytk. | Zakupy | Konwersja | Odrzucenia | Czas (min) | Strony | Przychód | AOV | RPU |
|---|---|---|---|---|---|---|---|---|---|
| Desktop | 5 | **4** | **80%** | **29%** | **3,98** | **9,4** | **650 € (70%)** | **162,50 €** | **130 €** |
| Mobile | 7 | 3 | 43% | **67%** | 1,87 | 3,9 | 200 € (21%) | 66,67 € | 28,57 € |
| Tablet | 3 | 1 | 33% | 57% | 2,40 | 4,7 | 85 € (9%) | 85,00 € | 28,33 € |

- **Najwięcej zakupów:** **Desktop**, 4 z 8 zakupów i 70% przychodu, choć to tylko 1/3 ruchu.
- **Wskaźnik odrzuceń:** Desktop 29%, Tablet 57%, Mobile 67%. **Optymalizacji wymaga Mobile.** To największa grupa (47% ruchu), ma najwyższe odrzucenia, najkrótszy czas na stronie i najniższy koszyk. Najgorzej wypadają **nowi użytkownicy na mobile**: 5 osób (1/3 całego ruchu), 79% odrzuceń, 1,1 min na stronie i tylko 1 zakup za 30 €. Tablet też wypada słabo, ale przy 3 użytkownikach to tylko sygnał do obserwacji.
- **Najwyższa wartość koszyka:** **Desktop**, 162,50 €. Nawet powracający kupują na mobile mniej niż na desktopie (85 € wobec 162,50 €), więc na mobile warto też zwiększać wartość koszyka.

## 5. Filtr: kraj

| Kraj | Użytk. | Zakupy | Konwersja | Odrzucenia | Czas (min) | Przychód | AOV | RPU |
|---|---|---|---|---|---|---|---|---|
| Niemcy | 5 | **4** | **80%** | **50,0%** | **2,90** | 350 € (37%) | 87,50 € | **70 €** |
| Polska | 6 | 3 | 50% | 50,8% | 2,58 | **385 € (41%)** | 128,33 € | 64 € |
| Wielka Brytania | 4 | 1 | 25% | 57,5% | 2,55 | 200 € (21%) | **200,00 €** | 50 € |

- **Konwersja i koszyk:** najlepiej konwertują Niemcy (80%), ale z najniższym koszykiem (87,50 €). Polska daje najwięcej przychodu (385 €) przy średniej konwersji. Wielka Brytania ma najwyższy koszyk (200 €), ale to jeden zakup i najniższa konwersja (25%).
- **Najniższy wskaźnik odrzuceń:** **Niemcy (50,0%) i Polska (50,8%)**, praktycznie na równi. Najwyższy ma Wielka Brytania (57,5%).
- **Nowi a powracający w krajach:**

| Kraj | Konwersja: nowi | Konwersja: powracający | Odrzucenia: nowi / powracający | Czas: nowi / powracający | Przychód: nowi / powracający |
|---|---|---|---|---|---|
| Niemcy | 1 z 2 (50%) | 3 z 3 (100%) | 78% / 32% | 1,2 / 4,0 min | 30 € / 320 € |
| Polska | 0 z 3 (0%) | 3 z 3 (100%) | 68% / 33% | 1,6 / 3,6 min | 0 € / 385 € |
| Wielka Brytania | 0 z 3 (0%) | 1 z 1 (100%) | 72% / 15% | 1,7 / 5,2 min | 0 € / 200 € |

  We wszystkich krajach powracający kupują zawsze, a nowi prawie nigdy. Jedyny zakup nowego użytkownika w danych to Niemcy, Social Media, mobile, 30 €. W Wielkiej Brytanii luka jest największa: powracający klient ma 15% odrzuceń i koszyk 200 €, a 3 nowych nie kupiło nic. To sygnał do sprawdzenia lokalizacji oferty (waluta, dostawa, cła, zwroty).

## 6. Nowi a powracający użytkownicy

| Typ | Użytk. | Zakupy | Konwersja | Odrzucenia | Czas (min) | Strony | Przychód | AOV |
|---|---|---|---|---|---|---|---|---|
| Nowy | 8 | 1 | 12,5% | 72% | 1,51 | 3,1 | 30 € (3%) | 30 € |
| Powracający | 7 | 7 | **100%** | **30%** | **4,01** | **9,0** | **905 € (97%)** | **129,29 €** |

- **Czy powracający konwertują lepiej?** Tak, zdecydowanie: 100% wobec 12,5%. Generują też 97% przychodu.
- **Kto spędza więcej czasu na stronie?** Powracający: 4,0 wobec 1,5 min (ok. 2,7 razy dłużej) i 9 stron wobec 3.
- **Strategie zachęcające nowych do częstszych zakupów:**
  - **Oferta na pierwsze zamówienie:** rabat lub darmowa dostawa w zamian za zapis do newslettera. Zbieramy przy tym kontakt do dalszej komunikacji.
  - **Remarketing do nowych, którzy nie kupili** (Google Ads i social, reklamy dynamiczne z oglądanymi produktami). Powracający kupują w 100%, więc **zamiana nowego w powracającego** to najmocniejsza dźwignia.
  - **Sekwencja e-maili powitalnych** po zapisie lub porzuceniu koszyka.
  - **Budowanie zaufania na stronie wejścia:** opinie i oceny, jasne zasady dostawy i zwrotów, widoczne CTA.
  - **Krótszy zakup:** zakup bez rejestracji, szybkie płatności mobilne (BLIK, Apple Pay, Google Pay).
  - **Dopasowanie strony docelowej do reklamy lub zapytania**, żeby użytkownik od razu widział to, czego szukał.

---

## 7. Podsumowanie: jak zoptymalizować działania marketingowe

**Grupy klientów wymagające uwagi (wg priorytetu):**
1. **Nowi użytkownicy na mobile:** 1/3 ruchu, 79% odrzuceń, prawie zero sprzedaży. Tu jest największa strata.
2. **Nowi z Organic Search:** 4 osoby, 0 zakupów, 71% odrzuceń. Ruch jest „darmowy", więc poprawa konwersji ma tu najwyższy zwrot.
3. **Nowi z Wielkiej Brytanii:** 0 z 3 zakupów przy najwyższym potencjale koszyka (200 €).
4. **Powracający na mobile:** kupują, ale za mniej niż połowę wartości koszyka z desktopu.

**Działania do przetestowania (hipoteza → test → KPI):**

| # | Działanie | Test | KPI |
|---|---|---|---|
| 1 | Szybsza, uproszczona strona wejścia i checkout na mobile, płatności mobilne | A/B: obecna vs nowa wersja mobile | odrzucenia i konwersja na mobile |
| 2 | Oferta powitalna dla nowych (np. −10% lub darmowa dostawa na 1. zamówienie) | A/B: z ofertą vs bez (pilnując marży) | konwersja nowych, AOV |
| 3 | Remarketing do nowych, którzy nie kupili | grupa objęta remarketingiem vs grupa kontrolna | odsetek powrotów, konwersja |
| 4 | Optymalizacja stron wejścia z SEO pod intencję wyszukiwania (treść, CTA, polecane produkty) | A/B na najczęstszych stronach wejścia | odrzucenia i konwersja Organic |
| 5 | Rozbicie Google Ads na brand, remarketing i prospecting, policzenie ROAS. Test przesunięcia części budżetu na pozyskanie nowych | porównanie kampanii | ROAS, CAC, konwersja nowych |
| 6 | Social Media jako kanał pozyskania (najwyższy koszyk powracających: 165 €, jedyna konwersja nowego użytkownika) | kampania lookalike / reklamy produktowe | konwersja nowych, RPU |
| 7 | Lokalizacja dla UK: ceny w GBP, jasna informacja o dostawie, cłach i zwrotach | A/B na ruchu z UK | konwersja nowych z UK |
| 8 | Cross-sell, up-sell i próg darmowej dostawy na mobile | A/B rekomendacji w koszyku | AOV na mobile |
| 9 | Program lojalnościowy i e-mail dla powracających (97% przychodu) | kohorty z programem vs bez | częstotliwość zakupów, LTV |

**Wniosek w jednym zdaniu:** sklep dobrze sprzedaje klientom, którzy już go znają (zwłaszcza na desktopie i z Google Ads), ale prawie nie konwertuje nowych użytkowników, szczególnie na mobile i z wyszukiwarki. Budżet i testy warto skierować na **pierwszy kontakt nowego klienta** oraz na **sprowadzenie go z powrotem**, a efektywność kanałów ocenić po dołożeniu danych o kosztach (ROAS, CAC).

---

## Pliki

- `dane.csv`: dane z karty warsztatu.
- `warsztat_1.xlsx`: arkusz **Dane** z filtrami w nagłówkach (do pokazania filtrowania na żywo) oraz arkusz **Podsumowanie** z formułami `COUNTIFS`, `AVERAGEIFS` i `SUMIFS` dla każdego filtra i filtrów krzyżowych z typem użytkownika. Zmiana danych przelicza podsumowanie.
- `analiza.py`: skrypt odtwarzający wszystkie zestawienia i budujący plik Excel (`python3 analiza.py`).
