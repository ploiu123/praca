"""Ankieta „Opinia klienta” - wykresy do raportu.

Uruchomienie:  python3 wykresy.py
Wczytuje odpowiedzi.csv i zapisuje wykresy PNG do katalogu wykresy/.
"""
from pathlib import Path

import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle

KAT = Path(__file__).resolve().parent
WYK = KAT / "wykresy"
WYK.mkdir(exist_ok=True)

NIEBIESKI = "#2a78d6"
POMARANCZ = "#eb6834"
TEKST = "#0b0b0b"
TEKST_2 = "#52514e"
SIATKA = "#e1e0d9"
OS = "#c3c2b7"
SKALA = ["#e34948", "#f2a2a1", "#d6d5cf", "#86b6ef", "#2a78d6"]  # 1 ... 5

plt.rcParams.update({
    "font.family": "Carlito", "font.size": 11, "text.color": TEKST,
    "axes.labelcolor": TEKST_2, "xtick.color": TEKST_2, "ytick.color": TEKST,
})

df = pd.read_csv(KAT / "odpowiedzi.csv")
df.columns = ["ts", "wiek", "zajecie", "czas", "sytuacje", "wazne", "cena", "styl", "marka", "gdzie", "brak"]
N = len(df)


def wielokrotny(kol):
    return df[kol].str.split(";").explode().str.strip().value_counts()


def slupek(ax, y, v, h, kolor):
    if v > 0:
        ax.barh(y, v, height=h, color=kolor, edgecolor="none")


def poziomy(ax, etykiety, wartosci, tytul=None, sortuj=False):
    if sortuj:
        para = sorted(zip(etykiety, wartosci), key=lambda p: p[1], reverse=True)
        etykiety, wartosci = [p[0] for p in para], [p[1] for p in para]
    k = len(etykiety)
    ax.set_xlim(0, N * 1.22)
    ax.set_ylim(k - 0.4, -0.6)
    ax.set_yticks(range(k))
    ax.set_yticklabels(etykiety)
    for i, v in enumerate(wartosci):
        slupek(ax, i, v, 0.56, NIEBIESKI)
        ax.text(v + N * 0.02, i, f"{v} ({v / N:.0%})", va="center", fontsize=10.5, color=TEKST_2)
    ax.set_xticks([])
    for s in ("top", "right", "bottom"):
        ax.spines[s].set_visible(False)
    ax.spines["left"].set_color(OS)
    ax.tick_params(axis="y", length=0, pad=6)
    if tytul:
        ax.set_title(tytul, loc="left", fontsize=12, color=TEKST, pad=8)


def zapisz(fig, nazwa):
    fig.savefig(WYK / nazwa, dpi=200, facecolor="white", bbox_inches="tight")
    plt.close(fig)


def licz(kol, kolejnosc):
    vc = df[kol].value_counts()
    return [int(vc.get(k, 0)) for k in kolejnosc]


# 1. Profil respondentów
fig, (a1, a2) = plt.subplots(1, 2, figsize=(6.6, 2.0))
wiek = ["18–24", "25–34", "45 i więcej"]
poziomy(a1, wiek, licz("wiek", wiek), "Wiek")
zaj = ["Praca na etacie", "Nauka lub studia"]
poziomy(a2, zaj, licz("zajecie", zaj), "Czym się zajmują")
fig.subplots_adjust(wspace=0.75)
zapisz(fig, "1_profil.png")

# 2. Czas w słuchawkach
fig, ax = plt.subplots(figsize=(6.0, 1.6))
czas = ["Do godziny", "1-3 godziny", "Powyżej 3 godzin"]
poziomy(ax, ["Do godziny", "1–3 godziny", "Powyżej 3 godzin"], licz("czas", czas))
zapisz(fig, "2_czas.png")

# 3. Sytuacje użycia
s = wielokrotny("sytuacje")
fig, ax = plt.subplots(figsize=(6.0, 2.3))
poziomy(ax, list(s.index), [int(v) for v in s.values], sortuj=True)
zapisz(fig, "3_sytuacje.png")

# 4. Co najważniejsze przy wyborze
w = wielokrotny("wazne")
fig, ax = plt.subplots(figsize=(6.0, 2.6))
poziomy(ax, list(w.index), [int(v) for v in w.values], sortuj=True)
zapisz(fig, "4_wazne.png")

# 5. Budżet
fig, ax = plt.subplots(figsize=(6.0, 2.0))
cena = ["Do 200 zł", "201-400 zł", "401-700 zł", "Powyżej 700 zł"]
poziomy(ax, ["Do 200 zł", "201–400 zł", "401–700 zł", "Powyżej 700 zł"], licz("cena", cena))
zapisz(fig, "5_cena.png")

# 6. Skale 1-5: styl i zaufanie do marki (rozbieżny wykres skumulowany)
pytania = [("styl", "Słuchawki to element\nmojego stylu"), ("marka", "Ufam tylko sprawdzonym\nmarkom, np. JBL")]
fig, ax = plt.subplots(figsize=(6.6, 1.9))
for i, (kol, etyk) in enumerate(pytania):
    vc = df[kol].value_counts()
    liczby = [int(vc.get(k, 0)) for k in range(1, 6)]
    lewo = -(liczby[0] + liczby[1] + liczby[2] / 2)
    for k, ile in enumerate(liczby):
        if ile:
            ax.barh(i, ile, left=lewo, height=0.55, color=SKALA[k], edgecolor="white", linewidth=2)
            ax.text(lewo + ile / 2, i, str(ile), ha="center", va="center", fontsize=10.5,
                    color="white" if k in (0, 4) else TEKST)
        lewo += ile
    sr = df[kol].mean()
    ax.text(6.9, i, f"średnio {sr:.1f}".replace(".", ","), va="center", fontsize=10.5, color=TEKST_2)
ax.axvline(0, color=OS, lw=1, zorder=0)
ax.set_xlim(-6.5, 8.6)
ax.set_ylim(1.6, -0.6)
ax.set_yticks([0, 1])
ax.set_yticklabels([p[1] for p in pytania])
ax.set_xticks([])
for s_ in ("top", "right", "bottom", "left"):
    ax.spines[s_].set_visible(False)
ax.tick_params(axis="y", length=0, pad=6)
uchwyty = [Rectangle((0, 0), 1, 1, fc=c) for c in SKALA]
ax.legend(uchwyty, ["1 – zdecydowanie nie", "2", "3", "4", "5 – zdecydowanie tak"], ncol=5,
          loc="upper center", bbox_to_anchor=(0.45, -0.02), frameon=False, fontsize=10,
          handlelength=1.0, columnspacing=1.0)
zapisz(fig, "6_skale.png")

# 7. Gdzie kupują
g = df["gdzie"].value_counts()
fig, ax = plt.subplots(figsize=(6.0, 2.0))
poziomy(ax, [e.replace("Na platformach typu Allegro czy Amazon", "Allegro, Amazon itp.")
             .replace("W sklepach pośrednich online", "Sklepy internetowe (pośrednicy)")
             .replace("W sklepie stacjonarnym", "Sklep stacjonarny")
             .replace("Na stronie producenta", "Strona producenta") for e in g.index],
        [int(v) for v in g.values], sortuj=True)
zapisz(fig, "7_gdzie.png")

# --- Porównanie grup: studenci (seria 1) i pracujący (seria 2) ---
GRUPY = [("Nauka lub studia", "Studenci i uczniowie", NIEBIESKI), ("Praca na etacie", "Pracujący", POMARANCZ)]


def grupowy(nazwa, etykiety, wartosci_dla, wysokosc):
    """Poziome słupki dla dwóch grup; wartości jako % osób w grupie."""
    fig, ax = plt.subplots(figsize=(6.4, wysokosc))
    k = len(etykiety)
    h = 0.34
    for g, (klucz, opis, kolor) in enumerate(GRUPY):
        d = df[df.zajecie == klucz]
        n = len(d)
        vals = wartosci_dla(d)
        for i, v in enumerate(vals):
            y = i + (g - 0.5) * (h + 0.04)
            if v:
                ax.barh(y, v / n, height=h, color=kolor, edgecolor="none")
            ax.text(v / n + 0.015, y, f"{v / n:.0%} ({v} z {n})", va="center", fontsize=9.5, color=TEKST_2)
    ax.set_xlim(0, 1.32)
    ax.set_ylim(k - 0.4, -0.75)
    ax.set_yticks(range(k))
    ax.set_yticklabels(etykiety)
    ax.set_xticks([])
    for s_ in ("top", "right", "bottom"):
        ax.spines[s_].set_visible(False)
    ax.spines["left"].set_color(OS)
    ax.tick_params(axis="y", length=0, pad=6)
    uchwyty = [Rectangle((0, 0), 1, 1, fc=k_) for _, _, k_ in GRUPY]
    ax.legend(uchwyty, [f"{o} ({(df.zajecie == kl).sum()} osób)" for kl, o, _ in GRUPY], ncol=2,
              loc="lower left", bbox_to_anchor=(0, 1.0), frameon=False, fontsize=10, handlelength=1.0)
    zapisz(fig, nazwa)


def wiele(kol, kolejnosc):
    return lambda d: [int(d[kol].str.contains(k, regex=False).sum()) for k in kolejnosc]


cechy = list(wielokrotny("wazne").index)
grupowy("8_wazne_grupy.png", cechy, wiele("wazne", cechy), 3.4)
syt = list(wielokrotny("sytuacje").index)
grupowy("9_sytuacje_grupy.png", syt, wiele("sytuacje", syt), 3.0)
grupowy("10_cena_grupy.png", ["Do 200 zł", "201–400 zł", "401–700 zł", "Powyżej 700 zł"],
        lambda d: [int((d.cena == k).sum()) for k in cena], 2.6)

# Średnie ocen w skalach 1-5 według grupy (wykres kropkowy)
fig, ax = plt.subplots(figsize=(6.4, 1.7))
for i, (kol, etyk) in enumerate(pytania):
    sr = [df[df.zajecie == kl][kol].mean() for kl, _, _ in GRUPY]
    ax.plot(sr, [i, i], color=OS, lw=2, zorder=1)
    lewy = min(range(2), key=lambda j: sr[j])
    for j, ((kl, opis, kolor), v) in enumerate(zip(GRUPY, sr)):
        ax.scatter([v], [i], s=90, color=kolor, zorder=3, edgecolor="white", linewidth=2)
        przes, wyr = (-0.09, "right") if j == lewy else (0.09, "left")
        ax.text(v + przes, i, f"{v:.1f}".replace(".", ","), ha=wyr, va="center", fontsize=10, color=TEKST_2)
ax.set_xlim(0.8, 5.2)
ax.set_ylim(1.5, -0.7)
ax.set_xticks([1, 2, 3, 4, 5])
ax.set_xticklabels(["1\nzdecydowanie nie", "2", "3", "4", "5\nzdecydowanie tak"], fontsize=9.5)
ax.set_yticks([0, 1])
ax.set_yticklabels([p_[1] for p_ in pytania])
ax.grid(axis="x", color=SIATKA, lw=0.8)
ax.set_axisbelow(True)
for s_ in ("top", "right", "left"):
    ax.spines[s_].set_visible(False)
ax.spines["bottom"].set_color(OS)
ax.tick_params(length=0, pad=6)
uchwyty = [plt.Line2D([0], [0], marker="o", ls="", ms=8, color=k_) for _, _, k_ in GRUPY]
ax.legend(uchwyty, [o for _, o, _ in GRUPY], ncol=2, loc="lower left", bbox_to_anchor=(0, 1.0),
          frameon=False, fontsize=10, handletextpad=0.3)
zapisz(fig, "11_skale_grupy.png")

print("Zapisano wykresy w", WYK)
