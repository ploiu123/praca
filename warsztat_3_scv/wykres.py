"""Wykres zaangażowania i konwersji dla czterech postów (do Worda)."""
from pathlib import Path

import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

KAT = Path(__file__).resolve().parent
plt.rcParams.update({"font.family": "Carlito", "font.size": 11, "text.color": "#0b0b0b",
                     "ytick.color": "#0b0b0b", "xtick.color": "#52514e"})

d = pd.read_csv(KAT / "dane.csv")
d["zaang"] = ((d.polubienia + d.komentarze + d.udostepnienia) / d.wyswietlenia * 100).round(1)
d["konw"] = (d.zakupy / d.klikniecia * 100).round(1)
etykiety = [f"{r.id} {r.format}" for r in d.itertuples()]

fig, osie = plt.subplots(1, 2, figsize=(7.2, 2.6))
for ax, kol, tytul, kolor in ((osie[0], "zaang", "Zaangażowanie", "#4472C4"),
                              (osie[1], "konw", "Konwersja po kliknięciu", "#ED7D31")):
    wart = list(d[kol])
    ax.barh(range(len(wart)), wart, height=0.55, color=kolor)
    for i, v in enumerate(wart):
        ax.text(v + max(wart) * 0.02, i, f"{v:.1f}%".replace(".", ","), va="center", fontsize=10.5, color="#52514e")
    ax.set_yticks(range(len(wart)))
    ax.set_yticklabels(etykiety)
    ax.set_ylim(len(wart) - 0.4, -0.6)
    ax.set_xlim(0, max(wart) * 1.25)
    ax.set_xticks([])
    for s in ("top", "right", "bottom"):
        ax.spines[s].set_visible(False)
    ax.spines["left"].set_color("#c3c2b7")
    ax.tick_params(axis="y", length=0, pad=6)
    ax.set_title(tytul, loc="left", fontsize=12, pad=8)
fig.subplots_adjust(wspace=0.75)
fig.savefig(KAT / "wykres_wskazniki.png", dpi=200, facecolor="white", bbox_inches="tight")
print("Zapisano wykres_wskazniki.png")
