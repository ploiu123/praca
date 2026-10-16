import sys

import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch
from matplotlib.path import Path

OUT = sys.argv[1]
plt.rcParams["font.family"] = "Carlito"
plt.rcParams["svg.fonttype"] = "none"

WEZLY = {
    "pop": ("Popyt rynkowy", 0.56, 0.95),
    "sez": ("Sezonowość", 0.33, 0.73),
    "rab": ("Rabaty", 0.60, 0.64),
    "rek": ("Wydatki\nreklamowe", 0.07, 0.40),
    "ruch": ("Ruch na stronie", 0.50, 0.40),
    "spr": ("Sprzedaż", 0.93, 0.40),
    "kon": ("Działania\nkonkurencji", 0.50, 0.08),
}

ZIELONY, CZERWONY, NIEBIESKI = "#2E7D32", "#C62828", "#1565C0"

# (od, do, wygięcie łuku, kolor)
KRAWEDZIE = [
    ("rek", "ruch", 0.0, ZIELONY),
    ("ruch", "spr", 0.0, ZIELONY),
    ("sez", "rek", 0.0, CZERWONY),
    ("pop", "rek", 0.50, CZERWONY),
    ("rab", "rek", 0.0, CZERWONY),
    ("kon", "rek", 0.0, CZERWONY),
    ("pop", "spr", -0.30, NIEBIESKI),
    ("rab", "spr", 0.0, NIEBIESKI),
    ("kon", "spr", 0.0, NIEBIESKI),
    ("sez", "pop", 0.0, NIEBIESKI),
    ("sez", "rab", 0.0, NIEBIESKI),
]

LEGENDA = [
    (ZIELONY, "ścieżka przyczynowa (przez mediator)"),
    (CZERWONY, "zmienna zakłócająca → przyczyna (reklama)"),
    (NIEBIESKI, "zmienna zakłócająca → skutek (sprzedaż)"),
]

fig, ax = plt.subplots(figsize=(8, 5.4), dpi=200)
ax.set_xlim(-0.10, 1.04)
ax.set_ylim(-0.02, 1.04)
ax.axis("off")

teksty = {k: ax.text(x, y, e, ha="center", va="center", fontsize=12, zorder=3,
                     bbox=dict(boxstyle="round,pad=0.45", fc="white", ec="black", lw=1.2))
          for k, (e, x, y) in WEZLY.items()}
fig.canvas.draw()

odwr = ax.transData.inverted()
pudelka = {}
for k, tx in teksty.items():
    bb = tx.get_bbox_patch().get_window_extent()
    (x0, y0), (x1, y1) = odwr.transform([(bb.x0 - 3, bb.y0 - 3), (bb.x1 + 3, bb.y1 + 3)])
    pudelka[k] = (x0, y0, x1, y1)


def w_pudelku(p, b):
    return b[0] <= p[0] <= b[2] and b[1] <= p[1] <= b[3]


for a, b, rad, kolor in KRAWEDZIE:
    p0 = np.array(WEZLY[a][1:]); p2 = np.array(WEZLY[b][1:])
    d = p2 - p0
    c = (p0 + p2) / 2 + rad * np.array([d[1], -d[0]])
    bez = lambda t: (1 - t) ** 2 * p0 + 2 * (1 - t) * t * c + t ** 2 * p2
    ts = np.linspace(0, 1, 2001)
    t0 = next(t for t in ts if not w_pudelku(bez(t), pudelka[a]))
    t1 = next(t for t in ts[::-1] if not w_pudelku(bez(t), pudelka[b]))
    q0, q2 = bez(t0), bez(t1)
    q1 = (1 - t0) * (1 - t1) * p0 + ((1 - t0) * t1 + t0 * (1 - t1)) * c + t0 * t1 * p2
    sciezka = Path([q0, q1, q2], [Path.MOVETO, Path.CURVE3, Path.CURVE3])
    ax.add_patch(FancyArrowPatch(path=sciezka, arrowstyle="-|>", mutation_scale=16,
                                 lw=2.0 if kolor == ZIELONY else 1.4, color=kolor, zorder=2))

from matplotlib.lines import Line2D
ax.legend([Line2D([0], [0], color=k, lw=2.5) for k, _ in LEGENDA], [o for _, o in LEGENDA],
          loc="upper center", bbox_to_anchor=(0.5, 0.0), ncol=1, frameon=False, fontsize=11)

fig.savefig(OUT, bbox_inches="tight", facecolor="white")
print("Zapisano", OUT)
