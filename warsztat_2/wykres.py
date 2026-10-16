import sys
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

plt.rcParams["font.family"] = "Carlito"
df = pd.read_excel(sys.argv[1], sheet_name="Dane syntetyczne")
x, y = df.iloc[:, 1].to_numpy(float), df.iloc[:, 2].to_numpy(float)
a, b = np.polyfit(x, y, 1)
r2 = np.corrcoef(x, y)[0, 1] ** 2

fig, ax = plt.subplots(figsize=(7, 4.2), dpi=200)
ax.scatter(x, y, s=28, color="#4472C4", zorder=3)
xs = np.linspace(x.min() - 2, x.max() + 2, 50)
ax.plot(xs, a * xs + b, color="black", lw=1, ls=":", zorder=2)
for xi, yi, m in zip(x, y, df.iloc[:, 0]):
    if m in ("Lis", "Gru"):
        ax.annotate(m, (xi, yi), textcoords="offset points", xytext=(-22, 2), fontsize=10)
ax.text(0.04, 0.92, f"y = {a:.2f}x + {b:.1f}".replace(".", ","), transform=ax.transAxes, fontsize=11)
ax.text(0.04, 0.84, f"R² = {r2:.3f}".replace(".", ","), transform=ax.transAxes, fontsize=11)
ax.set_title("Wydatki na reklamę a sprzedaż", fontsize=13)
ax.set_xlabel("Reklama (tys. zł)")
ax.set_ylabel("Sprzedaż (tys. zł)")
ax.grid(color="#D9D9D9", lw=0.8)
ax.set_axisbelow(True)
for s in ("top", "right"):
    ax.spines[s].set_visible(False)
fig.tight_layout()
fig.savefig(sys.argv[2], facecolor="white")
print(round(a, 4), round(b, 4), round(r2, 4))
