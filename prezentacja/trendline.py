"""Dodaje natywną linię trendu (z równaniem i R²) do wykresu rozrzutu w prezentacji.

Uruchomienie:  python3 trendline.py Analityka_warsztaty.pptx
"""
import re
import shutil
import sys
import zipfile
from pathlib import Path

plik = Path(sys.argv[1])
tmp = plik.with_suffix(".tmp.pptx")

TREND = ('<c:trendline><c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="1E2B37"/></a:solidFill>'
         '<a:prstDash val="dash"/></a:ln></c:spPr><c:trendlineType val="linear"/>'
         '<c:dispRSqr val="1"/><c:dispEq val="1"/>'
         '<c:trendlineLbl><c:numFmt formatCode="0.00" sourceLinked="0"/></c:trendlineLbl></c:trendline>')

zmienione = 0
with zipfile.ZipFile(plik) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        dane = zin.read(item.filename)
        if re.match(r"ppt/charts/chart\d+\.xml$", item.filename) and b"<c:scatterChart>" in dane:
            xml = dane.decode("utf-8")
            if "<c:trendline>" not in xml:
                # pierwsza seria wykresu rozrzutu: linia trendu przed <c:xVal>
                xml = xml.replace("<c:xVal>", TREND + "<c:xVal>", 1)
                zmienione += 1
            dane = xml.encode("utf-8")
        zout.writestr(item, dane)
shutil.move(tmp, plik)
print("Dodano linię trendu w", zmienione, "wykresie")
