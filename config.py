from pathlib import Path
import sys

if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent.parent
PROCESS_DIR = BASE_DIR/"Pengolahan"
TEMPLATE_DIR = BASE_DIR/"Bahan"
OUTPUT_DIR = PROCESS_DIR / "output"

DIRASAKAN_FILE = PROCESS_DIR / "dirasakan.xlsx"
DATALENGKAP_FILE = PROCESS_DIR/ "datalengkap.csv"
INPUT_PETA =PROCESS_DIR / "inputpeta.dat"
CROSS_SECTION = PROCESS_DIR / "crsx.dat"
QUAKES_CPT = PROCESS_DIR/ "quakes.cpt"


CITY = TEMPLATE_DIR / "kabkota.txt"
TEMPLATE_DOC = TEMPLATE_DIR / "template.docx"
SESAR_1 = TEMPLATE_DIR / "sesar2.gmt"
SESAR_2 = TEMPLATE_DIR / "sesar3.gmt"
LOGO_BMKG = TEMPLATE_DIR / "LogoBMKG.png"
MTANGAN = TEMPLATE_DIR / "mtangn.png"
EARTH_RELIEF = TEMPLATE_DIR / "earth_relief_15s.grd"
LEGENDA = TEMPLATE_DIR / "legenda.gmt"