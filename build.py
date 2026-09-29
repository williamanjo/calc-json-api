"""Packs the extension into dist/calc-json-api.oxt."""

import pathlib
import zipfile

ROOT = pathlib.Path(__file__).parent
FILES = [
    "META-INF/manifest.xml",
    "description.xml",
    "description/description_en.txt",
    "description/description_pt-BR.txt",
    "images/icon.png",
    "Addons.xcu",
    "python/json_api.py",
    "LICENSE",
]

out = ROOT / "dist" / "calc-json-api.oxt"
out.parent.mkdir(exist_ok=True)
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as oxt:
    for name in FILES:
        oxt.write(ROOT / name, name)
print(out)
