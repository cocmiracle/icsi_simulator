"""Сглобява images/images.js от снимките в images/ и списъка images/list.csv.

Защо е нужен: снимките се вграждат като данни (data URL) в images.js, за да работи
симулаторът и когато index.html е отворен с двоен клик (без сървър) - иначе браузърът
не позволява обработката на снимката на holding пипетата.

Нова или сменена снимка: сложете файла (.webp, .png или .jpg) в images/, добавете или
поправете реда му в list.csv и пуснете:  python tools/build_images.py
"""
import base64, csv, json, os, sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IMG = os.path.join(ROOT, "images")
MIME = {".webp": "image/webp", ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg"}

def main():
    rows = list(csv.DictReader(open(os.path.join(IMG, "list.csv"), encoding="utf-8"), delimiter=";"))
    data, meta = {}, []
    for r in rows:
        f = r["file"].strip()
        path = os.path.join(IMG, f)
        if not os.path.exists(path):
            sys.exit(f"Липсва файл: images/{f}")
        ext = os.path.splitext(f)[1].lower()
        key = os.path.splitext(f)[0]
        data[key] = f"data:{MIME[ext]};base64," + base64.b64encode(open(path, "rb").read()).decode()
        meta.append({k: r[k].strip() for k in ("file", "group", "what", "source", "license", "url")})
    out = ["// ГЕНЕРИРАН ФАЙЛ - не се пипа на ръка. Източник: images/*.webp и images/list.csv; сглобява се с tools/build_images.py",
           "window.IMG_LIST = " + json.dumps(meta, ensure_ascii=False, indent=1) + ";",
           "window.IMG = {"]
    out += [f'  "{k}": "{v}",' for k, v in data.items()]
    out.append("};")
    open(os.path.join(IMG, "images.js"), "w", encoding="utf-8").write("\n".join(out) + "\n")
    print(f"images.js: {len(data)} снимки")

if __name__ == "__main__":
    main()
