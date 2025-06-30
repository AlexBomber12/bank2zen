import json, pathlib, sys

def dedup(path: str):
    p = pathlib.Path(path)
    if not p.exists():
        print(f"{path}: not found"); return
    data = json.load(open(p, encoding='utf-8'))
    changed = False
    for k, lst in data.items():
        unique = sorted(set(lst), key=lst.index)   # сохраняем порядок
        if len(unique) != len(lst):
            data[k] = unique
            changed = True
    if changed:
        json.dump(data, open(p, 'w', encoding='utf-8'),
                  ensure_ascii=False, indent=2)
        print(f"{path}: duplicates removed")
    else:
        print(f"{path}: already clean")

for f in ("categories.json", "accounts_to.json"):
    dedup(f)
