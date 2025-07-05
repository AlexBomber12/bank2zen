import json, pathlib, shutil, sys


def load_safe(path: pathlib.Path):
    try:
        return json.load(open(path, encoding='utf-8'))
    except json.JSONDecodeError:
        bak = path.with_suffix('.bad.json')
        shutil.copy(path, bak)
        print(f"[WARN] {path.name} corrupted → backup to {bak.name}")
        return {}


def dedup_file(path: str):
    p = pathlib.Path(path)
    data = load_safe(p)
    changed = False
    for k, lst in data.items():
        uniq = list(dict.fromkeys(lst))          # preserve order
        if len(uniq) != len(lst):
            data[k] = uniq
            changed = True
    if changed:
        json.dump(
            data,
            open(p, 'w', encoding='utf-8'),
            ensure_ascii=False,
            indent=2,
        )
        print(f"{p.name}: duplicates removed")
