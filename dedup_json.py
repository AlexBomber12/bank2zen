import json
from pathlib import Path

FILES = ("categories.json", "accounts_to.json")


def _dedup_file(path: str) -> int:
    p = Path(path)
    if not p.exists():
        print(f"{path}: not found")
        return 0
    data = json.load(open(p, encoding="utf-8"))
    removed = 0
    for k, lst in data.items():
        uniq = []
        for item in lst:
            if item not in uniq:
                uniq.append(item)
        removed += len(lst) - len(uniq)
        data[k] = uniq
    if removed:
        json.dump(data, open(p, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
        print(f"{path}: {removed} duplicates removed")
    else:
        print(f"{path}: already clean")
    return removed


def dedup_json() -> int:
    total = 0
    for f in FILES:
        total += _dedup_file(f)
    print(f"Total cleaned: {total}")
    return total


if __name__ == "__main__":
    dedup_json()
