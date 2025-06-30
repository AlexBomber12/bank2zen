# bank2zen

Конвертер банковской выгрузки Fineco (`.xlsx`) в CSV‑формат ZenMoney c автоподстановкой категорий и счётов.

---

## 1. Что делает скрипт

- Читает `movements_*.xlsx` из Fineco.
- Ставит **Category** и **AccountTo** по self‑обучающим JSON‑словарам (`categories.json`, `accounts_to.json`).
- Расходы по кредитке (`MONOFUNZIONE CONTACTLESS CHIP 5100 **** **** 3142`) попадают на счёт **Fineco Credit**.
- Снятие наличных `Prelievo Bancomat Unicredit` → перевод на **Наличные Евро**.
- Выписывает `out_zenmoney.csv` (`;` разделитель, `utf‑8‑sig`).

---

## 2. Быстрый старт (Python)

```bash
python -m venv venv && venv\Scripts\activate
pip install pandas openpyxl pyinstaller
python gui_tk.py            # графический интерфейс
```

### Обучение категорий

При первом запуске появится `new_desc.xlsx` — назначьте категории в GUI → далее всё проставляется автоматически.

---

## 3. Импорт в ZenMoney (веб‑интерфейс)

1. **Файл**: `out_zenmoney.csv`, формат — *CSV‑выписка*.
2. **Пропустить первые** → `1` строку (заголовки).
3. Сопоставление колонок:
   | CSV‑колонка | Поле ZenMoney            |
   | ----------- | ------------------------ |
   | `Data`      | Дата (дд/мм/гггг)        |
   | `Category`  | Категория                |
   | `Account`   | Счёт                     |
   | `AccountTo` | Счёт‑получатель перевода |
   | `Income`    | Сумма (доход)            |
   | `Expense`   | Сумма (расход)           |
   | остальные   | — (пропустить)           |
4. Оставьте поле «Импортировать на счёт» пустым (скрипт уже указал счёт).
5. Поставьте галку **«Применить правила обработки транзакций»** (по желанию).



---

## 4. Сборка standalone EXE (Windows)

```bash
pyinstaller gui_tk.py \
  --onefile --noconsole \
  --add-data "categories.json;." \
  --add-data "accounts_to.json;." \
  --add-data "bank2zen.py;." \
  --name bank2zen
```

*EXE появится в **`dist\bank2zen.exe`** – можно запускать без Python.*

---

## 5. Обновление / очистка словарей

- Дубликаты можно убрать скриптом `dedup_json.py` (см. инструкции в чате).
- Для полного сброса обучения удалите оба JSON‑файла – при следующем запуске они создадутся заново.

