import os
import sys
import subprocess
import datetime as dt
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd, json, pathlib
from pathlib import Path
from bank2zen import convert, normalize, CATS_FILE, ACC_FILE, _read_movements_xlsx
from history_index import db_path, ensure_db, seen_lookup, mark_seen
from dedup_json import dedup_file

for f in ("categories.json", "accounts_to.json"):
    dedup_file(f)

# ── справочники ───────────────────────────────────────────────────────────────
CATS = sorted([
    "Аренда","Благотворительность","Возврат","Дети","ЖКХ","Забота о себе",
    "Зарплата","Инвестиции","Карманные расходы","Кафе и рестораны","Переводы",
    "Командировки и работа","Компенсация затрат","Корректировка","Кредиты",
    "Лекарства и медицина","Машина","Налоги","Наследство","Образование",
    "Общественный транспорт","Отдых и развлечения",
    "Платежи, комиссии, штрафы","Подарки","Покупки ПО, подписки",
    "Покупки: мебель и обстановка","Покупки: одежда","Покупки: техника",
    "Покупки: хозяйственное","Продажа имущества","Продукты","Проценты",
    "Связь","Соцвыплата","Страховка","Строительство и ремонт",
    "Услуги: фото, ремонт и т. д.","Хобби",
    "Забота о себе: Массаж","Забота о себе: Салоны красоты",
    "Карманные расходы: Ирина","Карманные расходы: Петр",
    "Кафе и рестораны: Обеды в школе","Кафе и рестораны: Обеды на работе",
    "Машина: Парковка","Машина: Платные дороги"
], key=str.lower)

ACCTS = sorted(["Fineco Credit","Fineco Debit","Наличные Евро","Revolut"], key=str.lower)

# ── date helpers ──────────────────────────────────────────────────────────────
def _parse_date_value(value):
    if value is None:
        return None
    if isinstance(value, pd.Timestamp):
        if pd.isna(value):
            return None
        return value.normalize()
    if isinstance(value, dt.datetime):
        return pd.Timestamp(value).normalize()
    if isinstance(value, dt.date):
        return pd.Timestamp(value).normalize()
    text = str(value).strip()
    if not text or text.lower() in {"nan", "nat"}:
        return None
    try:
        parsed = pd.to_datetime(text, dayfirst=True, errors="coerce")
    except Exception:
        return None
    if pd.isna(parsed):
        return None
    return parsed.normalize()


def _date_only_display(value) -> str:
    parsed = _parse_date_value(value)
    if parsed is not None:
        return parsed.date().isoformat()
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "nat"}:
        return ""
    return text

# ── json utils ────────────────────────────────────────────────────────────────
def jload(p):
    if pathlib.Path(p).exists():
        try: return json.load(open(p,encoding='utf-8'))
        except: pass
    return {}
def jsave(p,d): json.dump(d,open(p,'w',encoding='utf-8'),ensure_ascii=False,indent=2)
def add_pat(mp,k,p):                      # без дублей
    if k: mp.setdefault(k,[]); (p in mp[k]) or mp[k].append(p)

# ── Combobox c live-поиском ───────────────────────────────────────────────────
class AutoCombo(ttk.Combobox):
    def __init__(self,*a,**kw):
        super().__init__(*a,**kw); self._base=list(self['values'])
        self.bind('<KeyRelease>',self._filter)
    def _filter(self,_):
        try:
            if self.cget('state')=='readonly':
                self['values']=self._base
                return
        except Exception:
            pass
        q=self.get().lower(); self['values']=[v for v in self._base if q in v.lower()] if q else self._base

# ── Review окно ───────────────────────────────────────────────────────────────
class Review(tk.Toplevel):
    COLS=('Date','Amount','Description','Category','AccountTo')
    WIDTHS=(90,90,620,180,120)
    def __init__(self,master,df=None):
        super().__init__(master)
        if df is None:
            df=pd.read_csv('out_zenmoney.csv',sep=';')
        else:
            df=df.copy()
        df=df.where(pd.notna(df),"")
        df=df.replace({"nan":""})
        self.df=df
        self.sort_state={}
        self.title('Review'); self.geometry('1180x560')
        self._ui(); self._refresh()
    # UI
    def _ui(self):
        top=tk.Frame(self); top.pack(fill=tk.X,padx=10,pady=5)
        tk.Label(top,text='Filter:').pack(side=tk.LEFT)
        self.q=tk.StringVar(); self.q.trace_add('write',self._filter)
        tk.Entry(top,textvariable=self.q,width=26).pack(side=tk.LEFT,padx=5)
        ttk.Button(top,text='Clear',command=lambda:self.q.set('')).pack(side=tk.LEFT,padx=4)
        self.cb_cat=AutoCombo(top,values=['']+CATS,width=40); self.cb_cat.pack(side=tk.LEFT,padx=4)
        self.cb_acc=AutoCombo(top,values=['']+ACCTS,width=25); self.cb_acc.pack(side=tk.LEFT,padx=4)
        ttk.Button(top,text='Set Cat',command=self._set_cat).pack(side=tk.LEFT,padx=4)
        ttk.Button(top,text='Set AccTo',command=self._set_acc).pack(side=tk.LEFT,padx=4)
        ttk.Button(top,text='Save & Close',command=self._save_exit).pack(side=tk.RIGHT)
        self.tree=ttk.Treeview(self,columns=self.COLS,show='headings',height=22)
        for c,w in zip(self.COLS,self.WIDTHS):
            self.tree.heading(c,text=c,command=lambda col=c:self._sort(col))
            self.tree.column(c,width=w,anchor='w')
        vs=tk.Scrollbar(self,orient='vertical',command=self.tree.yview)
        self.tree.configure(yscroll=vs.set)
        self.tree.pack(side=tk.LEFT,fill=tk.BOTH,expand=True,padx=(10,0),pady=5); vs.pack(side=tk.RIGHT,fill=tk.Y,pady=5)
    # helpers
    def _fmt_cell(self,v):
        if v is None:
            return ""
        if isinstance(v,float) and pd.isna(v):
            return ""
        if isinstance(v,str) and v.lower()=="nan":
            return ""
        return str(v)
    def _amount_val(self,v):
        if v is None:
            return None
        if isinstance(v,str):
            s=v.strip()
            if not s or s.lower()=="nan":
                return None
            try:
                return float(s.replace(',','.'))
            except ValueError:
                return None
        if isinstance(v,(int,float)):
            return None if pd.isna(v) else float(v)
        return None
    def _amount(self,r):
        inc=self._amount_val(r.get('Income'))
        out=self._amount_val(r.get('Expense'))
        inc = inc if inc is not None else 0
        out = out if out is not None else 0
        return inc - out
    def _refresh(self,sub=None):
        self.tree.delete(*self.tree.get_children())
        data=sub if sub is not None else self.df
        for idx,r in data.iterrows():
            self.tree.insert('',tk.END,iid=str(idx),
                values=(_date_only_display(r['Data']),f"{self._amount(r):+.2f}",
                    self._fmt_cell(r['Descrizione_Completa']),
                    self._fmt_cell(r['Category']),
                    self._fmt_cell(r['AccountTo'])))
    def _filter(self,*_):
        q=self.q.get().lower()
        if not q: self._refresh(); return
        sub=self.df[self.df['Descrizione_Completa'].str.lower().str.contains(q)|self.df['Category'].str.lower().str.contains(q)]
        self._refresh(sub)
    def _update(self,iids,col,val):
        for iid in iids: self.tree.set(iid,col,val)
    # batch edit
    def _set_cat(self):
        sel=self.tree.selection(); val=self.cb_cat.get()
        if sel and val:
            for iid in sel:self.df.at[int(iid),'Category']=val
            self._update(sel,'Category',val); self.tree.focus_set()
    def _set_acc(self):
        sel=self.tree.selection(); val=self.cb_acc.get()
        if sel and val:
            for iid in sel:self.df.at[int(iid),'AccountTo']=val
            self._update(sel,'AccountTo',val); self.tree.focus_set()
    # sort
    def _sort(self,col):
        rev=self.sort_state.get(col,False)
        rows=[(self.tree.set(iid,col),iid,idx) for idx,iid in enumerate(self.tree.get_children(''))]
        def key(t):
            v=t[0]
            if col=='Amount': return float(v.replace('+','').replace(',','.'))
            if col=='Date':
                parsed=_parse_date_value(v)
                return parsed if parsed is not None else pd.Timestamp.max
            return v.lower()
        if col=='Date':
            valid=[]; invalid=[]
            for val,iid,order in rows:
                parsed=_parse_date_value(val)
                if parsed is None:
                    invalid.append((order,iid))
                else:
                    valid.append((parsed,order,iid))
            valid.sort(key=lambda t:(t[0],t[1]))
            if rev:
                valid.reverse()
            ordered=[iid for _,_,iid in valid]
            ordered.extend(iid for _,iid in sorted(invalid,key=lambda t:t[0]))
            for pos,iid in enumerate(ordered):
                self.tree.move(iid,'',pos)
        else:
            rows.sort(key=key,reverse=rev)
            for i,(_,iid,_) in enumerate(rows): self.tree.move(iid,'',i)
        self.sort_state[col]=not rev
    # save
    def _save_exit(self):
        try:
            cols=['Data','Category','Descrizione_Completa','Account','AccountTo','Income','Expense']
            df_to_save=self.df.copy()
            df_to_save=df_to_save.where(pd.notna(df_to_save),"")
            df_to_save=df_to_save.replace({"nan":""})
            df_to_save.to_csv('out_zenmoney.csv',columns=cols,index=False,encoding='utf-8-sig',sep=';')
            cm,am=jload(CATS_FILE),jload(ACC_FILE)
            for _,r in self.df.iterrows():
                p=normalize(r['Descrizione_Completa'])
                add_pat(cm,r['Category'],p); add_pat(am,r['AccountTo'],p)
            jsave(CATS_FILE,cm); jsave(ACC_FILE,am)
            self.master.destroy()
        except PermissionError:
            messagebox.showerror('Ошибка','Файл out_zenmoney.csv открыт.')

# AssignWin без изменений (опущено для краткости) ─────────────────────────────
class AssignWin(tk.Toplevel):
    """Assign categories to unknown rows."""

    def __init__(self, master, df):
        super().__init__(master)
        self.title('Assign categories'); self.geometry('980x540')
        self.df = df.reset_index(drop=True)
        drop_time=[c for c in self.df.columns if str(c).strip().lower()=="time"]
        if drop_time:
            self.df = self.df.drop(columns=drop_time)
        self.pairs = []
        # listbox
        fr = tk.Frame(self); fr.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        sb = tk.Scrollbar(fr); sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.lb = tk.Listbox(
            fr,
            selectmode=tk.EXTENDED,
            exportselection=False,
            width=140,
            yscrollcommand=sb.set,
        )
        self.lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True); sb.config(command=self.lb.yview)
        for _, r in self.df.iterrows():
            inc_raw = r.get('Entrate', None)
            exp_raw = r.get('Uscite', None)
            if inc_raw is None:
                inc_raw = r.get('Income', 0)
            if exp_raw is None:
                exp_raw = r.get('Expense', 0)

            inc = pd.to_numeric(inc_raw, errors='coerce')
            exp = pd.to_numeric(exp_raw, errors='coerce')
            inc = 0.0 if pd.isna(inc) else float(abs(inc))
            exp = 0.0 if pd.isna(exp) else float(abs(exp))

            if inc == 0 and exp == 0:
                amt = 0.0
            elif inc > 0 and exp > 0:
                amt = inc - exp
            elif inc > 0:
                amt = inc
            else:
                amt = -exp

            amount_str = f"{amt:+.2f}"
            date_str = _date_only_display(r.get('Data'))
            self.lb.insert(
                tk.END,
                f"{date_str} | {amount_str} | {r['Descrizione_Completa']}"
            )
        self.lb.bind("<<ListboxSelect>>", self._on_list_select)
        self.lb.bind("<Double-Button-1>", self._assign_selected)
        self.lb.bind("<Return>", self._assign_selected)
        self.lb.bind("<Control-a>", self._select_all)
        # controls
        bottom = tk.Frame(self); bottom.pack(fill=tk.X, padx=10, pady=(0,10))
        self.sel_count = tk.StringVar(value="Selected: 0")
        tk.Label(bottom, textvariable=self.sel_count).pack(side=tk.LEFT, padx=8)
        controls = tk.Frame(bottom); controls.pack(side=tk.RIGHT)
        self.cat_box = AutoCombo(controls, values=CATS, width=70); self.cat_box.grid(row=0, column=0, padx=5)
        self.cat_box.state(["readonly"])
        self.cat_box.bind("<Return>", self._assign_selected)
        self.cat_box.bind("<<ComboboxSelected>>", self._on_category_change)
        last_cat = getattr(self.master, "_last_assign_category", None)
        if last_cat and last_cat in self.cat_box['values']:
            self.cat_box.set(last_cat)
            self.master._last_assign_category = last_cat
        ttk.Button(controls, text='Assign', command=self._assign_selected).grid(row=0, column=1, padx=5)
        ttk.Button(controls, text='Finish', command=self.finish).grid(row=0, column=2, padx=5)

    def _on_list_select(self, *_):
        self.sel_count.set(f"Selected: {len(self.lb.curselection())}")

    def _select_all(self, event=None):
        self.lb.selection_set(0, tk.END); self._on_list_select()
        return "break"

    def _on_category_change(self, *_):
        cat = self._current_category()
        if cat:
            self.master._last_assign_category = cat
            try:
                self.cat_box['values'] = self.cat_box._base
            except Exception:
                pass

    def _current_category(self):
        cat = self.cat_box.get().strip()
        return cat or None

    def _assign_selected(self, event=None):
        cat = self._current_category()
        if not cat:
            messagebox.showwarning("bank2zen", "Сначала выберите категорию.")
            return "break" if event else None
        sel = list(self.lb.curselection())
        if not sel:
            messagebox.showwarning("bank2zen", "Выберите одну или несколько строк.")
            return "break" if event else None
        self.master._last_assign_category = cat
        for pos in sorted(sel, reverse=True):
            patt = normalize(self.df.iloc[pos]['Descrizione_Completa'])
            self.pairs.append((cat, patt))
            self.lb.delete(pos)
            self.df.drop(index=self.df.index[pos], inplace=True)
        self.df.reset_index(drop=True, inplace=True)
        self.lb.selection_clear(0, tk.END)
        self._on_list_select()
        if not self.lb.size():
            self.finish()
        else:
            self.lb.focus_set()
        try:
            self.cat_box['values'] = self.cat_box._base
        except Exception:
            pass
        if getattr(self.master, "_last_assign_category", None):
            self.cat_box.set(self.master._last_assign_category)
        return "break" if event else None

    def finish(self):
        data = jload(CATS_FILE)
        for cat, patt in self.pairs:
            add_pat(data, cat, patt)
        jsave(CATS_FILE, data)
        self.destroy()

# ── Main ──────────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('bank2zen — Fineco → ZenMoney'); self.geometry('770x360')
        frm=ttk.Frame(self,padding=10); frm.grid(sticky='nsew')
        self.columnconfigure(0,weight=1); frm.columnconfigure(1,weight=1)
        ttk.Label(frm,text='Bank export:').grid(row=0,column=0,sticky='e')
        self.path=tk.StringVar()
        ttk.Entry(frm,textvariable=self.path).grid(row=0,column=1,sticky='ew')
        ttk.Button(frm,text='Browse',command=self.browse).grid(row=0,column=2,sticky='ew',padx=5)
        ttk.Button(frm,text='Assign and Convert',command=self.assign_and_convert).grid(row=1,column=1,sticky='ew',pady=8)
        ttk.Button(frm,text='Exit',command=self.destroy).grid(row=1,column=2,sticky='ew',padx=5)
        self.log_box=tk.Text(frm,height=10,state='disabled')
        self.log_box.grid(row=2,column=0,columnspan=3,sticky='nsew'); frm.rowconfigure(2,weight=1)
        self._make_menubar()

    def log(self,msg):
        self.log_box['state']='normal'; self.log_box.insert(tk.END,msg+'\n')
        self.log_box.see(tk.END); self.log_box['state']='disabled'

    def browse(self):
        f=filedialog.askopenfilename(filetypes=[('Excel','*.xlsx')]); self.path.set(f or '')
    def assign_and_convert(self):
        fn=self.path.get().strip()
        if not fn:
            messagebox.showwarning('Нет файла','Выберите .xlsx')
            return
        try:
            res=convert(fn)
        except PermissionError:
            messagebox.showerror('Ошибка','Файл выгрузки открыт.')
            return
        if res=='no_new':
            self.log('Новых операций нет.')
            return
        if res=='ok':
            self.log('CSV создан → Review')
            Review(self,pd.read_csv('out_zenmoney.csv',sep=';'))
            return
        if not pathlib.Path('new_desc.xlsx').exists():
            self.log('new_desc.xlsx не найден')
            return
        df=pd.read_excel('new_desc.xlsx')
        n=len(df)
        self.log(f'Обнаружены неизвестные категории. Нужно разметить {n} строк.')
        assign_win=AssignWin(self,df)
        self.wait_window(assign_win)
        try:
            res2=convert(fn)
        except PermissionError:
            messagebox.showerror('Ошибка','Файл выгрузки открыт.')
            return
        if res2=='no_new':
            self.log('Новых операций нет.')
            return
        if res2=='ok':
            self.log('CSV создан → Review')
            Review(self,pd.read_csv('out_zenmoney.csv',sep=';'))
        else:
            if pathlib.Path('new_desc.xlsx').exists():
                left=len(pd.read_excel('new_desc.xlsx'))
                self.log(f'Осталось разметить {left} строк.')
            else:
                self.log('new_desc.xlsx не найден')

    def _make_menubar(self):
        try:
            m = tk.Menu(self)
            tools = tk.Menu(m, tearoff=0)
            tools.add_command(label="Open history folder", command=self._open_history_dir)
            tools.add_separator()
            tools.add_command(label="Seed history from folder…", command=self._seed_history_from_folder)
            tools.add_separator()
            tools.add_command(label="Reset history…", command=self._reset_history)
            m.add_cascade(label="Tools", menu=tools)
            self.config(menu=m)
        except Exception as e:
            self.log(f"Menu init error: {e}")

    def _history_dir(self) -> Path:
        try:
            return Path(db_path()).parent
        except Exception:
            return Path.home() / ".bank2zen"

    def _open_history_dir(self):
        p = self._history_dir()
        p.mkdir(parents=True, exist_ok=True)
        try:
            if sys.platform.startswith('win'):
                os.startfile(str(p))  # type: ignore[attr-defined]
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', str(p)])
            else:
                subprocess.Popen(['xdg-open', str(p)])
            self.log(f"Opened history folder: {p}")
        except Exception as e:
            self.log(f"Cannot open folder: {p} ({e})")
            messagebox.showinfo('bank2zen', f"History folder:\n{p}")

    def _reset_history(self):
        if not messagebox.askyesno('bank2zen', 'Reset history index?\nThis will forget all previously imported transactions.'):
            return
        p = Path(db_path())
        try:
            if p.exists():
                try:
                    p.unlink()
                    self.log('History reset: file deleted.')
                    messagebox.showinfo('bank2zen', 'History reset.\nFile removed.')
                    return
                except Exception:
                    pass
            con = ensure_db()
            try:
                con.execute('DELETE FROM seen')
                con.execute('VACUUM')
                con.commit()
            finally:
                con.close()
            self.log('History reset: table cleared.')
            messagebox.showinfo('bank2zen', 'History reset.\nTable cleared.')
        except Exception as e:
            self.log(f'History reset error: {e}')
            messagebox.showerror('bank2zen', f'Error resetting history:\n{e}')

    def _seed_history_from_folder(self):
        folder = filedialog.askdirectory(title='Select folder with XLSX exports')
        if not folder:
            return
        only90 = messagebox.askyesno('bank2zen', 'Seed only last 90 days?\nYes = last 90 days, No = all history.')
        from datetime import date, timedelta
        since = date.today() - timedelta(days=90) if only90 else None

        con = ensure_db()
        folder = Path(folder)
        files = sorted(folder.glob('*.xlsx'))
        total = dupes = new = 0

        for xf in files:
            try:
                df = _read_movements_xlsx(str(xf))
                if df.empty:
                    continue
                df['Data_Valuta'] = pd.to_datetime(df['Data_Valuta'], dayfirst=True, errors='coerce')
                df = df[df['Data_Valuta'].notna()].copy()
                if since is not None:
                    df = df[df['Data_Valuta'].dt.date >= since]
                if df.empty:
                    continue

                entr = pd.to_numeric(df.get('Entrate', 0), errors='coerce').fillna(0).abs()
                usct = pd.to_numeric(df.get('Uscite', 0), errors='coerce').fillna(0).abs()

                direction = pd.Series('I', index=df.index)
                direction.loc[usct > 0] = 'O'
                amount = entr.copy()
                amount.loc[usct > 0] = usct[usct > 0]
                amount = amount.round(2)

                if 'Descrizione_Completa' in df.columns:
                    desc_full = df['Descrizione_Completa']
                else:
                    desc_full = df.get('Descrizione', '')
                account = df.get('Account', '')

                keys = []
                from history_index import fingerprint  # local import avoids circulars
                for i in df.index:
                    date_iso = df.at[i, 'Data_Valuta'].date().isoformat()
                    dirc = direction.at[i]
                    cents = int((amount.at[i] * 100))
                    if isinstance(desc_full, pd.Series):
                        text = str(desc_full.at[i]) if pd.notna(desc_full.at[i]) else ''
                    else:
                        text = str(desc_full) if desc_full is not None else ''
                    if isinstance(account, pd.Series):
                        acc_val = str(account.at[i]) if i in account.index else ''
                    else:
                        acc_val = str(account) if account is not None else ''
                    keys.append(fingerprint(date_iso, dirc, cents, text, acc_val))

                found = seen_lookup(con, keys)
                seen_local = set(found)
                to_ins = []
                for j, key in enumerate(keys):
                    total += 1
                    if key in seen_local:
                        dupes += 1
                        continue
                    i = df.index[j]
                    d = df.at[i, 'Data_Valuta'].date().isoformat()
                    dirc = direction.at[i]
                    amt = float(amount.at[i])
                    if isinstance(desc_full, pd.Series):
                        text = str(desc_full.at[i]) if pd.notna(desc_full.at[i]) else ''
                    else:
                        text = str(desc_full) if desc_full is not None else ''
                    if isinstance(account, pd.Series):
                        acc_val = str(account.at[i]) if i in account.index else ''
                    else:
                        acc_val = str(account) if account is not None else ''
                    to_ins.append((key, d, dirc, amt, acc_val, '', '', xf.name))
                    seen_local.add(key)
                if to_ins:
                    mark_seen(con, to_ins)
                    new += len(to_ins)
                self.log(f'Seeded from {xf.name}: +{len(to_ins)} new, {len(found)} seen total.')
            except Exception as e:
                self.log(f'Seed error {xf.name}: {e}')

        messagebox.showinfo('bank2zen', f'Seed done.\nFiles: {len(files)}\nNew: {new}\nDuplicates: {dupes}\nScanned rows: {total}')
        self.log(f'Seed done. Files={len(files)} New={new} Duplicates={dupes} TotalRows={total}')
        con.close()

if __name__ == '__main__':
    import pandas as pd
    App().mainloop()
