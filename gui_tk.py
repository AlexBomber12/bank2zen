import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd, json, pathlib
from bank2zen import convert, normalize, CATS_FILE, ACC_FILE
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
        q=self.get().lower(); self['values']=[v for v in self._base if q in v.lower()] if q else self._base

# ── Review окно ───────────────────────────────────────────────────────────────
class Review(tk.Toplevel):
    COLS=('Date','Amount','Description','Category','AccountTo')
    WIDTHS=(90,90,620,180,120)
    def __init__(self,master,df):
        super().__init__(master); self.df=df; self.sort_state={}
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
    def _amount(self,r):
        inc=r['Income'] if pd.notna(r['Income']) else None
        out=r['Expense'] if pd.notna(r['Expense']) else None
        return inc if inc is not None else (-out if out is not None else 0)
    def _refresh(self,sub=None):
        self.tree.delete(*self.tree.get_children())
        data=sub if sub is not None else self.df
        for idx,r in data.iterrows():
            self.tree.insert('',tk.END,iid=str(idx),
                values=(r['Data'],f"{self._amount(r):+.2f}",r['Descrizione_Completa'],r['Category'],r['AccountTo']))
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
        rows=[(self.tree.set(iid,col),iid) for iid in self.tree.get_children('')]
        def key(t):
            v=t[0]
            if col=='Amount': return float(v.replace('+','').replace(',','.'))
            if col=='Date':   return pd.to_datetime(v,dayfirst=True,errors='coerce')
            return v.lower()
        rows.sort(key=key,reverse=rev)
        for i,(_,iid) in enumerate(rows): self.tree.move(iid,'',i)
        self.sort_state[col]=not rev
    # save
    def _save_exit(self):
        try:
            cols=['Data','Category','Descrizione_Completa','Account','AccountTo','Income','Expense']
            self.df.to_csv('out_zenmoney.csv',columns=cols,index=False,encoding='utf-8-sig',sep=';')
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
        self.df = df.reset_index(drop=True); self.pairs = []
        # listbox
        fr = tk.Frame(self); fr.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        sb = tk.Scrollbar(fr); sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.lb = tk.Listbox(
            fr,
            selectmode=tk.EXTENDED,
            width=140,
            yscrollcommand=sb.set,
        )
        self.lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True); sb.config(command=self.lb.yview)
        for _, r in self.df.iterrows():
            self.lb.insert(
                tk.END,
                f"{r['Data']} | {r['Amount']:+.2f} | {r['Descrizione_Completa']}"
            )
        # controls
        bar = tk.Frame(self); bar.pack(pady=6)
        self.cmb = AutoCombo(bar, values=CATS, width=70); self.cmb.grid(row=0, column=0, padx=5)
        ttk.Button(bar, text='Assign', command=self.assign).grid(row=0, column=1, padx=5)
        ttk.Button(bar, text='Finish', command=self.finish).grid(row=0, column=2, padx=5)

    def assign(self):
        cat = self.cmb.get(); sel = list(self.lb.curselection())
        if not sel or not cat:
            return
        for idx in reversed(sel):
            patt = normalize(self.df.loc[idx, 'Descrizione_Completa'])
            self.pairs.append((cat, patt))
            self.lb.delete(idx); self.df.drop(index=idx, inplace=True)
        if not self.lb.size():
            self.finish()

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
        ttk.Button(frm,text='Convert',command=self.convert).grid(row=1,column=1,sticky='ew',pady=8)
        ttk.Button(frm,text='Exit',command=self.destroy).grid(row=1,column=2,sticky='ew',padx=5)
        self.log_box=tk.Text(frm,height=10,state='disabled')
        self.log_box.grid(row=2,column=0,columnspan=3,sticky='nsew'); frm.rowconfigure(2,weight=1)

    def log(self,msg):
        self.log_box['state']='normal'; self.log_box.insert(tk.END,msg+'\n')
        self.log_box.see(tk.END); self.log_box['state']='disabled'

    def browse(self):
        f=filedialog.askopenfilename(filetypes=[('Excel','*.xlsx')]); self.path.set(f or '')
    def convert(self):
        fn=self.path.get()
        if not fn: messagebox.showwarning('Нет файла','Выберите .xlsx'); return
        try: rc=convert(fn)
        except PermissionError: messagebox.showerror('Ошибка','Файл выгрузки открыт.'); return
        if rc=='ok':
            self.log('CSV создан → Review')
            Review(self,pd.read_csv('out_zenmoney.csv',sep=';'))
        else:
            if not pathlib.Path('new_desc.xlsx').exists():
                self.log('new_desc.xlsx не найден'); return
            df=pd.read_excel('new_desc.xlsx'); self.log(f'Нужно разметить {len(df)} строк')
            AssignWin(self,df); self.wait_window()
            if convert(fn)=='ok':
                self.log('CSV создан → Review')
                Review(self,pd.read_csv('out_zenmoney.csv',sep=';'))

if __name__ == '__main__':
    import pandas as pd
    App().mainloop()
