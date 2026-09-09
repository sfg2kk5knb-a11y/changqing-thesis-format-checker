import json, os, sys, tkinter as tk
from tkinter import ttk, messagebox, simpledialog

PASSWORD = 'changqingwenchuang'
APP_NAME = '常青文创设计报价系统'
BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
CONFIG = os.path.join(BASE_DIR, '报价设置.json')

DEFAULT = {
 'education': {'函授本科（非学位）':'10-15','函授本科学位':'15-20','全日制本科':'25-30','硕士':'50-100','博士':'100+','全日制本科（外文）':'30','硕士（外文）':'100+','博士（外文）':'150+'},
 'services': {'普通数据分析':'100','做实证分析':'200','代码+截图':'100','系统/建模/仿真':'350'},
 'multiplier': 2
}

def load_cfg():
    try:
        with open(CONFIG, 'r', encoding='utf-8') as f: return json.load(f)
    except Exception: return DEFAULT.copy()

def parse_price(s):
    s = str(s).strip().replace('元','').replace('＋','+')
    if not s or s == '/': return None
    if '-' in s:
        a,b=s.split('-',1); return float(a),float(b)
    if '+' in s: return float(s.replace('+','')), None
    return float(s), float(s)

class App(tk.Tk):
    def __init__(self):
        super().__init__(); self.title(APP_NAME); self.geometry('920x650'); self.minsize(820,580)
        self.cfg=load_cfg(); self.configure(bg='#f4f7fb'); self.logo=None
        try:
            from PIL import Image, ImageTk
            p=os.path.join(BASE_DIR,'LOGO.png')
            if os.path.exists(p): self.logo=ImageTk.PhotoImage(Image.open(p).resize((310,92)))
        except Exception: pass
        self.build()
    def build(self):
        style=ttk.Style(self); style.theme_use('clam'); style.configure('TButton',font=('Microsoft YaHei',11),padding=8); style.configure('TLabel',background='#f4f7fb',font=('Microsoft YaHei',11)); style.configure('Header.TLabel',background='#10275b',foreground='white',font=('Microsoft YaHei',20,'bold')); style.configure('Card.TLabelframe',background='white'); style.configure('Card.TLabelframe.Label',background='white',foreground='#10275b',font=('Microsoft YaHei',12,'bold'))
        top=tk.Frame(self,bg='#10275b',height=115); top.pack(fill='x'); top.pack_propagate(False)
        if self.logo: tk.Label(top,image=self.logo,bg='#10275b').pack(side='left',padx=26,pady=10)
        else: ttk.Label(top,text=APP_NAME,style='Header.TLabel').pack(side='left',padx=28)
        tk.Label(top,text='专业 · 高效 · 明细透明',bg='#10275b',fg='#dbe6ff',font=('Microsoft YaHei',11)).pack(side='right',padx=32)
        body=tk.Frame(self,bg='#f4f7fb'); body.pack(fill='both',expand=True,padx=25,pady=20)
        left=ttk.LabelFrame(body,text='报价信息',style='Card.TLabelframe',padding=18); left.pack(side='left',fill='y',padx=(0,15)); right=ttk.LabelFrame(body,text='报价明细',style='Card.TLabelframe',padding=18); right.pack(side='left',fill='both',expand=True)
        self.kind=tk.StringVar(value='外包单'); ttk.Label(left,text='报价类型').pack(anchor='w'); f=tk.Frame(left,bg='white'); f.pack(anchor='w',pady=8); ttk.Radiobutton(f,text='外包单（最低价）',variable=self.kind,value='外包单').pack(side='left'); ttk.Radiobutton(f,text='一手单（翻倍）',variable=self.kind,value='一手单').pack(side='left',padx=10)
        ttk.Label(left,text='学历类型').pack(anchor='w',pady=(12,3)); self.edu=ttk.Combobox(left,values=list(self.cfg['education']),state='readonly',width=25); self.edu.current(0); self.edu.pack(anchor='w')
        ttk.Label(left,text='字数（千字）').pack(anchor='w',pady=(15,3)); self.words=ttk.Entry(left,width=27); self.words.insert(0,'10'); self.words.pack(anchor='w')
        ttk.Label(left,text='附加服务（可多选）').pack(anchor='w',pady=(18,3)); self.vars={}
        for n in self.cfg['services']:
            v=tk.BooleanVar(); self.vars[n]=v; ttk.Checkbutton(left,text=n,variable=v).pack(anchor='w',pady=2)
        bf=tk.Frame(left,bg='#f4f7fb'); bf.pack(fill='x',pady=(25,0)); ttk.Button(bf,text='开始计算',command=self.calculate).pack(side='left'); ttk.Button(bf,text='清空',command=self.clear).pack(side='left',padx=8)
        ttk.Button(left,text='价格设置',command=self.settings).pack(anchor='w',pady=12)
        self.detail=tk.Text(right,font=('Microsoft YaHei',12),bg='white',relief='flat',padx=12,pady=12,state='disabled'); self.detail.pack(fill='both',expand=True)
        self.copybtn=ttk.Button(right,text='复制报价结果',command=self.copy_result,state='disabled'); self.copybtn.pack(anchor='e',pady=(12,0)); self.result=''
    def clear(self): self.words.delete(0,'end'); self.words.insert(0,'10'); [v.set(False) for v in self.vars.values()]; self.show('')
    def show(self,s): self.detail.config(state='normal'); self.detail.delete('1.0','end'); self.detail.insert('1.0',s); self.detail.config(state='disabled')
    def calculate(self):
        try: w=float(self.words.get().strip())
        except: messagebox.showwarning('输入提示','请输入有效的千字数。'); return
        p=parse_price(self.cfg['education'][self.edu.get()]);
        if not p: messagebox.showwarning('价格提示','当前学历暂无基础价格。'); return
        base=w*p[0]; lines=[f'报价类型：{self.kind.get()}',f'学历类型：{self.edu.get()}',f'字数：{w:g} 千字',f'基础单价：{p[0]:g} 元/千字',f'基础费用：{base:.2f} 元','']
        total=base
        for n,v in self.vars.items():
            if v.get():
                sp=parse_price(self.cfg['services'][n]); fee=sp[0] if sp else 0; total+=fee; lines.append(f'{n}：{fee:.2f} 元')
        lines += ['',f'附加服务合计：{total-base:.2f} 元']
        if self.kind.get()=='一手单': total*=float(self.cfg.get('multiplier',2)); lines.append(f'一手单倍数：×{self.cfg.get("multiplier",2)}')
        lines += ['',f'最终报价：{total:.2f} 元']; self.result='\n'.join(lines); self.show(self.result); self.copybtn.config(state='normal')
    def copy_result(self): self.clipboard_clear(); self.clipboard_append(self.result); messagebox.showinfo('完成','报价明细已复制。')
    def settings(self):
        if simpledialog.askstring('价格设置验证','请输入设置密码：',show='*',parent=self)!=PASSWORD: messagebox.showerror('验证失败','密码错误。'); return
        win=tk.Toplevel(self); win.title('价格设置'); win.geometry('560x560'); win.transient(self); win.grab_set(); win.configure(bg='#f4f7fb')
        entries=[]; ttk.Label(win,text='学历价格（支持10-15、100+、固定值）').pack(anchor='w',padx=20,pady=12)
        for n,val in self.cfg['education'].items():
            f=tk.Frame(win,bg='#f4f7fb'); f.pack(fill='x',padx=20,pady=3); tk.Label(f,text=n,width=22,anchor='w',bg='#f4f7fb').pack(side='left'); e=ttk.Entry(f,width=18); e.insert(0,val); e.pack(side='left'); entries.append((self.cfg['education'],n,e))
        ttk.Label(win,text='附加服务价格').pack(anchor='w',padx=20,pady=(15,5))
        for n,val in self.cfg['services'].items():
            f=tk.Frame(win,bg='#f4f7fb'); f.pack(fill='x',padx=20,pady=3); tk.Label(f,text=n,width=22,anchor='w',bg='#f4f7fb').pack(side='left'); e=ttk.Entry(f,width=18); e.insert(0,val); e.pack(side='left'); entries.append((self.cfg['services'],n,e))
        def save():
            for d,n,e in entries: d[n]=e.get().strip()
            with open(CONFIG,'w',encoding='utf-8') as f: json.dump(self.cfg,f,ensure_ascii=False,indent=2)
            win.destroy(); messagebox.showinfo('保存成功','价格设置已保存。')
        ttk.Button(win,text='保存设置',command=save).pack(pady=22)

if __name__=='__main__': App().mainloop()
