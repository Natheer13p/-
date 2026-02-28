from flask import Flask, render_template, request, session, redirect, url_for
import pandas as pd
import json, os, glob, re

app = Flask(__name__)
app.secret_key = 'ebidi_school_portal_2024_xK9m'

DATA_DIR       = os.path.join(os.path.dirname(__file__), 'data')
PASSWORDS_FILE = os.path.join(DATA_DIR, 'passwords.json')
PERIODS = ['الفصل 1','نصف السنة','الفصل 2','السعي','الامتحان النهائي','الدرجة النهائية']
ADMIN_PASS = 'admin@ebidi2024'

def normalize(text):
    text = text.strip()
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'[إأآ]', 'ا', text)
    text = re.sub(r'ى', 'ي', text)
    text = re.sub(r'[\u064B-\u065F\u0670]', '', text)
    return text

def load_pw():
    if os.path.exists(PASSWORDS_FILE):
        with open(PASSWORDS_FILE,'r',encoding='utf-8') as f: return json.load(f)
    return {}

def save_pw(d):
    with open(PASSWORDS_FILE,'w',encoding='utf-8') as f:
        json.dump(d, f, ensure_ascii=False, indent=2)

def parse_excel(path):
    try: df = pd.read_excel(path, sheet_name=0, header=None)
    except: return [], [], '', ''
    school = str(df.iloc[1][0]).strip() if pd.notna(df.iloc[1][0]) else 'المدرسة'
    cls = next((str(df.iloc[1][c]).strip() for c in range(df.shape[1])
                if pd.notna(df.iloc[1][c]) and 'الصف' in str(df.iloc[1][c])), '')
    subjects = [str(df.iloc[2][c]).strip() for c in range(3, df.shape[1]-1)
                if pd.notna(df.iloc[2][c]) and str(df.iloc[2][c]).strip() not in ('nan','')]
    students, i = [], 3
    while i < len(df):
        row = df.iloc[i]
        if pd.notna(row[0]) and pd.notna(row[1]):
            try: num = int(float(str(row[0])))
            except: i+=1; continue
            grades = {}
            for j, period in enumerate(PERIODS):
                if i+j >= len(df): break
                r = df.iloc[i+j]
                pg = {s: (round(float(r[k+3]),2) if k+3<len(r) and pd.notna(r[k+3]) else None)
                      for k,s in enumerate(subjects)}
                last = df.shape[1]-1
                if last < len(r) and pd.notna(r[last]): pg['__total__'] = round(float(r[last]),2)
                grades[period] = pg
            students.append({'num':num,'name':str(row[1]).strip(),'grades':grades})
            i += 6
        else: i += 1
    return students, subjects, school, cls

def get_db():
    passwords = load_pw()
    db = {}
    for fpath in sorted(glob.glob(os.path.join(DATA_DIR,'*.xlsx'))):
        fname = os.path.basename(fpath)
        label = fname.replace('.xlsx','')
        students, subjects, school, cls = parse_excel(fpath)
        if not students: continue
        for s in students:
            key = f"{label}_{s['num']}"
            db[key] = {**s, 'key':key, 'password':passwords.get(key, str(s['num'])),
                       'file':fname, 'label':label, 'school':school, 'class':cls,
                       'subjects':subjects, 'name_norm':normalize(s['name'])}
    return db

def find_student(full_name, password):
    n = normalize(full_name)
    for s in get_db().values():
        if s['name_norm'] == n and s['password'] == password:
            return s
    return None

@app.route('/', methods=['GET','POST'])
def login():
    error = None
    if request.method == 'POST':
        name = request.form.get('fullname','').strip()
        pw   = request.form.get('password','').strip()
        if not name or not pw:
            error = 'يرجى إدخال الاسم الثلاثي وكلمة المرور'
        else:
            student = find_student(name, pw)
            if student:
                session['key'] = student['key']
                return redirect(url_for('dashboard'))
            error = 'الاسم أو كلمة المرور غير صحيحة — تأكد من كتابة اسمك الثلاثي كاملاً'
    return render_template('login.html', error=error)

@app.route('/dashboard')
def dashboard():
    if 'key' not in session: return redirect(url_for('login'))
    student = get_db().get(session['key'])
    if not student: session.clear(); return redirect(url_for('login'))
    return render_template('dashboard.html', student=student, periods=PERIODS)

@app.route('/logout')
def logout():
    session.clear(); return redirect(url_for('login'))

@app.route('/admin', methods=['GET','POST'])
def admin():
    if not session.get('is_admin'):
        if request.method=='POST' and request.form.get('admin_pass')==ADMIN_PASS:
            session['is_admin'] = True; return redirect(url_for('admin'))
        return render_template('admin_login.html', error=('كلمة المرور خاطئة' if request.method=='POST' else None))
    db = get_db(); pw = load_pw(); msg = None
    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'change_pw':
            key = request.form.get('key'); new_pw = request.form.get('new_pw','').strip()
            if key and new_pw and key in db:
                pw[key] = new_pw; save_pw(pw); msg = f'تم تغيير كلمة مرور {db[key]["name"]}'
        elif action == 'bulk_pw':
            label = request.form.get('label'); new_pw = request.form.get('bulk_new_pw','').strip()
            if label and new_pw:
                count = 0
                for k,s in db.items():
                    if s['label']==label: pw[k]=new_pw; count+=1
                save_pw(pw); msg = f'تم تغيير كلمة مرور {count} طالب'
    groups = {}
    for key,s in db.items():
        lbl = s['label']
        if lbl not in groups: groups[lbl]={'school':s['school'],'class':s['class'],'subjects':s['subjects'],'students':[]}
        groups[lbl]['students'].append({**s,'cur_pw':pw.get(s['key'],str(s['num']))})
    return render_template('admin.html', groups=groups, msg=msg)

@app.route('/admin/logout')
def admin_logout():
    session.pop('is_admin',None); return redirect(url_for('login'))

if __name__ == '__main__':
    os.makedirs(DATA_DIR, exist_ok=True)
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
