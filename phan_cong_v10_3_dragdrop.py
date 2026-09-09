# -*- coding: utf-8 -*-
"""
PHẦN MỀM PHÂN CÔNG CHUYÊN MÔN V10.3
Mô hình chính:
- Mỗi lớp có các thẻ môn.
- Mỗi giáo viên có các thẻ chuyên môn tương ứng.
- Kéo thẻ chuyên môn của GV vào thẻ môn của lớp để gán.
- Nếu vượt chuẩn / không đủ hiệu lực, giao diện cảnh báo trước khi chốt.
- Chuyên đề mặc định đi cùng môn; có thể "Tách CĐ" để gán riêng.
"""
import os, json, sqlite3, threading, webbrowser, urllib.parse, io, zipfile, base64, sys, datetime, math, re
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from xml.sax.saxutils import escape
import xml.etree.ElementTree as ET
try:
    import pymupdf as fitz
except Exception:
    try:
        import fitz
    except Exception:
        fitz=None

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if getattr(sys, "frozen", False):
    RESOURCE_DIR = getattr(sys, "_MEIPASS", BASE_DIR)
    DATA_DIR = os.path.dirname(sys.executable)
else:
    RESOURCE_DIR = BASE_DIR
    DATA_DIR = BASE_DIR
DB_PATH = os.path.join(DATA_DIR, "phan_cong_v10_3.db")
SEED_PATH = os.path.join(RESOURCE_DIR, "v10_3_seed.json")
INDEX_PATH = os.path.join(RESOURCE_DIR, "index_v10_3.html")
TKB_TEMPLATE_PATH = os.path.join(RESOURCE_DIR, "MAU_NHAP_TKB_EXCEL_V10_3.xlsx")
HOST = os.environ.get("HOST", "0.0.0.0")
PORT = int(os.environ.get("PORT", "8812"))

with open(SEED_PATH, "r", encoding="utf-8") as f:
    SEED = json.load(f)

def num(v, default=0.0):
    if v in (None, ""):
        return default
    try:
        return float(str(v).replace(",", ".").strip())
    except Exception:
        return default

def canonical_subject_name(s):
    x=(s or "").strip()
    k=x.lower().replace("–","-")
    aliases={
        "tn-hn":"TNHN","tnhn":"TNHN","hđtn":"TNHN","hdtn":"TNHN",
        "toán":"Toán học","toán học":"Toán học",
        "văn":"Ngữ văn","ngữ văn":"Ngữ văn",
        "lý":"Vật lý","vật lý":"Vật lý",
        "hóa":"Hóa học","hoá":"Hóa học","hóa học":"Hóa học","hoá học":"Hóa học",
        "sinh":"Sinh học","sinh học":"Sinh học",
        "tin":"Tin học","tin học":"Tin học",
        "sử":"Lịch sử","lịch sử":"Lịch sử",
        "địa":"Địa lí","địa lí":"Địa lí","địa lý":"Địa lí",
        "ktpl":"GDKTPL","gdktpl":"GDKTPL","gdkt-pl":"GDKTPL","gdkt - pl":"GDKTPL"
    }
    return aliases.get(k,x)

def grade_of(c):
    return str(c)[:2] if c else ""


def _xlsx_cell_value(cell, shared):
    t=cell.attrib.get("t")
    ns="{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    if t=="inlineStr":
        texts=[x.text or "" for x in cell.findall(".//"+ns+"t")]
        return "".join(texts)
    v=cell.find(ns+"v")
    if v is None:return ""
    raw=v.text or ""
    if t=="s":
        try:return shared[int(raw)]
        except:return raw
    try:
        if "." in raw:
            f=float(raw)
            return int(f) if f.is_integer() else f
        return int(raw)
    except:return raw

def parse_xlsx_rows(data):
    """Đọc XLSX bằng thư viện chuẩn Python; trả dict sheet -> rows."""
    z=zipfile.ZipFile(io.BytesIO(data))
    ns="{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    relns="{http://schemas.openxmlformats.org/package/2006/relationships}"
    shared=[]
    if "xl/sharedStrings.xml" in z.namelist():
        root=ET.fromstring(z.read("xl/sharedStrings.xml"))
        for si in root.findall(ns+"si"):
            shared.append("".join((t.text or "") for t in si.findall(".//"+ns+"t")))
    wb=ET.fromstring(z.read("xl/workbook.xml"))
    rel=ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid_to_target={r.attrib["Id"]:r.attrib["Target"] for r in rel.findall(relns+"Relationship")}
    out={}
    for sh in wb.find(ns+"sheets"):
        name=sh.attrib.get("name","Sheet")
        rid=sh.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        target=rid_to_target.get(rid,"")
        if not target:continue
        if target.startswith("/xl/"):
            target=target.lstrip("/")
        elif not target.startswith("xl/"):
            target="xl/"+target.lstrip("/")
        if target not in z.namelist():continue
        root=ET.fromstring(z.read(target))
        rows=[]
        for row in root.findall(".//"+ns+"row"):
            vals={}
            maxcol=0
            for c in row.findall(ns+"c"):
                ref=c.attrib.get("r","A1")
                letters="".join(ch for ch in ref if ch.isalpha())
                col=0
                for ch in letters:col=col*26+(ord(ch.upper())-64)
                maxcol=max(maxcol,col)
                vals[col-1]=_xlsx_cell_value(c,shared)
            arr=[""]*maxcol
            for i,v in vals.items():
                if i<len(arr):arr[i]=v
            rows.append(arr)
        out[name]=rows
    return out

def _norm_text(s):
    return re.sub(r"\s+"," ",str(s or "").strip()).lower()

def preview_import_xlsx(data):
    sheets=parse_xlsx_rows(data)
    # Chọn sheet có nhiều dòng chứa STT + họ tên nhất.
    best_name=None;best_rows=[];best_score=-1
    for name,rows in sheets.items():
        score=0
        for r in rows:
            if len(r)>=2 and str(r[0]).strip().isdigit() and len(str(r[1]).strip())>=4:
                score+=1
        if score>best_score:
            best_name=name;best_rows=rows;best_score=score

    teachers=[];classes=set()
    # Ưu tiên cấu trúc bảng phân công hiện tại của người dùng.
    for r in best_rows:
        if len(r)<2:continue
        stt=str(r[0]).strip()
        name=str(r[1]).strip()
        if not stt.isdigit() or len(name)<4:continue
        role=str(r[2]).strip() if len(r)>2 else ""
        role_credit=r[3] if len(r)>3 else 0
        homeroom=str(r[4]).strip().upper() if len(r)>4 else ""
        homeroom_credit=r[5] if len(r)>5 else 0
        subject=canonical_subject_name(str(r[6]).strip() if len(r)>6 else "")
        class_text=str(r[7]).strip() if len(r)>7 else ""
        standard=r[10] if len(r)>10 else 17
        note=str(r[12]).strip() if len(r)>12 else ""
        emp="Thỉnh giảng" if "thỉnh giảng" in role.lower() else "Cơ hữu"
        active_end=18 if "hki" in role.lower() and emp=="Thỉnh giảng" else 35
        can=[subject] if subject else []
        if re.search(r"TNHN|TN-HN",class_text,re.I) and "TNHN" not in can:can.append("TNHN")
        teachers.append({
            "stt":int(stt),"name":name,"primary_subject":subject,"can_teach":can,
            "standard":num(standard,17),"employment_type":emp,"active_start":1,"active_end":active_end,
            "role":role,"role_credit":num(role_credit),"homeroom":homeroom,
            "homeroom_credit":num(homeroom_credit),"note":note
        })
        if homeroom and re.match(r"^(10|11|12)[A-Z]$",homeroom):classes.add(homeroom)
        for code in re.findall(r"\b(?:10|11|12)[A-G]\b",class_text.upper()):
            classes.add(code)

    # Nếu không nhận được giáo viên theo mẫu, thử dò header đơn giản.
    if not teachers:
        for rows in sheets.values():
            header_i=None;name_col=None
            for i,r in enumerate(rows[:20]):
                for j,v in enumerate(r):
                    if "họ và tên" in _norm_text(v) or "ho va ten" in _norm_text(v):
                        header_i=i;name_col=j;break
                if header_i is not None:break
            if header_i is None:continue
            for r in rows[header_i+1:]:
                if name_col>=len(r):continue
                name=str(r[name_col]).strip()
                if len(name)<4:continue
                teachers.append({"stt":len(teachers)+1,"name":name,"primary_subject":"",
                    "can_teach":[],"standard":17,"employment_type":"Cơ hữu","active_start":1,
                    "active_end":35,"role":"","role_credit":0,"homeroom":"",
                    "homeroom_credit":0,"note":""})
            if teachers:break

    return {
        "sheet":best_name or "",
        "teachers":teachers,
        "classes":sorted(classes,key=lambda x:(int(x[:2]),x[2:])),
        "teacher_count":len(teachers),"class_count":len(classes),
        "sheet_names":list(sheets.keys())
    }


def preview_timetable_xlsx(data):
    sheets=parse_xlsx_rows(data)
    synonyms={
        "class_name":["lớp","lop","class","class_name"],
        "teacher":["giáo viên","giao vien","gv","teacher","họ tên gv"],
        "subject":["môn","mon","subject"],
        "session":["buổi","buoi","session"],
        "day":["thứ","thu","ngày","ngay","day"],
        "period":["tiết","tiet","period"],
        "start_week":["từ tuần","tu tuan","start_week"],
        "end_week":["đến tuần","den tuan","end_week"]
    }
    def find_col(headers,names):
        hs=[_norm_text(x) for x in headers]
        for i,h in enumerate(hs):
            if any(n in h for n in names): return i
        return None

    best=None
    for sname,rows in sheets.items():
        for hi,r in enumerate(rows[:25]):
            cols={k:find_col(r,v) for k,v in synonyms.items()}
            score=sum(cols[k] is not None for k in ("class_name","teacher","subject","day","period"))
            if score>=4 and (best is None or score>best[0]):
                best=(score,sname,rows,hi,cols)
    entries=[]
    if best:
        _,sname,rows,hi,cols=best
        for r in rows[hi+1:]:
            def val(k,default=""):
                c=cols.get(k)
                return r[c] if c is not None and c<len(r) else default
            cname=str(val("class_name","")).strip().upper()
            teacher=str(val("teacher","")).strip()
            subject=canonical_subject_name(str(val("subject","")).strip())
            session=str(val("session","")).strip()
            day=str(val("day","")).strip()
            period=int(num(val("period",0),0))
            sw=int(num(val("start_week",1),1)); ew=int(num(val("end_week",35),35))
            if not cname or not teacher or not subject or not day or period<=0: continue
            # Khớp tên GV bị Excel rút gọn, ví dụ "Nguyễn Thị Huyền Tr" -> "Nguyễn Thị Huyền Trang".
            try:
                names=[t["name"] for t in STORE.teachers()]
                exact=[n for n in names if n.lower()==teacher.lower()]
                prefix=[n for n in names if n.lower().startswith(teacher.lower()) or teacher.lower().startswith(n.lower())]
                if exact: teacher=exact[0]
                elif len(prefix)==1: teacher=prefix[0]
            except Exception:
                pass
            entries.append({"class_name":cname,"teacher":teacher,"subject":subject,
                            "session":session,"day":day,"period":period,"start_week":sw,"end_week":ew,
                            "source":sname,"note":""})
        return {"sheet":sname,"entries":entries,"count":len(entries),
                "sheet_names":list(sheets.keys()),"mode":"header"}
    return {"sheet":"","entries":[],"count":0,"sheet_names":list(sheets.keys()),"mode":"not_found"}


def _fold_vn(s):
    import unicodedata
    x=unicodedata.normalize("NFD",str(s or "").lower())
    x="".join(c for c in x if unicodedata.category(c)!="Mn").replace("đ","d")
    return re.sub(r"[^a-z0-9]+"," ",x).strip()

def _strip_teacher_title(s):
    x=_fold_vn(s)
    return re.sub(r"^(co|thay)\s+","",x).strip()

def _is_tkb_admin_subject(subject):
    k=_fold_vn(subject)
    return k in {"chao co","shl","sinh hoat lop","sinh hoat","sinh hoat duoi co","shdc"}

def _resolve_pdf_teacher(display_name, subject_names):
    """Khớp tên ngắn trên TKB PDF với tên đầy đủ trong đội ngũ."""
    try:
        staff=STORE.teachers()
    except Exception:
        return display_name, []
    key=_strip_teacher_title(display_name); toks=key.split()
    ranked=[]
    for t in staff:
        nt=_fold_vn(t.get("name","")); ntt=nt.split()
        rank=0
        if toks and len(ntt)>=len(toks) and ntt[-len(toks):]==toks:
            rank=3
        elif toks and all(tok in ntt for tok in toks):
            rank=2
        elif len(toks)==1 and ntt and toks[0]==ntt[-1]:
            rank=1
        if rank:
            ranked.append((rank,t))
    if not ranked:
        return display_name, []
    mr=max(r for r,_ in ranked); cand=[t for r,t in ranked if r==mr]
    if len(cand)>1:
        teaching=[canonical_subject_name(s) for s in subject_names if not _is_tkb_admin_subject(s)]
        teaching=[s for s in teaching if s!="TNHN"] or teaching
        dom=max(set(teaching),key=teaching.count) if teaching else ""
        if dom:
            c2=[]
            for t in cand:
                can=[canonical_subject_name(x) for x in t.get("can_teach_list",[]) or []]
                if canonical_subject_name(t.get("primary_subject",""))==dom or dom in can:
                    c2.append(t)
            if len(c2)==1:cand=c2
    if len(cand)==1:
        return cand[0]["name"], [cand[0]["name"]]
    return display_name,[t["name"] for t in cand]

def preview_timetable_pdf(data,start_week=1,end_week=35):
    """Đọc trực tiếp PDF TKB dạng bảng theo giáo viên như mẫu AD 07-09."""
    if fitz is None:
        raise ValueError("Chưa có thư viện đọc PDF. Hãy chạy lại CHAY_PHAN_CONG_V10_3.bat để cài PyMuPDF.")
    sw=int(start_week or 1);ew=int(end_week or 35)
    if not 1<=sw<=ew<=35:raise ValueError("Khoảng tuần áp dụng PDF không hợp lệ.")
    doc=fitz.open(stream=data,filetype="pdf")
    raw_entries=[];block_names=[];effective_date=""
    date_re=re.compile(r"\b(\d{2}-\d{2}-\d{4})\b")
    for pi,page in enumerate(doc):
        words=page.get_text("words")
        if not words:continue
        if not effective_date:
            m=date_re.search(page.get_text("text"));effective_date=m.group(1) if m else ""
        scale=page.rect.width/595.28 if page.rect.width else 1.0
        title_ys=sorted({w[1] for w in words if w[4]=="THỜI" and 55*scale<w[0]<115*scale})
        header_ys=sorted({w[1] for w in words if w[4]=="Giáo" and w[0]<125*scale})
        for hy in header_ys:
            next_titles=[y for y in title_ys if y>hy+2]
            endy=(min(next_titles)-1.0 if next_titles else page.rect.height-2)
            body=[w for w in words if w[1]>hy+7 and w[1]<endy]
            tw=[w for w in body if w[0]<145*scale and not re.fullmatch(r"[SC1-5]",w[4])]
            display=" ".join(w[4] for w in sorted(tw,key=lambda x:(round(x[1],1),x[0]))).strip()
            if not display:continue
            block_names.append(display)
            nums=sorted([w for w in body if 178*scale<=w[0]<=210*scale and re.fullmatch(r"[1-5]",w[4])],key=lambda x:x[1])
            rows=[];session_idx=0;prev=None
            for nw in nums:
                per=int(nw[4])
                if prev is not None and per<=prev:session_idx+=1
                session="Sáng" if session_idx==0 else "Chiều"
                rows.append(((nw[1]+nw[3])/2,session,per));prev=per
            ranges=[("Thứ 2",205,266),("Thứ 3",266,326),("Thứ 4",326,386),("Thứ 5",386,446),("Thứ 6",446,507),("Thứ 7",507,570)]
            for yc,session,period in rows:
                for day,x0,x1 in ranges:
                    cell=[w for w in body if x0*scale<=((w[0]+w[2])/2)<x1*scale and abs(((w[1]+w[3])/2)-yc)<3.8]
                    if not cell:continue
                    txt=" ".join(w[4] for w in sorted(cell,key=lambda x:x[0])).strip()
                    mm=re.match(r"^((?:10|11|12)[A-Z])-(.+)$",txt)
                    if not mm:continue
                    raw_entries.append({"teacher_display":display,"class_name":mm.group(1),
                        "subject":canonical_subject_name(mm.group(2).strip()),"session":session,
                        "day":day,"period":period,"start_week":sw,"end_week":ew,
                        "source":f"PDF GV {effective_date or 'TKB'}","note":""})
    if not raw_entries:
        raise ValueError("Không nhận diện được TKB PDF. Mẫu PDF cần có cột Giáo Viên - Buổi - Tiết - Thứ 2...Thứ 7 và ô dạng 12A-Toán.")

    by_teacher={}
    for e in raw_entries:by_teacher.setdefault(e["teacher_display"],[]).append(e["subject"])
    mapping=[];resolved={};unresolved=[]
    for display,subs in by_teacher.items():
        full,candidates=_resolve_pdf_teacher(display,subs)
        resolved[display]=full
        ok=full!=display or any(t.get("name")==display for t in STORE.teachers())
        if not ok:unresolved.append(display)
        mapping.append({"pdf_name":display,"resolved_name":full if ok else "",
                        "candidates":candidates,"status":"resolved" if ok else "unresolved"})
    entries=[]
    for e in raw_entries:
        x=dict(e);x["teacher"]=resolved.get(e["teacher_display"],e["teacher_display"]);x.pop("teacher_display",None)
        entries.append(x)
    return {"sheet":"PDF TKB giáo viên","entries":entries,"count":len(entries),
            "mode":"pdf_teacher_grid","effective_date":effective_date,
            "teacher_count":len(by_teacher),"teacher_mapping":mapping,
            "unresolved_teachers":unresolved,"start_week":sw,"end_week":ew,
            "page_count":len(doc)}


class Store:
    def __init__(self):
        self.cx = sqlite3.connect(DB_PATH, check_same_thread=False)
        self.cx.row_factory = sqlite3.Row
        self.lock = threading.RLock()
        self._program_rules_cache = None
        self._program_rules_cache_changes = -1
        self._workload_matrix_cache = {}
        self.setup()

    def setup(self):
        with self.lock:
            c = self.cx.cursor()
            c.execute("""CREATE TABLE IF NOT EXISTS academic_years(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                is_active INTEGER DEFAULT 0
            )""")
            c.execute("""CREATE TABLE IF NOT EXISTS settings(
                key TEXT PRIMARY KEY, value TEXT
            )""")
            c.execute("""CREATE TABLE IF NOT EXISTS teachers(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                year TEXT NOT NULL,
                stt INTEGER,
                name TEXT NOT NULL,
                primary_subject TEXT,
                can_teach TEXT,
                standard REAL DEFAULT 17,
                employment_type TEXT DEFAULT 'Cơ hữu',
                active_start INTEGER DEFAULT 1,
                active_end INTEGER DEFAULT 35,
                role TEXT,
                role_credit REAL DEFAULT 0,
                homeroom TEXT,
                homeroom_credit REAL DEFAULT 0,
                note TEXT,
                active INTEGER DEFAULT 1,
                UNIQUE(year,name)
            )""")
            c.execute("""CREATE TABLE IF NOT EXISTS classes(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                year TEXT NOT NULL,
                class_name TEXT NOT NULL,
                grade INTEGER NOT NULL,
                program_name TEXT,
                template_name TEXT,
                customized INTEGER DEFAULT 0,
                homeroom_teacher TEXT,
                active INTEGER DEFAULT 1,
                UNIQUE(year,class_name)
            )""")
            c.execute("""CREATE TABLE IF NOT EXISTS templates(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                year TEXT NOT NULL,
                name TEXT NOT NULL,
                grade INTEGER NOT NULL,
                active INTEGER DEFAULT 1,
                note TEXT DEFAULT '',
                UNIQUE(year,name)
            )""")
            c.execute("""CREATE TABLE IF NOT EXISTS template_items(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                template_id INTEGER NOT NULL,
                subject TEXT NOT NULL,
                core INTEGER DEFAULT 1,
                specialty INTEGER DEFAULT 0,
                UNIQUE(template_id,subject),
                FOREIGN KEY(template_id) REFERENCES templates(id)
            )""")
            c.execute("""CREATE TABLE IF NOT EXISTS program_rules(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                subject TEXT UNIQUE NOT NULL,
                base_hours REAL DEFAULT 0,
                specialty_hours REAL DEFAULT 0,
                can_specialty INTEGER DEFAULT 0,
                source TEXT,
                core_schedule_json TEXT DEFAULT '[]',
                specialty_schedule_json TEXT DEFAULT '[]',
                legal_locked INTEGER DEFAULT 1
            )""")
            c.execute("""CREATE TABLE IF NOT EXISTS requirements(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                year TEXT NOT NULL,
                class_name TEXT NOT NULL,
                subject TEXT NOT NULL,
                req_type TEXT NOT NULL,
                annual_hours REAL DEFAULT 0,
                teacher TEXT DEFAULT '',
                start_week INTEGER DEFAULT 1,
                end_week INTEGER DEFAULT 35,
                note TEXT DEFAULT '',
                UNIQUE(year,class_name,subject,req_type)
            )""")
            c.execute("""CREATE TABLE IF NOT EXISTS review_jobs(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                year TEXT NOT NULL,
                teacher TEXT NOT NULL,
                subject TEXT DEFAULT '',
                classes_json TEXT DEFAULT '[]',
                periods_per_week REAL DEFAULT 3,
                start_week INTEGER,
                end_week INTEGER DEFAULT 35,
                note TEXT
            )""")
            c.execute("""CREATE TABLE IF NOT EXISTS duties(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                year TEXT NOT NULL,
                teacher TEXT NOT NULL,
                duty_type TEXT NOT NULL,
                duty_name TEXT NOT NULL,
                class_name TEXT DEFAULT '',
                credit REAL DEFAULT 0,
                standard_override REAL,
                start_week INTEGER DEFAULT 1,
                end_week INTEGER DEFAULT 35,
                note TEXT DEFAULT ''
            )""")
            c.execute("""CREATE TABLE IF NOT EXISTS locks(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                year TEXT NOT NULL,
                scope_type TEXT NOT NULL,
                scope_key TEXT NOT NULL,
                note TEXT DEFAULT '',
                created_at TEXT DEFAULT '',
                UNIQUE(year,scope_type,scope_key)
            )""")
            c.execute("""CREATE TABLE IF NOT EXISTS audit_log(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                year TEXT NOT NULL,
                scenario_id INTEGER,
                created_at TEXT NOT NULL,
                actor TEXT DEFAULT '',
                action TEXT NOT NULL,
                before_json TEXT NOT NULL,
                after_json TEXT NOT NULL,
                undone INTEGER DEFAULT 0
            )""")
            c.execute("""CREATE TABLE IF NOT EXISTS scenarios(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                year TEXT NOT NULL,
                name TEXT NOT NULL,
                description TEXT DEFAULT '',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                data_json TEXT NOT NULL,
                metrics_json TEXT DEFAULT '{}',
                is_baseline INTEGER DEFAULT 0,
                workflow_status TEXT DEFAULT 'DRAFT',
                approved_by TEXT DEFAULT '',
                approved_at TEXT DEFAULT '',
                UNIQUE(year,name)
            )""")
            c.execute("""CREATE TABLE IF NOT EXISTS staff_events(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                year TEXT NOT NULL,
                teacher TEXT NOT NULL,
                event_type TEXT NOT NULL,
                start_week INTEGER DEFAULT 1,
                end_week INTEGER DEFAULT 35,
                blocks_teaching INTEGER DEFAULT 1,
                replacement_teacher TEXT DEFAULT '',
                note TEXT DEFAULT ''
            )""")
            c.execute("""CREATE TABLE IF NOT EXISTS timetable_entries(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                year TEXT NOT NULL,
                class_name TEXT NOT NULL,
                teacher TEXT NOT NULL,
                subject TEXT NOT NULL,
                session TEXT DEFAULT '',
                day TEXT NOT NULL,
                period INTEGER NOT NULL,
                start_week INTEGER DEFAULT 1,
                end_week INTEGER DEFAULT 35,
                source TEXT DEFAULT '',
                note TEXT DEFAULT ''
            )""")
            c.execute("""CREATE TABLE IF NOT EXISTS workflow_log(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                year TEXT NOT NULL,
                scenario_id INTEGER NOT NULL,
                created_at TEXT NOT NULL,
                actor TEXT DEFAULT '',
                old_status TEXT DEFAULT '',
                new_status TEXT NOT NULL,
                note TEXT DEFAULT ''
            )""")

            c.execute("INSERT OR IGNORE INTO settings(key,value) VALUES('active_scenario_id','null')")
            c.execute("INSERT OR IGNORE INTO settings(key,value) VALUES('last_import_preview','null')")
            c.execute("INSERT OR IGNORE INTO settings(key,value) VALUES('last_timetable_preview','null')")
            c.execute("INSERT OR IGNORE INTO settings(key,value) VALUES('operator_name','')")
            c.execute("INSERT OR IGNORE INTO settings(key,value) VALUES('hdtn_tkb_slots','2')")
            self.cx.commit()

            if c.execute("SELECT COUNT(*) FROM academic_years").fetchone()[0] == 0:
                year = SEED["active_year"]
                c.execute("INSERT INTO academic_years(name,is_active) VALUES(?,1)", (year,))
                for k,v in SEED["settings"].items():
                    c.execute("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)",
                              (k,json.dumps(v,ensure_ascii=False)))
                for t in SEED["teachers"]:
                    c.execute("""INSERT INTO teachers(
                        year,stt,name,primary_subject,can_teach,standard,employment_type,
                        active_start,active_end,role,role_credit,homeroom,homeroom_credit,note,active
                    ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(
                        year,t["stt"],t["name"],t["primary_subject"],
                        json.dumps(t["can_teach"],ensure_ascii=False),
                        num(t["standard"],17),t["employment_type"],t["active_start"],t["active_end"],
                        t["role"],num(t["role_credit"]),t["homeroom"],num(t["homeroom_credit"]),
                        t["note"],int(t["active"])
                    ))
                for tp in SEED.get("templates", []):
                    c.execute("""INSERT INTO templates(year,name,grade,active,note)
                        VALUES(?,?,?,?,?)""",(year,tp["name"],tp["grade"],int(tp.get("active",1)),tp.get("note","")))
                    tid=c.lastrowid
                    for it in tp.get("items",[]):
                        c.execute("""INSERT INTO template_items(template_id,subject,core,specialty)
                            VALUES(?,?,?,?)""",(tid,it["subject"],int(it.get("core",1)),int(it.get("specialty",0))))
                for cl in SEED["classes"]:
                    c.execute("""INSERT INTO classes(
                        year,class_name,grade,program_name,template_name,customized,homeroom_teacher,active
                    ) VALUES(?,?,?,?,?,?,?,?)""",(
                        year,cl["class_name"],cl["grade"],cl.get("program_name",""),
                        cl.get("template_name",cl.get("program_name","")),0,
                        cl["homeroom_teacher"],int(cl["active"])
                    ))
                for r in SEED["program_rules"]:
                    c.execute("""INSERT INTO program_rules(
                        subject,base_hours,specialty_hours,can_specialty,source,
                        core_schedule_json,specialty_schedule_json,legal_locked
                    ) VALUES(?,?,?,?,?,?,?,?)""",(r["subject"],r["base_hours"],r["specialty_hours"],
                                           int(r["can_specialty"]),r["source"],
                                           json.dumps(r.get("core_schedule",[]),ensure_ascii=False),
                                           json.dumps(r.get("specialty_schedule",[]),ensure_ascii=False),
                                           int(r.get("legal_locked",1))))
                for req in SEED["requirements"]:
                    c.execute("""INSERT OR IGNORE INTO requirements(
                        year,class_name,subject,req_type,annual_hours,teacher,start_week,end_week,note
                    ) VALUES(?,?,?,?,?,?,?,?,?)""",(
                        year,req["class_name"],req["subject"],req["req_type"],req["annual_hours"],
                        req["teacher"],req["start_week"],req["end_week"],req["note"]
                    ))
                for r in SEED["review_jobs"]:
                    c.execute("""INSERT INTO review_jobs(
                        year,teacher,subject,classes_json,periods_per_week,start_week,end_week,note
                    ) VALUES(?,?,?,?,?,?,?,?)""",(year,r["teacher"],r.get("subject",""),
                                                 json.dumps(r.get("classes",[]),ensure_ascii=False),
                                                 num(r.get("periods_per_week"),3),
                                                 r.get("start_week"),r.get("end_week",35),r.get("note","")))

                # Chuyển CN/KN cũ thành lịch sử nhiệm vụ tuần 1-35.
                for t in SEED["teachers"]:
                    role=(t.get("role") or "").strip()
                    role_credit=num(t.get("role_credit"))
                    if role and "thỉnh giảng" not in role.lower():
                        base_std=num(t.get("standard"),17)
                        duty_type="Kiêm nhiệm" if role_credit else "Chức vụ"
                        standard_override=base_std if (not role_credit and base_std<17) else None
                        c.execute("""INSERT INTO duties(
                            year,teacher,duty_type,duty_name,class_name,credit,standard_override,
                            start_week,end_week,note
                        ) VALUES(?,?,?,?,?,?,?,?,?,?)""",
                        (year,t["name"],duty_type,role,"",role_credit,standard_override,
                         1,35,"Khởi tạo từ dữ liệu cũ"))
                    homeroom=(t.get("homeroom") or "").strip()
                    homeroom_credit=num(t.get("homeroom_credit"))
                    if homeroom and homeroom_credit:
                        c.execute("""INSERT INTO duties(
                            year,teacher,duty_type,duty_name,class_name,credit,standard_override,
                            start_week,end_week,note
                        ) VALUES(?,?,?,?,?,?,?,?,?,?)""",
                        (year,t["name"],"Chủ nhiệm","Chủ nhiệm "+homeroom,homeroom,
                         homeroom_credit,None,1,35,"Khởi tạo từ dữ liệu cũ"))
                self.cx.commit()

            # Bảo đảm quy tắc Lịch sử đúng cho cả DB mới và DB được nâng cấp.
            self.ensure_v10_3_history_schedule()

    def ensure_v10_3_history_schedule(self):
        """Quy tắc Trung tâm: Lịch sử T1-17 = 2 tiết/tuần; T18-35 = 1 tiết/tuần."""
        desired=[{"start":1,"end":17,"periods":2},{"start":18,"end":35,"periods":1}]
        with self.lock:
            row=self.cx.execute("SELECT core_schedule_json FROM program_rules WHERE subject='Lịch sử'").fetchone()
            if row:
                try:cur=json.loads(row["core_schedule_json"] or "[]")
                except:cur=[]
                if cur!=desired:
                    self.cx.execute("UPDATE program_rules SET core_schedule_json=? WHERE subject='Lịch sử'",
                                    (json.dumps(desired,ensure_ascii=False),))
                    self.cx.commit()

    def setting(self,key,default=None):
        with self.lock:
            row=self.cx.execute("SELECT value FROM settings WHERE key=?",(key,)).fetchone()
            if not row:return default
            try:return json.loads(row["value"])
            except:return row["value"]

    def set_setting(self,key,value):
        with self.lock:
            self.cx.execute("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)",
                            (key,json.dumps(value,ensure_ascii=False)))
            self.cx.commit()

    def _rows(self,sql,params=()):
        with self.lock:
            return [dict(r) for r in self.cx.execute(sql,params)]

    def _insert_rows(self,table,rows):
        if not rows:return
        for row in rows:
            cols=list(row.keys())
            sql=f'INSERT INTO {table}('+",".join(cols)+') VALUES('+",".join("?" for _ in cols)+')'
            self.cx.execute(sql,tuple(row[c] for c in cols))

    def full_backup(self):
        tables=["academic_years","settings","teachers","templates","template_items","classes",
                "program_rules","requirements","review_jobs","duties","locks","audit_log","scenarios",
                "staff_events","timetable_entries","workflow_log"]
        data={}
        with self.lock:
            for table in tables:
                data[table]=[dict(r) for r in self.cx.execute(f"SELECT * FROM {table}")]
        return {"schema":"PHAN_CONG_V10_BACKUP_1","created_at":datetime.datetime.now().isoformat(timespec="seconds"),
                "active_year":self.active_year(),"tables":data}

    def restore_full_backup(self,backup):
        if not isinstance(backup,dict) or backup.get("schema")!="PHAN_CONG_V10_BACKUP_1":
            raise ValueError("Bản sao lưu toàn hệ thống không hợp lệ.")
        data=backup.get("tables") or {}
        delete_order=["workflow_log","audit_log","scenarios","timetable_entries","staff_events","locks",
                      "duties","review_jobs","requirements","classes","template_items","templates","teachers",
                      "program_rules","settings","academic_years"]
        insert_order=["academic_years","settings","teachers","templates","template_items","classes",
                      "program_rules","requirements","review_jobs","duties","locks","audit_log","scenarios",
                      "staff_events","timetable_entries","workflow_log"]
        with self.lock:
            self.cx.execute("PRAGMA foreign_keys=OFF")
            try:
                for table in delete_order:self.cx.execute(f"DELETE FROM {table}")
                for table in insert_order:self._insert_rows(table,data.get(table,[]))
                self.cx.commit()
            finally:
                self.cx.execute("PRAGMA foreign_keys=ON")
        self.ensure_v10_3_history_schedule()
        return {"active_year":self.active_year(),"years":len(self.years())}

    def full_snapshot(self,year=None):
        year=year or self.active_year()
        with self.lock:
            templates=[dict(r) for r in self.cx.execute("SELECT * FROM templates WHERE year=?",(year,))]
            tids=[r["id"] for r in templates]
            template_items=[]
            if tids:
                marks=",".join("?" for _ in tids)
                template_items=[dict(r) for r in self.cx.execute(
                    f"SELECT * FROM template_items WHERE template_id IN ({marks})",tids)]
            return {
                "year":year,
                "teachers":[dict(r) for r in self.cx.execute("SELECT * FROM teachers WHERE year=?",(year,))],
                "classes":[dict(r) for r in self.cx.execute("SELECT * FROM classes WHERE year=?",(year,))],
                "templates":templates,
                "template_items":template_items,
                "program_rules":[dict(r) for r in self.cx.execute("SELECT * FROM program_rules")],
                "requirements":[dict(r) for r in self.cx.execute("SELECT * FROM requirements WHERE year=?",(year,))],
                "review_jobs":[dict(r) for r in self.cx.execute("SELECT * FROM review_jobs WHERE year=?",(year,))],
                "duties":[dict(r) for r in self.cx.execute("SELECT * FROM duties WHERE year=?",(year,))],
                "locks":[dict(r) for r in self.cx.execute("SELECT * FROM locks WHERE year=?",(year,))],
                "staff_events":[dict(r) for r in self.cx.execute("SELECT * FROM staff_events WHERE year=?",(year,))],
                "timetable_entries":[dict(r) for r in self.cx.execute("SELECT * FROM timetable_entries WHERE year=?",(year,))]
            }

    def restore_snapshot(self,snapshot):
        year=snapshot.get("year") or self.active_year()
        with self.lock:
            old_tids=[r["id"] for r in self.cx.execute("SELECT id FROM templates WHERE year=?",(year,))]
            if old_tids:
                marks=",".join("?" for _ in old_tids)
                self.cx.execute(f"DELETE FROM template_items WHERE template_id IN ({marks})",old_tids)
            for table in ["requirements","review_jobs","duties","locks","staff_events","timetable_entries","classes","teachers","templates"]:
                self.cx.execute(f"DELETE FROM {table} WHERE year=?",(year,))
            self.cx.execute("DELETE FROM program_rules")
            self._insert_rows("teachers",snapshot.get("teachers",[]))
            self._insert_rows("classes",snapshot.get("classes",[]))
            self._insert_rows("templates",snapshot.get("templates",[]))
            self._insert_rows("template_items",snapshot.get("template_items",[]))
            self._insert_rows("program_rules",snapshot.get("program_rules",[]))
            self._insert_rows("requirements",snapshot.get("requirements",[]))
            self._insert_rows("review_jobs",snapshot.get("review_jobs",[]))
            self._insert_rows("duties",snapshot.get("duties",[]))
            self._insert_rows("locks",snapshot.get("locks",[]))
            self._insert_rows("staff_events",snapshot.get("staff_events",[]))
            self._insert_rows("timetable_entries",snapshot.get("timetable_entries",[]))
            self.cx.commit()
        self.invalidate_caches()

    def log_action(self,action,before,after):
        year=self.active_year()
        self.ensure_baseline()
        sid=self.active_scenario_id()
        now=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with self.lock:
            actor=str(self.setting("operator_name","") or "")
            self.cx.execute("""INSERT INTO audit_log(
                year,scenario_id,created_at,actor,action,before_json,after_json,undone
            ) VALUES(?,?,?,?,?,?,?,0)""",(year,sid,now,actor,action,
                json.dumps(before,ensure_ascii=False),
                json.dumps(after,ensure_ascii=False)))
            self.cx.commit()

    def history(self,limit=100):
        year=self.active_year()
        self.ensure_baseline()
        sid=self.active_scenario_id()
        with self.lock:
            rows=[dict(r) for r in self.cx.execute("""SELECT id,year,scenario_id,created_at,actor,action,undone
                FROM audit_log WHERE year=? AND scenario_id=? ORDER BY id DESC LIMIT ?""",
                (year,sid,int(limit)))]
        return rows

    def undo_last(self):
        year=self.active_year()
        self.ensure_baseline()
        sid=self.active_scenario_id()
        with self.lock:
            row=self.cx.execute("""SELECT * FROM audit_log
                WHERE year=? AND scenario_id=? AND undone=0 ORDER BY id DESC LIMIT 1""",
                (year,sid)).fetchone()
        if not row:
            raise ValueError("Không còn thao tác nào để hoàn tác.")
        snap=json.loads(row["before_json"])
        self.restore_snapshot(snap)
        with self.lock:
            self.cx.execute("UPDATE audit_log SET undone=1 WHERE id=?",(row["id"],))
            self.cx.commit()
        self.save_active_scenario()
        return {"action":row["action"],"created_at":row["created_at"]}

    # ---------- Khóa phân công ----------
    def locks(self,year=None):
        year=year or self.active_year()
        return self._rows("SELECT * FROM locks WHERE year=? ORDER BY scope_type,scope_key",(year,))

    def toggle_lock(self,scope_type,scope_key,note=""):
        year=self.active_year()
        scope_type=(scope_type or "").strip()
        scope_key=(scope_key or "").strip()
        if scope_type not in ("class","subject","requirement"):
            raise ValueError("Loại khóa không hợp lệ.")
        if not scope_key:raise ValueError("Thiếu đối tượng cần khóa.")
        with self.lock:
            row=self.cx.execute("""SELECT id FROM locks
                WHERE year=? AND scope_type=? AND scope_key=?""",(year,scope_type,scope_key)).fetchone()
            if row:
                self.cx.execute("DELETE FROM locks WHERE id=?",(row["id"],))
                locked=False
            else:
                self.cx.execute("""INSERT INTO locks(year,scope_type,scope_key,note,created_at)
                    VALUES(?,?,?,?,?)""",(year,scope_type,scope_key,note,
                    datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                locked=True
            self.cx.commit()
        return {"locked":locked}

    def requirement_lock_key(self,class_name,subject,req_type):
        return f"{class_name}|{subject}|{req_type}"

    def is_locked(self,class_name,subject=None,req_type=None):
        year=self.active_year()
        with self.lock:
            if self.cx.execute("""SELECT 1 FROM locks
                WHERE year=? AND scope_type='class' AND scope_key=?""",(year,class_name)).fetchone():
                return True
            if subject:
                skey=f"{class_name}|{subject}"
                if self.cx.execute("""SELECT 1 FROM locks
                    WHERE year=? AND scope_type='subject' AND scope_key=?""",(year,skey)).fetchone():
                    return True
            if subject and req_type:
                key=self.requirement_lock_key(class_name,subject,req_type)
                if self.cx.execute("""SELECT 1 FROM locks
                    WHERE year=? AND scope_type='requirement' AND scope_key=?""",(year,key)).fetchone():
                    return True
        return False

    # ---------- Phương án thử ----------
    def plan_metrics(self):
        year=self.active_year()
        reqs=self.requirements(year)
        teachers=self.teachers(year,active_only=True)
        unassigned=sum(1 for r in reqs if not (r.get("teacher") or "").strip())
        wrong=0
        thinh_giang_reqs=0
        tmap={t["name"]:t for t in self.teachers(year)}
        for r in reqs:
            tn=(r.get("teacher") or "").strip()
            if not tn:continue
            t=tmap.get(tn)
            if not t:
                wrong+=1;continue
            if r["subject"] not in t.get("can_teach_list",[]) and r["subject"]!=t.get("primary_subject"):
                wrong+=1
            if "Thỉnh giảng" in (t.get("employment_type") or ""):
                thinh_giang_reqs+=1
        over_count=0;sum_over=0;max_over=0
        max_loads=[]
        for t in teachers:
            mx=self.teacher_max_load(t["name"],year)
            if not mx:continue
            max_loads.append(num(mx["total"]))
            if mx["diff"]>1e-6:
                over_count+=1
                sum_over+=mx["diff"]
                max_over=max(max_over,mx["diff"])
        spread=(max(max_loads)-min(max_loads)) if len(max_loads)>=2 else 0
        return {
            "coverage_percent":((len(reqs)-unassigned)*100/len(reqs) if reqs else 0),
            "unassigned":unassigned,
            "over_count":over_count,
            "sum_over":sum_over,
            "max_over":max_over,
            "wrong_subject":wrong,
            "thinh_giang_requirements":thinh_giang_reqs,
            "load_spread":spread
        }

    def active_scenario_id(self):
        v=self.setting("active_scenario_id",None)
        return int(v) if v not in (None,"","null") else None

    def ensure_baseline(self):
        year=self.active_year()
        with self.lock:
            row=self.cx.execute("""SELECT * FROM scenarios
                WHERE year=? AND is_baseline=1 LIMIT 1""",(year,)).fetchone()
        if row:
            active=self.active_scenario_id()
            with self.lock:
                active_row=self.cx.execute("SELECT id FROM scenarios WHERE id=? AND year=?",
                                           (active or -1,year)).fetchone()
            if not active_row:self.set_setting("active_scenario_id",row["id"])
            return dict(row)
        now=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        snap=self.full_snapshot(year);metrics=self.plan_metrics()
        with self.lock:
            self.cx.execute("""INSERT INTO scenarios(
                year,name,description,created_at,updated_at,data_json,metrics_json,is_baseline
            ) VALUES(?,?,?,?,?,?,?,1)""",(year,"BẢN CHÍNH",
                "Phương án chính được bảo vệ",now,now,
                json.dumps(snap,ensure_ascii=False),json.dumps(metrics,ensure_ascii=False)))
            sid=self.cx.execute("SELECT last_insert_rowid()").fetchone()[0]
            self.cx.commit()
        self.set_setting("active_scenario_id",sid)
        return dict(self.cx.execute("SELECT * FROM scenarios WHERE id=?",(sid,)).fetchone())

    def scenarios(self):
        year=self.active_year()
        self.ensure_baseline()
        active=self.active_scenario_id()
        with self.lock:
            rows=[dict(r) for r in self.cx.execute("""SELECT id,year,name,description,created_at,
                updated_at,metrics_json,is_baseline,workflow_status,approved_by,approved_at FROM scenarios
                WHERE year=? ORDER BY is_baseline DESC,id""",(year,))]
        for r in rows:
            try:r["metrics"]=json.loads(r["metrics_json"] or "{}")
            except:r["metrics"]={}
            r["active"]=int(r["id"])==int(active or -1)
            r.pop("metrics_json",None)
        return rows

    def save_active_scenario(self):
        sid=self.active_scenario_id()
        if not sid:return
        now=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        snap=self.full_snapshot();metrics=self.plan_metrics()
        with self.lock:
            self.cx.execute("""UPDATE scenarios SET updated_at=?,data_json=?,metrics_json=?
                WHERE id=?""",(now,json.dumps(snap,ensure_ascii=False),
                json.dumps(metrics,ensure_ascii=False),sid))
            self.cx.commit()

    def create_scenario(self,name,description=""):
        name=(name or "").strip()
        if not name:raise ValueError("Chưa nhập tên phương án.")
        base=self.ensure_baseline()
        self.save_active_scenario()
        now=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        snap=self.full_snapshot();metrics=self.plan_metrics()
        with self.lock:
            self.cx.execute("""INSERT INTO scenarios(
                year,name,description,created_at,updated_at,data_json,metrics_json,is_baseline
            ) VALUES(?,?,?,?,?,?,?,0)""",(self.active_year(),name,description,now,now,
                json.dumps(snap,ensure_ascii=False),json.dumps(metrics,ensure_ascii=False)))
            sid=self.cx.execute("SELECT last_insert_rowid()").fetchone()[0]
            self.cx.commit()
        self.set_setting("active_scenario_id",sid)
        return {"id":sid,"name":name}

    def switch_scenario(self,sid):
        self.ensure_baseline()
        self.save_active_scenario()
        with self.lock:
            row=self.cx.execute("SELECT * FROM scenarios WHERE id=? AND year=?",
                                (int(sid),self.active_year())).fetchone()
        if not row:raise ValueError("Không tìm thấy phương án.")
        self.restore_snapshot(json.loads(row["data_json"]))
        self.set_setting("active_scenario_id",int(sid))
        return {"id":row["id"],"name":row["name"]}

    def apply_scenario_as_main(self,sid):
        self.save_active_scenario()
        with self.lock:
            src=self.cx.execute("SELECT * FROM scenarios WHERE id=? AND year=?",
                                (int(sid),self.active_year())).fetchone()
            base=self.cx.execute("""SELECT * FROM scenarios
                WHERE year=? AND is_baseline=1 LIMIT 1""",(self.active_year(),)).fetchone()
        if not src or not base:raise ValueError("Không tìm thấy phương án.")
        if (base["workflow_status"] or "DRAFT")=="APPROVED":
            raise ValueError("APPROVED_LOCK|BẢN CHÍNH đã được Ban Giám đốc duyệt. Hãy mở lại trạng thái Soạn thảo trước khi áp dụng phương án mới.")
        snap=json.loads(src["data_json"]);metrics=json.loads(src["metrics_json"] or "{}")
        now=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with self.lock:
            self.cx.execute("""UPDATE scenarios SET data_json=?,metrics_json=?,updated_at=?,
                workflow_status='DRAFT',approved_by='',approved_at='' WHERE id=?""",
                (json.dumps(snap,ensure_ascii=False),json.dumps(metrics,ensure_ascii=False),now,base["id"]))
            self.cx.commit()
        self.restore_snapshot(snap)
        self.set_setting("active_scenario_id",base["id"])
        return {"id":base["id"],"name":"BẢN CHÍNH"}

    def delete_scenario(self,sid):
        self.ensure_baseline()
        with self.lock:
            row=self.cx.execute("SELECT * FROM scenarios WHERE id=?",(int(sid),)).fetchone()
        if not row:return
        if row["is_baseline"]:raise ValueError("Không thể xóa BẢN CHÍNH.")
        if self.active_scenario_id()==int(sid):
            base=self.cx.execute("""SELECT id FROM scenarios
                WHERE year=? AND is_baseline=1""",(self.active_year(),)).fetchone()
            self.switch_scenario(base["id"])
        with self.lock:
            self.cx.execute("DELETE FROM scenarios WHERE id=?",(int(sid),))
            self.cx.commit()

    # ---------- Gợi ý giáo viên ----------
    def recommend_teachers(self,class_name,subject,req_type=None,week=1,limit=10):
        rows=[]
        for t in self.teachers(active_only=True):
            try:
                p=self.preview_drop(class_name,subject,t["name"],req_type,week)
            except Exception:
                continue
            compatible=bool(p["compatible"]);covers=bool(p["covers"])
            maxdiff=num(p["max_after"]["diff"]);afterdiff=num(p["after"]["diff"])
            employment=t.get("employment_type","")
            parts=self._target_parts(class_name,subject,req_type)
            sw=min(x["start_week"] for x in parts);ew=max(x["end_week"] for x in parts)
            tkb_conflicts=self.timetable_conflicts_for_teacher(t["name"],class_name,subject,sw,ew)
            score=0
            if compatible:score+=1000
            if covers:score+=300
            if maxdiff<=0:score+=200
            score-=max(0,maxdiff)*35
            score-=abs(min(0,afterdiff))*2
            score-=len(tkb_conflicts)*700
            if "Cơ hữu" in employment:score+=15
            elif "Thỉnh giảng" in employment:score-=15
            rows.append({
                "teacher":t["name"],"subject":t.get("primary_subject",""),
                "employment_type":employment,"compatible":compatible,"covers":covers,
                "timetable_conflicts":len(tkb_conflicts),
                "current":p["current"],"after":p["after"],"max_after":p["max_after"],
                "score":score
            })
        rows.sort(key=lambda x:(-x["score"],x["teacher"]))
        return rows[:int(limit)]

    # ---------- Trung tâm kiểm tra lỗi ----------
    def validate_all(self):
        year=self.active_year()
        issues=[]
        reqs=self.requirements(year)
        teachers=self.teachers(year)
        classes=self.classes(year)
        tmap={t["name"]:t for t in teachers}
        rules={r["subject"]:r for r in self.program_rules()}

        def add(severity,code,title,detail,ref=""):
            issues.append({"severity":severity,"code":code,"title":title,"detail":detail,"ref":ref})

        # 1. Nhu cầu chưa có GV / sai chuyên môn / ngoài hiệu lực
        for r in reqs:
            teacher=(r.get("teacher") or "").strip()
            ref=f'{r["class_name"]} · {r["subject"]} · {r["req_type"]}'
            if not teacher:
                add("error","UNASSIGNED","Chưa phân công giáo viên",
                    f'{r["class_name"]} {r["subject"]} ({r["req_type"]}) chưa có giáo viên.',ref)
                continue
            t=tmap.get(teacher)
            if not t:
                add("error","TEACHER_MISSING","Giáo viên không tồn tại",
                    f'{ref} đang gán cho "{teacher}" nhưng giáo viên không có trong đội ngũ.',ref)
                continue
            if r["subject"] not in t.get("can_teach_list",[]) and r["subject"]!=t.get("primary_subject"):
                add("error","WRONG_SUBJECT","Không đúng chuyên môn",
                    f'{teacher} đang dạy {r["subject"]} ở {r["class_name"]} nhưng môn này không nằm trong thẻ chuyên môn.',teacher)
            if int(t["active_start"])>int(r["start_week"]) or int(t["active_end"])<int(r["end_week"]):
                add("warning","TEACHER_PERIOD","Hiệu lực giáo viên không phủ nhiệm vụ",
                    f'{teacher} hiệu lực T{t["active_start"]}-{t["active_end"]}, nhưng nhiệm vụ {ref} có hiệu lực T{r["start_week"]}-{r["end_week"]}.',teacher)

        # 2. Tổng phân bổ chương trình
        for rule in self.program_rules():
            if abs(num(rule["core_schedule_total"])-num(rule["base_hours"]))>1e-6:
                add("error","PROGRAM_TOTAL","Sai tổng phân bổ cốt lõi",
                    f'{rule["subject"]}: phân bổ {self._fmt_hours_vi(rule["core_schedule_total"])} / {self._fmt_hours_vi(rule["base_hours"])} tiết.',rule["subject"])
            if rule["can_specialty"] and abs(num(rule["specialty_schedule_total"])-num(rule["specialty_hours"]))>1e-6:
                add("error","SPECIALTY_TOTAL","Sai tổng phân bổ chuyên đề",
                    f'{rule["subject"]}: phân bổ CĐ {self._fmt_hours_vi(rule["specialty_schedule_total"])} / {self._fmt_hours_vi(rule["specialty_hours"])} tiết.',rule["subject"])

        # 3. GV vượt chuẩn ở bất kỳ tuần nào
        for t in self.teachers(year,active_only=True):
            mx=self.teacher_max_load(t["name"],year)
            if mx and mx["diff"]>1e-6:
                add("warning","OVERLOAD","Giáo viên vượt chuẩn",
                    f'{t["name"]} cao nhất T{mx["week"]}: {self._fmt_hours_vi(mx["total"])}/{self._fmt_hours_vi(mx["standard"])}; dư {self._fmt_hours_vi(mx["diff"])} tiết.',t["name"])
            if not (t.get("primary_subject") or "").strip():
                add("warning","NO_PRIMARY_SUBJECT","Chưa khai chuyên môn chính",
                    f'{t["name"]} chưa có chuyên môn chính.',t["name"])

        # 4. Chủ nhiệm trùng cùng lớp theo thời gian
        cn=[d for d in self.duties(year) if d["duty_type"]=="Chủ nhiệm"]
        for i,a in enumerate(cn):
            for b in cn[i+1:]:
                if a["class_name"]!=b["class_name"] or a["teacher"]==b["teacher"]:continue
                if not (int(a["end_week"])<int(b["start_week"]) or int(a["start_week"])>int(b["end_week"])):
                    add("error","HOMEROOM_OVERLAP","Trùng chủ nhiệm",
                        f'{a["class_name"]} có {a["teacher"]} và {b["teacher"]} cùng hiệu lực chồng nhau.',a["class_name"])

        # 5. Nhiệm vụ cùng loại của 1 GV chồng nhau
        duties=self.duties(year)
        for teacher in {d["teacher"] for d in duties}:
            td=[d for d in duties if d["teacher"]==teacher]
            for dtype in ("Chủ nhiệm","Kiêm nhiệm","Chức vụ"):
                arr=[d for d in td if d["duty_type"]==dtype]
                for i,a in enumerate(arr):
                    for b in arr[i+1:]:
                        if not (int(a["end_week"])<int(b["start_week"]) or int(a["start_week"])>int(b["end_week"])):
                            add("warning","DUTY_OVERLAP","Nhiệm vụ cùng loại chồng tuần",
                                f'{teacher}: {a["duty_name"]} T{a["start_week"]}-{a["end_week"]} chồng {b["duty_name"]} T{b["start_week"]}-{b["end_week"]}.',teacher)

        # 6. Ôn TN
        for r in self.review_jobs(year):
            if r["start_week"] is None:
                add("warning","REVIEW_NO_START","Ôn TN chưa có tuần bắt đầu",
                    f'{r["teacher"]} · {r["subject"]} · {r["classes_label"] or "chưa chọn lớp"} chưa xác định tuần bắt đầu.',r["teacher"])
            with self.lock:
                eligible={x["class_name"] for x in self.cx.execute("""SELECT DISTINCT class_name FROM requirements
                    WHERE year=? AND teacher=? AND subject=? AND class_name LIKE '12%'""",
                    (year,r["teacher"],r["subject"])).fetchall()}
            invalid=[c for c in r["classes"] if c not in eligible]
            if invalid:
                add("error","REVIEW_CLASS","Lớp ôn không thuộc phân công",
                    f'{r["teacher"]} đang ôn {", ".join(invalid)} nhưng không dạy {r["subject"]} ở các lớp này.',r["teacher"])

        # 7. Lớp điều chỉnh riêng so với tổ hợp
        for c in classes:
            if c.get("customized"):
                add("info","CLASS_CUSTOMIZED","Lớp có cấu trúc riêng",
                    f'{c["class_name"]} đang điều chỉnh riêng, không hoàn toàn theo {c.get("template_name") or "tổ hợp"}.',c["class_name"])

        # 8. Thiếu nhân sự theo môn
        staffing=self.staffing_summary(year)
        for subject,d in staffing["by_subject"].items():
            need=num(d.get("fte17"));available=num(d.get("teachers"))
            if available+1e-9<math.ceil(need):
                add("warning","STAFF_SHORTAGE","Có nguy cơ thiếu giáo viên",
                    f'{subject}: nhu cầu quy đổi {need:.2f} GV, hiện có {int(available)} GV có thể dạy.',subject)

        # 9. Biến động nhân sự làm gián đoạn phân công
        for ev in self.staff_events(year=year):
            if not ev["blocks_teaching"]: continue
            impacted=[r for r in reqs if r["teacher"]==ev["teacher"] and
                      not (int(r["end_week"])<int(ev["start_week"]) or int(r["start_week"])>int(ev["end_week"]))]
            if impacted:
                sample=", ".join(f'{r["class_name"]} {r["subject"]}' for r in impacted[:5])
                add("error","STAFF_EVENT_CONFLICT","GV không khả dụng trong thời gian được phân công",
                    f'{ev["teacher"]} · {ev["event_type"]} T{ev["start_week"]}-{ev["end_week"]}; ảnh hưởng {len(impacted)} nhiệm vụ: {sample}.',ev["teacher"])

        # 10. TKB
        tv=self.validate_timetable()
        for x in tv["issues"]:
            add(x["severity"],"TKB_"+x["code"],"Vấn đề thời khóa biểu",x["detail"],"TKB")

        counts={"error":0,"warning":0,"info":0}
        for x in issues:counts[x["severity"]]=counts.get(x["severity"],0)+1
        return {"counts":counts,"issues":issues,"total":len(issues),"timetable":tv}

    # ---------- Nhập dữ liệu đầu năm ----------
    def apply_import_preview(self,preview,import_teachers=True,import_classes=True):
        year=self.active_year()
        added_t=updated_t=added_c=0
        with self.lock:
            if import_teachers:
                for t in preview.get("teachers",[]):
                    row=self.cx.execute("SELECT id FROM teachers WHERE year=? AND name=?",
                                        (year,t["name"])).fetchone()
                    can=json.dumps(t.get("can_teach",[]),ensure_ascii=False)
                    if row:
                        self.cx.execute("""UPDATE teachers SET stt=?,primary_subject=?,can_teach=?,
                            standard=?,employment_type=?,active_start=?,active_end=?,note=?,active=1
                            WHERE id=?""",(t.get("stt",0),t.get("primary_subject",""),can,
                            num(t.get("standard"),17),t.get("employment_type","Cơ hữu"),
                            int(t.get("active_start",1)),int(t.get("active_end",35)),
                            t.get("note",""),row["id"]))
                        updated_t+=1
                    else:
                        self.cx.execute("""INSERT INTO teachers(
                            year,stt,name,primary_subject,can_teach,standard,employment_type,
                            active_start,active_end,role,role_credit,homeroom,homeroom_credit,note,active
                        ) VALUES(?,?,?,?,?,?,?,?,?,'',0,'',0,?,1)""",(year,t.get("stt",0),t["name"],
                            t.get("primary_subject",""),can,num(t.get("standard"),17),
                            t.get("employment_type","Cơ hữu"),int(t.get("active_start",1)),
                            int(t.get("active_end",35)),t.get("note","")))
                        added_t+=1
            if import_classes:
                homeroom_map={t.get("homeroom",""):t["name"] for t in preview.get("teachers",[]) if t.get("homeroom")}
                for cname in preview.get("classes",[]):
                    row=self.cx.execute("SELECT id FROM classes WHERE year=? AND class_name=?",
                                       (year,cname)).fetchone()
                    if row:
                        if cname in homeroom_map:
                            self.cx.execute("UPDATE classes SET homeroom_teacher=? WHERE id=?",
                                            (homeroom_map[cname],row["id"]))
                        continue
                    grade=int(cname[:2])
                    self.cx.execute("""INSERT INTO classes(
                        year,class_name,grade,program_name,template_name,customized,homeroom_teacher,active
                    ) VALUES(?,?,?,'','',0,?,1)""",(year,cname,grade,homeroom_map.get(cname,"")))
                    added_c+=1

            # Đồng bộ CN/KN/chức vụ từ mẫu Excel nhưng không tạo trùng.
            if import_teachers:
                for t in preview.get("teachers",[]):
                    role=(t.get("role") or "").strip()
                    if role and "thỉnh giảng" not in role.lower():
                        role_credit=num(t.get("role_credit"))
                        dtype="Kiêm nhiệm" if role_credit else "Chức vụ"
                        std_override=num(t.get("standard"),17) if (dtype=="Chức vụ" and num(t.get("standard"),17)<17) else None
                        exists=self.cx.execute("""SELECT 1 FROM duties WHERE year=? AND teacher=? AND duty_type=?
                            AND duty_name=? AND start_week=1 AND end_week=35""",
                            (year,t["name"],dtype,role)).fetchone()
                        if not exists:
                            self.cx.execute("""INSERT INTO duties(
                                year,teacher,duty_type,duty_name,class_name,credit,standard_override,
                                start_week,end_week,note
                            ) VALUES(?,?,?,?,?,?,?,?,?,?)""",
                            (year,t["name"],dtype,role,"",role_credit,std_override,1,35,"Nhập từ Excel"))
                    hom=(t.get("homeroom") or "").strip().upper()
                    hcredit=num(t.get("homeroom_credit"))
                    if hom and hcredit:
                        exists=self.cx.execute("""SELECT 1 FROM duties WHERE year=? AND teacher=? AND duty_type='Chủ nhiệm'
                            AND class_name=? AND start_week=1 AND end_week=35""",
                            (year,t["name"],hom)).fetchone()
                        if not exists:
                            self.cx.execute("""INSERT INTO duties(
                                year,teacher,duty_type,duty_name,class_name,credit,standard_override,
                                start_week,end_week,note
                            ) VALUES(?,?,?,?,?,?,?,?,?,?)""",
                            (year,t["name"],"Chủ nhiệm","Chủ nhiệm "+hom,hom,hcredit,None,1,35,"Nhập từ Excel"))
            self.cx.commit()
        return {"teachers_added":added_t,"teachers_updated":updated_t,"classes_added":added_c}

    # ---------- Quy trình duyệt ----------
    def active_scenario_row(self):
        self.ensure_baseline()
        sid=self.active_scenario_id()
        with self.lock:
            r=self.cx.execute("SELECT * FROM scenarios WHERE id=?",(sid,)).fetchone()
        return dict(r) if r else None

    def workflow_status(self):
        r=self.active_scenario_row()
        return (r or {}).get("workflow_status","DRAFT")

    def assert_editable(self):
        r=self.active_scenario_row()
        if r and r.get("workflow_status")=="APPROVED":
            raise ValueError("APPROVED_LOCK|Phương án đã được Ban Giám đốc duyệt. Hãy mở lại trạng thái Soạn thảo trước khi sửa.")

    def set_workflow(self,status,actor="",note="",sid=None):
        allowed={"DRAFT","REVIEWED","APPROVED"}
        status=(status or "").strip().upper()
        if status not in allowed: raise ValueError("Trạng thái duyệt không hợp lệ.")
        if status=="APPROVED":
            guard=self.validate_all()
            if guard["counts"]["error"]>0:
                sample="; ".join(x["detail"] for x in guard["issues"] if x["severity"]=="error")[:600]
                raise ValueError(f'VALIDATION_BLOCK|Không thể duyệt: còn {guard["counts"]["error"]} lỗi bắt buộc phải xử lý. {sample}')
        sid=int(sid or self.active_scenario_id() or self.ensure_baseline()["id"])
        with self.lock:
            row=self.cx.execute("SELECT * FROM scenarios WHERE id=? AND year=?",(sid,self.active_year())).fetchone()
            if not row: raise ValueError("Không tìm thấy phương án.")
            old=row["workflow_status"] or "DRAFT"
            now=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            approved_by=actor if status=="APPROVED" else ""
            approved_at=now if status=="APPROVED" else ""
            self.cx.execute("""UPDATE scenarios SET workflow_status=?,approved_by=?,approved_at=?,updated_at=?
                WHERE id=?""",(status,approved_by,approved_at,now,sid))
            self.cx.execute("""INSERT INTO workflow_log(year,scenario_id,created_at,actor,old_status,new_status,note)
                VALUES(?,?,?,?,?,?,?)""",(self.active_year(),sid,now,actor,old,status,note))
            self.cx.commit()
        if actor: self.set_setting("operator_name",actor)
        return {"id":sid,"old_status":old,"new_status":status}

    def workflow_history(self,sid=None):
        sid=int(sid or self.active_scenario_id() or -1)
        return self._rows("""SELECT * FROM workflow_log WHERE year=? AND scenario_id=?
            ORDER BY id DESC""",(self.active_year(),sid))

    # ---------- Sự kiện nhân sự ----------
    def staff_events(self,teacher=None,year=None):
        year=year or self.active_year()
        q="SELECT * FROM staff_events WHERE year=?";params=[year]
        if teacher:q+=" AND teacher=?";params.append(teacher)
        q+=" ORDER BY start_week,end_week,teacher"
        return self._rows(q,params)

    def save_staff_event(self,d):
        teacher=(d.get("teacher") or "").strip()
        et=(d.get("event_type") or "").strip()
        sw=int(d.get("start_week") or 1);ew=int(d.get("end_week") or self.setting("weeks",35))
        if not teacher or not et:raise ValueError("Thiếu giáo viên hoặc loại biến động.")
        if not 1<=sw<=ew<=int(self.setting("weeks",35)):raise ValueError("Khoảng tuần không hợp lệ.")
        vals=(teacher,et,sw,ew,int(bool(d.get("blocks_teaching",True))),
              (d.get("replacement_teacher") or "").strip(),d.get("note",""))
        with self.lock:
            if d.get("id"):
                self.cx.execute("""UPDATE staff_events SET teacher=?,event_type=?,start_week=?,end_week=?,
                    blocks_teaching=?,replacement_teacher=?,note=? WHERE id=?""",vals+(int(d["id"]),))
            else:
                self.cx.execute("""INSERT INTO staff_events(
                    year,teacher,event_type,start_week,end_week,blocks_teaching,replacement_teacher,note
                ) VALUES(?,?,?,?,?,?,?,?)""",(self.active_year(),)+vals)
            self.cx.commit()

    def delete_staff_event(self,rid):
        with self.lock:
            self.cx.execute("DELETE FROM staff_events WHERE id=?",(int(rid),));self.cx.commit()

    def blocking_events(self,teacher,start_week,end_week,year=None):
        year=year or self.active_year()
        return self._rows("""SELECT * FROM staff_events WHERE year=? AND teacher=? AND blocks_teaching=1
            AND NOT(end_week<? OR start_week>?) ORDER BY start_week""",
            (year,teacher,int(start_week),int(end_week)))

    def event_replacement_plan(self,event_id):
        with self.lock:
            ev=self.cx.execute("SELECT * FROM staff_events WHERE id=?",(int(event_id),)).fetchone()
        if not ev:raise ValueError("Không tìm thấy biến động nhân sự.")
        ev=dict(ev);reqs=self.requirements()
        impacted=[r for r in reqs if r["teacher"]==ev["teacher"] and
                  not (int(r["end_week"])<int(ev["start_week"]) or int(r["start_week"])>int(ev["end_week"]))]
        groups={}
        for r in impacted:groups.setdefault((r["class_name"],r["subject"]),[]).append(r)
        rows=[]
        for (cname,subject),parts in sorted(groups.items()):
            rt=None if len(parts)>1 else parts[0]["req_type"]
            rec=[x for x in self.recommend_teachers(cname,subject,rt,ev["start_week"],5)
                 if x["teacher"]!=ev["teacher"]]
            rows.append({"class_name":cname,"subject":subject,
                         "req_type":"Tất cả" if len(parts)>1 else parts[0]["req_type"],
                         "recommendations":rec[:3]})
        return {"event":ev,"impacted_count":len(impacted),"rows":rows}

    # ---------- Thời khóa biểu ----------
    def timetable_entries(self,year=None):
        year=year or self.active_year()
        return self._rows("""SELECT * FROM timetable_entries WHERE year=?
            ORDER BY session,day,period,class_name,teacher""",(year,))

    def save_timetable_entry(self,d):
        cname=(d.get("class_name") or "").strip().upper()
        teacher=(d.get("teacher") or "").strip()
        subject=canonical_subject_name(d.get("subject",""))
        session=(d.get("session") or "").strip()
        day=(d.get("day") or "").strip()
        period=int(d.get("period") or 0)
        sw=int(d.get("start_week") or 1);ew=int(d.get("end_week") or self.setting("weeks",35))
        if not cname or not teacher or not subject or not day or period<=0:
            raise ValueError("Thiếu lớp, giáo viên, môn, thứ hoặc tiết.")
        vals=(cname,teacher,subject,session,day,period,sw,ew,d.get("source","Nhập tay"),d.get("note",""))
        with self.lock:
            if d.get("id"):
                self.cx.execute("""UPDATE timetable_entries SET class_name=?,teacher=?,subject=?,session=?,day=?,
                    period=?,start_week=?,end_week=?,source=?,note=? WHERE id=?""",vals+(int(d["id"]),))
            else:
                self.cx.execute("""INSERT INTO timetable_entries(
                    year,class_name,teacher,subject,session,day,period,start_week,end_week,source,note
                ) VALUES(?,?,?,?,?,?,?,?,?,?,?)""",(self.active_year(),)+vals)
            self.cx.commit()

    def delete_timetable_entry(self,rid):
        with self.lock:
            self.cx.execute("DELETE FROM timetable_entries WHERE id=?",(int(rid),));self.cx.commit()

    def _replace_timetable_range(self,start_week,end_week):
        sw=int(start_week);ew=int(end_week);year=self.active_year()
        rows=[dict(r) for r in self.cx.execute("SELECT * FROM timetable_entries WHERE year=?",(year,))]
        for e in rows:
            s=int(e["start_week"]);t=int(e["end_week"])
            if t<sw or s>ew:continue
            if s<sw and t>ew:
                # Giữ cả đoạn trước và sau khoảng thay thế.
                self.cx.execute("UPDATE timetable_entries SET end_week=? WHERE id=?",(sw-1,e["id"]))
                self.cx.execute("""INSERT INTO timetable_entries(year,class_name,teacher,subject,session,day,period,start_week,end_week,source,note)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?)""",(year,e["class_name"],e["teacher"],e["subject"],e.get("session",""),e["day"],e["period"],ew+1,t,e["source"],e["note"]))
            elif s<sw<=t<=ew:
                self.cx.execute("UPDATE timetable_entries SET end_week=? WHERE id=?",(sw-1,e["id"]))
            elif sw<=s<=ew<t:
                self.cx.execute("UPDATE timetable_entries SET start_week=? WHERE id=?",(ew+1,e["id"]))
            else:
                self.cx.execute("DELETE FROM timetable_entries WHERE id=?",(e["id"],))

    def apply_timetable_preview(self,preview,replace=True,replace_mode=None):
        entries=preview.get("entries",[])
        mode=replace_mode or ("all" if replace else "append")
        with self.lock:
            if mode=="all":
                self.cx.execute("DELETE FROM timetable_entries WHERE year=?",(self.active_year(),))
            elif mode=="range" and entries:
                sw=min(int(e.get("start_week",1)) for e in entries);ew=max(int(e.get("end_week",35)) for e in entries)
                self._replace_timetable_range(sw,ew)
            for e in entries:
                self.cx.execute("""INSERT INTO timetable_entries(
                    year,class_name,teacher,subject,session,day,period,start_week,end_week,source,note
                ) VALUES(?,?,?,?,?,?,?,?,?,?,?)""",(self.active_year(),e["class_name"],e["teacher"],
                    canonical_subject_name(e["subject"]),e.get("session",""),e["day"],int(e["period"]),
                    int(e.get("start_week",1)),int(e.get("end_week",35)),e.get("source","Excel"),e.get("note","")))
            self.cx.commit()
        return {"imported":len(entries),"replace_mode":mode}

    def timetable_conflicts_for_teacher(self,teacher,class_name,subject,start_week,end_week):
        target=[e for e in self.timetable_entries() if e["class_name"]==class_name and
                canonical_subject_name(e["subject"])==canonical_subject_name(subject)]
        allrows=self.timetable_entries()
        out=[]
        for slot in target:
            for e in allrows:
                if e["teacher"]!=teacher or e["class_name"]==class_name:continue
                if (e.get("session") or "")==(slot.get("session") or "") and e["day"]==slot["day"] and int(e["period"])==int(slot["period"]) and not (
                    int(e["end_week"])<max(int(slot["start_week"]),int(start_week)) or
                    int(e["start_week"])>min(int(slot["end_week"]),int(end_week))):
                    out.append({"session":slot.get("session",""),"day":slot["day"],"period":slot["period"],"other_class":e["class_name"],
                                "other_subject":e["subject"]})
        return out

    def timetable_sync_preview(self):
        reqs=self.requirements()
        entries=self.timetable_entries()

        # Mỗi lớp/môn chỉ được tự đồng bộ khi phân công chuyên môn có đúng 1 GV.
        assignment={}
        for r in reqs:
            if r["teacher"]:
                assignment.setdefault((r["class_name"],canonical_subject_name(r["subject"])),set()).add(r["teacher"])
        target_map={k:next(iter(v)) for k,v in assignment.items() if len(v)==1}

        known={t["name"] for t in self.teachers()}
        unknown=sorted({e["teacher"] for e in entries if e["teacher"] and e["teacher"] not in known})

        candidate_keys=set()
        rows_by_key={}
        for key,target in sorted(target_map.items()):
            es=[e for e in entries if (e["class_name"],canonical_subject_name(e["subject"]))==key]
            mismatch=[e for e in es if e["teacher"]!=target]
            if not mismatch:
                continue
            candidate_keys.add(key)
            rows_by_key[key]={
                "class_name":key[0],"subject":key[1],"target_teacher":target,
                "current_teachers":sorted({e["teacher"] for e in mismatch}),
                "row_count":len(mismatch),"conflicts":[],"conflict_count":0,"safe":True
            }

        # Mô phỏng TOÀN BỘ thay đổi cùng lúc để tránh trường hợp từng nhóm riêng lẻ
        # đều an toàn nhưng khi đồng bộ hàng loạt lại tạo trùng giáo viên.
        proposed=[]
        for e in entries:
            key=(e["class_name"],canonical_subject_name(e["subject"]))
            pe=dict(e)
            if key in candidate_keys:
                pe["_old_teacher"]=e["teacher"]
                pe["_candidate_key"]=key
                pe["teacher"]=target_map[key]
            else:
                pe["_old_teacher"]=e["teacher"]
                pe["_candidate_key"]=None
            proposed.append(pe)

        unsafe=set()
        conflicts_by_key={k:[] for k in candidate_keys}
        for i,a in enumerate(proposed):
            for b in proposed[i+1:]:
                if a["teacher"]!=b["teacher"] or a["class_name"]==b["class_name"]:
                    continue
                if (a.get("session") or "")!=(b.get("session") or ""):
                    continue
                if a["day"]!=b["day"] or int(a["period"])!=int(b["period"]):
                    continue
                if int(a["end_week"])<int(b["start_week"]) or int(a["start_week"])>int(b["end_week"]):
                    continue
                # Chỉ xem là rủi ro đồng bộ nếu ít nhất một phía là nhóm đang đổi.
                keys=[x for x in (a.get("_candidate_key"),b.get("_candidate_key")) if x in candidate_keys]
                if not keys:
                    continue
                for k in keys:
                    unsafe.add(k)
                    other=b if a.get("_candidate_key")==k else a
                    conflicts_by_key[k].append({
                        "session":a.get("session",""),"day":a["day"],"period":a["period"],
                        "other_class":other["class_name"],"other_subject":other["subject"],
                        "teacher":a["teacher"]
                    })

        rows=[]
        for key in sorted(candidate_keys):
            x=rows_by_key[key]
            unique=[];seen=set()
            for c in conflicts_by_key.get(key,[]):
                ck=(c["session"],c["day"],c["period"],c["other_class"],c["other_subject"],c["teacher"])
                if ck not in seen:
                    seen.add(ck);unique.append(c)
            x["conflicts"]=unique
            x["conflict_count"]=len(unique)
            x["safe"]=key not in unsafe
            rows.append(x)

        return {
            "rows":rows,
            "mismatch_rows":sum(x["row_count"] for x in rows),
            "safe_rows":sum(x["row_count"] for x in rows if x["safe"]),
            "unsafe_rows":sum(x["row_count"] for x in rows if not x["safe"]),
            "unknown_teachers":unknown
        }

    def sync_timetable_to_assignment(self):
        preview=self.timetable_sync_preview()
        changed=0;skipped=0
        with self.lock:
            for x in preview["rows"]:
                if not x["safe"]:
                    skipped+=x["row_count"]
                    continue
                cur=self.cx.execute("""UPDATE timetable_entries SET teacher=?
                    WHERE year=? AND class_name=? AND subject=? AND teacher<>?""",
                    (x["target_teacher"],self.active_year(),x["class_name"],x["subject"],x["target_teacher"]))
                if cur.rowcount and cur.rowcount>0:
                    changed+=cur.rowcount
            self.cx.commit()
        return {"changed":changed,"skipped":skipped,"before":preview}

    def validate_timetable(self):
        rows=self.timetable_entries();issues=[]
        if not rows:
            return {"total":0,"counts":{"error":0,"warning":0,"info":0},"issues":[],"available":False}
        def add(sev,code,detail):issues.append({"severity":sev,"code":code,"detail":detail})
        # Trùng GV / trùng lớp
        for i,a in enumerate(rows):
            for b in rows[i+1:]:
                if (a.get("session") or "")!=(b.get("session") or "") or a["day"]!=b["day"] or int(a["period"])!=int(b["period"]):continue
                overlap=not (int(a["end_week"])<int(b["start_week"]) or int(a["start_week"])>int(b["end_week"]))
                if not overlap:continue
                if a["teacher"]==b["teacher"] and a["class_name"]!=b["class_name"]:
                    add("error","TEACHER_CLASH",f'{a["teacher"]}: {a.get("session","")} {a["day"]} tiết {a["period"]} trùng {a["class_name"]} và {b["class_name"]}.')
                if a["class_name"]==b["class_name"] and (a["teacher"]!=b["teacher"] or a["subject"]!=b["subject"]):
                    add("error","CLASS_CLASH",f'{a["class_name"]}: {a.get("session","")} {a["day"]} tiết {a["period"]} có hai phân công TKB.')
        # TKB vs phân công chuyên môn
        reqs=self.requirements();rset={(r["class_name"],canonical_subject_name(r["subject"]),r["teacher"]) for r in reqs if r["teacher"]}
        known_teachers={t["name"] for t in self.teachers()}
        mismatch_seen=set();unknown_seen=set()
        for e in rows:
            if e["teacher"] not in known_teachers and e["teacher"] not in unknown_seen:
                unknown_seen.add(e["teacher"])
                add("warning","UNKNOWN_TKB_TEACHER",f'TKB có tên/vị trí "{e["teacher"]}" nhưng chưa có trong danh sách đội ngũ. Có thể là giáo viên hợp đồng/vị trí chưa định danh.')
            k=(e["class_name"],canonical_subject_name(e["subject"]),e["teacher"])
            if not _is_tkb_admin_subject(e["subject"]) and k not in rset and k not in mismatch_seen:
                mismatch_seen.add(k)
                add("warning","ASSIGNMENT_MISMATCH",f'{e["teacher"]} có TKB {e["class_name"]} {e["subject"]} nhưng không khớp phân công chuyên môn.')
        # So số slot từng tuần, HĐTN quy đổi 3 tiết phân công -> 2 slot TKB.
        by_cs={}
        for e in rows:by_cs.setdefault((e["class_name"],canonical_subject_name(e["subject"])),[]).append(e)
        for c in self.classes(active_only=True):
            cname=c["class_name"]
            subjects={r["subject"] for r in reqs if r["class_name"]==cname}
            for subject in subjects:
                rr=[r for r in reqs if r["class_name"]==cname and r["subject"]==subject]
                slots=by_cs.get((cname,canonical_subject_name(subject)),[])
                bad=[]
                for w in range(1,int(self.setting("weeks",35))+1):
                    expected=sum(self.requirement_weekly_hours(r,w) for r in rr)
                    if canonical_subject_name(subject)=="TNHN" and expected>0:
                        expected=num(self.setting("hdtn_tkb_slots",2),2)
                    actual=sum(1 for e in slots if int(e["start_week"])<=w<=int(e["end_week"]))
                    if abs(actual-expected)>0.01:bad.append(w)
                if bad:
                    # Chỉ nêu gọn khoảng đầu-cuối.
                    add("warning","SLOT_MISMATCH",f'{cname} {subject}: số tiết TKB không khớp phân bổ ở {len(bad)}/35 tuần (ví dụ T{bad[0]}).')
        counts={"error":0,"warning":0,"info":0}
        for x in issues:counts[x["severity"]]+=1
        return {"total":len(issues),"counts":counts,"issues":issues,"available":True}

    # ---------- Phân tích tác động ----------
    def impact_analysis(self,class_name,subject,teacher,req_type=None,week=1):
        parts=self._target_parts(class_name,subject,req_type)
        if not parts:raise ValueError("Không tìm thấy thẻ môn.")
        p=self.preview_drop(class_name,subject,teacher,req_type,week)
        old_teachers=sorted({x["teacher"] for x in parts if x["teacher"] and x["teacher"]!=teacher})
        old_impacts=[]
        ids=[x["id"] for x in parts]
        for ot in old_teachers:
            before=self.teacher_load_at_week(ot,week)
            after=self.teacher_load_at_week(ot,week,exclude_ids=ids)
            mx=None
            for w in range(1,int(self.setting("weeks",35))+1):
                d=self.teacher_load_at_week(ot,w,exclude_ids=ids)
                if d and (mx is None or d["diff"]>mx["diff"]):
                    mx=dict(d);mx["week"]=w
            old_impacts.append({"teacher":ot,"before":before,"after":after,"max_after":mx})
        min_start=min(x["start_week"] for x in parts);max_end=max(x["end_week"] for x in parts)
        events=self.blocking_events(teacher,min_start,max_end)
        tkb_conflicts=self.timetable_conflicts_for_teacher(teacher,class_name,subject,min_start,max_end)
        return {"target":p,"old_teachers":old_impacts,"blocking_events":events,
                "timetable_conflicts":tkb_conflicts,
                "safe":bool(p["compatible"] and p["covers"] and not events and not tkb_conflicts and num(p["max_after"]["diff"])<=0)}

    # ---------- Tối ưu phân công tự động ----------
    def _strategy_assignment(self,strategy):
        """Tối ưu cục bộ từ phương án hiện tại; chỉ nhận thay đổi làm mục tiêu tốt hơn."""
        year=self.active_year();weeks=int(self.setting("weeks",35))
        teachers=self.teachers(year,active_only=True);tmap={t["name"]:t for t in teachers}
        reqs=self.requirements(year);rules={r["subject"]:r for r in self.program_rules()}

        load={};std={}
        for t in teachers:
            name=t["name"]
            load[name]=[];std[name]=[]
            for w in range(1,weeks+1):
                d=self.teacher_load_at_week(name,w,year)
                load[name].append(num(d["total"]));std[name].append(num(d["standard"]))

        # Unit = cốt lõi + CĐ cùng lớp/môn nếu hiện cùng GV; nếu đang tách thì giữ từng phần.
        raw={}
        for r in reqs:raw.setdefault((r["class_name"],r["subject"]),[]).append(r)
        units=[]
        for key,parts in raw.items():
            teachers_now={r["teacher"] for r in parts}
            if len(teachers_now)<=1:
                units.append({"key":key,"parts":parts,"ids":[r["id"] for r in parts],
                              "old_teacher":parts[0]["teacher"] if parts else ""})
            else:
                for r in parts:
                    units.append({"key":(r["class_name"],r["subject"],r["req_type"]),"parts":[r],
                                  "ids":[r["id"]],"old_teacher":r["teacher"]})
        for u in units:
            u["class_name"]=u["parts"][0]["class_name"];u["subject"]=u["parts"][0]["subject"]
            u["req_type"]=None if len(u["parts"])>1 else u["parts"][0]["req_type"]
            u["gh"]=[sum(self.requirement_weekly_hours(r,w,rules) for r in u["parts"]) for w in range(1,weeks+1)]
            u["annual"]=sum(num(r["annual_hours"]) for r in u["parts"])
            u["locked"]=self.is_locked(u["class_name"]) or any(
                self.is_locked(u["class_name"],u["subject"],r["req_type"]) for r in u["parts"])

        assignment={i:u["old_teacher"] for i,u in enumerate(units)}
        original=dict(assignment)

        def eligible(name,u):
            if not name or name not in tmap:return False
            t=tmap[name]
            if u["subject"] not in t["can_teach_list"] and u["subject"]!=t["primary_subject"]:return False
            active=[w for w,h in enumerate(u["gh"],1) if h>0]
            if active and (int(t["active_start"])>min(active) or int(t["active_end"])<max(active)):return False
            if active and self.blocking_events(name,min(active),max(active)):return False
            if active and self.timetable_conflicts_for_teacher(name,u["class_name"],u["subject"],min(active),max(active)):return False
            return True

        def objective():
            maxdiff=[];maxtotal=[]
            for name in load:
                diffs=[load[name][i]-std[name][i] for i in range(weeks)]
                maxdiff.append(max(diffs) if diffs else 0);maxtotal.append(max(load[name]) if load[name] else 0)
            over_count=sum(1 for d in maxdiff if d>1e-6)
            sum_over=sum(max(0,d) for d in maxdiff)
            max_over=max([0]+maxdiff)
            spread=(max(maxtotal)-min(maxtotal)) if maxtotal else 0
            changes=sum(1 for i,a in assignment.items() if a!=original[i])
            unassigned=sum(1 for a in assignment.values() if not a)
            external=sum(1 for a in assignment.values() if a and "Thỉnh giảng" in (tmap.get(a,{}).get("employment_type","")))
            if strategy=="stable":
                return unassigned*1e8 + changes*2500 + over_count*60000 + sum_over*4000 + max_over*1500 + spread*50 + external*100
            if strategy=="least_over":
                return unassigned*1e8 + over_count*120000 + sum_over*12000 + max_over*3500 + spread*20 + changes*5 + external*100
            return unassigned*1e8 + over_count*90000 + sum_over*7000 + max_over*2200 + spread*450 + changes*20 + external*100

        def apply_move(i,new_teacher):
            u=units[i];old=assignment[i]
            if old and old in load:
                load[old]=[load[old][w]-u["gh"][w] for w in range(weeks)]
            if new_teacher and new_teacher in load:
                load[new_teacher]=[load[new_teacher][w]+u["gh"][w] for w in range(weeks)]
            assignment[i]=new_teacher

        # Ưu tiên xét các unit đang nằm trên GV quá tải.
        def overload_of(name):
            if not name or name not in load:return -999
            return max(load[name][i]-std[name][i] for i in range(weeks))
        order=sorted(range(len(units)),key=lambda i:(-overload_of(assignment[i]),-units[i]["annual"],units[i]["class_name"]))

        current_obj=objective()
        for _pass in range(4):
            improved=False
            for i in order:
                u=units[i]
                if u["locked"]:continue
                old=assignment[i]
                best_teacher=old;best_obj=current_obj
                candidates=[t["name"] for t in teachers if eligible(t["name"],u)]
                if not old and not candidates:continue
                for cand in candidates:
                    if cand==old:continue
                    apply_move(i,cand);o=objective();apply_move(i,old)
                    if o+1e-6<best_obj:
                        best_obj=o;best_teacher=cand
                if best_teacher!=old:
                    apply_move(i,best_teacher);current_obj=best_obj;improved=True
            if not improved:break

        # Thử hoán đổi hai lớp cùng môn để thoát một số cực trị cục bộ.
        for subject in sorted({u["subject"] for u in units}):
            idxs=[i for i,u in enumerate(units) if u["subject"]==subject and not u["locked"]]
            pair_improved=True;loops=0
            while pair_improved and loops<2:
                pair_improved=False;loops+=1
                for a_pos in range(len(idxs)):
                    i=idxs[a_pos];ai=assignment[i]
                    if not ai:continue
                    for b_pos in range(a_pos+1,len(idxs)):
                        j=idxs[b_pos];aj=assignment[j]
                        if not aj or ai==aj:continue
                        if not eligible(aj,units[i]) or not eligible(ai,units[j]):continue
                        # swap
                        apply_move(i,aj);apply_move(j,ai);o=objective()
                        if o+1e-6<current_obj:
                            current_obj=o;pair_improved=True
                            ai=assignment[i];break
                        # rollback
                        apply_move(j,aj);apply_move(i,ai)
                    if pair_improved:break

        result={}
        for i,u in enumerate(units):
            for rid in u["ids"]:result[rid]=assignment[i]
        return result

    def generate_optimized_scenarios(self):
        self.ensure_baseline();self.save_active_scenario()
        active=self.active_scenario_id();base_snap=self.full_snapshot()
        names=[("balanced","AUTO · Cân bằng tải"),("least_over","AUTO · Ít vượt giờ"),("stable","AUTO · Ít thay đổi")]
        created=[]
        stamp=datetime.datetime.now().strftime("%H%M%S%f")
        try:
            for strategy,label in names:
                self.restore_snapshot(base_snap)
                assignment=self._strategy_assignment(strategy)

                # Áp dụng theo từng lớp/môn và kiểm tra TKB ngay trước mỗi thay đổi.
                # Nếu một thay đổi tạo trùng giờ, giữ phân công cũ cho nhóm đó.
                req_now={r["id"]:r for r in self.requirements()}
                group_targets={}
                for rid,target in assignment.items():
                    r=req_now.get(rid)
                    if not r:continue
                    group_targets.setdefault((r["class_name"],r["subject"]),[]).append((rid,target,r))

                applied_groups=0;skipped_tkb=0
                for (cname,subject),items in sorted(group_targets.items()):
                    targets={t or "" for _,t,_ in items}
                    # Không tự đổi nhóm đang tách CĐ thành nhiều GV vì TKB không phân biệt cốt lõi/CĐ.
                    if len(targets)!=1:
                        continue
                    target=next(iter(targets))
                    current={r["teacher"] or "" for _,_,r in items}
                    if len(current)==1 and next(iter(current))==target:
                        continue
                    if not target:
                        continue
                    sw=min(int(r["start_week"]) for _,_,r in items)
                    ew=max(int(r["end_week"]) for _,_,r in items)
                    if self.timetable_conflicts_for_teacher(target,cname,subject,sw,ew):
                        skipped_tkb+=1
                        continue
                    with self.lock:
                        for rid,_,_ in items:
                            self.cx.execute("UPDATE requirements SET teacher=? WHERE id=?",(target,rid))
                        self.cx.execute("""UPDATE timetable_entries SET teacher=?
                            WHERE year=? AND class_name=? AND subject=?""",
                            (target,self.active_year(),cname,subject))
                        self.cx.commit()
                    applied_groups+=1

                self.invalidate_caches()
                snap=self.full_snapshot();metrics=self.plan_metrics()
                metrics["optimizer_applied_groups"]=applied_groups
                metrics["optimizer_skipped_tkb"]=skipped_tkb
                tv=self.validate_timetable()
                metrics["timetable_errors"]=tv["counts"]["error"]
                metrics["timetable_warnings"]=tv["counts"]["warning"]
                now=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                name=f"{label} {stamp}"
                with self.lock:
                    self.cx.execute("""INSERT INTO scenarios(year,name,description,created_at,updated_at,
                        data_json,metrics_json,is_baseline,workflow_status) VALUES(?,?,?,?,?,?,?,0,'DRAFT')""",
                        (self.active_year(),name,"Sinh tự động bởi bộ tối ưu V10.3",now,now,
                         json.dumps(snap,ensure_ascii=False),json.dumps(metrics,ensure_ascii=False)))
                    sid=self.cx.execute("SELECT last_insert_rowid()").fetchone()[0];self.cx.commit()
                created.append({"id":sid,"name":name,"metrics":metrics})
        finally:
            self.restore_snapshot(base_snap);self.set_setting("active_scenario_id",active)
        return created

    def suggest_relief(self,teacher,target_total=19):
        year=self.active_year();weeks=int(self.setting("weeks",35))
        curmax=self.teacher_max_load(teacher,year)
        reqs=[r for r in self.requirements(year) if r["teacher"]==teacher]
        groups={}
        for r in reqs:groups.setdefault((r["class_name"],r["subject"]),[]).append(r)
        options=[]
        for (cname,subject),parts in groups.items():
            if self.is_locked(cname) or any(self.is_locked(cname,subject,r["req_type"]) for r in parts):continue
            ids=[r["id"] for r in parts]
            oldmx=None
            for w in range(1,weeks+1):
                d=self.teacher_load_at_week(teacher,w,year,exclude_ids=ids)
                if d and (oldmx is None or d["total"]>oldmx["total"]):
                    oldmx=dict(d);oldmx["week"]=w
            rec=[x for x in self.recommend_teachers(cname,subject,None,1,8) if x["teacher"]!=teacher and x["compatible"] and x["covers"] and not x.get("timetable_conflicts")]
            if not rec:continue
            cand=rec[0]
            reduction=num(curmax["total"])-num(oldmx["total"])
            options.append({"class_name":cname,"subject":subject,"from_teacher":teacher,
                            "to_teacher":cand["teacher"],"teacher_max_after":oldmx,
                            "candidate_max_after":cand["max_after"],"reduction":reduction,
                            "target_met":num(oldmx["total"])<=num(target_total)})
        options.sort(key=lambda x:(not x["target_met"],-x["reduction"],num(x["candidate_max_after"]["diff"])))
        return {"teacher":teacher,"target":num(target_total),"current_max":curmax,"options":options[:8]}

    # ---------- Trợ lý phân tích nội bộ ----------
    def assistant_query(self,text,week=1):
        q=(text or "").strip();ql=q.lower()
        if not q:return {"answer":"Nhập câu hỏi về tải giáo viên, lớp hoặc phương án phân công.","actions":[]}
        # Ai dư/thiếu trên X tiết
        m=re.search(r"(dư|vượt|thiếu).{0,12}(\d+(?:[,.]\d+)?)",ql)
        if ("ai" in ql or "giáo viên" in ql) and m:
            mode=m.group(1);thr=num(m.group(2))
            rows=[]
            for t in self.teachers(active_only=True):
                d=self.teacher_load_at_week(t["name"],week)
                diff=num(d["diff"])
                ok=(mode in ("dư","vượt") and diff>=thr) or (mode=="thiếu" and -diff>=thr)
                if ok:rows.append((t["name"],diff,d["total"],d["standard"]))
            rows.sort(key=lambda x:-abs(x[1]))
            ans="; ".join(f'{n}: {self._fmt_hours_vi(t)}/{self._fmt_hours_vi(s)} ({("+" if d>0 else "")}{self._fmt_hours_vi(d)})' for n,d,t,s in rows)
            return {"answer":ans or "Không có giáo viên phù hợp điều kiện ở tuần đang xem.","actions":[]}
        # "Tìm cách giảm cô Lan còn tối đa 19 tiết"
        reduce_match=re.search(r"(?:giảm|hạ).{0,50}(?:còn\s*(?:tối đa\s*)?|xuống\s*|tối đa\s*)(\d+(?:[,.]\d+)?)",ql)
        if reduce_match:
            target=num(reduce_match.group(1),19)
            chosen=None
            for t in self.teachers(active_only=True):
                last=t["name"].split()[-1].lower()
                if t["name"].lower() in ql or re.search(r"\b"+re.escape(last)+r"\b",ql):
                    chosen=t;break
            if chosen:
                plan=self.suggest_relief(chosen["name"],target)
                if not plan["options"]:
                    return {"answer":f'Chưa tìm được chuyển lớp phù hợp để giảm {chosen["name"]} xuống {self._fmt_hours_vi(target)} tiết mà vẫn đúng chuyên môn/hiệu lực/TKB.',"actions":[]}
                lines=[]
                for x in plan["options"][:5]:
                    lines.append(f'{x["class_name"]} {x["subject"]} → {x["to_teacher"]}; tải cao nhất của {chosen["name"]} còn {self._fmt_hours_vi(x["teacher_max_after"]["total"])}; người nhận max {self._fmt_hours_vi(x["candidate_max_after"]["total"])}/{self._fmt_hours_vi(x["candidate_max_after"]["standard"])}')
                return {"answer":f'Hiện {chosen["name"]} cao nhất {self._fmt_hours_vi(plan["current_max"]["total"])} tiết. Các phương án giảm tải:\n'+"\n".join(lines),"actions":[]}

        # Hỏi giả định: "nếu giao 11D cho cô Tâm thì sao?"
        class_match=re.search(r"\b(10|11|12)[A-G]\b",q.upper())
        if ("nếu" in ql or "giao" in ql or "phân công" in ql) and class_match:
            cname=class_match.group(0)
            chosen=None
            for t in self.teachers(active_only=True):
                last=t["name"].split()[-1].lower()
                if t["name"].lower() in ql or re.search(r"\b"+re.escape(last)+r"\b",ql):
                    chosen=t;break
            if chosen:
                subject=chosen["primary_subject"]
                parts=self._target_parts(cname,subject,None)
                if parts:
                    im=self.impact_analysis(cname,subject,chosen["name"],None,week)
                    p=im["target"]
                    risks=[]
                    if p["max_after"]["diff"]>0:risks.append(f'cao nhất vượt {self._fmt_hours_vi(p["max_after"]["diff"])} tiết')
                    if im["timetable_conflicts"]:risks.append(f'{len(im["timetable_conflicts"])} xung đột TKB')
                    if im["blocking_events"]:risks.append("có biến động nhân sự")
                    return {"answer":f'Nếu giao {cname} {subject} cho {chosen["name"]}: tuần {week} tải {self._fmt_hours_vi(p["current"]["total"])} → {self._fmt_hours_vi(p["after"]["total"])}; cao nhất năm {self._fmt_hours_vi(p["max_after"]["total"])}/{self._fmt_hours_vi(p["max_after"]["standard"])} tại T{p["max_after"]["week"]}. '+("Rủi ro: "+", ".join(risks) if risks else "Không phát hiện rủi ro lớn."),
                            "actions":[]}
        # Hỏi tên GV cụ thể
        for t in self.teachers(active_only=True):
            if t["name"].lower() in ql or t["name"].split()[-1].lower() in ql:
                d=self.teacher_load_at_week(t["name"],week);mx=self.teacher_max_load(t["name"])
                return {"answer":f'{t["name"]}: tuần {week} {self._fmt_hours_vi(d["total"])}/{self._fmt_hours_vi(d["standard"])} tiết, dư/thiếu {self._fmt_hours_vi(d["diff"])}. Cao nhất năm T{mx["week"]}: {self._fmt_hours_vi(mx["total"])}/{self._fmt_hours_vi(mx["standard"])}.',
                        "actions":[]}
        if "chưa phân công" in ql or "chưa gán" in ql:
            rows=[r for r in self.requirements() if not r["teacher"]]
            return {"answer":f"Còn {len(rows)} nhu cầu chưa phân công. "+("; ".join(f'{r["class_name"]} {r["subject"]} {r["req_type"]}' for r in rows[:12])),"actions":[]}
        if "lỗi" in ql or "kiểm tra" in ql:
            v=self.validate_all()
            return {"answer":f'Kiểm tra hiện có {v["counts"]["error"]} lỗi, {v["counts"]["warning"]} cảnh báo, {v["counts"]["info"]} thông tin.',"actions":[{"type":"open_page","page":"validation"}]}
        return {"answer":"Tôi có thể trả lời nhanh các câu như: “Ai dư trên 5 tiết?”, “Thầy Dũng bao nhiêu tiết?”, “Còn lớp nào chưa phân công?”, “Kiểm tra lỗi hiện tại”. Với thay đổi cụ thể, dùng nút ★ hoặc Phân tích tác động trên thẻ môn.","actions":[]}

    # ---------- Chuyển năm học thông minh ----------
    def rollover_year(self,new_year,grade10_count=5):
        new_year=(new_year or "").strip()
        if not new_year:raise ValueError("Chưa nhập năm học mới.")
        old=self.active_year()
        if any(y["name"]==new_year for y in self.years()):
            raise ValueError("Năm học này đã tồn tại.")
        # Ghi nhận mẫu tổ hợp đang dùng ở từng tên lớp của năm cũ.
        oldclasses=self.classes(old)
        template_by_class={c["class_name"]:c.get("template_name","") for c in oldclasses}

        self.create_year(new_year,copy_staff=True,copy_classes=False)

        with self.lock:
            # Sao chép toàn bộ mẫu tổ hợp.
            for tp in self.cx.execute("SELECT * FROM templates WHERE year=?",(old,)).fetchall():
                self.cx.execute("INSERT INTO templates(year,name,grade,active,note) VALUES(?,?,?,?,?)",
                                (new_year,tp["name"],tp["grade"],tp["active"],tp["note"]))
                ntid=self.cx.execute("SELECT last_insert_rowid()").fetchone()[0]
                for it in self.cx.execute("SELECT * FROM template_items WHERE template_id=?",(tp["id"],)).fetchall():
                    self.cx.execute("INSERT INTO template_items(template_id,subject,core,specialty) VALUES(?,?,?,?)",
                                    (ntid,it["subject"],it["core"],it["specialty"]))

            new_classes=[]
            # Học viên khối 10 -> 11; 11 -> 12. Template lấy theo lớp đích đang dùng
            # ở năm cũ (11A dùng mẫu 11A, 12A dùng mẫu 12A...), tránh mang template sai khối.
            for c in oldclasses:
                if int(c["grade"]) not in (10,11):continue
                ng=int(c["grade"])+1;name=f'{ng}{c["class_name"][2:]}'
                tp=template_by_class.get(name,"")
                self.cx.execute("""INSERT OR IGNORE INTO classes(
                    year,class_name,grade,program_name,template_name,customized,homeroom_teacher,active
                ) VALUES(?,?,?,?,?,0,'',1)""",(new_year,name,ng,tp,tp))
                new_classes.append((name,tp))

            # Khối 10 mới dùng mẫu đang áp dụng cho 10A/10B/... của năm trước.
            for i in range(int(grade10_count)):
                letter=chr(ord("A")+i);name="10"+letter
                tp=template_by_class.get(name,"")
                self.cx.execute("""INSERT OR IGNORE INTO classes(
                    year,class_name,grade,program_name,template_name,customized,homeroom_teacher,active
                ) VALUES(?,?,?,?,?,0,'',1)""",(new_year,name,10,tp,tp))
                new_classes.append((name,tp))
            self.cx.commit()

        # Sinh lại nhu cầu môn theo template, không sao chép giáo viên phân công cũ.
        for cname,tp in new_classes:
            if tp:
                self.sync_class_to_template(cname,tp,replace=True)

        self.ensure_baseline()
        return {"year":new_year,"classes":len(self.classes(new_year)),
                "message":"Đã lên lớp 10→11, 11→12; tạo khối 10 mới; phân công/CN-KN/ôn TN được để trống để duyệt lại."}

    # ---------- Cache hiệu năng V10.3 ----------
    def invalidate_caches(self):
        self._program_rules_cache=None
        self._program_rules_cache_changes=-1
        self._workload_matrix_cache.clear()

    def workload_matrix(self,year=None):
        """Tính tải tất cả GV × tất cả tuần một lần rồi tái sử dụng."""
        year=year or self.active_year()
        stamp=self.cx.total_changes
        cached=self._workload_matrix_cache.get(year)
        if cached is not None and cached[0]==stamp:
            return cached[1]
        weeks=int(self.setting("weeks",35))
        rules={r["subject"]:r for r in self.program_rules()}
        teachers=self.teachers(year)
        with self.lock:
            reqs=[dict(r) for r in self.cx.execute("SELECT * FROM requirements WHERE year=?",(year,))]
            reviews=[dict(r) for r in self.cx.execute("SELECT * FROM review_jobs WHERE year=?",(year,))]
            duties=[dict(r) for r in self.cx.execute("SELECT * FROM duties WHERE year=?",(year,))]
            events=[dict(r) for r in self.cx.execute("SELECT * FROM staff_events WHERE year=?",(year,))]
        req_by={};rev_by={};duty_by={};event_by={}
        for r in reqs:req_by.setdefault(r.get("teacher") or "",[]).append(r)
        for r in reviews:rev_by.setdefault(r.get("teacher") or "",[]).append(r)
        for r in duties:duty_by.setdefault(r.get("teacher") or "",[]).append(r)
        for r in events:event_by.setdefault(r.get("teacher") or "",[]).append(r)
        out={}
        for t in teachers:
            name=t["name"];arr=[None]*(weeks+1)
            treqs=req_by.get(name,[]);trev=rev_by.get(name,[]);tdut=duty_by.get(name,[]);tevents=event_by.get(name,[])
            for w in range(1,weeks+1):
                teaching=0.0
                for r in treqs:
                    if int(r.get("start_week",1))<=w<=int(r.get("end_week",weeks)):
                        teaching+=self.requirement_weekly_hours(r,w,rules)
                review=0.0
                for r in trev:
                    sw=r.get("start_week")
                    if sw is not None and int(sw)<=w<=int(r.get("end_week",weeks)):
                        review+=num(r.get("periods_per_week"))
                unavailable=[e for e in tevents if int(e.get("blocks_teaching",0)) and int(e.get("start_week",1))<=w<=int(e.get("end_week",weeks))]
                if unavailable:
                    teaching=0.0;review=0.0
                active_duties=[d for d in tdut if int(d.get("start_week",1))<=w<=int(d.get("end_week",weeks))]
                duty_credit=sum(num(d.get("credit")) for d in active_duties)
                homeroom=sum(num(d.get("credit")) for d in active_duties if d.get("duty_type")=="Chủ nhiệm")
                role=sum(num(d.get("credit")) for d in active_duties if d.get("duty_type") in ("Kiêm nhiệm","Chức vụ"))
                std=num(t.get("standard"))
                std_duties=[d for d in active_duties if d.get("duty_type")=="Chức vụ" and d.get("standard_override") not in (None,"")]
                if std_duties:
                    latest=max(std_duties,key=lambda d:(int(d.get("start_week",1)),int(d.get("id",0))))
                    std=num(latest.get("standard_override"),std)
                total=teaching+review+duty_credit
                arr[w]={"teaching":teaching,"review":review,"role":role,"homeroom":homeroom,
                        "duties":active_duties,"duty_credit":duty_credit,"unavailable":bool(unavailable),
                        "staff_events":unavailable,"total":total,"standard":std,"diff":total-std if std else total}
            out[name]=arr
        self._workload_matrix_cache[year]=(self.cx.total_changes,out)
        return out

    def active_year(self):
        with self.lock:
            row=self.cx.execute("SELECT name FROM academic_years WHERE is_active=1 LIMIT 1").fetchone()
            return row["name"] if row else SEED["active_year"]

    def years(self):
        with self.lock:
            return [dict(r) for r in self.cx.execute("SELECT * FROM academic_years ORDER BY name DESC")]

    def set_active_year(self,name):
        with self.lock:
            if not self.cx.execute("SELECT 1 FROM academic_years WHERE name=?",(name,)).fetchone():
                raise ValueError("Năm học không tồn tại.")
            self.cx.execute("UPDATE academic_years SET is_active=0")
            self.cx.execute("UPDATE academic_years SET is_active=1 WHERE name=?",(name,))
            self.cx.commit()
        self.invalidate_caches()

    def create_year(self,name,copy_staff=True,copy_classes=True):
        name=(name or "").strip()
        if not name:raise ValueError("Chưa nhập năm học.")
        with self.lock:
            old=self.active_year()
            self.cx.execute("INSERT INTO academic_years(name,is_active) VALUES(?,0)",(name,))
            if copy_staff:
                for t in self.cx.execute("SELECT * FROM teachers WHERE year=?",(old,)).fetchall():
                    self.cx.execute("""INSERT INTO teachers(
                        year,stt,name,primary_subject,can_teach,standard,employment_type,
                        active_start,active_end,role,role_credit,homeroom,homeroom_credit,note,active
                    ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(
                        name,t["stt"],t["name"],t["primary_subject"],t["can_teach"],t["standard"],
                        t["employment_type"],t["active_start"],t["active_end"],t["role"],
                        t["role_credit"],t["homeroom"],t["homeroom_credit"],t["note"],t["active"]
                    ))
            if copy_classes:
                # Sao chép mẫu tổ hợp của năm cũ sang năm mới.
                template_id_map={}
                for tp in self.cx.execute("SELECT * FROM templates WHERE year=?",(old,)).fetchall():
                    self.cx.execute("""INSERT INTO templates(year,name,grade,active,note)
                        VALUES(?,?,?,?,?)""",(name,tp["name"],tp["grade"],tp["active"],tp["note"]))
                    new_tid=self.cx.execute("SELECT last_insert_rowid()").fetchone()[0]
                    template_id_map[tp["id"]]=new_tid
                    for it in self.cx.execute("SELECT * FROM template_items WHERE template_id=?",(tp["id"],)).fetchall():
                        self.cx.execute("""INSERT INTO template_items(template_id,subject,core,specialty)
                            VALUES(?,?,?,?)""",(new_tid,it["subject"],it["core"],it["specialty"]))
                for cl in self.cx.execute("SELECT * FROM classes WHERE year=?",(old,)).fetchall():
                    self.cx.execute("""INSERT INTO classes(
                        year,class_name,grade,program_name,template_name,customized,homeroom_teacher,active
                    ) VALUES(?,?,?,?,?,?,?,?)""",(name,cl["class_name"],cl["grade"],cl["program_name"],
                                                  cl["template_name"],cl["customized"],
                                                  cl["homeroom_teacher"],cl["active"]))
                for r in self.cx.execute("SELECT * FROM requirements WHERE year=?",(old,)).fetchall():
                    self.cx.execute("""INSERT INTO requirements(
                        year,class_name,subject,req_type,annual_hours,teacher,start_week,end_week,note
                    ) VALUES(?,?,?,?,?,'',?,?,?)""",(name,r["class_name"],r["subject"],r["req_type"],
                                                     r["annual_hours"],r["start_week"],r["end_week"],
                                                     "Sao chép cấu trúc từ "+old))
            self.cx.commit()
            self.set_active_year(name)

    def teachers(self,year=None,active_only=False):
        year=year or self.active_year()
        q="SELECT * FROM teachers WHERE year=?";params=[year]
        if active_only:q+=" AND active=1"
        q+=" ORDER BY stt,name"
        with self.lock:
            rows=[dict(r) for r in self.cx.execute(q,params)]
        for r in rows:
            r["primary_subject"]=canonical_subject_name(r.get("primary_subject",""))
            try:r["can_teach_list"]=[canonical_subject_name(x) for x in json.loads(r["can_teach"] or "[]")]
            except:r["can_teach_list"]=[r["primary_subject"]] if r["primary_subject"] else []
            r["can_teach_list"]=list(dict.fromkeys(x for x in r["can_teach_list"] if x))
        return rows

    def save_teacher(self,d):
        year=d.get("year") or self.active_year()
        name=(d.get("name") or "").strip()
        if not name:raise ValueError("Chưa nhập họ tên.")
        can=d.get("can_teach") or []
        if isinstance(can,str):can=[x.strip() for x in can.split(",") if x.strip()]
        can=[canonical_subject_name(x) for x in can]
        primary=canonical_subject_name(d.get("primary_subject",""))
        if primary and primary not in can:can.insert(0,primary)
        vals=(int(d.get("stt") or 0),name,primary,json.dumps(can,ensure_ascii=False),
              num(d.get("standard"),17),d.get("employment_type","Cơ hữu"),
              int(d.get("active_start") or 1),int(d.get("active_end") or 35),
              d.get("role",""),num(d.get("role_credit")),d.get("homeroom",""),
              num(d.get("homeroom_credit")),d.get("note",""),int(d.get("active",1)))
        with self.lock:
            tid=d.get("id")
            if tid:
                self.cx.execute("""UPDATE teachers SET
                    stt=?,name=?,primary_subject=?,can_teach=?,standard=?,employment_type=?,
                    active_start=?,active_end=?,role=?,role_credit=?,homeroom=?,homeroom_credit=?,
                    note=?,active=? WHERE id=?""",vals+(tid,))
            else:
                self.cx.execute("""INSERT INTO teachers(
                    year,stt,name,primary_subject,can_teach,standard,employment_type,
                    active_start,active_end,role,role_credit,homeroom,homeroom_credit,note,active
                ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(year,)+vals)
            self.cx.commit()

    def toggle_teacher(self,tid):
        with self.lock:
            self.cx.execute("UPDATE teachers SET active=1-active WHERE id=?",(tid,))
            self.cx.commit()

    def templates(self,year=None,active_only=False):
        year=year or self.active_year()
        q="SELECT * FROM templates WHERE year=?";params=[year]
        if active_only:q+=" AND active=1"
        q+=" ORDER BY grade,name"
        with self.lock:
            rows=[dict(r) for r in self.cx.execute(q,params)]
            for tp in rows:
                tp["items"]=[dict(x) for x in self.cx.execute(
                    "SELECT * FROM template_items WHERE template_id=? ORDER BY subject",(tp["id"],))]
                tp["class_count"]=self.cx.execute(
                    "SELECT COUNT(*) n FROM classes WHERE year=? AND template_name=? AND active=1",
                    (year,tp["name"])).fetchone()["n"]
            return rows

    def save_template(self,d):
        year=d.get("year") or self.active_year()
        name=(d.get("name") or "").strip()
        if not name:raise ValueError("Chưa nhập tên tổ hợp.")
        grade=int(d.get("grade") or 10)
        items=d.get("items") or []
        if not any(int(x.get("core",0)) for x in items):
            raise ValueError("Tổ hợp phải có ít nhất một môn cốt lõi.")
        with self.lock:
            tid=d.get("id")
            if tid:
                old=self.cx.execute("SELECT * FROM templates WHERE id=?",(tid,)).fetchone()
                if not old:raise ValueError("Không tìm thấy tổ hợp.")
                old_name=old["name"]
                self.cx.execute("UPDATE templates SET name=?,grade=?,active=?,note=? WHERE id=?",
                                (name,grade,int(d.get("active",1)),d.get("note",""),tid))
                if old_name!=name:
                    self.cx.execute("""UPDATE classes SET template_name=?,program_name=?
                        WHERE year=? AND template_name=?""",(name,name,year,old_name))
                self.cx.execute("DELETE FROM template_items WHERE template_id=?",(tid,))
            else:
                self.cx.execute("""INSERT INTO templates(year,name,grade,active,note)
                    VALUES(?,?,?,?,?)""",(year,name,grade,int(d.get("active",1)),d.get("note","")))
                tid=self.cx.execute("SELECT last_insert_rowid()").fetchone()[0]
            for it in items:
                if not int(it.get("core",0)) and not int(it.get("specialty",0)):
                    continue
                self.cx.execute("""INSERT INTO template_items(template_id,subject,core,specialty)
                    VALUES(?,?,?,?)""",(tid,it["subject"],int(it.get("core",0)),int(it.get("specialty",0))))
            self.cx.commit()
            if d.get("update_classes"):
                for cl in self.cx.execute("""SELECT class_name FROM classes
                    WHERE year=? AND template_name=?""",(year,name)).fetchall():
                    self.sync_class_to_template(cl["class_name"],name,replace=True)

    def toggle_template(self,tid):
        with self.lock:
            self.cx.execute("UPDATE templates SET active=1-active WHERE id=?",(tid,))
            self.cx.commit()

    def copy_template(self,tid,new_name=None):
        year=self.active_year()
        with self.lock:
            tp=self.cx.execute("SELECT * FROM templates WHERE id=?",(tid,)).fetchone()
            if not tp:raise ValueError("Không tìm thấy tổ hợp.")
            name=(new_name or (tp["name"]+" - Bản sao")).strip()
            base=name;i=2
            while self.cx.execute("SELECT 1 FROM templates WHERE year=? AND name=?",(year,name)).fetchone():
                name=f"{base} {i}";i+=1
            self.cx.execute("""INSERT INTO templates(year,name,grade,active,note)
                VALUES(?,?,?,?,?)""",(year,name,tp["grade"],1,"Sao chép từ "+tp["name"]))
            new_id=self.cx.execute("SELECT last_insert_rowid()").fetchone()[0]
            for it in self.cx.execute("SELECT * FROM template_items WHERE template_id=?",(tid,)).fetchall():
                self.cx.execute("""INSERT INTO template_items(template_id,subject,core,specialty)
                    VALUES(?,?,?,?)""",(new_id,it["subject"],it["core"],it["specialty"]))
            self.cx.commit()
            return {"id":new_id,"name":name}

    def template_by_name(self,name,year=None):
        year=year or self.active_year()
        with self.lock:
            tp=self.cx.execute("SELECT * FROM templates WHERE year=? AND name=?",(year,name)).fetchone()
            if not tp:return None
            d=dict(tp)
            d["items"]=[dict(x) for x in self.cx.execute(
                "SELECT * FROM template_items WHERE template_id=? ORDER BY subject",(tp["id"],))]
            return d

    def sync_class_to_template(self,class_name,template_name,replace=True):
        year=self.active_year()
        tp=self.template_by_name(template_name,year)
        if not tp:raise ValueError("Tổ hợp không tồn tại.")
        with self.lock:
            target=set()
            rules={r["subject"]:r for r in self.program_rules()}
            for it in tp["items"]:
                if int(it["core"]):
                    target.add((it["subject"],"Cốt lõi"))
                if int(it["specialty"]):
                    target.add((it["subject"],"Chuyên đề"))
            existing=[dict(r) for r in self.cx.execute("""SELECT * FROM requirements
                WHERE year=? AND class_name=?""",(year,class_name))]
            existing_map={(r["subject"],r["req_type"]):r for r in existing}
            for subject,typ in target:
                if (subject,typ) in existing_map:continue
                rule=rules.get(subject)
                if not rule:continue
                hours=rule["specialty_hours"] if typ=="Chuyên đề" else rule["base_hours"]
                self.cx.execute("""INSERT OR IGNORE INTO requirements(
                    year,class_name,subject,req_type,annual_hours,teacher,start_week,end_week,note
                ) VALUES(?,?,?,?,?,'',1,35,?)""",(year,class_name,subject,typ,hours,
                                                  "Sinh từ tổ hợp "+template_name))
            if replace:
                for key,r in existing_map.items():
                    if key not in target:
                        self.cx.execute("DELETE FROM requirements WHERE id=?",(r["id"],))
            self.cx.execute("""UPDATE classes SET template_name=?,program_name=?,customized=0
                WHERE year=? AND class_name=?""",(template_name,template_name,year,class_name))
            self.cx.commit()

    def classes(self,year=None,active_only=False):
        year=year or self.active_year()
        q="SELECT * FROM classes WHERE year=?";params=[year]
        if active_only:q+=" AND active=1"
        q+=" ORDER BY grade,class_name"
        with self.lock:
            rows=[dict(r) for r in self.cx.execute(q,params)]
        reqs=self.requirements(year)
        for c in rows:
            parts=[r for r in reqs if r["class_name"]==c["class_name"]]
            grouped={}
            for r in parts:
                g=grouped.setdefault(r["subject"],{"core":False,"specialty":False})
                if r["req_type"]=="Chuyên đề":g["specialty"]=True
                else:g["core"]=True
            c["structure"]=[{"subject":s,**v} for s,v in sorted(grouped.items())]
        return rows

    def save_class(self,d):
        year=d.get("year") or self.active_year()
        cname=(d.get("class_name") or "").strip().upper()
        if not cname:raise ValueError("Chưa nhập tên lớp.")
        grade=int(d.get("grade") or cname[:2])
        template_name=(d.get("template_name") or "").strip()
        with self.lock:
            cid=d.get("id")
            old_name=None;old_template=None
            if cid:
                old=self.cx.execute("SELECT * FROM classes WHERE id=?",(cid,)).fetchone()
                if not old:raise ValueError("Không tìm thấy lớp.")
                old_name=old["class_name"];old_template=old["template_name"] or ""
                if old_name!=cname:
                    self.cx.execute("""UPDATE requirements SET class_name=?
                        WHERE year=? AND class_name=?""",(cname,year,old_name))
                self.cx.execute("""UPDATE classes SET
                    class_name=?,grade=?,program_name=?,template_name=?,homeroom_teacher=?,active=?
                    WHERE id=?""",(cname,grade,template_name,template_name,
                                  d.get("homeroom_teacher",""),int(d.get("active",1)),cid))
            else:
                self.cx.execute("""INSERT INTO classes(
                    year,class_name,grade,program_name,template_name,customized,homeroom_teacher,active
                ) VALUES(?,?,?,?,?,0,?,1)""",(year,cname,grade,template_name,template_name,
                                             d.get("homeroom_teacher","")))
            self.cx.commit()
        # Khi tạo mới hoặc đổi tổ hợp: tự sinh lại cấu trúc môn.
        if template_name and (not cid or template_name!=old_template or d.get("force_sync")):
            self.sync_class_to_template(cname,template_name,replace=True)

    def save_class_structure(self,d):
        year=d.get("year") or self.active_year()
        cname=(d.get("class_name") or "").strip().upper()
        items=d.get("items") or []
        rules={r["subject"]:r for r in self.program_rules()}
        target=set()
        for it in items:
            if int(it.get("core",0)):target.add((it["subject"],"Cốt lõi"))
            if int(it.get("specialty",0)):target.add((it["subject"],"Chuyên đề"))
        with self.lock:
            existing=[dict(r) for r in self.cx.execute(
                "SELECT * FROM requirements WHERE year=? AND class_name=?",(year,cname))]
            emap={(r["subject"],r["req_type"]):r for r in existing}
            for subject,typ in target:
                if (subject,typ) in emap:continue
                rule=rules.get(subject)
                if not rule:continue
                hours=rule["specialty_hours"] if typ=="Chuyên đề" else rule["base_hours"]
                self.cx.execute("""INSERT OR IGNORE INTO requirements(
                    year,class_name,subject,req_type,annual_hours,teacher,start_week,end_week,note
                ) VALUES(?,?,?,?,?,'',1,35,'Điều chỉnh riêng của lớp')""",
                (year,cname,subject,typ,hours))
            for key,r in emap.items():
                if key not in target:self.cx.execute("DELETE FROM requirements WHERE id=?",(r["id"],))
            self.cx.execute("""UPDATE classes SET customized=1
                WHERE year=? AND class_name=?""",(year,cname))
            self.cx.commit()

    def restore_class_template(self,class_name):
        year=self.active_year()
        with self.lock:
            cl=self.cx.execute("SELECT * FROM classes WHERE year=? AND class_name=?",
                               (year,class_name)).fetchone()
            if not cl:raise ValueError("Không tìm thấy lớp.")
            if not cl["template_name"]:raise ValueError("Lớp chưa gán tổ hợp.")
            template_name=cl["template_name"]
        self.sync_class_to_template(class_name,template_name,replace=True)

    def toggle_class(self,cid):
        with self.lock:
            self.cx.execute("UPDATE classes SET active=1-active WHERE id=?",(cid,))
            self.cx.commit()

    def _normalize_schedule(self,schedule):
        weeks=int(self.setting("weeks",35))
        out=[]
        for seg in (schedule or []):
            a=int(seg.get("start") or 0); b=int(seg.get("end") or 0); p=num(seg.get("periods"))
            if not (1<=a<=b<=weeks):
                raise ValueError(f"Khoảng tuần {a}-{b} không hợp lệ.")
            if p<0:
                raise ValueError("Số tiết/tuần không được âm.")
            if p==0:
                continue
            out.append({"start":a,"end":b,"periods":p})
        out.sort(key=lambda x:(x["start"],x["end"]))
        prev_end=0
        for seg in out:
            if seg["start"]<=prev_end:
                raise ValueError("Các khoảng tuần phân bổ không được chồng lấn.")
            prev_end=seg["end"]
        return out

    def _schedule_total(self,schedule):
        return sum((int(s["end"])-int(s["start"])+1)*num(s["periods"]) for s in (schedule or []))

    def _schedule_hours_at_week(self,schedule,week):
        for s in (schedule or []):
            if int(s["start"])<=int(week)<=int(s["end"]):
                return num(s["periods"])
        return 0.0

    def _schedule_summary(self,schedule):
        if not schedule:
            return "Không bố trí"
        parts=[]
        for s in schedule:
            a,b=int(s["start"]),int(s["end"])
            p=self._fmt_hours_vi(s["periods"])
            if a==1 and b==int(self.setting("weeks",35)):
                parts.append(f"{p} × {b-a+1} tuần")
            else:
                parts.append(f"T{a}-{b}: {p} tiết/tuần")
        return "; ".join(parts)

    def program_rules(self):
        if self._program_rules_cache is not None and self._program_rules_cache_changes==self.cx.total_changes:
            return self._program_rules_cache
        with self.lock:
            rows=[dict(r) for r in self.cx.execute("SELECT * FROM program_rules ORDER BY subject")]
        for r in rows:
            try:r["core_schedule"]=json.loads(r.get("core_schedule_json") or "[]")
            except:r["core_schedule"]=[]
            try:r["specialty_schedule"]=json.loads(r.get("specialty_schedule_json") or "[]")
            except:r["specialty_schedule"]=[]
            r["core_schedule_summary"]=self._schedule_summary(r["core_schedule"])
            r["specialty_schedule_summary"]=self._schedule_summary(r["specialty_schedule"]) if num(r["specialty_hours"]) else "—"
            r["total_with_specialty"]=num(r["base_hours"])+(num(r["specialty_hours"]) if int(r["can_specialty"]) else 0)
            r["core_schedule_total"]=self._schedule_total(r["core_schedule"])
            r["specialty_schedule_total"]=self._schedule_total(r["specialty_schedule"])
        self._program_rules_cache=rows
        self._program_rules_cache_changes=self.cx.total_changes
        return rows

    def save_program_rule(self,d):
        rid=int(d["id"])
        base=num(d.get("base_hours"))
        specialty=num(d.get("specialty_hours"))
        can=int(d.get("can_specialty",0))
        core=self._normalize_schedule(d.get("core_schedule") or [])
        spec=self._normalize_schedule(d.get("specialty_schedule") or [])
        if abs(self._schedule_total(core)-base)>1e-6:
            raise ValueError(
                f"Phân bổ cốt lõi đang có {self._fmt_hours_vi(self._schedule_total(core))} tiết, "
                f"phải bằng {self._fmt_hours_vi(base)} tiết/năm."
            )
        if can:
            if abs(self._schedule_total(spec)-specialty)>1e-6:
                raise ValueError(
                    f"Phân bổ chuyên đề đang có {self._fmt_hours_vi(self._schedule_total(spec))} tiết, "
                    f"phải bằng {self._fmt_hours_vi(specialty)} tiết/năm."
                )
        else:
            specialty=0
            spec=[]
        with self.lock:
            row=self.cx.execute("SELECT subject FROM program_rules WHERE id=?",(rid,)).fetchone()
            if not row:raise ValueError("Không tìm thấy môn.")
            subject=row["subject"]
            self.cx.execute("""UPDATE program_rules SET
                base_hours=?,specialty_hours=?,can_specialty=?,source=?,
                core_schedule_json=?,specialty_schedule_json=?,legal_locked=?
                WHERE id=?""",(
                    base,specialty,can,d.get("source",""),
                    json.dumps(core,ensure_ascii=False),json.dumps(spec,ensure_ascii=False),
                    int(d.get("legal_locked",1)),rid
                ))
            # Đồng bộ số tiết/năm của các nhu cầu đang tồn tại.
            self.cx.execute("""UPDATE requirements SET annual_hours=?
                WHERE subject=? AND req_type='Cốt lõi'""",(base,subject))
            self.cx.execute("""UPDATE requirements SET annual_hours=?
                WHERE subject=? AND req_type='Chuyên đề'""",(specialty,subject))
            self.cx.commit()

    def requirement_weekly_hours(self,req,week,rules_map=None):
        if not (int(req.get("start_week",1))<=int(week)<=int(req.get("end_week",self.setting("weeks",35)))):
            return 0.0
        if rules_map is None:
            rules_map={r["subject"]:r for r in self.program_rules()}
        rule=rules_map.get(req["subject"])
        if not rule:
            return num(req.get("annual_hours"))/int(self.setting("weeks",35))
        schedule=rule["specialty_schedule"] if req.get("req_type")=="Chuyên đề" else rule["core_schedule"]
        return self._schedule_hours_at_week(schedule,week)

    def requirement_schedule_summary(self,req,rules_map=None):
        if rules_map is None:
            rules_map={r["subject"]:r for r in self.program_rules()}
        rule=rules_map.get(req["subject"])
        if not rule:return ""
        return rule["specialty_schedule_summary"] if req.get("req_type")=="Chuyên đề" else rule["core_schedule_summary"]

    def requirements(self,year=None):
        year=year or self.active_year()
        weeks=int(self.setting("weeks",35))
        rules={r["subject"]:r for r in self.program_rules()}
        with self.lock:
            rows=[dict(r) for r in self.cx.execute("""SELECT * FROM requirements
                WHERE year=? ORDER BY class_name,subject,req_type""",(year,))]
        for r in rows:
            r["weekly_equiv"]=num(r["annual_hours"])/(weeks or 35)
            r["schedule_summary"]=self.requirement_schedule_summary(r,rules)
        return rows

    def save_requirement(self,d):
        year=d.get("year") or self.active_year()
        cname=(d.get("class_name") or "").strip().upper()
        subject=d.get("subject");typ=d.get("req_type","Cốt lõi")
        if not cname or not subject:raise ValueError("Thiếu lớp hoặc môn.")
        with self.lock:
            rule=self.cx.execute("SELECT * FROM program_rules WHERE subject=?",(subject,)).fetchone()
            if not rule:raise ValueError("Chưa có định mức môn.")
            hours=rule["specialty_hours"] if typ=="Chuyên đề" else rule["base_hours"]
            self.cx.execute("""INSERT INTO requirements(
                year,class_name,subject,req_type,annual_hours,teacher,start_week,end_week,note
            ) VALUES(?,?,?,?,?,'',1,35,'')
            ON CONFLICT(year,class_name,subject,req_type) DO UPDATE SET annual_hours=excluded.annual_hours""",
            (year,cname,subject,typ,hours))
            self.cx.commit()

    def delete_requirement(self,rid):
        with self.lock:self.cx.execute("DELETE FROM requirements WHERE id=?",(rid,));self.cx.commit()

    def subject_groups(self,year=None):
        year=year or self.active_year()
        reqs=self.requirements(year)
        groups={}
        for r in reqs:
            key=(r["class_name"],r["subject"])
            g=groups.setdefault(key,{
                "class_name":r["class_name"],"subject":r["subject"],"parts":[],
                "total_hours":0,"weekly_equiv":0,"has_specialty":False
            })
            g["parts"].append(r)
            g["total_hours"]+=num(r["annual_hours"])
            g["weekly_equiv"]+=num(r["weekly_equiv"])
            if r["req_type"]=="Chuyên đề":g["has_specialty"]=True
        out=[]
        for g in groups.values():
            ts=[p["teacher"] for p in g["parts"] if p["teacher"]]
            uniq=sorted(set(ts))
            if not ts:g["assignment_status"]="unassigned";g["teacher_label"]=""
            elif len(uniq)==1 and len(ts)==len(g["parts"]):
                g["assignment_status"]="same";g["teacher_label"]=uniq[0]
            else:
                g["assignment_status"]="split"
                g["teacher_label"]=" / ".join(f'{p["req_type"]}: {p["teacher"] or "—"}' for p in g["parts"])
            out.append(g)
        out.sort(key=lambda x:(x["class_name"],x["subject"]))
        return out

    def _target_parts(self,class_name,subject,req_type=None):
        q="SELECT * FROM requirements WHERE year=? AND class_name=? AND subject=?"
        params=[self.active_year(),class_name,subject]
        if req_type:
            q+=" AND req_type=?";params.append(req_type)
        q+=" ORDER BY req_type"
        with self.lock:return [dict(r) for r in self.cx.execute(q,params)]

    def teacher_load_at_week(self,teacher,week,year=None,add_parts=None,exclude_ids=None,
                             exclude_review_ids=None,add_review_periods=0,add_review_start=None,add_review_end=None):
        year=year or self.active_year()
        if not add_parts and not exclude_ids and not exclude_review_ids and not add_review_periods:
            arr=self.workload_matrix(year).get(teacher)
            if arr and 1<=int(week)<len(arr):
                return dict(arr[int(week)])
        exclude_ids=set(exclude_ids or [])
        exclude_review_ids=set(exclude_review_ids or [])
        with self.lock:
            t=self.cx.execute("SELECT * FROM teachers WHERE year=? AND name=?",(year,teacher)).fetchone()
            if not t:return None
            rows=self.cx.execute("""SELECT * FROM requirements
                WHERE year=? AND teacher=? AND start_week<=? AND end_week>=?""",
                (year,teacher,week,week)).fetchall()
            rules={r["subject"]:r for r in self.program_rules()}
            teaching=sum(self.requirement_weekly_hours(dict(r),week,rules) for r in rows if r["id"] not in exclude_ids)
            if add_parts:
                for p in add_parts:
                    if p["id"] not in exclude_ids and p["teacher"]!=teacher and p["start_week"]<=week<=p["end_week"]:
                        teaching+=self.requirement_weekly_hours(p,week,rules)
            rev_rows=self.cx.execute("""SELECT id,periods_per_week FROM review_jobs
                WHERE year=? AND teacher=? AND start_week IS NOT NULL
                AND start_week<=? AND end_week>=?""",(year,teacher,week,week)).fetchall()
            review=sum(num(r["periods_per_week"]) for r in rev_rows if r["id"] not in exclude_review_ids)
            if add_review_periods and add_review_start is not None and add_review_start<=week<=int(add_review_end or self.setting("weeks",35)):
                review+=num(add_review_periods)

            # Nếu có biến động nhân sự "chặn khả năng dạy", phần tải dạy và ôn
            # của giáo viên được xem là không thực hiện trong giai đoạn đó.
            unavailable_events=self.blocking_events(teacher,week,week,year)
            if unavailable_events:
                teaching=0.0
                review=0.0

            duty_rows=self.duties_at_week(teacher,week,year)
            duty_credit=sum(num(x["credit"]) for x in duty_rows)
            homeroom_credit=sum(num(x["credit"]) for x in duty_rows if x["duty_type"]=="Chủ nhiệm")
            role_credit=sum(num(x["credit"]) for x in duty_rows if x["duty_type"] in ("Kiêm nhiệm","Chức vụ"))
            total=teaching+review+duty_credit
            std=num(t["standard"])
            std_duties=[x for x in duty_rows if x["duty_type"]=="Chức vụ" and x.get("standard_override")]
            if std_duties:
                latest=max(std_duties,key=lambda x:(int(x["start_week"]),int(x["id"])))
                std=num(latest["standard_override"],std)
            return {"teaching":teaching,"review":review,"role":role_credit,
                    "homeroom":homeroom_credit,"duties":duty_rows,"duty_credit":duty_credit,
                    "unavailable":bool(unavailable_events),"staff_events":unavailable_events,
                    "total":total,"standard":std,"diff":total-std if std else total}

    def teacher_max_load(self,teacher,year=None,add_parts=None,exclude_ids=None,
                         exclude_review_ids=None,add_review_periods=0,add_review_start=None,add_review_end=None):
        year=year or self.active_year()
        if not add_parts and not exclude_ids and not exclude_review_ids and not add_review_periods:
            arr=self.workload_matrix(year).get(teacher) or []
            rows=[(w,arr[w]) for w in range(1,len(arr)) if arr[w] is not None]
            if not rows:return None
            w,d=max(rows,key=lambda wd:(num(wd[1]["diff"]),num(wd[1]["total"])))
            out=dict(d);out["week"]=w;return out
        best=None
        for w in range(1,int(self.setting("weeks",35))+1):
            d=self.teacher_load_at_week(
                teacher,w,year,add_parts,exclude_ids,exclude_review_ids,
                add_review_periods,add_review_start,add_review_end
            )
            if d and (best is None or d["diff"]>best["diff"] or
                      (abs(d["diff"]-best["diff"])<1e-9 and d["total"]>best["total"])):
                best=dict(d);best["week"]=w
        return best

    def _fmt_hours_vi(self, value):
        v=round(num(value),6)
        if abs(v-round(v))<1e-9:
            return str(int(round(v)))
        s=f"{v:.6f}".rstrip("0").rstrip(".")
        return s.replace(".",",")

    def workload_change_note(self, teacher_name, year=None):
        """
        Nén tải 35 tuần thành các giai đoạn liên tiếp có cùng mức dư/thiếu.
        Ví dụ:
        Tuần 1-2 thiếu 1 tiết; Tuần 3-35 dư 2 tiết
        """
        year=year or self.active_year()
        weeks=int(self.setting("weeks",35))
        with self.lock:
            t=self.cx.execute("SELECT * FROM teachers WHERE year=? AND name=?",
                              (year,teacher_name)).fetchone()
        if not t:
            return ""
        if num(t["standard"])<=0:
            return "Chưa có chuẩn để tính dư/thiếu"

        values=[]
        for w in range(1,weeks+1):
            d=self.teacher_load_at_week(teacher_name,w,year)
            diff=round(num(d["diff"]),6) if d else 0
            # Chuẩn hóa sai số số thực
            if abs(diff)<1e-6: diff=0.0
            values.append((w,diff))

        segments=[]
        start=values[0][0]
        prev_diff=values[0][1]
        prev_week=values[0][0]

        for w,diff in values[1:]:
            if abs(diff-prev_diff)>1e-6:
                segments.append((start,prev_week,prev_diff))
                start=w
                prev_diff=diff
            prev_week=w
        segments.append((start,prev_week,prev_diff))

        parts=[]
        for a,b,diff in segments:
            week_text=f"Tuần {a}" if a==b else f"Tuần {a}-{b}"
            if diff>0:
                status=f"dư {self._fmt_hours_vi(diff)} tiết"
            elif diff<0:
                status=f"thiếu {self._fmt_hours_vi(abs(diff))} tiết"
            else:
                status="đủ chuẩn"
            parts.append(f"{week_text} {status}")
        return "; ".join(parts)

    def workload_change_note_range(self,teacher_name,start_week,end_week,year=None):
        year=year or self.active_year();a=int(start_week);b=int(end_week)
        values=[]
        for w in range(a,b+1):
            d=self.teacher_load_at_week(teacher_name,w,year)
            diff=round(num(d["diff"]),6) if d else 0.0
            if abs(diff)<1e-6:diff=0.0
            values.append((w,diff))
        if not values:return ""
        seg=[];s=values[0][0];pd=values[0][1];pw=values[0][0]
        for w,diff in values[1:]:
            if abs(diff-pd)>1e-6:
                seg.append((s,pw,pd));s=w;pd=diff
            pw=w
        seg.append((s,pw,pd))
        parts=[]
        for x,y,diff in seg:
            wt=f"Tuần {x}" if x==y else f"Tuần {x}-{y}"
            st=(f"dư {self._fmt_hours_vi(diff)} tiết" if diff>0 else
                f"thiếu {self._fmt_hours_vi(abs(diff))} tiết" if diff<0 else "đủ chuẩn")
            parts.append(f"{wt} {st}")
        return "; ".join(parts)

    def teacher_period_summary(self,teacher,start_week,end_week,year=None):
        year=year or self.active_year();a=int(start_week);b=int(end_week)
        rows=[]
        for w in range(a,b+1):
            d=self.teacher_load_at_week(teacher,w,year)
            if d:
                x=dict(d);x["week"]=w;rows.append(x)
        if not rows:return {"total":0,"standard":0,"diff":0,"average":0,"average_standard":0,"max":None,"weeks":0}
        teaching=round(sum(num(x["teaching"]) for x in rows),6);review=round(sum(num(x["review"]) for x in rows),6)
        homeroom=round(sum(num(x["homeroom"]) for x in rows),6);role=round(sum(num(x["role"]) for x in rows),6)
        total=round(sum(num(x["total"]) for x in rows),6);standard=round(sum(num(x["standard"]) for x in rows),6)
        mx=max(rows,key=lambda x:(num(x["diff"]),num(x["total"])))
        return {"teaching":teaching,"review":review,"homeroom":homeroom,"role":role,
                "total":total,"standard":standard,"diff":round(total-standard,6),
                "average":round(total/len(rows),6),"average_standard":round(standard/len(rows),6),
                "max":mx,"weeks":len(rows),"note":self.workload_change_note_range(teacher,a,b,year)}

    def teacher_loads_summary(self,week):
        year=self.active_year()
        out={}
        for t in self.teachers(year,active_only=True):
            cur=self.teacher_load_at_week(t["name"],week,year)
            mx=self.teacher_max_load(t["name"],year)
            out[t["name"]]={"current":cur,"max":mx}
        return out

    def preview_drop(self,class_name,subject,teacher,req_type=None,week=1):
        parts=self._target_parts(class_name,subject,req_type)
        if not parts:raise ValueError("Không tìm thấy thẻ môn.")
        with self.lock:
            t=self.cx.execute("SELECT * FROM teachers WHERE year=? AND name=?",
                              (self.active_year(),teacher)).fetchone()
            if not t:raise ValueError("Giáo viên không tồn tại.")
            try:can=json.loads(t["can_teach"] or "[]")
            except:can=[t["primary_subject"]] if t["primary_subject"] else []
        compatible=subject in can or subject==t["primary_subject"]
        min_start=min(p["start_week"] for p in parts);max_end=max(p["end_week"] for p in parts)
        blocking=self.blocking_events(teacher,min_start,max_end)
        covers=int(t["active_start"])<=min_start and int(t["active_end"])>=max_end and not blocking
        current=self.teacher_load_at_week(teacher,week)
        after=self.teacher_load_at_week(teacher,week,add_parts=parts)
        mx=self.teacher_max_load(teacher,add_parts=parts)
        return {
            "class_name":class_name,"subject":subject,"req_type":req_type or "Tất cả",
            "teacher":teacher,"compatible":compatible,"covers":covers,"blocking_events":blocking,
            "current":current,"after":after,"max_after":mx,
            "added_weekly":sum(self.requirement_weekly_hours(p,week) for p in parts if p["teacher"]!=teacher),
            "parts":[{"id":p["id"],"req_type":p["req_type"],"annual_hours":p["annual_hours"],"old_teacher":p["teacher"]} for p in parts]
        }

    def assign_drop(self,class_name,subject,teacher,req_type=None):
        parts=self._target_parts(class_name,subject,req_type)
        if not parts:raise ValueError("Không tìm thấy nhu cầu.")
        if self.is_locked(class_name):
            raise ValueError("LOCKED|Lớp "+class_name+" đang bị khóa.")
        for p in parts:
            if self.is_locked(class_name,subject,p["req_type"]):
                raise ValueError("LOCKED|"+self.requirement_lock_key(class_name,subject,p["req_type"])+" đang bị khóa.")
        with self.lock:
            for p in parts:
                self.cx.execute("UPDATE requirements SET teacher=? WHERE id=?",(teacher or "",p["id"]))
            self.cx.commit()

    def duties(self,year=None,teacher=None):
        year=year or self.active_year()
        q="SELECT * FROM duties WHERE year=?";params=[year]
        if teacher:
            q+=" AND teacher=?";params.append(teacher)
        q+=" ORDER BY teacher,start_week,end_week,duty_type,duty_name"
        with self.lock:
            return [dict(r) for r in self.cx.execute(q,params)]

    def duties_at_week(self,teacher,week,year=None,exclude_ids=None):
        year=year or self.active_year()
        exclude_ids=set(exclude_ids or [])
        with self.lock:
            rows=[dict(r) for r in self.cx.execute("""SELECT * FROM duties
                WHERE year=? AND teacher=? AND start_week<=? AND end_week>=?
                ORDER BY duty_type,duty_name""",(year,teacher,week,week))]
        return [r for r in rows if r["id"] not in exclude_ids]

    def duty_credit_at_week(self,teacher,week,year=None,exclude_ids=None):
        return sum(num(d["credit"]) for d in self.duties_at_week(teacher,week,year,exclude_ids))

    def duty_conflicts(self,teacher,start_week,end_week,year=None,exclude_id=None,duty_type=None):
        year=year or self.active_year()
        q="""SELECT * FROM duties
             WHERE year=? AND teacher=? AND NOT(end_week<? OR start_week>?)"""
        params=[year,teacher,start_week,end_week]
        if duty_type:
            q+=" AND duty_type=?";params.append(duty_type)
        q+=" ORDER BY start_week,end_week"
        with self.lock:
            rows=[dict(r) for r in self.cx.execute(q,params)]
        if exclude_id:
            rows=[r for r in rows if int(r["id"])!=int(exclude_id)]
        return rows

    def save_duty(self,d):
        year=d.get("year") or self.active_year()
        teacher=(d.get("teacher") or "").strip()
        duty_type=(d.get("duty_type") or "").strip()
        duty_name=(d.get("duty_name") or "").strip()
        class_name=(d.get("class_name") or "").strip().upper()
        credit=num(d.get("credit"))
        standard_override=d.get("standard_override")
        standard_override=None if standard_override in (None,"","0",0) else num(standard_override)
        sw=int(d.get("start_week") or 1)
        ew=int(d.get("end_week") or self.setting("weeks",35))
        rid=d.get("id")
        if not teacher:raise ValueError("Chưa chọn giáo viên.")
        if duty_type not in ("Chủ nhiệm","Kiêm nhiệm","Chức vụ"):
            raise ValueError("Loại nhiệm vụ phải là Chức vụ, Kiêm nhiệm hoặc Chủ nhiệm.")
        if duty_type=="Chủ nhiệm":
            if not class_name:raise ValueError("Chưa chọn lớp chủ nhiệm.")
            duty_name="Chủ nhiệm "+class_name
            standard_override=None
        if not duty_name:raise ValueError("Chưa nhập nội dung nhiệm vụ.")
        if credit<0:raise ValueError("Tiết quy đổi không được âm.")
        if standard_override is not None and standard_override<=0:
            raise ValueError("Chuẩn áp dụng phải lớn hơn 0.")
        if not (1<=sw<=ew<=int(self.setting("weeks",35))):
            raise ValueError("Khoảng tuần không hợp lệ.")

        conflicts=self.duty_conflicts(teacher,sw,ew,year,rid,duty_type)
        if conflicts and not d.get("allow_overlap"):
            desc="; ".join(
                f'{x["duty_name"]} (T{x["start_week"]}-{x["end_week"]}, +{self._fmt_hours_vi(x["credit"])})'
                for x in conflicts
            )
            raise ValueError("OVERLAP|"+desc)

        with self.lock:
            vals=(teacher,duty_type,duty_name,class_name,credit,standard_override,sw,ew,d.get("note",""))
            if rid:
                self.cx.execute("""UPDATE duties SET
                    teacher=?,duty_type=?,duty_name=?,class_name=?,credit=?,standard_override=?,
                    start_week=?,end_week=?,note=? WHERE id=?""",vals+(rid,))
            else:
                self.cx.execute("""INSERT INTO duties(
                    year,teacher,duty_type,duty_name,class_name,credit,standard_override,
                    start_week,end_week,note
                ) VALUES(?,?,?,?,?,?,?,?,?,?)""",(year,)+vals)

            # Lớp học hiển thị GVCN mới nhất theo tuần bắt đầu.
            if duty_type=="Chủ nhiệm" and class_name:
                latest=self.cx.execute("""SELECT teacher FROM duties
                    WHERE year=? AND duty_type='Chủ nhiệm' AND class_name=?
                    ORDER BY start_week DESC,id DESC LIMIT 1""",(year,class_name)).fetchone()
                if latest:
                    self.cx.execute("""UPDATE classes SET homeroom_teacher=?
                        WHERE year=? AND class_name=?""",(latest["teacher"],year,class_name))
            self.cx.commit()

    def delete_duty(self,rid):
        with self.lock:
            row=self.cx.execute("SELECT * FROM duties WHERE id=?",(rid,)).fetchone()
            if not row:return
            year=row["year"];class_name=row["class_name"];duty_type=row["duty_type"]
            self.cx.execute("DELETE FROM duties WHERE id=?",(rid,))
            if duty_type=="Chủ nhiệm" and class_name:
                latest=self.cx.execute("""SELECT teacher FROM duties
                    WHERE year=? AND duty_type='Chủ nhiệm' AND class_name=?
                    ORDER BY start_week DESC,id DESC LIMIT 1""",(year,class_name)).fetchone()
                self.cx.execute("""UPDATE classes SET homeroom_teacher=?
                    WHERE year=? AND class_name=?""",
                    ((latest["teacher"] if latest else ""),year,class_name))
            self.cx.commit()

    def review_jobs(self,year=None):
        year=year or self.active_year()
        with self.lock:
            rows=[dict(r) for r in self.cx.execute(
                "SELECT * FROM review_jobs WHERE year=? ORDER BY teacher,id",(year,))]
        for r in rows:
            try:r["classes"]=json.loads(r.get("classes_json") or "[]")
            except:r["classes"]=[]
            r["classes_label"]=", ".join(r["classes"])
        return rows

    def preview_review(self,teacher,periods_per_week,start_week=None,end_week=None,rid=None,view_week=1):
        year=self.active_year()
        periods=num(periods_per_week)
        if periods<=0:raise ValueError("Số tiết ôn phải lớn hơn 0.")
        weeks=int(self.setting("weeks",35))
        sw=None if start_week in(None,"","0",0) else int(start_week)
        ew=int(end_week or weeks)
        if sw is not None and not(1<=sw<=ew<=weeks):
            raise ValueError("Tuần bắt đầu/kết thúc không hợp lệ.")
        exclude={int(rid)} if rid not in(None,"","0",0) else set()
        current=self.teacher_load_at_week(teacher,int(view_week),year,exclude_review_ids=exclude)
        if current is None:raise ValueError("Không tìm thấy giáo viên.")
        # Nếu chưa xác định tuần bắt đầu: chỉ tính phương án giả định bắt đầu ở tuần đang xem.
        hypothetical_start=int(view_week) if sw is None else sw
        at_start=self.teacher_load_at_week(
            teacher,hypothetical_start,year,exclude_review_ids=exclude,
            add_review_periods=periods,add_review_start=hypothetical_start,add_review_end=ew
        )
        max_after=self.teacher_max_load(
            teacher,year,exclude_review_ids=exclude,
            add_review_periods=periods,add_review_start=hypothetical_start,add_review_end=ew
        )
        return {
            "teacher":teacher,"periods_per_week":periods,"start_week":sw,"end_week":ew,
            "view_week":int(view_week),"current":current,"after_start":at_start,"max_after":max_after,
            "hypothetical":sw is None
        }

    def save_review(self,d):
        year=d.get("year") or self.active_year()
        teacher=(d.get("teacher") or "").strip()
        if not teacher:raise ValueError("Chưa chọn giáo viên.")
        classes=d.get("classes") or []
        if isinstance(classes,str):
            classes=[x.strip() for x in classes.split(",") if x.strip()]
        if not classes:raise ValueError("Hãy chọn ít nhất một lớp ôn.")
        periods=num(d.get("periods_per_week"),3)
        if periods<=0:raise ValueError("Số tiết/tuần phải lớn hơn 0.")
        sw=d.get("start_week");sw=None if sw in(None,"","0",0) else int(sw)
        ew=int(d.get("end_week") or self.setting("weeks",35))
        if sw is not None and not(1<=sw<=ew<=int(self.setting("weeks",35))):
            raise ValueError("Tuần không hợp lệ.")
        subject=(d.get("subject") or "").strip()
        if not subject:
            with self.lock:
                t=self.cx.execute("SELECT primary_subject FROM teachers WHERE year=? AND name=?",
                                  (year,teacher)).fetchone()
            subject=t["primary_subject"] if t else ""

        # Chỉ cho phép chọn lớp 12 mà chính giáo viên đang được phân công dạy môn ôn đó.
        with self.lock:
            eligible_rows=self.cx.execute("""SELECT DISTINCT class_name FROM requirements
                WHERE year=? AND teacher=? AND subject=? AND class_name LIKE '12%'""",
                (year,teacher,subject)).fetchall()
        eligible={r["class_name"] for r in eligible_rows}
        invalid=[c for c in classes if c not in eligible]
        if invalid:
            raise ValueError("Có lớp không thuộc phân công hiện tại của giáo viên: "+", ".join(invalid))

        vals=(teacher,subject,json.dumps(classes,ensure_ascii=False),periods,sw,ew,d.get("note",""))
        with self.lock:
            rid=d.get("id")
            if rid:self.cx.execute("""UPDATE review_jobs SET
                teacher=?,subject=?,classes_json=?,periods_per_week=?,start_week=?,end_week=?,note=?
                WHERE id=?""",vals+(rid,))
            else:self.cx.execute("""INSERT INTO review_jobs(
                year,teacher,subject,classes_json,periods_per_week,start_week,end_week,note
            ) VALUES(?,?,?,?,?,?,?,?)""",(year,)+vals)
            self.cx.commit()

    def delete_review(self,rid):
        with self.lock:self.cx.execute("DELETE FROM review_jobs WHERE id=?",(rid,));self.cx.commit()

    def staffing_summary(self,year=None):
        year=year or self.active_year()
        weeks=int(self.setting("weeks",35));classes=self.classes(year,True);teachers=self.teachers(year,True)
        by={}
        for r in self.requirements(year):
            d=by.setdefault(r["subject"],{"annual_hours":0,"weekly_need":0,"fte17":0,"teachers":0})
            d["annual_hours"]+=num(r["annual_hours"])
        std=num(self.setting("default_teacher_standard",17))
        for subj,d in by.items():
            d["weekly_need"]=d["annual_hours"]/(weeks or 35)
            d["fte17"]=d["weekly_need"]/(std or 17)
            d["teachers"]=sum(1 for t in teachers if subj in t["can_teach_list"] or subj==t["primary_subject"])
        return {"class_count":len(classes),"legal_teacher_equiv":len(classes)*num(self.setting("teacher_per_class",2.25)),
                "active_teacher_count":len(teachers),"by_subject":by}

STORE=Store()

# ----------------- XLSX export (không cần thư viện ngoài) -----------------
def col_letter(n):
    s=""
    while n:
        n,r=divmod(n-1,26);s=chr(65+r)+s
    return s

def xml_cell(ref,v,style=0):
    if v is None:return f'<c r="{ref}" s="{style}"/>'
    if isinstance(v,(int,float)) and not isinstance(v,bool):
        return f'<c r="{ref}" s="{style}"><v>{v}</v></c>'
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{escape(str(v))}</t></is></c>'

def sheet_xml(rows,widths):
    p=['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
       '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
       '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>']
    for i,w in enumerate(widths,1):p.append(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>')
    p.append('</cols><sheetData>')
    for ri,row in enumerate(rows,1):
        p.append(f'<row r="{ri}">')
        for ci,v in enumerate(row,1):p.append(xml_cell(f"{col_letter(ci)}{ri}",v,1 if ri==1 else 0))
        p.append('</row>')
    p.append('</sheetData></worksheet>')
    return ''.join(p)


def teaching_classes_text(teacher_name, reqs):
    """
    Gộp các phân công của một GV theo dạng:
    10A CĐ (Hóa), 10B (Hóa), 10C (HĐTN)
    Nếu GV có phần Chuyên đề của lớp/môn đó thì thêm 'CĐ'.
    """
    subject_label = {
        "Toán học": "Toán",
        "Ngữ văn": "Văn",
        "Vật lý": "Lý",
        "Hóa học": "Hóa",
        "Sinh học": "Sinh",
        "Tin học": "Tin",
        "Lịch sử": "Sử",
        "Địa lí": "Địa",
        "GDKTPL": "KTPL",
        "TNHN": "HĐTN",
        "GD địa phương": "GDĐP",
        "Ngoại ngữ": "Ngoại ngữ",
    }

    grouped = {}
    for r in reqs:
        if (r.get("teacher") or "") != teacher_name:
            continue
        key = (r["class_name"], r["subject"])
        g = grouped.setdefault(key, {"has_specialty": False})
        if r["req_type"] == "Chuyên đề":
            g["has_specialty"] = True

    def class_sort_key(item):
        cname, subject = item[0]
        grade = int(cname[:2]) if cname[:2].isdigit() else 99
        letter = cname[2:] if len(cname) > 2 else ""
        order = {
            "Ngữ văn": 1, "Toán học": 2, "Lịch sử": 3, "Địa lí": 4,
            "GDKTPL": 5, "Vật lý": 6, "Hóa học": 7, "Sinh học": 8,
            "Tin học": 9, "TNHN": 10, "GD địa phương": 11, "Ngoại ngữ": 12,
        }.get(subject, 99)
        return (grade, letter, order, subject)

    parts = []
    for (cname, subject), meta in sorted(grouped.items(), key=class_sort_key):
        label = subject_label.get(subject, subject)
        cd = " CĐ" if meta["has_specialty"] else ""
        parts.append(f"{cname}{cd} ({label})")
    return ", ".join(parts)

def resolve_export_range(scope="all",start_week=None,end_week=None):
    weeks=int(STORE.setting("weeks",35));s1=int(STORE.setting("semester1_weeks",18))
    scope=(scope or "all").lower()
    if scope=="hk1":return 1,min(s1,weeks),"Học kỳ I"
    if scope=="hk2":return min(s1+1,weeks),weeks,"Học kỳ II"
    if scope=="custom":
        a=max(1,min(weeks,int(start_week or 1)));b=max(1,min(weeks,int(end_week or weeks)))
        if a>b:a,b=b,a
        return a,b,f"Tuần {a}-{b}"
    return 1,weeks,"Cả năm"

def _range_overlap(start,end,a,b):
    s=1 if start in (None,"") else int(start);e=int(STORE.setting("weeks",35)) if end in (None,"") else int(end)
    return not (e<a or s>b)

def make_xlsx(year, week=1, scope="all", start_week=None, end_week=None):
    weeks=int(STORE.setting("weeks",35));range_start,range_end,range_label=resolve_export_range(scope,start_week,end_week)
    # Khi xuất theo HK/khoảng tuần, tuần tham chiếu phải nằm trong chính phạm vi đó.
    # HKII vì vậy không còn lấy tải tuần 1/HKI nếu ô "Tuần xem tải" vẫn đang là 1.
    requested_week=max(1,min(weeks,int(week or 1)))
    reference_week=requested_week if range_start<=requested_week<=range_end else range_start
    week=reference_week
    staff=STORE.teachers(year);classes=STORE.classes(year);reqs=STORE.requirements(year)
    reviews=STORE.review_jobs(year);rules=STORE.program_rules();summary=STORE.staffing_summary(year)

    overview=[["Năm học",year],["Phạm vi xuất",range_label],["Từ tuần",range_start],["Đến tuần",range_end],
              ["Số lớp hoạt động",summary["class_count"]],
              ["Định mức 2,25 GV/lớp",summary["legal_teacher_equiv"]],
              ["GV hoạt động",summary["active_teacher_count"]],["Số tuần năm học",weeks],
              ["Tuần tham chiếu trong phạm vi",week]]

    staff_rows=[["STT","Họ tên","Chuyên môn chính","Có thể dạy","Các lớp dạy",f"Chuẩn T{week}",f"Tổng tiết T{week}",f"Dư giờ T{week}",f"Cao nhất {range_label}","Tuần cao nhất","Loại GV","Hiệu lực từ","Đến","Chức vụ/KN","Tiết CV/KN","Chủ nhiệm","Tiết CN","Trạng thái","Ghi chú"]]
    for t in staff:
        load = STORE.teacher_load_at_week(t["name"], week, year)
        period_summary = STORE.teacher_period_summary(t["name"],range_start,range_end,year)
        max_load = period_summary.get("max") or STORE.teacher_max_load(t["name"], year)
        total_periods = load["total"] if load else 0
        current_standard = load["standard"] if load else num(t["standard"])
        excess_hours = load["diff"] if load else ""
        max_total = max_load["total"] if max_load else 0
        max_week = max_load["week"] if max_load else ""
        active_duties=STORE.duties_at_week(t["name"],week,year)
        roles=[x for x in active_duties if x["duty_type"] in ("Kiêm nhiệm","Chức vụ")]
        homerooms=[x for x in active_duties if x["duty_type"]=="Chủ nhiệm"]
        role_text="; ".join(x["duty_name"] for x in roles)
        role_credit=sum(num(x["credit"]) for x in roles)
        homeroom_text="; ".join(x["class_name"] for x in homerooms)
        homeroom_credit=sum(num(x["credit"]) for x in homerooms)
        auto_note = STORE.workload_change_note_range(t["name"],range_start,range_end,year)
        manual_note = (t["note"] or "").strip()
        final_note = auto_note
        if manual_note:
            final_note = (auto_note + ". Ghi chú khác: " + manual_note) if auto_note else manual_note
        staff_rows.append([t["stt"],t["name"],t["primary_subject"],", ".join(t["can_teach_list"]),
                           teaching_classes_text(t["name"], reqs),
                           current_standard,total_periods,excess_hours,max_total,max_week,
                           t["employment_type"],t["active_start"],t["active_end"],
                           role_text,role_credit,homeroom_text,homeroom_credit,
                           "Hoạt động" if t["active"] else "Ngừng",final_note])

    period_rows=[["STT","Họ tên","Chuyên môn","Phạm vi","Số tuần","Tiết dạy","Ôn TN","Chủ nhiệm","Chức vụ/KN","Tổng tiết phạm vi","Chuẩn phạm vi","Dư/thiếu phạm vi","BQ tiết/tuần","Cao nhất tuần","Tuần cao nhất","Ghi chú biến động"]]
    for t in staff:
        ps=STORE.teacher_period_summary(t["name"],range_start,range_end,year)
        mx=ps.get("max") or {}
        period_rows.append([t["stt"],t["name"],t["primary_subject"],range_label,ps["weeks"],ps["teaching"],ps["review"],ps["homeroom"],ps["role"],
                            ps["total"],ps["standard"],ps["diff"],ps["average"],mx.get("total",0),mx.get("week",""),ps.get("note","")])

    class_rows=[["Lớp","Khối","Tổ hợp","Cấu trúc môn","GVCN","Điều chỉnh riêng","Trạng thái"]]
    for c in classes:
        structure=", ".join(
            (x["subject"]+(" CĐ" if x.get("specialty") else ""))
            for x in c.get("structure",[])
        )
        class_rows.append([c["class_name"],c["grade"],c.get("template_name",""),structure,
                           c["homeroom_teacher"],"Có" if c.get("customized") else "",
                           "Hoạt động" if c["active"] else "Ngừng"])
    template_rows=[["Tổ hợp","Khối","Môn cốt lõi","Chuyên đề","Số lớp dùng","Trạng thái","Ghi chú"]]
    for tp in STORE.templates(year):
        core=", ".join(x["subject"] for x in tp["items"] if x["core"])
        cd=", ".join(x["subject"] for x in tp["items"] if x["specialty"])
        template_rows.append([tp["name"],tp["grade"],core,cd,tp["class_count"],
                              "Hoạt động" if tp["active"] else "Ngừng",tp["note"]])

    need_rows=[["Lớp","Môn","Loại","Tiết/năm","Phân bổ thực dạy","Tiết tuần xuất","Giáo viên","Tuần đầu","Tuần cuối","Ghi chú"]]
    for r in reqs:
        need_rows.append([r["class_name"],r["subject"],r["req_type"],r["annual_hours"],
                          r.get("schedule_summary",""),STORE.requirement_weekly_hours(r,week),
                          r["teacher"] or "CHƯA PHÂN CÔNG",
                          r["start_week"],r["end_week"],r["note"]])

    assign_rows=[["Giáo viên","Môn","Lớp","Loại","Tiết/năm","Phân bổ thực dạy","Tiết tuần xuất"]]
    for r in reqs:
        if r["teacher"]:
            assign_rows.append([r["teacher"],r["subject"],r["class_name"],r["req_type"],
                                r["annual_hours"],r.get("schedule_summary",""),
                                STORE.requirement_weekly_hours(r,week)])

    review_rows=[["Giáo viên","Môn ôn","Lớp ôn","Tiết/tuần","Tuần bắt đầu","Tuần kết thúc","Tải dự kiến","Dư giờ dự kiến","Ghi chú"]]
    for r in reviews:
        if r["start_week"] is not None and not _range_overlap(r.get("start_week"),r.get("end_week"),range_start,range_end):continue
        pv=STORE.preview_review(r["teacher"],r["periods_per_week"],r["start_week"],r["end_week"],r["id"],week)
        after=pv["after_start"]
        review_rows.append([r["teacher"],r["subject"],r["classes_label"],r["periods_per_week"],
                            r["start_week"] if r["start_week"] is not None else "",
                            r["end_week"],after["total"],after["diff"],r["note"]])

    rule_rows=[["Môn","Cốt lõi tiết/năm","Chuyên đề tiết/năm","Tổng nếu có CĐ","Phân bổ cốt lõi","Phân bổ CĐ","Cho phép CĐ","Nguồn"]]
    for r in rules:
        rule_rows.append([r["subject"],r["base_hours"],r["specialty_hours"],r["total_with_specialty"],
                          r["core_schedule_summary"],r["specialty_schedule_summary"],
                          "Có" if r["can_specialty"] else "Không",r["source"]])

    export_weeks=list(range(range_start,range_end+1))
    weekly_rows=[["Môn","Loại","Tiết/năm"]+[f"Tuần {w}" for w in export_weeks]+["Tổng phạm vi"]]
    for r in rules:
        core_week=[STORE._schedule_hours_at_week(r["core_schedule"],w) for w in export_weeks]
        weekly_rows.append([r["subject"],"Cốt lõi",r["base_hours"]]+core_week+[sum(core_week)])
        if r["can_specialty"]:
            cd_week=[STORE._schedule_hours_at_week(r["specialty_schedule"],w) for w in export_weeks]
            weekly_rows.append([r["subject"],"Chuyên đề",r["specialty_hours"]]+cd_week+[sum(cd_week)])

    staffing_rows=[["Môn","Tổng tiết/năm","Nhu cầu BQ/tuần","GV quy đổi chuẩn 17","GV có thể dạy"]]
    for subj,d in sorted(summary["by_subject"].items()):
        staffing_rows.append([subj,d["annual_hours"],d["weekly_need"],d["fte17"],d["teachers"]])

    duty_rows=[["Giáo viên","Loại","Nội dung","Lớp","Tiết quy đổi","Chuẩn áp dụng","Từ tuần","Đến tuần","Ghi chú"]]
    for d in STORE.duties(year):
        if not _range_overlap(d.get("start_week"),d.get("end_week"),range_start,range_end):continue
        duty_rows.append([d["teacher"],d["duty_type"],d["duty_name"],d["class_name"],
                          d["credit"],d["standard_override"] if d["standard_override"] is not None else "",
                          d["start_week"],d["end_week"],d["note"]])

    event_rows=[["Giáo viên","Loại biến động","Từ tuần","Đến tuần","Chặn dạy","GV thay thế","Ghi chú"]]
    for e in STORE.staff_events(year=year):
        if not _range_overlap(e.get("start_week"),e.get("end_week"),range_start,range_end):continue
        event_rows.append([e["teacher"],e["event_type"],e["start_week"],e["end_week"],
                           "Có" if e["blocks_teaching"] else "Không",e["replacement_teacher"],e["note"]])
    tkb_rows=[["Lớp","Giáo viên","Môn","Buổi","Thứ","Tiết","Từ tuần","Đến tuần","Nguồn","Ghi chú"]]
    for e in STORE.timetable_entries(year):
        if not _range_overlap(e.get("start_week"),e.get("end_week"),range_start,range_end):continue
        tkb_rows.append([e["class_name"],e["teacher"],e["subject"],e.get("session",""),e["day"],e["period"],
                         e["start_week"],e["end_week"],e["source"],e["note"]])
    scenario=STORE.active_scenario_row() or {}
    workflow_rows=[["Phương án","Trạng thái","Người duyệt","Thời gian duyệt","Người thao tác"]]
    workflow_rows.append([scenario.get("name",""),scenario.get("workflow_status",""),
                          scenario.get("approved_by",""),scenario.get("approved_at",""),
                          STORE.setting("operator_name","")])

    audit_rows=[["Thời gian","Người thao tác","Thao tác","Đã hoàn tác"]]
    for a in reversed(STORE.history(300)):
        audit_rows.append([a.get("created_at",""),a.get("actor",""),a.get("action",""),
                           "Có" if a.get("undone") else "Không"])
    workflow_history_rows=[["Thời gian","Người thực hiện","Trạng thái cũ","Trạng thái mới","Ghi chú"]]
    for w in reversed(STORE.workflow_history()):
        workflow_history_rows.append([w.get("created_at",""),w.get("actor",""),w.get("old_status",""),
                                      w.get("new_status",""),w.get("note","")])

    sheets=[
        ("Tong quan",overview,[26,18]),
        ("Doi ngu",staff_rows,[6,24,16,28,50,9,11,11,13,11,14,10,8,26,10,16,10,12,62]),
        ("Tong hop pham vi",period_rows,[6,24,16,18,10,12,10,12,14,18,16,18,14,14,12,62]),
        ("Lich su CN-KN",duty_rows,[24,14,28,10,12,14,10,10,32]),
        ("Bien dong nhan su",event_rows,[24,20,10,10,10,24,36]),
        ("Thoi khoa bieu",tkb_rows,[10,24,16,10,12,8,10,10,20,30]),
        ("Phe duyet",workflow_rows,[28,20,24,22,24]),
        ("Lich su thay doi",audit_rows,[22,24,60,14]),
        ("Lich su phe duyet",workflow_history_rows,[22,24,18,18,48]),
        ("To hop chuong trinh",template_rows,[22,8,55,42,12,12,28]),
        ("Lop hoc",class_rows,[10,8,22,65,24,14,12]),
        ("Nhu cau",need_rows,[10,15,12,12,30,14,24,10,10,30]),
        ("Phan cong",assign_rows,[24,15,10,12,12,30,14]),
        ("On tot nghiep",review_rows,[24,16,28,10,12,12,14,14,32]),
        ("Chuong trinh",rule_rows,[18,18,20,18,30,30,14,26]),
        ("Phan bo tuan",weekly_rows,[18,12,12]+[9]*len(export_weeks)+[14]),
        ("Dinh muc nhan su",staffing_rows,[18,16,22,22,16]),
    ]

    out=io.BytesIO()
    with zipfile.ZipFile(out,"w",zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'+''.join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(1,len(sheets)+1))+'</Types>')
        z.writestr("_rels/.rels",'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        z.writestr("xl/workbook.xml",'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'+''.join(f'<sheet name="{escape(n)}" sheetId="{i}" r:id="rId{i}"/>' for i,(n,_,_) in enumerate(sheets,1))+'</sheets></workbook>')
        z.writestr("xl/_rels/workbook.xml.rels",'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+''.join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1,len(sheets)+1))+f'<Relationship Id="rId{len(sheets)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
        z.writestr("xl/styles.xml",'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFill="1" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
        for i,(_,rows,widths) in enumerate(sheets,1):
            z.writestr(f"xl/worksheets/sheet{i}.xml",sheet_xml(rows,widths))
    return out.getvalue()



def make_report_html(week=1,scope="all",start_week=None,end_week=None):
    year=STORE.active_year();m=STORE.plan_metrics();v=STORE.validate_all()
    a,b,range_label=resolve_export_range(scope,start_week,end_week)
    scenario=STORE.active_scenario_row() or {}
    rows=[]
    for t in STORE.teachers(active_only=True):
        ps=STORE.teacher_period_summary(t["name"],a,b);mx=ps.get("max") or {}
        rows.append(f"<tr><td>{escape(t['name'])}</td><td>{escape(t['primary_subject'])}</td>"
                    f"<td>{STORE._fmt_hours_vi(ps['standard'])}</td><td>{STORE._fmt_hours_vi(ps['total'])}</td>"
                    f"<td>{STORE._fmt_hours_vi(ps['diff'])}</td><td>T{mx.get('week','')} · {STORE._fmt_hours_vi(mx.get('total',0))}</td></tr>")
    status={"DRAFT":"Đang soạn","REVIEWED":"Tổ trưởng đã kiểm tra","APPROVED":"Ban Giám đốc đã duyệt"}.get(scenario.get("workflow_status"),scenario.get("workflow_status",""))
    html=f'''<!doctype html><html><head><meta charset="utf-8"><title>Báo cáo phân công {escape(year)}</title>
    <style>body{{font-family:"Times New Roman";font-size:14px;margin:28mm 22mm}}h1,h2{{text-align:center}}table{{width:100%;border-collapse:collapse;margin-top:10px}}th,td{{border:1px solid #333;padding:5px}}th{{background:#eee}}.meta{{margin:10px 0}}</style></head>
    <body><h1>BÁO CÁO PHÂN CÔNG CHUYÊN MÔN</h1><h2>Năm học {escape(year)}</h2>
    <div class="meta">Phương án: <b>{escape(scenario.get('name','BẢN CHÍNH'))}</b> · Trạng thái: <b>{escape(status)}</b> · Phạm vi: <b>{escape(range_label)}</b> (T{a}-T{b})</div>
    <p>Hoàn thành phân công: <b>{m['coverage_percent']:.1f}%</b>; chưa gán: <b>{m['unassigned']}</b>; giáo viên vượt chuẩn: <b>{m['over_count']}</b>; lỗi kiểm tra: <b>{v['counts']['error']}</b>; cảnh báo: <b>{v['counts']['warning']}</b>.</p>
    <table><thead><tr><th>Giáo viên</th><th>Chuyên môn</th><th>Chuẩn phạm vi</th><th>Tổng tiết phạm vi</th><th>Dư/thiếu</th><th>Cao nhất phạm vi</th></tr></thead>
    <tbody>{''.join(rows)}</tbody></table>
    <p style="margin-top:45px;text-align:right">Người lập báo cáo: {escape(str(STORE.setting('operator_name','') or ''))}</p>
    </body></html>'''
    return html.encode("utf-8")

def make_report_docx(week=1,scope="all",start_week=None,end_week=None):
    year=STORE.active_year();scenario=STORE.active_scenario_row() or {};m=STORE.plan_metrics();v=STORE.validate_all();a,b,range_label=resolve_export_range(scope,start_week,end_week)
    week=int(week or 1);week=week if a<=week<=b else a
    def tx(s):return escape(str(s or ""))
    def para(text,bold=False,size=22,align="left"):
        return f'<w:p><w:pPr><w:jc w:val="{align}"/></w:pPr><w:r><w:rPr>{"<w:b/>" if bold else ""}<w:sz w:val="{size}"/></w:rPr><w:t xml:space="preserve">{tx(text)}</w:t></w:r></w:p>'
    body=para("BÁO CÁO PHÂN CÔNG CHUYÊN MÔN",True,28,"center")
    body+=para(f"Năm học {year}",True,24,"center")
    status={"DRAFT":"Đang soạn","REVIEWED":"Tổ trưởng đã kiểm tra","APPROVED":"Ban Giám đốc đã duyệt"}.get(scenario.get("workflow_status"),scenario.get("workflow_status",""))
    body+=para(f"Phương án: {scenario.get('name','BẢN CHÍNH')} · Trạng thái: {status} · Phạm vi: {range_label} (T{a}-T{b})")
    if scenario.get("approved_by"):
        body+=para(f"Người duyệt: {scenario.get('approved_by')} · Thời gian: {scenario.get('approved_at','')}")
    body+=para(f"Hoàn thành: {m['coverage_percent']:.1f}% · Chưa gán: {m['unassigned']} · GV vượt chuẩn: {m['over_count']} · Lỗi: {v['counts']['error']} · Cảnh báo: {v['counts']['warning']}")
    headers=["Giáo viên","Chuyên môn","Chuẩn phạm vi","Tổng phạm vi","Dư/thiếu","Cao nhất"]
    table_rows=[headers]
    for t in STORE.teachers(active_only=True):
        ps=STORE.teacher_period_summary(t["name"],a,b);mx=ps.get("max") or {}
        table_rows.append([t["name"],t["primary_subject"],STORE._fmt_hours_vi(ps["standard"]),
                           STORE._fmt_hours_vi(ps["total"]),STORE._fmt_hours_vi(ps["diff"]),
                           f'T{mx.get("week","")}: {STORE._fmt_hours_vi(mx.get("total",0))}'])
    def cell(text,bold=False):
        return f'<w:tc><w:tcPr><w:tcW w:w="1800" w:type="dxa"/></w:tcPr><w:p><w:r><w:rPr>{"<w:b/>" if bold else ""}<w:sz w:val="18"/></w:rPr><w:t>{tx(text)}</w:t></w:r></w:p></w:tc>'
    body+='<w:tbl><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>'
    for i,row in enumerate(table_rows):
        body+='<w:tr>'+''.join(cell(x,i==0) for x in row)+'</w:tr>'
    body+='</w:tbl>'+para(f"Người lập báo cáo: {STORE.setting('operator_name','')}",False,22,"right")
    doc='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'+body+'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/></w:sectPr></w:body></w:document>'
    out=io.BytesIO()
    with zipfile.ZipFile(out,"w",zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",'<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        z.writestr("_rels/.rels",'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr("word/document.xml",doc)
    return out.getvalue()

def dashboard_data(week=1):
    year=STORE.active_year()
    teachers=STORE.teachers(year,active_only=True)
    classes=STORE.classes(year,active_only=True)
    reqs=STORE.requirements(year)
    loads=STORE.teacher_loads_summary(week)
    staffing=STORE.staffing_summary(year)
    matrix=STORE.workload_matrix(year)

    # Tình trạng tải giáo viên ở tuần đang xem.
    status={"under":0,"equal":0,"over":0}
    teacher_rows=[]
    for t in teachers:
        cur=loads.get(t["name"],{}).get("current") or {}
        total=num(cur.get("total"))
        standard=num(cur.get("standard"),num(t.get("standard"),17))
        diff=total-standard
        if abs(diff)<1e-6: status["equal"]+=1
        elif diff>0: status["over"]+=1
        else: status["under"]+=1
        teacher_rows.append({
            "teacher":t["name"],"subject":t.get("primary_subject",""),
            "total":total,"standard":standard,"diff":diff
        })

    # Mức hoàn thành phân công theo khối.
    grade_rows=[]
    for grade in (10,11,12):
        class_names={c["class_name"] for c in classes if int(c["grade"])==grade}
        grade_reqs=[r for r in reqs if r["class_name"] in class_names]
        assigned=sum(1 for r in grade_reqs if (r.get("teacher") or "").strip())
        total=len(grade_reqs)
        grade_rows.append({
            "grade":str(grade),"assigned":assigned,"unassigned":total-assigned,
            "percent":(assigned*100/total if total else 0)
        })

    # Nhu cầu nhân sự theo môn.
    staffing_rows=[]
    for subject,d in sorted(staffing["by_subject"].items()):
        staffing_rows.append({
            "subject":subject,
            "need":num(d.get("fte17")),
            "available":num(d.get("teachers")),
            "weekly_need":num(d.get("weekly_need"))
        })

    total_reqs=len(reqs)
    assigned_reqs=sum(1 for r in reqs if (r.get("teacher") or "").strip())

    return {
        "week":week,
        "kpis":{
            "teachers":len(teachers),
            "classes":len(classes),
            "assignment_percent":(assigned_reqs*100/total_reqs if total_reqs else 0),
            "over_count":status["over"],
            "unassigned":total_reqs-assigned_reqs
        },
        "teacher_loads":teacher_rows,
        "load_status":status,
        "grades":grade_rows,
        "staffing":staffing_rows,
        "subject_gaps":[
            {"subject":x["subject"],"gap":x["available"]-x["need"],
             "need":x["need"],"available":x["available"]}
            for x in staffing_rows
        ],
        "heatmap":[
            {"teacher":t["name"],"values":[
                {"week":w,"diff":num(matrix[t["name"]][w]["diff"]),
                 "total":num(matrix[t["name"]][w]["total"]),
                 "standard":num(matrix[t["name"]][w]["standard"])}
                for w in range(1,int(STORE.setting("weeks",35))+1)
            ]}
            for t in teachers
        ]
    }

def teacher_trend_data(teacher):
    year=STORE.active_year()
    rows=[]
    for w in range(1,int(STORE.setting("weeks",35))+1):
        d=STORE.teacher_load_at_week(teacher,w,year)
        if d:
            rows.append({
                "week":w,"total":num(d["total"]),"standard":num(d["standard"]),
                "diff":num(d["diff"]),"teaching":num(d["teaching"]),
                "review":num(d["review"]),"duties":num(d.get("duty_credit"))
            })
    return {"teacher":teacher,"rows":rows}


def mutation_label(path,d):
    labels={
        "/api/drop":"Phân công / đổi giáo viên",
        "/api/teacher/save":"Thêm / sửa giáo viên",
        "/api/teacher/toggle":"Đổi trạng thái giáo viên",
        "/api/template/save":"Sửa tổ hợp chương trình",
        "/api/template/toggle":"Đổi trạng thái tổ hợp",
        "/api/template/copy":"Sao chép tổ hợp",
        "/api/class/save":"Thêm / sửa lớp",
        "/api/class/toggle":"Đổi trạng thái lớp",
        "/api/class/structure":"Sửa cấu trúc lớp",
        "/api/class/restore_template":"Khôi phục lớp theo tổ hợp",
        "/api/program/save":"Sửa chương trình / phân bổ tuần",
        "/api/requirement/save":"Thêm nhu cầu môn",
        "/api/requirement/delete":"Xóa nhu cầu môn",
        "/api/duty/save":"Thêm / sửa Chức vụ-CN-KN",
        "/api/duty/delete":"Xóa Chức vụ-CN-KN",
        "/api/review/save":"Thêm / sửa ôn tốt nghiệp",
        "/api/review/delete":"Xóa ôn tốt nghiệp",
        "/api/lock/toggle":"Khóa / mở khóa phân công",
        "/api/import/apply":"Nhập dữ liệu đầu năm",
        "/api/staff_event/save":"Thêm / sửa biến động nhân sự",
        "/api/staff_event/delete":"Xóa biến động nhân sự",
        "/api/timetable/apply":"Nhập thời khóa biểu",
        "/api/timetable/save":"Thêm / sửa dòng thời khóa biểu",
        "/api/timetable/delete":"Xóa dòng thời khóa biểu",
        "/api/timetable/sync":"Đồng bộ TKB theo phân công chuyên môn"
    }
    base=labels.get(path,path)
    extra=""
    if path=="/api/drop":
        extra=f' {d.get("class_name","")} {d.get("subject","")} → {d.get("teacher","") or "Bỏ gán"}'
    elif path=="/api/duty/save":
        extra=f' · {d.get("teacher","")} · {d.get("duty_name") or d.get("class_name","")}'
    elif path=="/api/review/save":
        extra=f' · {d.get("teacher","")} · {d.get("subject","")}'
    return base+extra

class Handler(BaseHTTPRequestHandler):
    def log_message(self,*args):pass

    def send_json(self,d,status=200):
        b=json.dumps(d,ensure_ascii=False).encode("utf-8")
        self.send_response(status);self.send_header("Content-Type","application/json; charset=utf-8")
        self.send_header("Content-Length",str(len(b)));self.end_headers();self.wfile.write(b)

    def read_json(self):
        n=int(self.headers.get("Content-Length","0") or 0)
        return json.loads(self.rfile.read(n).decode("utf-8")) if n else {}

    def do_GET(self):
        try:
            u=urllib.parse.urlparse(self.path)
            if u.path=="/":
                with open(INDEX_PATH,"rb") as f:b=f.read()
                self.send_response(200);self.send_header("Content-Type","text/html; charset=utf-8")
                self.send_header("Content-Length",str(len(b)));self.end_headers();self.wfile.write(b);return
            if u.path=="/template_tkb.xlsx":
                with open(TKB_TEMPLATE_PATH,"rb") as f:b=f.read()
                self.send_response(200)
                self.send_header("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                self.send_header("Content-Disposition",'attachment; filename="MAU_NHAP_TKB_EXCEL_V10_3.xlsx"')
                self.send_header("Content-Length",str(len(b)));self.end_headers();self.wfile.write(b);return
            if u.path=="/api/state":
                q=urllib.parse.parse_qs(u.query)
                year=(q.get("year") or [STORE.active_year()])[0]
                if year!=STORE.active_year():STORE.set_active_year(year)
                week=max(1,min(int(STORE.setting("weeks",35)),int((q.get("week") or ["1"])[0])))
                self.send_json({
                    "active_year":STORE.active_year(),"years":STORE.years(),
                    "settings":{"weeks":STORE.setting("weeks",35),
                                "teacher_per_class":STORE.setting("teacher_per_class",2.25),
                                "default_teacher_standard":STORE.setting("default_teacher_standard",17),
                                "review_periods_per_week":STORE.setting("review_periods_per_week",3),
                                "program_source_title":STORE.setting("program_source_title","Chương trình GDTX cấp THPT"),
                                "program_source":STORE.setting("program_source","TT 12/2022/TT-BGDĐT"),
                                "semester1_weeks":STORE.setting("semester1_weeks",18),
                                "semester2_weeks":STORE.setting("semester2_weeks",17),
                                "operator_name":STORE.setting("operator_name",""),
                                "hdtn_tkb_slots":STORE.setting("hdtn_tkb_slots",2),
                                "sources":STORE.setting("sources",{})},
                    "teachers":STORE.teachers(),"classes":STORE.classes(),
                    "templates":STORE.templates(),
                    "program_rules":STORE.program_rules(),"groups":STORE.subject_groups(),
                    "requirements":STORE.requirements(),
                    "review_jobs":STORE.review_jobs(),
                    "review_previews":{str(r["id"]):STORE.preview_review(
                        r["teacher"],r["periods_per_week"],r["start_week"],r["end_week"],r["id"],week
                    ) for r in STORE.review_jobs()},
                    "duties":STORE.duties(),
                    "staff_events":STORE.staff_events(),
                    "timetable_entries":STORE.timetable_entries(),
                    "locks":STORE.locks(),
                    "active_scenario":STORE.active_scenario_row(),
                    "workload_notes":{t["name"]:STORE.workload_change_note(t["name"]) for t in STORE.teachers()},
                    "loads":STORE.teacher_loads_summary(week),"staffing":STORE.staffing_summary()
                });return
            if u.path=="/api/preview_drop":
                q=urllib.parse.parse_qs(u.query)
                self.send_json(STORE.preview_drop(
                    (q.get("class") or [""])[0],(q.get("subject") or [""])[0],
                    (q.get("teacher") or [""])[0],(q.get("req_type") or [""])[0] or None,
                    int((q.get("week") or ["1"])[0])
                ));return
            if u.path=="/api/preview_review":
                q=urllib.parse.parse_qs(u.query)
                self.send_json(STORE.preview_review(
                    (q.get("teacher") or [""])[0],
                    (q.get("periods") or ["3"])[0],
                    (q.get("start") or [""])[0] or None,
                    (q.get("end") or [str(STORE.setting("weeks",35))])[0],
                    (q.get("id") or [""])[0] or None,
                    int((q.get("week") or ["1"])[0])
                ));return
            if u.path=="/api/impact":
                q=urllib.parse.parse_qs(u.query)
                self.send_json(STORE.impact_analysis(
                    (q.get("class") or [""])[0],(q.get("subject") or [""])[0],
                    (q.get("teacher") or [""])[0],(q.get("req_type") or [""])[0] or None,
                    int((q.get("week") or ["1"])[0])
                ));return
            if u.path=="/api/staff_event/recommend":
                q=urllib.parse.parse_qs(u.query)
                self.send_json(STORE.event_replacement_plan(int((q.get("id") or ["0"])[0])));return
            if u.path=="/api/timetable/sync_preview":
                self.send_json(STORE.timetable_sync_preview());return
            if u.path=="/api/timetable/validation":
                self.send_json(STORE.validate_timetable());return
            if u.path=="/api/workflow/history":
                self.send_json({"rows":STORE.workflow_history()});return
            if u.path=="/api/validation":
                self.send_json(STORE.validate_all());return
            if u.path=="/api/history":
                self.send_json({"rows":STORE.history(int((urllib.parse.parse_qs(u.query).get("limit") or ["100"])[0]))});return
            if u.path=="/api/scenarios":
                self.send_json({"rows":STORE.scenarios(),"active_id":STORE.active_scenario_id()});return
            if u.path=="/api/recommend":
                q=urllib.parse.parse_qs(u.query)
                self.send_json({"rows":STORE.recommend_teachers(
                    (q.get("class") or [""])[0],(q.get("subject") or [""])[0],
                    (q.get("req_type") or [""])[0] or None,
                    int((q.get("week") or ["1"])[0]),int((q.get("limit") or ["10"])[0])
                )});return
            if u.path=="/backup.json":
                payload={"app":"PHAN_CONG_V10_3","database":STORE.full_backup()}
                b=json.dumps(payload,ensure_ascii=False,indent=2).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type","application/json; charset=utf-8")
                self.send_header("Content-Disposition",f'attachment; filename="sao_luu_phan_cong_{STORE.active_year()}.json"')
                self.send_header("Content-Length",str(len(b)));self.end_headers();self.wfile.write(b);return
            if u.path=="/api/dashboard":
                q=urllib.parse.parse_qs(u.query)
                week=max(1,min(int(STORE.setting("weeks",35)),int((q.get("week") or ["1"])[0])))
                self.send_json(dashboard_data(week));return
            if u.path=="/api/teacher_trend":
                q=urllib.parse.parse_qs(u.query)
                teacher=(q.get("teacher") or [""])[0]
                if not teacher:
                    self.send_json({"error":"Chưa chọn giáo viên"},400);return
                self.send_json(teacher_trend_data(teacher));return
            if u.path=="/report.html":
                q=urllib.parse.parse_qs(u.query);week=int((q.get("week") or ["1"])[0]);scope=(q.get("scope") or ["all"])[0]
                start=(q.get("start") or [None])[0];end=(q.get("end") or [None])[0]
                b=make_report_html(week,scope,start,end);self.send_response(200)
                self.send_header("Content-Type","text/html; charset=utf-8")
                self.send_header("Content-Length",str(len(b)));self.end_headers();self.wfile.write(b);return
            if u.path=="/report.docx":
                q=urllib.parse.parse_qs(u.query);week=int((q.get("week") or ["1"])[0]);scope=(q.get("scope") or ["all"])[0]
                start=(q.get("start") or [None])[0];end=(q.get("end") or [None])[0]
                b=make_report_docx(week,scope,start,end);self.send_response(200)
                a,bw,label=resolve_export_range(scope,start,end);slug={"hk1":"HKI","hk2":"HKII","all":"CA_NAM"}.get(scope,f"T{a}_{bw}")
                self.send_header("Content-Type","application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                self.send_header("Content-Disposition",f'attachment; filename="bao_cao_phan_cong_{STORE.active_year()}_{slug}.docx"')
                self.send_header("Content-Length",str(len(b)));self.end_headers();self.wfile.write(b);return
            if u.path=="/export.xlsx":
                q=urllib.parse.parse_qs(u.query)
                week=max(1,min(int(STORE.setting("weeks",35)),int((q.get("week") or ["1"])[0])))
                scope=(q.get("scope") or ["all"])[0];start=(q.get("start") or [None])[0];end=(q.get("end") or [None])[0]
                a,b,label=resolve_export_range(scope,start,end)
                data=make_xlsx(STORE.active_year(),week,scope,start,end)
                slug={"hk1":"HKI","hk2":"HKII","all":"CA_NAM"}.get(scope,f"T{a}_{b}")
                self.send_response(200)
                self.send_header("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                self.send_header("Content-Disposition",f'attachment; filename="phan_cong_V10_3_{STORE.active_year()}_{slug}.xlsx"')
                self.send_header("Content-Length",str(len(data)));self.end_headers();self.wfile.write(data);return
            self.send_error(404)
        except Exception as e:self.send_json({"error":str(e)},400)

    def do_POST(self):
        try:
            d=self.read_json();p=self.path

            # Các lệnh quản trị không ghi vào lịch sử thay đổi nghiệp vụ.
            if p=="/api/shutdown":
                self.send_json({"ok":True})
                threading.Thread(target=self.server.shutdown,daemon=True).start()
                return
            if p=="/api/undo":
                self.send_json({"ok":True,"undone":STORE.undo_last()});return
            if p=="/api/scenario/create":
                self.send_json({"ok":True,"scenario":STORE.create_scenario(d.get("name"),d.get("description",""))});return
            if p=="/api/scenario/switch":
                self.send_json({"ok":True,"scenario":STORE.switch_scenario(int(d["id"]))});return
            if p=="/api/scenario/apply":
                self.send_json({"ok":True,"scenario":STORE.apply_scenario_as_main(int(d["id"]))});return
            if p=="/api/scenario/delete":
                STORE.delete_scenario(int(d["id"]));self.send_json({"ok":True});return
            if p=="/api/optimizer/generate":
                self.send_json({"ok":True,"rows":STORE.generate_optimized_scenarios()});return
            if p=="/api/workflow/set":
                self.send_json({"ok":True,"workflow":STORE.set_workflow(d.get("status"),d.get("actor",""),d.get("note",""),d.get("id"))});return
            if p=="/api/assistant":
                self.send_json(STORE.assistant_query(d.get("text",""),int(d.get("week") or 1)));return
            if p=="/api/settings/operator":
                STORE.set_setting("operator_name",d.get("name",""));self.send_json({"ok":True});return
            if p=="/api/year/rollover":
                self.send_json({"ok":True,"result":STORE.rollover_year(d.get("name"),int(d.get("grade10_count") or 5))});return
            if p=="/api/import/preview":
                raw=base64.b64decode(d.get("data_base64",""))
                preview=preview_import_xlsx(raw)
                STORE.set_setting("last_import_preview",preview)
                self.send_json({"ok":True,"preview":preview});return
            if p=="/api/timetable/preview":
                raw=base64.b64decode(d.get("data_base64",""))
                filename=(d.get("filename") or "").lower()
                if filename.endswith(".pdf"):
                    preview=preview_timetable_pdf(raw,int(d.get("start_week") or 1),int(d.get("end_week") or STORE.setting("weeks",35)))
                else:
                    preview=preview_timetable_xlsx(raw)
                STORE.set_setting("last_timetable_preview",preview)
                self.send_json({"ok":True,"preview":preview});return
            if p=="/api/restore_backup":
                payload=d.get("payload") or {}
                if not isinstance(payload,dict):raise ValueError("File sao lưu không hợp lệ.")
                if payload.get("database"):
                    STORE.restore_full_backup(payload["database"])
                elif payload.get("snapshot"):
                    before=STORE.full_snapshot()
                    STORE.restore_snapshot(payload["snapshot"])
                    after=STORE.full_snapshot()
                    STORE.log_action("Phục hồi bản sao lưu cũ",before,after)
                    STORE.save_active_scenario()
                else:
                    raise ValueError("File sao lưu không hợp lệ.")
                self.send_json({"ok":True});return

            # Chuyển năm học không dùng snapshot/undo.
            if p=="/api/year/new":
                STORE.create_year(d.get("name"),d.get("copy_staff",True),d.get("copy_classes",True))
                self.send_json({"ok":True});return
            if p=="/api/year/active":
                STORE.set_active_year(d["name"]);self.send_json({"ok":True});return

            auditable={
                "/api/drop","/api/teacher/save","/api/teacher/toggle",
                "/api/template/save","/api/template/toggle","/api/template/copy",
                "/api/class/save","/api/class/toggle","/api/class/structure","/api/class/restore_template",
                "/api/program/save","/api/requirement/save","/api/requirement/delete",
                "/api/duty/save","/api/duty/delete","/api/review/save","/api/review/delete",
                "/api/lock/toggle","/api/import/apply",
                "/api/staff_event/save","/api/staff_event/delete",
                "/api/timetable/apply","/api/timetable/save","/api/timetable/delete",
                "/api/timetable/sync"
            }
            if p in auditable:
                STORE.ensure_baseline()
                STORE.assert_editable()
            before=STORE.full_snapshot() if p in auditable else None
            result={"ok":True}

            if p=="/api/drop":
                STORE.assign_drop(d["class_name"],d["subject"],d.get("teacher",""),d.get("req_type") or None)
            elif p=="/api/teacher/save":STORE.save_teacher(d)
            elif p=="/api/teacher/toggle":STORE.toggle_teacher(int(d["id"]))
            elif p=="/api/template/save":STORE.save_template(d)
            elif p=="/api/template/toggle":STORE.toggle_template(int(d["id"]))
            elif p=="/api/template/copy":result["template"]=STORE.copy_template(int(d["id"]),d.get("name"))
            elif p=="/api/class/save":STORE.save_class(d)
            elif p=="/api/class/toggle":STORE.toggle_class(int(d["id"]))
            elif p=="/api/class/structure":STORE.save_class_structure(d)
            elif p=="/api/class/restore_template":STORE.restore_class_template(d["class_name"])
            elif p=="/api/program/save":STORE.save_program_rule(d)
            elif p=="/api/requirement/save":STORE.save_requirement(d)
            elif p=="/api/requirement/delete":STORE.delete_requirement(int(d["id"]))
            elif p=="/api/duty/save":STORE.save_duty(d)
            elif p=="/api/duty/delete":STORE.delete_duty(int(d["id"]))
            elif p=="/api/review/save":STORE.save_review(d)
            elif p=="/api/review/delete":STORE.delete_review(int(d["id"]))
            elif p=="/api/lock/toggle":
                result.update(STORE.toggle_lock(d["scope_type"],d["scope_key"],d.get("note","")))
            elif p=="/api/import/apply":
                preview=STORE.setting("last_import_preview",None)
                if not preview:raise ValueError("Chưa có dữ liệu xem trước để nhập.")
                result["imported"]=STORE.apply_import_preview(
                    preview,bool(d.get("import_teachers",True)),bool(d.get("import_classes",True)))
            elif p=="/api/staff_event/save":STORE.save_staff_event(d)
            elif p=="/api/staff_event/delete":STORE.delete_staff_event(int(d["id"]))
            elif p=="/api/timetable/apply":
                preview=STORE.setting("last_timetable_preview",None)
                if not preview:raise ValueError("Chưa có dữ liệu TKB xem trước.")
                mode=(d.get("replace_mode") or ("all" if d.get("replace",True) else "append"))
                result["timetable"]=STORE.apply_timetable_preview(preview,bool(d.get("replace",True)),mode)
            elif p=="/api/timetable/save":STORE.save_timetable_entry(d)
            elif p=="/api/timetable/delete":STORE.delete_timetable_entry(int(d["id"]))
            elif p=="/api/timetable/sync":result["sync"]=STORE.sync_timetable_to_assignment()
            else:
                self.send_json({"error":"API không tồn tại"},404);return

            if before is not None:
                STORE.invalidate_caches()
                after=STORE.full_snapshot()
                STORE.log_action(mutation_label(p,d),before,after)
                STORE.save_active_scenario()

            self.send_json(result)
        except Exception as e:self.send_json({"error":str(e)},400)

def main():
    server=ThreadingHTTPServer((HOST,PORT),Handler);url=f"http://{HOST}:{PORT}"
    print("="*72);print("PHẦN MỀM PHÂN CÔNG V10.3 PERFORMANCE — TỐI ƯU TỐC ĐỘ");print("Mở:",url)
    print("Dữ liệu:",DB_PATH);print("Ctrl+C để dừng");print("="*72)
    if HOST in ("127.0.0.1", "localhost"):
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:server.serve_forever()
    except KeyboardInterrupt:pass
    finally:server.server_close()

if __name__=="__main__":main()
