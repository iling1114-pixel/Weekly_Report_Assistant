"""週報小幫手核心：解析工作紀錄、管理待辦、依範本產生 Word 週報。"""
import copy
import datetime as dt
import json
import re
import shutil
import sys
from pathlib import Path

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# 打包成 exe 時：內建檔案在 _MEIPASS，使用者資料放在 exe 旁邊；用原始碼跑時兩者都是本資料夾
if getattr(sys, "frozen", False):
    RES = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    BASE = Path(sys.executable).resolve().parent
else:
    RES = BASE = Path(__file__).resolve().parent
TEMPLATE = RES / "template.docx"
SETTINGS = BASE / "settings.json"
TODOS = BASE / "todos.json"
DRAFT = BASE / "工作紀錄草稿.txt"
OUT_DIR = BASE / "週報"

# 範本中當作「模子」的位置（列, 欄）
WORK_SAMPLE = (2, 1)   # 本週工作欄：List Paragraph + 編號
TODO_SAMPLE = (2, 2)   # 待辦欄：Normal + 編號
EMPTY_SAMPLE = (4, 1)  # 空白格

# 版本與更新：發新版時改 VERSION，再跑 release.bat
VERSION = "1.0.0"
REPO = "iling1114-pixel/Weekly_Report_Assistant"  # GitHub 帳號/repo，程式會到這裡查有沒有新版

# 專案、簡稱、收件人都是個人設定（settings.json），程式碼裡不放實際名稱
FIXED_PROJECTS = ["其他"]
ALIASES = {}

LV0 = re.compile(r"^\s*\d+\s*[.、．)]\s*")
LV1 = re.compile(r"^\s*[-－–—•*‧·]\s*")
IGNORE = {"工作事項", "待辦事項", "本週工作", "工作紀錄"}


# ---------- 解析 ----------
def _scan(text, known):
    """逐行分類 → [(行號, 種類, 內容)]，種類：project / module / item / child。
    沒編號的行：第一個、前一行空白、或對得到已知專案 → 專案；否則是專案底下的模組。"""
    out, seen, blank = [], False, True
    for i, raw in enumerate(text.splitlines()):
        line = raw.strip()
        if not line:
            blank = True
            continue
        if line in IGNORE:
            continue
        if LV0.match(line):
            out.append((i, "item", LV0.sub("", line)))
        elif LV1.match(line):
            out.append((i, "child", LV1.sub("", line)))
        elif not seen or blank or match_project(line, known):
            out.append((i, "project", line))
            seen = True
        else:
            out.append((i, "module", line))
        blank = False
    return out


def parse(text, known=None):
    """文字 → [(專案名, [{'text', 'children': [...]}])]，保留出現順序。模組名會併成「模組_項目」。"""
    projects, cur, mod = [], None, None
    for _, kind, s in _scan(text, known or FIXED_PROJECTS):
        if kind == "project":
            cur, mod = (s, []), None
            projects.append(cur)
            continue
        if kind == "module":
            mod = s
            continue
        if not s:  # 只打了「1.」還沒寫內容
            continue
        if cur is None:
            cur = ("其他", [])
            projects.append(cur)
        if kind == "child" and cur[1]:
            cur[1][-1]["children"].append(s)
        else:
            cur[1].append({"text": f"{mod}_{s}" if mod else s, "children": []})
    return projects


def split_log(text, known):
    """工作紀錄 → {專案正式名: 內容}（照出現順序，同專案多段接在一起；專案前的項目歸「其他」）。"""
    heads = {i: s for i, kind, s in _scan(text, known) if kind == "project"}
    out, cur = {}, "其他"
    for i, raw in enumerate(text.splitlines()):
        if i in heads:
            cur = match_project(heads[i], known) or heads[i]
            out.setdefault(cur, [])
        elif raw.strip() and raw.strip() not in IGNORE:
            out.setdefault(cur, []).append(raw.rstrip())
    return {k: "\n".join(v) for k, v in out.items()}


def join_log(fields):
    """[(專案, 內容)] → 工作紀錄文字。空的專案略過；內容裡的空行拿掉，不然下一行會被當成新專案。"""
    blocks = []
    for proj, content in fields:
        lines = [l.rstrip() for l in content.splitlines() if l.strip()]
        if lines:
            blocks.append("\n".join([proj] + lines))
    return "\n\n".join(blocks)


def _block(text, project, known):
    """找 project 最後一段的範圍 → (插入行號, 已有幾個 1. 項目)；模組前的位置才不會被加上模組前綴。"""
    pos, n, inside, in_mod = None, 0, False, False
    for i, kind, s in _scan(text, known):
        if kind == "project":
            inside = (match_project(s, known) or s) == project
            if inside:
                pos, n, in_mod = i, 0, False
        elif inside and not in_mod:
            if kind == "module":
                in_mod = True
            else:
                pos = i
                n += kind == "item"
    return pos, n


def insert_log_line(text, project, item, known, children=()):
    """在工作紀錄的 project 段落加一行「n.item」（和底下的「-子項」）；沒有這段就加在最後。"""
    lines = text.rstrip().splitlines() if text.strip() else []
    pos, n = _block("\n".join(lines), project, known)
    new = [f"{n + 1}. {item}"] + [f"    - {c}" for c in children]  # 跟輸入框自動編號同格式
    if pos is None:
        lines += ([""] if lines else []) + [project] + new
    else:
        lines[pos + 1:pos + 1] = new
    return "\n".join(lines) + "\n"


def find_log_line(text, project, item, known, kinds=("item", "child"), under=None):
    """在 project 段落找 item 那行 → rows 的索引；先找完全相同，再找相似。找不到回傳 None。
    under：只找這個 rows 索引（「1.」項目）底下的「-」小項。"""
    rows = _scan(text, known)
    cands, inside = [], False
    for r, (i, kind, s) in enumerate(rows):
        if under is not None:
            if r <= under:
                continue
            if kind != "child":
                break
            cands.append((r, s))
            continue
        if kind == "project":
            inside = (match_project(s, known) or s) == project
        elif inside and kind in kinds:
            cands.append((r, s))
    key = re.sub(r"\s", "", item)
    for r, s in cands:
        if re.sub(r"\s", "", s) == key:
            return r
    for r, s in cands:
        if _similar(item, s):
            return r
    return None


def insert_child_log_line(text, project, parent, child, known):
    """待辦小項寫進紀錄：放在大項那行底下的「-」小項最後；紀錄沒有這個大項就連大項一起加。已經有就不動。"""
    p = find_log_line(text, project, parent, known, kinds=("item",))
    if p is None:
        return insert_log_line(text, project, parent, known, [child])
    if find_log_line(text, project, child, known, under=p) is not None:
        return text
    rows = _scan(text, known)
    end = rows[p][0] + 1
    for j, k2, _ in rows[p + 1:]:
        if k2 != "child":
            break
        end = j + 1
    lines = text.splitlines()
    lines.insert(end, f"    - {child}")
    return "\n".join(lines).rstrip() + "\n"


def remove_child_log_line(text, project, parent, child, known, drop_parent):
    """刪掉大項底下的那行「-」小項；drop_parent 且大項底下已經沒有小項 → 大項那行也刪。找不到就原樣回傳。"""
    p = find_log_line(text, project, parent, known, kinds=("item",))
    if p is None:
        return remove_log_line(text, project, child, known)
    c = find_log_line(text, project, child, known, under=p)
    if c is None:
        return text
    rows = _scan(text, known)
    lines = text.splitlines()
    alone = c == p + 1 and (c + 1 >= len(rows) or rows[c + 1][1] != "child")
    if drop_parent and alone:
        del lines[rows[p][0]:rows[c][0] + 1]
    else:
        del lines[rows[c][0]]
    _renumber(lines, project, known)
    return _drop_empty_projects("\n".join(lines), known) + "\n"


def _renumber(lines, project, known):
    """project 段落裡的「1.」項目重新照順序編號（模組底下重新從 1 開始）。"""
    inside, n = False, 0
    for i, kind, s in _scan("\n".join(lines), known):
        if kind == "project":
            inside, n = (match_project(s, known) or s) == project, 0
        elif inside and kind == "module":
            n = 0
        elif inside and kind == "item":
            n += 1
            indent = lines[i][:len(lines[i]) - len(lines[i].lstrip())]
            lines[i] = f"{indent}{n}. {s}"


def remove_log_line(text, project, item, known):
    """刪掉 project 段落裡 item 那行；是「1.」項目就連同緊接的「-子項」一起刪。找不到就原樣回傳。"""
    r = find_log_line(text, project, item, known)
    if r is None:
        return text
    lines = text.splitlines()
    rows = _scan(text, known)
    i, kind, _ = rows[r]
    end = i + 1
    if kind == "item":
        for j, k2, _ in rows[r + 1:]:
            if k2 != "child":
                break
            end = j + 1
    del lines[i:end]
    _renumber(lines, project, known)
    return _drop_empty_projects("\n".join(lines), known) + "\n"


def _drop_empty_projects(text, known):
    """刪掉底下已經沒有任何內容的專案標題（和它前面的空行）。"""
    lines = text.splitlines()
    rows = _scan(text, known)
    for r in reversed(range(len(rows))):
        i, kind, _ = rows[r]
        if kind == "project" and (r + 1 == len(rows) or rows[r + 1][1] == "project"):
            start = i
            while start > 0 and not lines[start - 1].strip():
                start -= 1
            del lines[start:i + 1]
    return "\n".join(lines).rstrip()


def match_project(name, known):
    """對到既有專案名（簡稱、相等或互相包含），對不到回傳 None。"""
    if ALIASES.get(name) in known:
        return ALIASES[name]
    if name in known:
        return name
    for k in known:
        if name in k or k in name:
            return k
    return None


# ---------- 待辦存檔 ----------
def _cell_items(cell):
    items = []
    for p in cell.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        num = p._p.pPr.numPr if p._p.pPr is not None else None
        lvl = num.ilvl.val if num is not None and num.ilvl is not None else 0
        if LV1.match(t):  # 手打的「-」也算第 2 層
            t, lvl = LV1.sub("", t), 1
        t = LV0.sub("", t)
        if not t:
            continue
        if lvl >= 1 and items:
            items[-1]["children"].append(t)
        else:
            items.append({"text": t, "children": []})
    return items


def normalize(state):
    """專案順序以 projects 為準（可新增刪除，不再強制固定標題）；只出現在 todos 的專案接在後面。"""
    names = list(dict.fromkeys(state.get("projects", [])))
    names += [k for k in state.get("todos", {}) if k not in names]
    todos = state.get("todos", {})
    return {"projects": names, "todos": {n: todos.get(n, []) for n in names}}


def load_state():
    if TODOS.exists():
        return normalize(json.loads(TODOS.read_text(encoding="utf-8")))
    # 第一次執行：用設定裡的專案清單，待辦全空
    if SETTINGS.exists():
        return normalize({"projects": list(FIXED_PROJECTS), "todos": {}})
    # 舊版沒有設定檔：從範本匯入專案順序與待辦
    d = docx.Document(TEMPLATE)
    rows = d.tables[0].rows[1:]
    state = {"projects": [], "todos": {}}
    for r in rows:
        name = r.cells[0].text.strip()
        state["projects"].append(name)
        state["todos"][name] = _cell_items(r.cells[2])
    state = normalize(state)
    save_state(state)
    return state


def draft_mtime():
    return DRAFT.stat().st_mtime if DRAFT.exists() else 0


def load_draft():
    return DRAFT.read_text(encoding="utf-8").rstrip() if DRAFT.exists() else ""


def save_draft(text):
    DRAFT.write_text(text.rstrip() + "\n", encoding="utf-8")


def clear_draft():
    DRAFT.unlink(missing_ok=True)


def import_report(path):
    """讀一份週報 Word 的待辦欄 → (專案順序, {專案: 待辦})。"""
    d = docx.Document(path)
    if not d.tables:
        raise ValueError("這份 Word 裡找不到週報表格")
    projects, todos = [], {}
    for r in d.tables[0].rows[1:]:
        cells = r.cells
        if len(cells) < 3:
            continue
        name = cells[0].text.strip()
        if not name:
            continue
        if name not in todos:
            projects.append(name)
            todos[name] = []
        todos[name].extend(_cell_items(cells[2]))
    if not projects:
        raise ValueError("這份 Word 的表格裡沒有專案列")
    state = normalize({"projects": projects, "todos": todos})
    return state["projects"], state["todos"]


def save_state(state):
    if TODOS.exists():
        shutil.copyfile(TODOS, TODOS.with_suffix(".json.bak"))
    TODOS.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def resolve(state, name):
    """回傳專案的正式名稱；新專案會加入清單。"""
    k = match_project(name, state["projects"])
    if k is None:
        state["projects"].append(name)
        state["todos"].setdefault(name, [])
        k = name
    return k


def add_todos(state, text):
    """把「專案名＋1.＋-」格式的文字併進待辦。回傳新增筆數。"""
    n = 0
    for name, items in parse(text, state["projects"]):
        k = resolve(state, name)
        n += _merge(state["todos"].setdefault(k, []), items)
    return n


def _merge(lst, items):
    """items 併進待辦清單 lst：同名項目只補沒有的小項。回傳新增筆數。"""
    n = 0
    for it in items:
        same = next((x for x in lst if x["text"] == it["text"]), None)
        if same:
            for c in it["children"]:
                if c not in same["children"]:
                    same["children"].append(c)
                    n += 1
        else:
            lst.append(it)
            n += 1 + len(it["children"])
    return n


def parse_items(text):
    """不含專案名的項目文字 → [{'text', 'children'}]。每行一筆，「1.」可省略；「-」開頭是上一筆的小項。"""
    items = []
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        if LV1.match(line) and items:
            items[-1]["children"].append(LV1.sub("", line))
        else:
            t = LV0.sub("", LV1.sub("", line))
            if t:
                items.append({"text": t, "children": []})
    return items


def add_items(state, project, text):
    """把項目文字加到指定專案（新專案會加入清單）。回傳新增筆數。
    專案是從下拉選單選或打的，只認完全相同或簡稱；不用「互相包含」，打「測試新專案」才不會併進「測試」。"""
    known = state["projects"]
    k = project if project in known else ALIASES.get(project) if ALIASES.get(project) in known else None
    if k is None:
        known.append(project)
        k = project
    return _merge(state["todos"].setdefault(k, []), parse_items(text))


def has_project_line(text, known):
    """文字裡有沒有一行剛好是已知專案名或簡稱（不用「互相包含」，免得項目被當成標題）。"""
    names = set(known) | {a for a, full in ALIASES.items() if full in known}
    return any(l.strip() in names for l in text.splitlines())


# ---------- 判斷待辦是否已完成 ----------
def _bigrams(s):
    s = re.sub(r"[\s\W_]", "", s)
    return {s[i:i + 2] for i in range(len(s) - 1)}


def _similar(a, b):
    """互相包含，或兩字詞重疊佔較短一方的 80% 以上。
    門檻原本 60%：「加費設定彈窗」和「更多設定彈窗」只差前兩字、共用「設定彈窗」就被當成同一件。"""
    a2, b2 = re.sub(r"\s", "", a), re.sub(r"\s", "", b)
    if len(a2) >= 2 and len(b2) >= 2 and (a2 in b2 or b2 in a2):
        return True
    ga, gb = _bigrams(a), _bigrams(b)
    if not ga or not gb:
        return False
    return len(ga & gb) / min(len(ga), len(gb)) >= 0.8


def suggest_done(state, work):
    """回傳建議刪除的待辦 key 集合：(專案, i) 或 (專案, i, j)。"""
    lines = {}
    for name, items in work:
        k = match_project(name, state["projects"]) or name
        bucket = lines.setdefault(k, [])
        for it in items:
            bucket.append(it["text"])
            bucket.extend(it["children"])
    done = set()
    for proj, items in state["todos"].items():
        logs = lines.get(proj, [])
        for i, it in enumerate(items):
            kids = [j for j, c in enumerate(it["children"])
                    if any(_similar(c, l) for l in logs)]  # 小項只看自己的文字，大項有寫不代表小項完成
            done.update((proj, i, j) for j in kids)
            # 大項：文字對得到，且小項全部都完成（沒有小項就只看文字）才自動勾
            if len(kids) == len(it["children"]) and any(_similar(it["text"], l) for l in logs):
                done.add((proj, i))
    return done


def remove_todos(state, keys):
    for proj, items in state["todos"].items():
        kept = []
        for i, it in enumerate(items):
            it = dict(it, children=[c for j, c in enumerate(it["children"])
                                    if (proj, i, j) not in keys])
            if (proj, i) in keys and not it["children"]:
                continue  # 大項勾了但還有沒勾的小項 → 留著，小項才不會一起被刪
            kept.append(it)
        state["todos"][proj] = kept


# ---------- 寄送 ----------
MAIL_TO = ""
MAIL_CC = ""
MAIL_BODY = """Dear （主管稱呼）,
附件為本週週報，敬請參考。
–
Best regards,
範例科技股份有限公司
（中文姓名） （英文姓名）
E-mail：（你的 Email）
TEL：(02) 1234-5678
FAX：(02) 1234-5679
100台北市範例區範例路1號2樓
"""
HEADER = ""  # 週報表頭（例：王小明 UI/UX設計師），空的就沿用範本

# 第一次開時的預設值：公司資訊、專案帶好，個人欄位留空；實際值存在 settings.json
DEFAULT_SETTINGS = {
    "header": "",
    "mail_to": MAIL_TO,
    "mail_cc": MAIL_CC,
    "mail_body": MAIL_BODY,
    "projects": FIXED_PROJECTS,
    "aliases": ALIASES,
}


# ---------- 個人設定 ----------
def load_settings():
    """讀 settings.json 套到模組變數；沒有檔案回傳 None（要請使用者先填）。"""
    global HEADER, MAIL_TO, MAIL_CC, MAIL_BODY, FIXED_PROJECTS, ALIASES
    if not SETTINGS.exists():
        return None
    s = dict(DEFAULT_SETTINGS, **json.loads(SETTINGS.read_text(encoding="utf-8")))
    HEADER, MAIL_TO, MAIL_CC, MAIL_BODY = s["header"], s["mail_to"], s["mail_cc"], s["mail_body"]
    FIXED_PROJECTS, ALIASES = list(s["projects"]) or ["其他"], dict(s["aliases"])
    return s


def save_settings(s):
    SETTINGS.write_text(json.dumps(s, ensure_ascii=False, indent=2), encoding="utf-8")
    return load_settings()


# ---------- 檢查新版 ----------
def _ver(s):
    return tuple(int(x) for x in re.findall(r"\d+", s)[:3])


def check_update():
    """到 GitHub 查最新版 → (版本, 下載頁網址)；已是最新、沒網路、查不到都回傳 None。"""
    import urllib.request
    try:
        req = urllib.request.Request(f"https://api.github.com/repos/{REPO}/releases/latest",
                                     headers={"Accept": "application/vnd.github+json",
                                              "User-Agent": "weekly-report-helper"})
        with urllib.request.urlopen(req, timeout=5) as r:
            data = json.loads(r.read().decode("utf-8"))
        tag = data.get("tag_name", "")
        if tag and _ver(tag) > _ver(VERSION):
            return tag.lstrip("vV"), data.get("html_url") or f"https://github.com/{REPO}/releases/latest"
    except Exception:
        pass
    return None


# 中文經環境變數傳給 PowerShell，腳本本身只有英文，避免編碼問題
_PS_MAIL = r"""
$ErrorActionPreference = 'Stop'
$ol = New-Object -ComObject Outlook.Application
$m = $ol.CreateItem(0)
$m.To = $env:WR_TO
$m.CC = $env:WR_CC
$m.Subject = $env:WR_SUBJECT
$m.Body = $env:WR_BODY
[void]$m.Attachments.Add($env:WR_FILE)
$m.Display()
"""


def mail_subject(date):
    return f"{date.year - 1911}-{date:%m-%d}範例科技週工作紀錄"


def open_mail(path, date):
    """用 Outlook（classic）開一封填好收件者、主旨、內文、附件的新郵件（不會自動寄出）。"""
    import os
    import subprocess
    env = dict(os.environ, WR_TO=MAIL_TO, WR_CC=MAIL_CC, WR_SUBJECT=mail_subject(date),
               WR_BODY=MAIL_BODY, WR_FILE=str(Path(path).resolve()))
    r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", _PS_MAIL],
                       env=env, capture_output=True, text=True, encoding="utf-8", errors="replace",
                       creationflags=subprocess.CREATE_NO_WINDOW, timeout=120)
    if r.returncode != 0:
        raise RuntimeError(r.stderr.strip() or "Outlook 沒有回應")


# ---------- 產生 Word ----------
def this_friday(today=None):
    today = today or dt.date.today()
    return today + dt.timedelta(days=4 - today.weekday())


def output_path(date):
    return OUT_DIR / f"{date.year - 1911}-{date:%m-%d} 範例科技週工作紀錄.docx"


def _new_num(numbering, abstract_id):
    """新增一個從 1 重新編號的 w:num，回傳 numId。"""
    ids = [int(n.get(qn("w:numId"))) for n in numbering.findall(qn("w:num"))]
    nid = max(ids) + 1
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(nid))
    a = OxmlElement("w:abstractNumId")
    a.set(qn("w:val"), str(abstract_id))
    num.append(a)
    for lvl in ("0", "1"):
        o = OxmlElement("w:lvlOverride")
        o.set(qn("w:ilvl"), lvl)
        s = OxmlElement("w:startOverride")
        s.set(qn("w:val"), "1")
        o.append(s)
        num.append(o)
    numbering.append(num)
    return nid


def _samples(cell):
    """從範本格取出第 1 層、第 2 層段落與其 abstractNumId。"""
    lv0 = lv1 = None
    for p in cell.paragraphs:
        np_ = p._p.pPr.numPr
        if np_.ilvl.val == 0 and lv0 is None:
            lv0 = p._p
        if np_.ilvl.val == 1 and lv1 is None:
            lv1 = p._p
    return lv0, lv1, int(lv0.pPr.numPr.numId.val)


def _set_text(p, text):
    runs = p.findall(qn("w:r"))
    for r in runs[1:]:
        p.remove(r)
    r = runs[0]
    for t in r.findall(qn("w:t")):
        r.remove(t)
    t = OxmlElement("w:t")
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.append(t)


def _fill(cell, items, sample, empty_p, numbering, abstract_map):
    tc = cell._tc
    for p in tc.findall(qn("w:p")):
        tc.remove(p)
    if not items:
        tc.append(copy.deepcopy(empty_p))
        return
    lv0, lv1, nid0 = sample
    nid = _new_num(numbering, abstract_map[nid0])
    for it in items:
        for text, src in [(it["text"], lv0)] + [(c, lv1) for c in it["children"]]:
            p = copy.deepcopy(src)
            p.pPr.numPr.numId.val = nid
            _set_text(p, text)
            tc.append(p)


def empty_projects(work, state):
    """本週工作紀錄和（刪掉完成項後的）待辦都沒有內容的專案，照清單順序。"""
    s = copy.deepcopy(state)  # resolve 會把新專案加進清單，別動到原本的
    has = {resolve(s, name) for name, items in work if items}
    return [p for p in state["projects"] if p not in has and not state["todos"].get(p)]


def build(work, state, date, path=None, skip=()):
    """work: parse() 的結果；state: 已刪除完成項的待辦；skip: 不放進週報的專案。回傳輸出路徑。"""
    d = docx.Document(TEMPLATE)
    numbering = d.part.numbering_part.element
    abstract_map = {int(n.get(qn("w:numId"))): int(n.find(qn("w:abstractNumId")).get(qn("w:val")))
                    for n in numbering.findall(qn("w:num"))}

    # 表頭姓名職稱（沒設定就沿用範本）
    if HEADER:
        head = d.tables[0].rows[0].cells[0].paragraphs[0]._p
        if head.findall(qn("w:r")):
            _set_text(head, HEADER)

    # 標題日期
    title = d.paragraphs[1]
    title.runs[0].text = f"本週工作狀態書（{date:%Y.%m.%d}）"
    for r in title.runs[1:]:
        r._r.getparent().remove(r._r)

    table = d.tables[0]
    rows = table.rows
    work_sample = _samples(rows[WORK_SAMPLE[0]].cells[WORK_SAMPLE[1]])
    todo_sample = _samples(rows[TODO_SAMPLE[0]].cells[TODO_SAMPLE[1]])
    empty_p = copy.deepcopy(rows[EMPTY_SAMPLE[0]].cells[EMPTY_SAMPLE[1]].paragraphs[0]._p)
    row_tpl = copy.deepcopy(rows[EMPTY_SAMPLE[0]]._tr)
    name_p = copy.deepcopy(rows[EMPTY_SAMPLE[0]].cells[0].paragraphs[0]._p)
    for r in list(rows)[1:]:
        table._tbl.remove(r._tr)

    work_by = {}
    for name, items in work:
        k = resolve(state, name)
        work_by.setdefault(k, []).extend(items)

    for proj in state["projects"]:
        if proj in skip:
            continue
        tr = copy.deepcopy(row_tpl)
        table._tbl.append(tr)
        cells = table.rows[-1].cells
        tc0 = cells[0]._tc
        for p in tc0.findall(qn("w:p")):
            tc0.remove(p)
        p = copy.deepcopy(name_p)
        _set_text(p, proj)
        tc0.append(p)
        _fill(cells[1], work_by.get(proj, []), work_sample, empty_p, numbering, abstract_map)
        _fill(cells[2], state["todos"].get(proj, []), todo_sample, empty_p, numbering, abstract_map)

    path = Path(path or output_path(date))
    path.parent.mkdir(parents=True, exist_ok=True)
    d.save(path)
    return path
