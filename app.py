import os, re, json, unicodedata, difflib, sqlite3, threading, hashlib, secrets
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file, session, redirect, url_for
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", secrets.token_hex(32))
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
MERE_PATH  = os.path.join(BASE_DIR, "mere.xlsx")
LOG_PATH   = os.path.join(BASE_DIR, "log.json")
USERS_PATH = os.path.join(BASE_DIR, "users.json")
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# Admin password (set via Railway env var ADMIN_PASSWORD)
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "admin2026")

NCOLS  = 9
C_NEW  = "FFF2CC"
thin   = Side(style="thin", color="CCCCCC")
BDR    = Border(left=thin, right=thin, top=thin, bottom=thin)

write_lock = threading.Lock()
_db = None

def strip_accents(s):
    s = unicodedata.normalize("NFD", str(s or ""))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower()

def normalize(s):
    s = strip_accents(s)
    return re.sub(r"[-_\s\'\u2019\u2018]+", " ", s).strip()

def clean_phone(p):
    # Supprime espaces, tirets, points, parenthèses ET étoiles (*01..., *56...)
    p = re.sub(r"[\s\-\.\(\)\*]+", "", str(p or ""))
    return re.sub(r"^\+229|^00229", "", p).strip()

def load_log():
    if os.path.exists(LOG_PATH):
        with open(LOG_PATH, encoding="utf-8") as f:
            return json.load(f)
    return {"sessions": [], "total_inserted": 0, "users": {}}

def save_log(log):
    with open(LOG_PATH, "w", encoding="utf-8") as f:
        json.dump(log, f, ensure_ascii=False, indent=2)

# ── User management ───────────────────────────────────────────────────────────

def hash_pw(pw):
    return hashlib.sha256(pw.encode("utf-8")).hexdigest()

def load_users():
    if os.path.exists(USERS_PATH):
        with open(USERS_PATH, encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_users(users):
    with open(USERS_PATH, "w", encoding="utf-8") as f:
        json.dump(users, f, ensure_ascii=False, indent=2)

def get_user(username):
    return load_users().get(username.lower().strip())

def create_user(username, password):
    users = load_users()
    key = username.lower().strip()
    if key in users:
        return False, "Utilisateur déjà existant"
    users[key] = {
        "display": username.strip(),
        "password": hash_pw(password),
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    save_users(users)
    return True, "Utilisateur créé"

def check_password(username, password):
    user = get_user(username)
    if not user: return False
    return user["password"] == hash_pw(password)

def is_admin():
    return session.get("role") == "admin"

def is_logged_in():
    return session.get("username") is not None or is_admin()

def current_username():
    if is_admin(): return "__admin__"
    return session.get("username", "")

def require_login(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if not is_logged_in():
            return redirect(url_for("login_page"))
        return f(*args, **kwargs)
    return decorated

def require_admin(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if not is_admin():
            return jsonify({"error": "Accès refusé"}), 403
        return f(*args, **kwargs)
    return decorated

def get_db():
    global _db
    if _db is None:
        _db = build_index_db()
    return _db

def invalidate_db():
    global _db
    _db = None

def build_index_db():
    print("\u26a1 Construction de l'index SQLite...", flush=True)
    t0 = datetime.now()
    wb = load_workbook(MERE_PATH, read_only=True)
    ws = wb.active
    con = sqlite3.connect(":memory:", check_same_thread=False)
    con.execute("""CREATE TABLE slots (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        village TEXT, village_norm TEXT, arrond_norm TEXT,
        row_start INTEGER, filled INTEGER, last_data_row INTEGER)""")
    con.execute("CREATE TABLE phones (phone TEXT PRIMARY KEY)")
    con.execute("CREATE TABLE arrond_last (arrond_norm TEXT PRIMARY KEY, last_data_row INTEGER)")
    slots = []
    phones = []
    arrond_rows = {}
    arrond_norm_cur = ""
    for r, row in enumerate(ws.iter_rows(values_only=True), 1):
        v1 = str(row[0] or "").strip()
        v5 = str(row[4] or "").strip() if len(row) > 4 else ""
        v2 = str(row[1] or "").strip() if len(row) > 1 else ""
        v9 = str(row[8] or "").strip() if len(row) > 8 else ""
        if "ARRONDISSEMENT" in v1 and v1[:1].isdigit():
            arrond_norm_cur = normalize(v1)
        if v1 == "1" and v2 and arrond_norm_cur:
            slots.append((v2, normalize(v2), arrond_norm_cur, r, 0, r+4))
        if v1 in ["1","2","3","4","5"] and arrond_norm_cur:
            arrond_rows[arrond_norm_cur] = r
        if v9:
            p = clean_phone(v9)
            if p: phones.append((p,))
    wb.close()
    con.executemany(
        "INSERT INTO slots(village,village_norm,arrond_norm,row_start,filled,last_data_row) VALUES(?,?,?,?,?,?)",
        slots)
    con.executemany("INSERT OR IGNORE INTO phones VALUES(?)", phones)
    con.executemany("INSERT OR REPLACE INTO arrond_last VALUES(?,?)", list(arrond_rows.items()))
    # Compute actual filled counts
    wb2 = load_workbook(MERE_PATH, read_only=True)
    ws2 = wb2.active
    row_to_id = {r: sid for sid, r in con.execute("SELECT id, row_start FROM slots").fetchall()}
    for r, row in enumerate(ws2.iter_rows(values_only=True), 1):
        v1 = str(row[0] or "").strip()
        v3 = str(row[2] or "").strip() if len(row) > 2 else ""
        if v1 in ["1","2","3","4","5"] and v3:
            slot_row = r - (int(v1) - 1)
            if slot_row in row_to_id:
                sid = row_to_id[slot_row]
                con.execute("UPDATE slots SET filled=MAX(filled,?), last_data_row=MAX(last_data_row,?) WHERE id=?",
                            (int(v1), r, sid))
    wb2.close()
    con.execute("CREATE INDEX idx_vn ON slots(village_norm)")
    con.execute("CREATE INDEX idx_an ON slots(arrond_norm)")
    con.commit()
    elapsed = (datetime.now() - t0).total_seconds()
    n = con.execute("SELECT COUNT(*) FROM slots").fetchone()[0]
    np = con.execute("SELECT COUNT(*) FROM phones").fetchone()[0]
    print(f"   \u2705 Index pr\u00eat en {elapsed:.1f}s \u2014 {n} villages, {np} t\u00e9l.", flush=True)
    return con

def find_slot(con, quartier):
    key = normalize(quartier)
    row = con.execute(
        "SELECT id,village,row_start,filled,last_data_row,arrond_norm FROM slots WHERE village_norm=? LIMIT 1",
        (key,)).fetchone()
    if row: return row
    all_keys = [r[0] for r in con.execute("SELECT DISTINCT village_norm FROM slots").fetchall()]
    matches = difflib.get_close_matches(key, all_keys, n=1, cutoff=0.85)
    if matches:
        return con.execute(
            "SELECT id,village,row_start,filled,last_data_row,arrond_norm FROM slots WHERE village_norm=? LIMIT 1",
            (matches[0],)).fetchone()
    return None

def style_new_row(ws, rn):
    ws.row_dimensions[rn].height = 15
    for c in range(1, NCOLS+1):
        cell = ws.cell(row=rn, column=c)
        cell.fill      = PatternFill("solid", start_color=C_NEW)
        cell.font      = Font(name="Arial", size=10)
        cell.border    = BDR
        cell.alignment = Alignment(vertical="center")
    ws.cell(row=rn, column=1).alignment = Alignment(horizontal="center", vertical="center")

def write_person_row(ws, rn, num, village, p):
    ws.cell(row=rn, column=1).value = num
    ws.cell(row=rn, column=2).value = village
    ws.cell(row=rn, column=3).value = p["nom"]
    ws.cell(row=rn, column=4).value = p["prenom"]
    ws.cell(row=rn, column=5).value = p.get("partis","")
    ws.cell(row=rn, column=6).value = p.get("profession","")
    ws.cell(row=rn, column=7).value = p.get("date_naissance","")
    ws.cell(row=rn, column=8).value = p.get("lieu_naissance","")
    ws.cell(row=rn, column=9).value = p["telephone"]

ALIASES = {
    "quartier":       ["quartier","village","localite","quartier/village","quartiers",
                       "villages / quartiers de ville"],
    "nom":            ["nom","noms","name"],
    "prenom":         ["prenom","prenoms","firstname"],
    "nom_prenom":     ["nom et prenoms","nom et prenom","noms et prenoms","responsable"],
    "telephone":      ["telephone","tel","phone","mobile","contact",
                       "numero de telephone","numero",
                       "adresse complete","adresse complete","adresse",
                       "adresse compete"],
    "partis":         ["partis","parti","party"],
    "profession":     ["profession","metier","emploi"],
    "date_naissance": ["date de naissance","date_naissance","naissance"],
    "lieu_naissance": ["lieu de naissance","lieu_naissance"],
}

def split_nom_prenom(full_name):
    """
    Premier mot = NOM, tout le reste = PRENOMS.
    Ex: 'LAWANI SENI Akim' → NOM='LAWANI', PRENOM='SENI Akim'
    Ex: 'GNAMMI Datonga'   → NOM='GNAMMI', PRENOM='Datonga'
    Si un seul mot : NOM=PRENOM=ce mot (évite le rejet).
    """
    parts = full_name.strip().split()
    if not parts:       return '', ''
    if len(parts) == 1: return parts[0], parts[0]
    return parts[0], ' '.join(parts[1:])

def detect_columns(ws_in):
    """
    Détecte les colonnes en scannant les premières lignes.
    Supporte aussi la colonne 'Nom et Prénoms' combinée.
    """
    for r in range(1, 10):
        mapping = {}
        for c in range(1, ws_in.max_column+1):
            h = normalize(str(ws_in.cell(row=r, column=c).value or ""))
            for field, aliases in ALIASES.items():
                if h in aliases:
                    mapping[field] = c
        # Cas normal : nom + prenom séparés
        if "nom" in mapping and "prenom" in mapping:
            return mapping, r
        # Cas spécial : colonne combinée "Nom et Prénoms" ou "RESPONSABLE"
        if "nom_prenom" in mapping:
            mapping["nom"]    = mapping["nom_prenom"]
            mapping["prenom"] = mapping["nom_prenom"]  # même colonne, sera splitée
            mapping["_split_nom_prenom"] = True
            return mapping, r
    return {}, None

def integrate_file(filepath, filename, username=None):
    # ── Support .xls : convertir en xlsx à la volée ───────────────────────────
    if filepath.lower().endswith(".xls"):
        try:
            import xlrd
            xls_wb  = xlrd.open_workbook(filepath)
            xls_ws  = xls_wb.sheet_by_index(0)
            from openpyxl import Workbook as WB
            new_wb  = WB()
            new_ws  = new_wb.active
            for rr in range(xls_ws.nrows):
                row_vals = []
                for cc in range(xls_ws.ncols):
                    cell = xls_ws.cell(rr, cc)
                    # xlrd type 2 = float, might be phone number
                    if cell.ctype == 2 and cell.value == int(cell.value):
                        row_vals.append(str(int(cell.value)))
                    else:
                        row_vals.append(cell.value if cell.value != '' else None)
                new_ws.append(row_vals)
            xlsx_path = filepath + ".xlsx"
            new_wb.save(xlsx_path)
            filepath = xlsx_path
        except ImportError:
            return {"filename": filename,
                    "error": "Fichier .xls non supporté sur ce serveur. Veuillez le convertir en .xlsx avec Excel ou LibreOffice."}
        except Exception as e:
            return {"filename": filename, "error": f"Erreur lecture .xls : {e}"}

    wb_in = load_workbook(filepath, data_only=True)
    ws_in = wb_in.active
    col_map, header_row = detect_columns(ws_in)
    if not col_map:
        return {"filename": filename, "error": "Colonnes non detectees (Nom + Prenom requis)"}

    def gv(r, field):
        col = col_map.get(field)
        return str(ws_in.cell(row=r, column=col).value or "").strip() if col else ""

    is_split = col_map.get("_split_nom_prenom", False)

    def find_tel(r, base_col):
        if not base_col: return ""
        for delta in [0, 1, -1, 2]:
            v = str(ws_in.cell(row=r, column=base_col+delta).value or "").strip()
            if v and re.search(r"\d{5,}", v.replace(" ", "").replace("-","")):
                return v
        return ""

    # Détecter alt_map (colonnes décalées de +1) seulement si pas de mode split
    alt_map = None
    if not is_split:
        alt_hits = 0
        for r in range(header_row+1, min(header_row+60, ws_in.max_row+1)):
            if not str(ws_in.cell(row=r, column=col_map.get("nom",99)).value or "").strip():
                if str(ws_in.cell(row=r, column=col_map.get("nom",99)+1).value or "").strip():
                    alt_hits += 1
        if alt_hits > 0:
            alt_map = {f: c+1 for f, c in col_map.items() if not f.startswith("_")}

    def get_row_data(r):
        # ── Quartier ──────────────────────────────────────────────────────────
        quartier = gv(r, "quartier")
        if not quartier:
            col_nom_val = col_map.get("nom", col_map.get("nom_prenom", 99))
            skip = {"N","NOM","NOMS","PRENOM","PRENOMS","PROFESSION",
                    "ADRESSE COMPLETE","ADRESSE COMPETE","CONTACT",
                    "QUARTIER","QUARTIERS","N° QUARTIERS",""}

            # Cas 7ème CE : quartier en col_nom-1 (col juste avant Nom et Prénoms)
            col_before = col_nom_val - 1
            if col_before >= 1:
                v = str(ws_in.cell(row=r, column=col_before).value or "").strip()
                norm_v = normalize(v).upper()
                if (v and norm_v not in [normalize(s) for s in skip]
                        and not re.match(r"^\d+$", v)
                        and len(v) > 1):
                    quartier = v

            # Cas général : chercher aussi col_quartier±1
            if not quartier and "quartier" in col_map:
                col_q = col_map["quartier"]
                for delta in [1, -1]:
                    v = str(ws_in.cell(row=r, column=col_q+delta).value or "").strip()
                    if v and normalize(v) not in [normalize(s) for s in skip] \
                            and not re.match(r"^\d+$", v):
                        quartier = v; break

        # ── Nom / Prénom ──────────────────────────────────────────────────────
        if is_split:
            # Mode "Nom et Prénoms" combiné ou "RESPONSABLE" : splitter le contenu
            full = gv(r, "nom").strip()
            if full:
                nom, prenom = split_nom_prenom(full)
            else:
                nom, prenom = "", ""
        else:
            nom    = gv(r, "nom").strip()
            prenom = gv(r, "prenom").strip()

            # Cas RoW 4ème VF : NOM vide mais PRENOM contient nom+prénom combinés
            if not nom and prenom:
                nom, prenom = split_nom_prenom(prenom)

            # Cas RoW 4ème VF : NOM contient nom+prénom combinés, PRENOM vide
            elif nom and not prenom and " " in nom:
                # Vérifier si c'est vraiment un nom combiné (plusieurs mots)
                col_n = col_map.get("nom")
                col_p = col_map.get("prenom")
                if col_n and col_p and col_n != col_p:
                    nom, prenom = split_nom_prenom(nom)

        # ── Téléphone ─────────────────────────────────────────────────────────
        tel_raw = find_tel(r, col_map.get("telephone"))

        # Alt_map : colonnes décalées si ligne principale vide
        if alt_map and not nom:
            n2 = str(ws_in.cell(row=r, column=alt_map.get("nom",99)).value or "").strip()
            if n2:
                q2 = str(ws_in.cell(row=r, column=alt_map.get("quartier",99)).value or "").strip()
                p2 = str(ws_in.cell(row=r, column=alt_map.get("prenom",99)).value or "").strip()
                if q2 and not quartier: quartier = q2
                if not p2 and " " in n2:
                    nom, prenom = split_nom_prenom(n2)
                else:
                    nom = n2; prenom = p2
                if not tel_raw: tel_raw = find_tel(r, alt_map.get("telephone"))

        return quartier, nom, prenom, tel_raw, {
            "partis":         gv(r, "partis"),
            "profession":     gv(r, "profession"),
            "date_naissance": gv(r, "date_naissance"),
            "lieu_naissance": gv(r, "lieu_naissance"),
        }

    persons = []; rejected = []; seen_phones = set(); last_quartier = None

    for r in range(1, ws_in.max_row+1):
        for c in range(1, min(ws_in.max_column+1, 6)):
            cv = str(ws_in.cell(row=r, column=c).value or "").strip()
            m = re.match(r"^QUARTIER\s*[:\-]\s*(.+)", cv, re.IGNORECASE)
            if m: last_quartier = m.group(1).strip(); break
            if re.match(r"^QUARTIER\s*[:\-]?\s*$", cv, re.IGNORECASE):
                nxt = str(ws_in.cell(row=r, column=c+1).value or "").strip()
                if nxt: last_quartier = nxt
                break
        if r < header_row+1: continue
        quartier, nom, prenom, tel_raw, extras = get_row_data(r)
        tel = clean_phone(tel_raw)
        if not any([quartier, nom, prenom, tel_raw]): continue
        if normalize(nom) in ["nom","noms","name"]: continue
        if normalize(prenom) in ["prenom","prenoms"]: continue
        first_cell = str(ws_in.cell(row=r, column=1).value or "").strip()
        if re.match(r"^QUARTIER\s*[:\-]", first_cell, re.IGNORECASE): continue
        skip_row = False
        for c in range(1, min(ws_in.max_column+1, 4)):
            cv = str(ws_in.cell(row=r, column=c).value or "").strip().upper()
            if cv.startswith("DEPARTEMENT") or "ARRONDISSEMENT" in cv or cv.startswith("COMMUNE"):
                skip_row = True; break
        if skip_row or not nom: continue
        if quartier: last_quartier = quartier
        elif not quartier and nom and (prenom or tel):
            if last_quartier: quartier = last_quartier
        missing = []
        if not quartier: missing.append("Quartier")
        if not nom:      missing.append("Nom")
        if not prenom:   missing.append("Prenom")
        if not tel:      missing.append("Telephone")
        if missing:
            reason_str = ", ".join(missing)
            rejected.append({"row": r, "quartier": quartier, "nom": nom,
                             "reason": "Manquant : " + reason_str})
            continue
        if tel in seen_phones:
            rejected.append({"row": r, "quartier": quartier, "nom": nom,
                             "reason": f"Doublon interne : {tel_raw}"}); continue
        seen_phones.add(tel)
        persons.append({"row": r, "quartier": quartier, "nom": nom, "prenom": prenom,
                        **extras, "telephone": tel_raw, "tel_clean": tel})

    with write_lock:
        con    = get_db()
        wb_out = load_workbook(MERE_PATH)
        ws_out = wb_out.active
        inserted = overflow = new_quartier = duplicates = 0

        # ── Phase 1 : Résoudre tous les placements AVANT toute écriture ──────
        direct_writes  = []   # (target_row, person) pour slots existants (filled<5)
        overflow_groups = {}  # sid -> {village, insert_after, persons_list, start_num}
        new_village_list = [] # (quartier, persons_list) pour nouveaux villages
        phones_to_add  = []

        # Snapshot complet des slots
        slot_state = {}
        for row in con.execute("SELECT id,village,row_start,filled,last_data_row,arrond_norm FROM slots").fetchall():
            sid, village, row_start, filled, last_data_row, arrond_n = row
            slot_state[sid] = {"village": village, "row_start": row_start,
                               "filled": filled, "last_data_row": last_data_row}

        slot_fill_count = {}  # sid -> filled count courant

        for p in persons:
            tel = p["tel_clean"]
            if con.execute("SELECT 1 FROM phones WHERE phone=?", (tel,)).fetchone():
                duplicates += 1
                rejected.append({"row": p["row"], "quartier": p["quartier"], "nom": p["nom"],
                                 "reason": f"Doublon tel : {p['telephone']}"}); continue

            slot = find_slot(con, p["quartier"])

            if slot:
                sid, village, row_start, filled, last_data_row, arrond_n = slot
                current_filled = slot_fill_count.get(sid, filled)

                if current_filled < 5:
                    # Écriture directe sur slot pré-alloué (N° 1 à 5)
                    target_row = row_start + current_filled
                    direct_writes.append((target_row, current_filled + 1, village, p))
                else:
                    # Débordement → regrouper par village pour une seule insertion groupée
                    if sid not in overflow_groups:
                        overflow_groups[sid] = {
                            "village":      village,
                            "insert_after": slot_state[sid]["last_data_row"],
                            "start_num":    current_filled + 1,
                            "persons":      []
                        }
                    overflow_groups[sid]["persons"].append(p)
                    overflow += 1

                slot_fill_count[sid] = current_filled + 1
                phones_to_add.append(tel)
                inserted += 1

            else:
                # Nouveau village → regrouper par nom de quartier
                key = normalize(p["quartier"])
                existing = next((g for g in new_village_list if normalize(g["quartier"]) == key), None)
                if existing:
                    existing["persons"].append(p)
                else:
                    new_village_list.append({"quartier": p["quartier"], "persons": [p]})
                new_quartier += 1
                phones_to_add.append(tel)
                inserted += 1

        # ── Phase 2 : Trier les insertions par position DÉCROISSANTE ─────────
        # overflow_groups triés par insert_after DESC → bottom-to-top → pas de décalage
        sorted_overflows = sorted(overflow_groups.values(), key=lambda g: g["insert_after"], reverse=True)

        # ── Phase 3 : Styles helper ───────────────────────────────────────────
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side as Sd
        thin2 = Sd(style="thin", color="CCCCCC")
        bdr2  = Border(left=thin2, right=thin2, top=thin2, bottom=thin2)

        def style_inserted_row(ws, rn, bg=C_NEW):
            ws.row_dimensions[rn].height = 15
            for c in range(1, NCOLS+1):
                cell = ws.cell(row=rn, column=c)
                cell.fill      = PatternFill("solid", start_color=bg)
                cell.font      = Font(name="Arial", size=10)
                cell.border    = bdr2
                cell.alignment = Alignment(vertical="center")
            ws.cell(row=rn, column=1).alignment = Alignment(horizontal="center", vertical="center")

        def write_to_row(ws, rn, num, village, p):
            ws.cell(row=rn, column=1).value = num
            ws.cell(row=rn, column=2).value = village
            ws.cell(row=rn, column=3).value = p["nom"]
            ws.cell(row=rn, column=4).value = p["prenom"]
            ws.cell(row=rn, column=5).value = p.get("partis","")
            ws.cell(row=rn, column=6).value = p.get("profession","")
            ws.cell(row=rn, column=7).value = p.get("date_naissance","")
            ws.cell(row=rn, column=8).value = p.get("lieu_naissance","")
            ws.cell(row=rn, column=9).value = p["telephone"]

        # ── Phase 4 : Insertions groupées bottom-to-top (overflow) ───────────
        for group in sorted_overflows:
            count        = len(group["persons"])
            insert_at    = group["insert_after"] + 1  # insérer juste après dernière ligne
            village      = group["village"]
            start_num    = group["start_num"]

            # UNE SEULE insertion de N lignes → bloc contigu garanti
            ws_out.insert_rows(insert_at, amount=count)

            for i, p in enumerate(group["persons"]):
                rn = insert_at + i
                style_inserted_row(ws_out, rn)
                write_to_row(ws_out, rn, start_num + i, village, p)

        # ── Phase 5 : Nouveaux villages ajoutés en fin de fichier ────────────
        # (fin de fichier = après tous les décalages dus aux inserts ci-dessus)
        for nv in new_village_list:
            count    = len(nv["persons"])
            base_row = ws_out.max_row + 1
            ws_out.insert_rows(base_row, amount=count)
            for i, p in enumerate(nv["persons"]):
                rn = base_row + i
                style_inserted_row(ws_out, rn)
                write_to_row(ws_out, rn, i + 1, nv["quartier"], p)

        # ── Phase 6 : Écriture directe sur slots existants ───────────────────
        # IMPORTANT : recalculer les positions après les insert_rows
        # Chaque insert_rows(pos, n) décale de +n toutes les lignes >= pos
        # On accumule les décalages pour corriger les positions directes

        # Reconstruire la liste des décalages provoqués par les inserts (bottom-to-top → ordre inverse pour recalcul)
        # Les inserts ont été faits du bas vers le haut, donc recalculer du haut vers le bas
        all_inserts = []  # (insert_at, count)
        for group in sorted_overflows:
            all_inserts.append((group["insert_after"] + 1, len(group["persons"])))

        def adjust_row(original_row):
            """Recalcule la position d'une ligne après tous les insert_rows."""
            adjusted = original_row
            for insert_at, count in all_inserts:
                if insert_at <= adjusted:
                    adjusted += count
            return adjusted

        for target_row, num, village, p in direct_writes:
            adjusted_row = adjust_row(target_row)
            ws_out.cell(row=adjusted_row, column=3).value = p["nom"]
            ws_out.cell(row=adjusted_row, column=4).value = p["prenom"]
            ws_out.cell(row=adjusted_row, column=5).value = p.get("partis","")
            ws_out.cell(row=adjusted_row, column=6).value = p.get("profession","")
            ws_out.cell(row=adjusted_row, column=7).value = p.get("date_naissance","")
            ws_out.cell(row=adjusted_row, column=8).value = p.get("lieu_naissance","")
            ws_out.cell(row=adjusted_row, column=9).value = p["telephone"]

        # ── Phase 7 : Téléphones + invalider cache ────────────────────────────
        for tel in phones_to_add:
            con.execute("INSERT OR IGNORE INTO phones VALUES(?)", (tel,))
        invalidate_db()

        con.commit()
        wb_out.save(MERE_PATH)

    log = load_log()
    log["total_inserted"] = log.get("total_inserted", 0) + inserted
    if "users" not in log: log["users"] = {}
    # Per-user stats
    uname = username or "inconnu"
    if uname not in log["users"]:
        log["users"][uname] = {"total_inserted": 0, "sessions": []}
    log["users"][uname]["total_inserted"] = log["users"][uname].get("total_inserted", 0) + inserted
    log["users"][uname]["sessions"].append({
        "file": filename, "date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "inserted": inserted, "rejected": len(rejected),
        "duplicates": duplicates, "overflow": overflow, "new_quartiers": new_quartier,
    })
    # Global sessions
    log["sessions"].append({
        "file": filename, "date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "inserted": inserted, "rejected": len(rejected),
        "duplicates": duplicates, "overflow": overflow, "new_quartiers": new_quartier,
        "user": uname,
    })
    save_log(log)
    con2 = get_db()
    total_personnes = con2.execute("SELECT SUM(filled) FROM slots").fetchone()[0] or 0
    return {
        "filename": filename, "inserted": inserted, "overflow": overflow,
        "new_quartiers": new_quartier, "duplicates": duplicates,
        "rejected_count": len(rejected), "rejected": rejected[:50],
        "total_mere": ws_out.max_row, "total_personnes": total_personnes,
        "total_cumule": log["total_inserted"],
    }


# ── Auth Routes ───────────────────────────────────────────────────────────────

@app.route("/login", methods=["GET"])
def login_page():
    if is_logged_in():
        return redirect(url_for("index"))
    return render_template("login.html")

@app.route("/login", methods=["POST"])
def login():
    data = request.get_json() or request.form
    username = str(data.get("username","")).strip()
    password = str(data.get("password","")).strip()

    # Check admin
    if password == ADMIN_PASSWORD and (username.lower() in ["admin","administrateur"] or password == ADMIN_PASSWORD):
        if username.lower() in ["admin","administrateur"]:
            session["role"] = "admin"
            session["username"] = "admin"
            return jsonify({"ok": True, "role": "admin"})

    # Check regular user
    if check_password(username, password):
        session["username"] = username.lower().strip()
        session["display"]  = get_user(username)["display"]
        session["role"]     = "user"
        return jsonify({"ok": True, "role": "user"})

    return jsonify({"ok": False, "error": "Identifiants incorrects"}), 401

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login_page"))

# ── Main App Routes ───────────────────────────────────────────────────────────

@app.route("/")
@require_login
def index():
    return render_template("index.html",
        username=session.get("display", session.get("username","")),
        is_admin=is_admin())

@app.route("/upload", methods=["POST"])
@require_login
def upload():
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "Aucun fichier recu"}), 400
    uname = current_username()
    results = []
    for f in files:
        if not f.filename.lower().endswith((".xlsx", ".xls")):
            results.append({"filename": f.filename, "error": "Format non supporte"}); continue

        # ── Nom unique par utilisateur + timestamp pour éviter collisions ──
        ext       = ".xls" if f.filename.lower().endswith(".xls") else ".xlsx"
        unique_id = f"{uname}_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}{ext}"
        filepath  = os.path.join(UPLOAD_DIR, unique_id)

        f.save(filepath)
        try:
            result = integrate_file(filepath, f.filename, username=uname)
            results.append(result)
        except Exception as e:
            results.append({"filename": f.filename, "error": str(e)})
        finally:
            try: os.remove(filepath)
            except: pass
    return jsonify({"results": results, "log": load_log()})

@app.route("/download")
@require_admin
def download():
    return send_file(MERE_PATH, as_attachment=True,
                     download_name="Villages_Quartiers_Benin_Final.xlsx")

@app.route("/stats")
@require_login
def stats():
    log = load_log()
    con = get_db()
    total_personnes = con.execute("SELECT SUM(filled) FROM slots").fetchone()[0] or 0
    wb = load_workbook(MERE_PATH, read_only=True)
    ws = wb.active
    total_lignes = ws.max_row
    wb.close()

    uname = current_username()
    user_stats = log.get("users", {}).get(uname, {})
    user_sessions = user_stats.get("sessions", [])[-10:]
    user_total = user_stats.get("total_inserted", 0)

    return jsonify({
        "total_lignes_mere":   total_lignes,
        "total_personnes":     total_personnes,
        "total_insere_cumule": log.get("total_inserted", 0),
        "nb_sessions":         len(log.get("sessions", [])),
        "sessions":            log.get("sessions", [])[-10:],
        # User-specific
        "user_sessions":       user_sessions,
        "user_total_inserted": user_total,
        "user_nb_sessions":    len(user_sessions),
        "is_admin":            is_admin(),
    })

@app.route("/admin/stats")
@require_admin
def admin_stats():
    """Stats complètes par utilisateur pour l'admin."""
    log = load_log()
    users_data = load_users()
    con = get_db()
    total_personnes = con.execute("SELECT SUM(filled) FROM slots").fetchone()[0] or 0

    user_summary = []
    for uname, udata in log.get("users", {}).items():
        display = users_data.get(uname, {}).get("display", uname) if uname != "__admin__" else "Admin"
        user_summary.append({
            "username":       uname,
            "display":        display,
            "total_inserted": udata.get("total_inserted", 0),
            "nb_sessions":    len(udata.get("sessions", [])),
            "last_session":   udata["sessions"][-1]["date"] if udata.get("sessions") else "—",
        })
    user_summary.sort(key=lambda x: x["total_inserted"], reverse=True)

    wb = load_workbook(MERE_PATH, read_only=True)
    ws = wb.active
    total_lignes = ws.max_row
    wb.close()

    return jsonify({
        "total_lignes_mere":   total_lignes,
        "total_personnes":     total_personnes,
        "total_insere_cumule": log.get("total_inserted", 0),
        "nb_sessions":         len(log.get("sessions", [])),
        "sessions":            log.get("sessions", [])[-20:],
        "users":               user_summary,
        "all_users":           list(users_data.keys()),
    })

@app.route("/admin/users", methods=["POST"])
@require_admin
def create_user_route():
    data = request.get_json() or {}
    username = str(data.get("username","")).strip()
    password = str(data.get("password","")).strip()
    if not username or not password:
        return jsonify({"ok": False, "error": "Prénom et mot de passe requis"}), 400
    ok, msg = create_user(username, password)
    return jsonify({"ok": ok, "message": msg})

@app.route("/admin/users/<username>", methods=["DELETE"])
@require_admin
def delete_user_route(username):
    users = load_users()
    key = username.lower().strip()
    if key not in users:
        return jsonify({"ok": False, "error": "Utilisateur introuvable"}), 404
    del users[key]
    save_users(users)
    return jsonify({"ok": True, "message": f"{username} supprimé"})

@app.route("/reset", methods=["POST"])
@require_admin
def reset():
    with write_lock:
        wb = load_workbook(MERE_PATH)
        ws = wb.active
        count = 0
        ORIGINAL_MAX = 28088
        for r in range(1, min(ORIGINAL_MAX+1, ws.max_row+1)):
            if str(ws.cell(row=r, column=1).value or "").strip() in ["1","2","3","4","5"]:
                for c in range(3, 10):
                    ws.cell(row=r, column=c).value = None
                count += 1
        if ws.max_row > ORIGINAL_MAX:
            ws.delete_rows(ORIGINAL_MAX+1, ws.max_row - ORIGINAL_MAX)
        wb.save(MERE_PATH)
        save_log({"sessions": [], "total_inserted": 0, "users": {}})
        invalidate_db()
    return jsonify({"message": f"{count} lignes reinitalisees"})

print("\n" + "="*55)
print("  \U0001f680 Villages & Quartiers Benin \u2014 Systeme v4")
print("  \u279c  http://localhost:5000")
print("="*55)

if not os.path.exists(MERE_PATH):
    print(f"\u274c Fichier mere introuvable : {MERE_PATH}")
else:
    print(f"\u2705 Fichier mere trouve : {MERE_PATH}")
    try:
        with app.app_context():
            get_db()
    except Exception as e:
        print(f"\u26a0\ufe0f  Index non construit : {e}")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
