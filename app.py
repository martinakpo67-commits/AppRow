"""
Villages & Quartiers Bénin — Système d'intégration v3
======================================================
- Index SQLite en mémoire (construit 1x au démarrage)
- Traitement ultra-rapide (recherche <0.001s vs 2s avant)
- Compatible Railway / Render / PythonAnywhere (mise en ligne)
- Multi-fichiers simultanés
"""

import os, re, json, unicodedata, difflib, sqlite3, shutil, threading
from datetime import datetime
from collections import defaultdict
from flask import Flask, render_template, request, jsonify, send_file
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
MERE_PATH  = os.path.join(BASE_DIR, 'mere.xlsx')
LOG_PATH   = os.path.join(BASE_DIR, 'log.json')
UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')
os.makedirs(UPLOAD_DIR, exist_ok=True)

NCOLS   = 9
C_WHITE = "FFFFFF"
C_NEW   = "FFF2CC"

thin = Side(style='thin', color='CCCCCC')
BDR  = Border(left=thin, right=thin, top=thin, bottom=thin)
wht  = Side(style='thin', color='FFFFFF')
WBDR = Border(left=wht, right=wht, top=wht, bottom=wht)

# ── Thread lock (one write at a time) ────────────────────────────────────────
write_lock = threading.Lock()

# ── SQLite in-memory index (rebuilt on demand) ────────────────────────────────
_db = None

def get_db():
    global _db
    if _db is None:
        _db = build_index_db()
    return _db

def invalidate_db():
    global _db
    _db = None

# ── String utilities ──────────────────────────────────────────────────────────

def strip_accents(s):
    s = unicodedata.normalize('NFD', str(s or ''))
    return ''.join(c for c in s if unicodedata.category(c) != 'Mn').lower()

def normalize(s):
    s = strip_accents(s)
    return re.sub(r'[-_\s\'\u2019\u2018]+', ' ', s).strip()

def clean_phone(p):
    p = re.sub(r'[\s\-\.\(\)]+', '', str(p or ''))
    return re.sub(r'^\+229|^00229', '', p).strip()

# ── Log helpers ───────────────────────────────────────────────────────────────

def load_log():
    if os.path.exists(LOG_PATH):
        with open(LOG_PATH, encoding='utf-8') as f:
            return json.load(f)
    return {'sessions': [], 'total_inserted': 0}

def save_log(log):
    with open(LOG_PATH, 'w', encoding='utf-8') as f:
        json.dump(log, f, ensure_ascii=False, indent=2)

# ── Build SQLite index from mere.xlsx ─────────────────────────────────────────

def build_index_db():
    print("⚡ Construction de l'index SQLite...", flush=True)
    t0 = datetime.now()

    wb = load_workbook(MERE_PATH, read_only=True)
    ws = wb.active

    con = sqlite3.connect(':memory:', check_same_thread=False)
    con.execute('PRAGMA journal_mode=WAL')
    con.execute('''CREATE TABLE slots (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        row_num     INTEGER,
        village     TEXT,
        village_norm TEXT,
        arrond      TEXT,
        arrond_norm TEXT,
        dept        TEXT,
        commune     TEXT,
        filled      INTEGER DEFAULT 0,
        last_row    INTEGER
    )''')
    con.execute('CREATE TABLE phones (phone TEXT PRIMARY KEY)')
    con.execute('CREATE TABLE arrond_last (arrond_norm TEXT PRIMARY KEY, last_row INTEGER)')

    dept_s = comm_s = arrond_lbl = arrond_norm_cur = ''
    slots, phones, arrond_rows = [], [], {}

    for r, row in enumerate(ws.iter_rows(values_only=True), 1):
        v1 = str(row[0] or '').strip()
        v5 = str(row[4] or '').strip() if len(row) > 4 else ''
        v9 = str(row[8] or '').strip() if len(row) > 8 else ''

        if v1.startswith('DEPARTEMENT:'):
            dept_s = v1.replace('DEPARTEMENT:', '').strip()
            comm_s = v5.replace('COMMUNE:', '').strip()
        elif 'ARRONDISSEMENT' in v1 and v1[:1].isdigit():
            arrond_lbl      = v1
            arrond_norm_cur = normalize(v1)
        elif v1 in ['1','2','3','4','5'] and arrond_norm_cur:
            village = str(row[1] or '').strip() if len(row) > 1 else ''
            arrond_rows[arrond_norm_cur] = r
            if v1 == '1' and village:
                slots.append((r, village, normalize(village),
                              arrond_lbl, arrond_norm_cur,
                              dept_s, comm_s, 0, r+4))
        if v9:
            p = clean_phone(v9)
            if p: phones.append((p,))

    con.executemany(
        'INSERT INTO slots(row_num,village,village_norm,arrond,arrond_norm,dept,commune,filled,last_row) VALUES(?,?,?,?,?,?,?,?,?)',
        slots)
    con.executemany('INSERT OR IGNORE INTO phones VALUES(?)', phones)
    con.executemany('INSERT OR REPLACE INTO arrond_last VALUES(?,?)',
                    list(arrond_rows.items()))
    con.execute('CREATE INDEX idx_vn ON slots(village_norm)')
    con.execute('CREATE INDEX idx_an ON slots(arrond_norm)')
    con.commit()
    wb.close()

    elapsed = (datetime.now() - t0).total_seconds()
    n_slots  = con.execute('SELECT COUNT(*) FROM slots').fetchone()[0]
    n_phones = con.execute('SELECT COUNT(*) FROM phones').fetchone()[0]
    print(f"   ✅ Index prêt en {elapsed:.1f}s — {n_slots} villages, {n_phones} tél.", flush=True)
    return con

# ── Column detection ──────────────────────────────────────────────────────────

ALIASES = {
    'quartier':       ['quartier','village','localite','localité','quartier/village'],
    'nom':            ['nom','noms','name'],
    'prenom':         ['prenom','prénom','prenoms','prénoms','firstname'],
    'telephone':      ['telephone','téléphone','tel','tél','phone','mobile','contact',
                       'numéro de téléphone','numero de telephone','numéro','numero'],
    'partis':         ['partis','parti','party'],
    'profession':     ['profession','métier','metier','emploi'],
    'date_naissance': ['date de naissance','date_naissance','naissance'],
    'lieu_naissance': ['lieu de naissance','lieu_naissance'],
    'adresse':        ['adresse','adresse complete','adresse complète'],
}

def detect_columns(ws_in):
    for r in range(1, 8):
        mapping = {}
        for c in range(1, ws_in.max_column + 1):
            h = normalize(str(ws_in.cell(row=r, column=c).value or ''))
            for field, aliases in ALIASES.items():
                if h in aliases:
                    mapping[field] = c
        if 'nom' in mapping and 'prenom' in mapping:
            return mapping, r
    return {}, None

# ── Slot finder with fuzzy matching ──────────────────────────────────────────

def find_slot_db(con, quartier):
    key = normalize(quartier)
    # Exact
    row = con.execute(
        'SELECT id,row_num,village,filled,last_row,arrond_norm FROM slots WHERE village_norm=? LIMIT 1',
        (key,)).fetchone()
    if row:
        return row
    # Fuzzy
    all_keys = [r[0] for r in con.execute('SELECT DISTINCT village_norm FROM slots').fetchall()]
    matches  = difflib.get_close_matches(key, all_keys, n=1, cutoff=0.85)
    if matches:
        row = con.execute(
            'SELECT id,row_num,village,filled,last_row,arrond_norm FROM slots WHERE village_norm=? LIMIT 1',
            (matches[0],)).fetchone()
        return row
    return None

# ── Excel row styling ─────────────────────────────────────────────────────────

def style_data_row(ws, rn, bg=C_WHITE):
    ws.row_dimensions[rn].height = 15
    for c in range(1, NCOLS + 1):
        cell = ws.cell(row=rn, column=c)
        cell.fill      = PatternFill('solid', start_color=bg)
        cell.font      = Font(name='Arial', size=10)
        cell.border    = BDR
        cell.alignment = Alignment(vertical='center')
    ws.cell(row=rn, column=1).alignment = Alignment(horizontal='center', vertical='center')

def write_person(ws, target_row, person):
    ws.cell(row=target_row, column=3).value = person['nom']
    ws.cell(row=target_row, column=4).value = person['prenom']
    ws.cell(row=target_row, column=5).value = person.get('partis', '')
    ws.cell(row=target_row, column=6).value = person.get('profession', '')
    ws.cell(row=target_row, column=7).value = person.get('date_naissance', '')
    ws.cell(row=target_row, column=8).value = person.get('lieu_naissance', '')
    ws.cell(row=target_row, column=9).value = person['telephone']

# ── Core integrate function ───────────────────────────────────────────────────

def integrate_file(filepath, filename):
    wb_in = load_workbook(filepath, data_only=True)
    ws_in = wb_in.active

    col_map, header_row = detect_columns(ws_in)
    if not col_map:
        return {'filename': filename, 'error': 'Colonnes non détectées (Nom + Prénom requis)'}

    def get(r, field):
        col = col_map.get(field)
        return str(ws_in.cell(row=r, column=col).value or '').strip() if col else ''

    # ── Détection du format double-colonne (colonnes décalées) ──────────────
    # Certains fichiers ont 2 jeux de colonnes décalés d'1 vers la droite
    # Ex: Nom col5 + Prénom col7 + Tél col12  ET  Nom col6 + Prénom col8 + Tél col13
    # On détecte en vérifiant si des lignes de données utilisent col+1
    def detect_alt_cols(col_map, ws_in, header_row):
        """Retourne un col_map alternatif décalé de +1 si ce format existe."""
        alt = {}
        for field, col in col_map.items():
            alt[field] = col + 1
        # Vérifier que le format alternatif existe dans les données
        sample_rows = 0
        alt_rows    = 0
        for r in range(header_row + 1, min(header_row + 50, ws_in.max_row + 1)):
            nom_main = str(ws_in.cell(row=r, column=col_map.get('nom', 99)).value or '').strip()
            nom_alt  = str(ws_in.cell(row=r, column=alt.get('nom', 99)).value or '').strip()
            if nom_main: sample_rows += 1
            if nom_alt and not nom_main: alt_rows += 1
        if alt_rows > 0:
            return alt
        return None

    alt_map = detect_alt_cols(col_map, ws_in, header_row)

    def get_row_data(r):
        """Extrait les données d'une ligne, en essayant le format principal puis alternatif."""
        def gv(cm, field):
            col = cm.get(field)
            return str(ws_in.cell(row=r, column=col).value or '').strip() if col else ''

        def find_tel(r, base_col):
            """Cherche le téléphone dans col, col+1, col-1, col+2."""
            if not base_col: return ''
            for delta in [0, 1, -1, 2]:
                v = str(ws_in.cell(row=r, column=base_col+delta).value or '').strip()
                if v and re.search(r'\d{5,}', v.replace(' ','')):
                    return v
            return ''

        # Format principal
        quartier = gv(col_map, 'quartier')
        # Quartier peut aussi être en col+1 (ex: col5 au lieu col4)
        if not quartier:
            col_q = col_map.get('quartier', 4)
            v_next = str(ws_in.cell(row=r, column=col_q+1).value or '').strip()
            skip_vals = {'N°','NOM','NOMS','PRENOM','PRENOMS','PROFESSION','ADRESSE COMPLETE','CONTACT','QUARTIER'}
            if v_next and v_next.upper() not in skip_vals and not re.match(r'^\d+$', v_next):
                quartier = v_next

        nom     = gv(col_map, 'nom').strip()
        prenom  = gv(col_map, 'prenom').strip()
        # Téléphone: chercher dans col_tel et colonnes adjacentes
        tel_col = col_map.get('telephone') or col_map.get('adresse')
        tel_raw = find_tel(r, tel_col)

        # Format alternatif décalé (si détecté et ligne principale vide)
        if alt_map and not nom:
            q2  = str(ws_in.cell(row=r, column=alt_map.get('quartier',99)).value or '').strip()
            n2  = str(ws_in.cell(row=r, column=alt_map.get('nom',99)).value or '').strip()
            p2  = str(ws_in.cell(row=r, column=alt_map.get('prenom',99)).value or '').strip()
            if n2:
                if q2 and not quartier: quartier = q2
                nom     = n2
                prenom  = p2
                if not tel_raw:
                    alt_tel_col = alt_map.get('telephone') or alt_map.get('adresse')
                    tel_raw = find_tel(r, alt_tel_col)

        return quartier, nom, prenom, tel_raw, {
            'partis':         gv(col_map, 'partis'),
            'profession':     gv(col_map, 'profession'),
            'date_naissance': gv(col_map, 'date_naissance'),
            'lieu_naissance': gv(col_map, 'lieu_naissance'),
        }
    # ────────────────────────────────────────────────────────────────────────

    # Collect all valid persons first (fast pass, no Excel writes)
    persons     = []
    rejected    = []
    seen_phones = set()
    last_quartier = None

    for r in range(1, ws_in.max_row + 1):

        # ── REGLE 1 : Ligne "QUARTIER : NOM" ─────────────────────────────
        for c in range(1, min(ws_in.max_column + 1, 6)):
            cell_val = str(ws_in.cell(row=r, column=c).value or '').strip()
            m = re.match(r'^QUARTIER\s*[:\-]\s*(.+)', cell_val, re.IGNORECASE)
            if m:
                last_quartier = m.group(1).strip()
                break
            if re.match(r'^QUARTIER\s*[:\-]?\s*$', cell_val, re.IGNORECASE):
                nxt = str(ws_in.cell(row=r, column=c+1).value or '').strip()
                if nxt: last_quartier = nxt
                break
        # ─────────────────────────────────────────────────────────────────

        if r < header_row + 1:
            continue

        quartier, nom, prenom, tel_raw, extras = get_row_data(r)
        tel = clean_phone(tel_raw)

        # Ignorer lignes vides
        if not any([quartier, nom, prenom, tel_raw]): continue
        # Ignorer en-têtes répétés
        if normalize(nom) in ['nom','noms','name']: continue
        if normalize(prenom) in ['prenom','prénom','prenoms','prénoms']: continue
        # Ignorer lignes QUARTIER: (déjà traitées)
        first_cell = str(ws_in.cell(row=r, column=1).value or '').strip()
        if re.match(r'^QUARTIER\s*[:\-]', first_cell, re.IGNORECASE): continue
        # Ignorer lignes de structure
        skip = False
        for c in range(1, min(ws_in.max_column+1, 4)):
            cv = str(ws_in.cell(row=r, column=c).value or '').strip().upper()
            if cv.startswith('DEPARTEMENT') or 'ARRONDISSEMENT' in cv or cv.startswith('COMMUNE'):
                skip = True; break
        if skip or not nom: continue

        # ── REGLE 2 : Propagation du quartier ────────────────────────────
        if quartier:
            last_quartier = quartier
        elif not quartier and nom and (prenom or tel):
            if last_quartier:
                quartier = last_quartier
        # ─────────────────────────────────────────────────────────────────

        missing = []
        if not quartier: missing.append('Quartier')
        if not nom:      missing.append('Nom')
        if not prenom:   missing.append('Prénom')
        if not tel:      missing.append('Téléphone')
        if missing:
            rejected.append({'row': r, 'quartier': quartier, 'nom': nom,
                             'reason': f"Manquant : {', '.join(missing)}"})
            continue
        if tel in seen_phones:
            rejected.append({'row': r, 'quartier': quartier, 'nom': nom,
                             'reason': f"Doublon interne : {tel_raw}"})
            continue
        seen_phones.add(tel)
        persons.append({'row': r, 'quartier': quartier, 'nom': nom, 'prenom': prenom,
                        'partis':         extras['partis'],
                        'profession':     extras['profession'],
                        'date_naissance': extras['date_naissance'],
                        'lieu_naissance': extras['lieu_naissance'],
                        'telephone': tel_raw, 'tel_clean': tel})

    # Now do all Excel writes in one lock
    with write_lock:
        con      = get_db()
        wb_out   = load_workbook(MERE_PATH)
        ws_out   = wb_out.active

        # Track rows inserted in this session (to adjust row numbers)
        row_offset = 0   # cumulative shift due to insert_rows()

        inserted = overflow = new_quartier = duplicates = 0

        # Build in-session row mapping adjustments
        insertions = []  # list of (target_row, person, bg)

        for p in persons:
            tel = p['tel_clean']

            # Check phone not already in DB
            exists = con.execute('SELECT 1 FROM phones WHERE phone=?', (tel,)).fetchone()
            if exists:
                duplicates += 1
                rejected.append({'row': p['row'], 'quartier': p['quartier'], 'nom': p['nom'],
                                 'reason': f"Doublon tél : {p['telephone']}"})
                continue

            slot = find_slot_db(con, p['quartier'])

            if slot:
                sid, row_num, village, filled, last_row, arrond_n = slot
                actual_row_num = row_num + row_offset
                actual_last    = last_row + row_offset

                if filled < 5:
                    # Use existing pre-formatted slot row
                    target_row    = actual_row_num + filled
                    overflow_flag = False
                else:
                    # Capacity exceeded → insert a brand new row
                    target_row = actual_last + 1
                    ws_out.insert_rows(target_row)
                    style_data_row(ws_out, target_row, C_NEW)
                    ws_out.cell(row=target_row, column=1).value = filled + 1
                    ws_out.cell(row=target_row, column=2).value = village
                    row_offset += 1
                    overflow_flag = True
                    overflow += 1

                write_person(ws_out, target_row, p)

                # Update DB slot
                con.execute('UPDATE slots SET filled=?, last_row=? WHERE id=?',
                            (filled + 1, last_row + (1 if overflow_flag else 0), sid))
                inserted += 1

            else:
                # New quartier — find arrond context from incoming file
                arrond_ctx = None
                for rr in range(p['row'], 0, -1):
                    v = str(ws_in.cell(row=rr, column=1).value or '').strip()
                    if 'ARRONDISSEMENT' in v.upper():
                        arrond_ctx = normalize(v)
                        break

                # Find last row of that arrond in DB
                arr_last = con.execute('SELECT last_row FROM arrond_last WHERE arrond_norm=?',
                                       (arrond_ctx,)).fetchone()
                if arr_last:
                    insert_after = arr_last[0] + row_offset
                else:
                    insert_after = ws_out.max_row

                new_row = insert_after + 1
                ws_out.insert_rows(new_row)
                style_data_row(ws_out, new_row, C_NEW)
                ws_out.cell(row=new_row, column=1).value = 1
                ws_out.cell(row=new_row, column=2).value = p['quartier']
                write_person(ws_out, new_row, p)
                row_offset  += 1
                new_quartier += 1
                inserted     += 1

                # Register in DB
                vn  = normalize(p['quartier'])
                an  = arrond_ctx or 'inconnu'
                con.execute(
                    'INSERT INTO slots(row_num,village,village_norm,arrond,arrond_norm,dept,commune,filled,last_row) VALUES(?,?,?,?,?,?,?,?,?)',
                    (new_row, p['quartier'], vn, an, an, '', '', 1, new_row))
                if arrond_ctx:
                    con.execute('INSERT OR REPLACE INTO arrond_last VALUES(?,?)',
                                (arrond_ctx, new_row))

            # Register phone
            con.execute('INSERT OR IGNORE INTO phones VALUES(?)', (tel,))

        con.commit()
        wb_out.save(MERE_PATH)

    # Update log
    log = load_log()
    log['total_inserted'] = log.get('total_inserted', 0) + inserted
    log['sessions'].append({
        'file':          filename,
        'date':          datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'inserted':      inserted,
        'rejected':      len(rejected),
        'duplicates':    duplicates,
        'overflow':      overflow,
        'new_quartiers': new_quartier,
    })
    save_log(log)

    # Count filled rows for stats
    con2 = get_db()
    total_personnes = con2.execute('SELECT SUM(filled) FROM slots').fetchone()[0] or 0

    return {
        'filename':        filename,
        'inserted':        inserted,
        'overflow':        overflow,
        'new_quartiers':   new_quartier,
        'duplicates':      duplicates,
        'rejected_count':  len(rejected),
        'rejected':        rejected[:50],
        'total_mere':      ws_out.max_row,
        'total_personnes': total_personnes,
        'total_cumule':    log['total_inserted'],
    }

# ── Routes ────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    files = request.files.getlist('files')
    if not files:
        return jsonify({'error': 'Aucun fichier reçu'}), 400

    results = []
    for f in files:
        if not f.filename.lower().endswith(('.xlsx', '.xls')):
            results.append({'filename': f.filename, 'error': 'Format non supporté (.xlsx/.xls)'})
            continue
        fname    = secure_filename(f.filename)
        filepath = os.path.join(UPLOAD_DIR, fname)
        f.save(filepath)
        try:
            result = integrate_file(filepath, f.filename)
            results.append(result)
        except Exception as e:
            results.append({'filename': f.filename, 'error': str(e)})
        finally:
            try: os.remove(filepath)
            except: pass

    return jsonify({'results': results, 'log': load_log()})

@app.route('/download')
def download():
    return send_file(MERE_PATH, as_attachment=True,
                     download_name='Villages_Quartiers_Benin_Final.xlsx')

@app.route('/stats')
def stats():
    log = load_log()
    con = get_db()
    total_personnes = con.execute('SELECT SUM(filled) FROM slots').fetchone()[0] or 0
    total_slots     = con.execute('SELECT COUNT(*) FROM slots').fetchone()[0]
    wb  = load_workbook(MERE_PATH, read_only=True)
    ws  = wb.active
    total_lignes = ws.max_row
    wb.close()
    return jsonify({
        'total_lignes_mere':   total_lignes,
        'total_personnes':     total_personnes,
        'total_insere_cumule': log.get('total_inserted', 0),
        'nb_sessions':         len(log.get('sessions', [])),
        'sessions':            log.get('sessions', [])[-10:],
    })

@app.route('/reset', methods=['POST'])
def reset():
    with write_lock:
        wb = load_workbook(MERE_PATH)
        ws = wb.active
        count = 0
        for r in range(1, ws.max_row + 1):
            if str(ws.cell(row=r, column=1).value or '').strip() in ['1','2','3','4','5']:
                for c in range(3, 10):
                    ws.cell(row=r, column=c).value = None
                count += 1
        wb.save(MERE_PATH)
        save_log({'sessions': [], 'total_inserted': 0})
        invalidate_db()
    return jsonify({'message': f'{count} lignes réinitialisées'})

# ── Startup ───────────────────────────────────────────────────────────────────

print("\n" + "="*55)
print("  🚀 Villages & Quartiers Bénin — Système v3")
print("  ➜  http://localhost:5000")
print("="*55)

# Verify mere.xlsx exists at startup
if not os.path.exists(MERE_PATH):
    print(f"❌ ERREUR : fichier mère introuvable : {MERE_PATH}")
    print(f"   Fichiers présents : {os.listdir(BASE_DIR)}")
else:
    print(f"✅ Fichier mère trouvé : {MERE_PATH}")
    # Pre-build index at startup
    try:
        with app.app_context():
            get_db()
    except Exception as e:
        print(f"⚠️  Index non construit au démarrage : {e}")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
