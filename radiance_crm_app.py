#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Radiance CRM — Client Loyalty Extractor
GUI Application wrapping DOCX extraction + Odoo CSV export pipeline
Author: Ramzi Lbs | Data Analyst Portfolio Project
"""

import os
import re
import csv
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from datetime import datetime

# ============================================================================
# DEPENDENCY CHECK
# ============================================================================

MISSING_DEPS = []
try:
    from docx2python import docx2python
except ImportError:
    MISSING_DEPS.append("docx2python")

try:
    import pandas as pd
except ImportError:
    MISSING_DEPS.append("pandas")

try:
    from rapidfuzz import fuzz
except ImportError:
    MISSING_DEPS.append("rapidfuzz")

try:
    import dateparser
except ImportError:
    MISSING_DEPS.append("dateparser")


# ============================================================================
# CORE LOGIC (from main.py + main2.py)
# ============================================================================

def normalize_phone(phone_str: str) -> str:
    if not phone_str:
        return ""
    digits = re.sub(r'\D', '', str(phone_str))
    return digits if len(digits) >= 9 else ""


def extract_phone_from_text(text: str):
    phone_patterns = [
        r'0\d{9}',
        r'0\d{2}[\s\.-]?\d{2}[\s\.-]?\d{2}[\s\.-]?\d{2}[\s\.-]?\d{2}',
        r'\+?\d{2,4}[\s\.-]?\d{2,4}[\s\.-]?\d{2,4}[\s\.-]?\d{2,4}',
        r'\d{9,}',
    ]
    phone = ""
    clean_text = text
    for pattern in phone_patterns:
        match = re.search(pattern, text)
        if match:
            phone_candidate = normalize_phone(match.group(0))
            if phone_candidate:
                phone = phone_candidate
                clean_text = re.sub(pattern, '', text).strip()
                break
    return clean_text, phone


def is_valid_name(text: str) -> bool:
    if not text or not isinstance(text, str):
        return False
    text = text.strip()
    if not text:
        return False
    if re.match(r'^[\d\s/\-\.]+$', text):
        return False
    if re.match(r'^\d{1,2}[/\-\.]\d{1,2}', text):
        return False
    if len(text.strip()) < 2:
        return False
    if re.search(r'[a-zA-ZÀ-ÿ]', text):
        return True
    return False


def parse_name(text: str):
    if not text or not isinstance(text, str):
        return "", "", ""
    text = text.strip()
    if not is_valid_name(text):
        return "", "", ""
    text_without_phone, phone = extract_phone_from_text(text)
    text_without_phone = re.sub(r'[^\w\s\-]', ' ', text_without_phone)
    text_without_phone = re.sub(r'\s+', ' ', text_without_phone).strip()
    if not text_without_phone:
        return "", "", phone
    if not is_valid_name(text_without_phone):
        return "", "", ""
    parts = text_without_phone.split()
    if len(parts) == 0:
        return "", "", phone
    elif len(parts) == 1:
        return parts[0].title(), "", phone
    else:
        return parts[0].title(), parts[1].title(), phone


def flatten_cell_content(cell_data) -> str:
    if isinstance(cell_data, str):
        return cell_data.strip()
    elif isinstance(cell_data, list):
        result = []
        for item in cell_data:
            flattened = flatten_cell_content(item)
            if flattened:
                result.append(flattened)
        return ' '.join(result)
    else:
        return str(cell_data).strip() if cell_data else ""


def parse_dates(date_text: str):
    if not date_text or not isinstance(date_text, str):
        return []
    for sep in [',', ';', '\n', '\r']:
        date_text = date_text.replace(sep, '|')
    date_parts = [d.strip() for d in date_text.split('|') if d.strip()]
    dates = []
    for date_str in date_parts:
        try:
            if dateparser:
                parsed_date = dateparser.parse(
                    date_str,
                    settings={'DATE_ORDER': 'DMY', 'PREFER_DAY_OF_MONTH': 'first', 'STRICT_PARSING': False}
                )
                if parsed_date:
                    year = parsed_date.year
                    if 2000 <= year <= 2030:
                        dates.append(parsed_date.strftime('%d/%m/%Y'))
                    continue
        except Exception:
            pass
        match = re.search(r'(\d{1,2})[/\-\.](\d{1,2})(?:[/\-\.](\d{2,4}))?', date_str)
        if match:
            day, month = int(match.group(1)), int(match.group(2))
            year = int(match.group(3)) if match.group(3) else datetime.now().year
            if year < 100:
                year += 2000
            if 1 <= day <= 31 and 1 <= month <= 12 and 2000 <= year <= 2030:
                dates.append(f"{day:02d}/{month:02d}/{year}")
    return list(set(dates))


def extract_clients_from_docx(filepath: str, log_fn=None):
    clients = []
    try:
        doc = docx2python(filepath)
        filename = Path(filepath).name
        for body_section in doc.body:
            for table in body_section:
                if len(table) < 2:
                    continue
                header_row = table[0] if table else []
                header_texts = [flatten_cell_content(cell).lower() for cell in header_row]
                name_col = next((i for i, h in enumerate(header_texts) if any(k in h for k in ['nom', 'prénom', 'prenom', 'name', 'client'])), None)
                date_col = next((i for i, h in enumerate(header_texts) if any(k in h for k in ['date', 'séance', 'seance', 'rendez', 'rdv'])), None)
                phone_col = next((i for i, h in enumerate(header_texts) if any(k in h for k in ['tel', 'tél', 'phone', 'portable', 'mobile'])), None)
                if name_col is None:
                    continue
                for row in table[1:]:
                    if len(row) <= name_col:
                        continue
                    name_text = flatten_cell_content(row[name_col])
                    nom, prenom, phone_from_name = parse_name(name_text)
                    if not nom:
                        continue
                    if phone_col is not None and len(row) > phone_col:
                        phone_text = flatten_cell_content(row[phone_col])
                        phone = normalize_phone(phone_text) or phone_from_name
                    else:
                        phone = phone_from_name
                    dates = []
                    if date_col is not None and len(row) > date_col:
                        date_text = flatten_cell_content(row[date_col])
                        dates = parse_dates(date_text)
                    clients.append({
                        'nom': nom, 'prenom': prenom, 'telephone': phone,
                        'dates': dates, 'nb_seances': len(dates), 'source_file': filename
                    })
        if log_fn:
            log_fn(f"  ✓ {filename}: {len(clients)} clients trouvés")
    except Exception as e:
        if log_fn:
            log_fn(f"  ✗ Erreur {Path(filepath).name}: {e}")
    return clients


def process_all_docx_files(input_dir: str, log_fn=None):
    if not os.path.exists(input_dir):
        if log_fn:
            log_fn(f"✗ Dossier introuvable: {input_dir}")
        return []
    docx_files = list(Path(input_dir).glob("*.docx"))
    if not docx_files:
        if log_fn:
            log_fn("✗ Aucun fichier .docx trouvé dans le dossier")
        return []
    if log_fn:
        log_fn(f"📁 {len(docx_files)} fichier(s) DOCX trouvé(s)")
    all_clients = []
    for f in docx_files:
        clients = extract_clients_from_docx(str(f), log_fn)
        all_clients.extend(clients)
    if log_fn:
        log_fn(f"\n📊 Total brut: {len(all_clients)} entrées clients")
    return all_clients


def calculate_similarity(c1, c2) -> float:
    name1 = f"{c1['nom']} {c1['prenom']}".strip()
    name2 = f"{c2['nom']} {c2['prenom']}".strip()
    name_score = fuzz.ratio(name1.lower(), name2.lower())
    if c1['telephone'] and c2['telephone'] and c1['telephone'] == c2['telephone']:
        return max(name_score, 95.0)
    return name_score


def merge_duplicate_clients(clients, threshold=85, log_fn=None):
    if not clients:
        return []
    if log_fn:
        log_fn(f"\n🔄 Fusion des doublons (seuil: {threshold}%)...")
    merged = []
    processed = set()
    for i, client1 in enumerate(clients):
        if i in processed:
            continue
        merged_client = {
            'nom': client1['nom'], 'prenom': client1['prenom'],
            'telephone': client1['telephone'],
            'dates': set(client1['dates']),
            'source_files': {client1['source_file']}
        }
        for j, client2 in enumerate(clients[i+1:], start=i+1):
            if j in processed:
                continue
            if calculate_similarity(client1, client2) >= threshold:
                merged_client['dates'].update(client2['dates'])
                merged_client['source_files'].add(client2['source_file'])
                if len(client2['telephone']) > len(merged_client['telephone']):
                    merged_client['telephone'] = client2['telephone']
                processed.add(j)
        processed.add(i)
        merged_client['dates'] = sorted(list(merged_client['dates']))
        merged_client['nb_seances'] = len(merged_client['dates'])
        merged.append(merged_client)
    if log_fn:
        log_fn(f"   ✓ {len(clients)} → {len(merged)} clients après fusion")
    return merged


def filter_loyal_clients(clients, min_sessions=2, log_fn=None):
    loyal = [c for c in clients if c['nb_seances'] >= min_sessions]
    if log_fn:
        log_fn(f"\n✅ {len(loyal)} clients fidèles (≥{min_sessions} séances) / {len(clients)} total")
    return loyal


def export_standard_csv(clients, output_file: str, log_fn=None):
    if not clients:
        return
    clients_sorted = sorted(clients, key=lambda c: (-c['nb_seances'], c['nom'], c['prenom']))
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    with open(output_file, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=['Nom', 'Prénom', 'Téléphone', 'Nombre de séances'])
        writer.writeheader()
        for c in clients_sorted:
            writer.writerow({'Nom': c['nom'], 'Prénom': c['prenom'],
                             'Téléphone': c['telephone'], 'Nombre de séances': c['nb_seances']})
    if log_fn:
        log_fn(f"💾 CSV standard exporté: {output_file}")


def export_odoo_csv(clients, output_file: str, include_tags: bool = True, log_fn=None):
    if not clients:
        return
    clients_sorted = sorted(clients, key=lambda c: (-c['nb_seances'], c['nom'], c['prenom']))
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    odoo_data = []
    for c in clients_sorted:
        name = f"{c['prenom']} {c['nom']}".strip() if c['prenom'] else c['nom']
        row = {'Name': name, 'Phone': c['telephone'], 'Notes': f"Nombre de séances: {c['nb_seances']}"}
        if include_tags:
            row['Tags'] = 'Client Fidèle'
        odoo_data.append(row)
    fieldnames = ['Name', 'Phone', 'Tags', 'Notes'] if include_tags else ['Name', 'Phone', 'Notes']
    with open(output_file, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(odoo_data)
    if log_fn:
        log_fn(f"💾 CSV Odoo exporté: {output_file}")


# ============================================================================
# GUI APPLICATION
# ============================================================================

class RadianceCRMApp(tk.Tk):
    DARK_BG = "#0f1117"
    PANEL_BG = "#1a1d27"
    CARD_BG = "#1e2130"
    ACCENT = "#7c6af7"
    ACCENT2 = "#4fc3f7"
    SUCCESS = "#4caf97"
    WARNING = "#f5a623"
    ERROR = "#e05c5c"
    TEXT = "#e8eaf0"
    MUTED = "#8890a4"
    BORDER = "#2a2f45"

    def __init__(self):
        super().__init__()
        self.title("Radiance CRM — Client Loyalty Extractor")
        self.geometry("1050x720")
        self.minsize(900, 620)
        self.configure(bg=self.DARK_BG)
        self.resizable(True, True)

        # State
        self.input_dir = tk.StringVar()
        self.output_dir = tk.StringVar(value=str(Path.home() / "Desktop" / "radiance_crm"))
        self.min_sessions = tk.IntVar(value=2)
        self.fuzzy_threshold = tk.IntVar(value=85)
        self.include_tags = tk.BooleanVar(value=True)
        self.export_standard = tk.BooleanVar(value=True)
        self.export_odoo = tk.BooleanVar(value=True)
        self.is_running = False

        self._setup_styles()
        self._build_ui()

        if MISSING_DEPS:
            self._show_deps_warning()

    def _setup_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background=self.DARK_BG)
        style.configure("Card.TFrame", background=self.CARD_BG)
        style.configure("TLabel", background=self.DARK_BG, foreground=self.TEXT,
                         font=("Courier New", 10))
        style.configure("Heading.TLabel", background=self.DARK_BG, foreground=self.TEXT,
                         font=("Courier New", 13, "bold"))
        style.configure("Sub.TLabel", background=self.DARK_BG, foreground=self.MUTED,
                         font=("Courier New", 9))
        style.configure("Card.TLabel", background=self.CARD_BG, foreground=self.TEXT,
                         font=("Courier New", 10))
        style.configure("CardHead.TLabel", background=self.CARD_BG, foreground=self.ACCENT,
                         font=("Courier New", 10, "bold"))
        style.configure("TCheckbutton", background=self.CARD_BG, foreground=self.TEXT,
                         font=("Courier New", 10))
        style.configure("TSpinbox", fieldbackground=self.PANEL_BG, background=self.PANEL_BG,
                         foreground=self.TEXT, insertcolor=self.TEXT,
                         font=("Courier New", 10))
        style.configure("Horizontal.TProgressbar", background=self.ACCENT,
                         troughcolor=self.BORDER, bordercolor=self.BORDER,
                         lightcolor=self.ACCENT, darkcolor=self.ACCENT)
        style.map("TCheckbutton",
                  background=[("active", self.CARD_BG)],
                  foreground=[("active", self.ACCENT)])

    def _build_ui(self):
        # ── Top Bar ─────────────────────────────────────────────────────────
        topbar = tk.Frame(self, bg=self.PANEL_BG, height=60)
        topbar.pack(fill="x", side="top")
        topbar.pack_propagate(False)

        tk.Label(topbar, text="◈  RADIANCE CRM", bg=self.PANEL_BG,
                 fg=self.ACCENT, font=("Courier New", 15, "bold")).pack(side="left", padx=22, pady=14)
        tk.Label(topbar, text="Client Loyalty Extractor  ·  Odoo-Ready Export",
                 bg=self.PANEL_BG, fg=self.MUTED, font=("Courier New", 9)).pack(side="left", pady=14)
        tk.Label(topbar, text="v1.0", bg=self.PANEL_BG, fg=self.BORDER,
                 font=("Courier New", 9)).pack(side="right", padx=20)

        # ── Main Container ───────────────────────────────────────────────────
        main = tk.Frame(self, bg=self.DARK_BG)
        main.pack(fill="both", expand=True, padx=18, pady=14)
        main.columnconfigure(0, weight=1)
        main.columnconfigure(1, weight=2)
        main.rowconfigure(0, weight=1)

        # ── LEFT COLUMN: Config ──────────────────────────────────────────────
        left = tk.Frame(main, bg=self.DARK_BG)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        self._card_paths(left)
        self._card_params(left)
        self._card_export(left)
        self._run_button(left)

        # ── RIGHT COLUMN: Log + Stats ────────────────────────────────────────
        right = tk.Frame(main, bg=self.DARK_BG)
        right.grid(row=0, column=1, sticky="nsew")
        right.rowconfigure(0, weight=3)
        right.rowconfigure(1, weight=1)
        right.columnconfigure(0, weight=1)

        self._card_log(right)
        self._card_stats(right)

        # ── Bottom Bar ───────────────────────────────────────────────────────
        self.progress = ttk.Progressbar(self, orient="horizontal",
                                        mode="indeterminate", style="Horizontal.TProgressbar")
        self.progress.pack(fill="x", side="bottom", padx=18, pady=(0, 10))
        self.status_var = tk.StringVar(value="Prêt")
        tk.Label(self, textvariable=self.status_var, bg=self.DARK_BG,
                 fg=self.MUTED, font=("Courier New", 9), anchor="w").pack(
                     fill="x", side="bottom", padx=20, pady=(2, 0))

    def _make_card(self, parent, title, pady=(0, 12)):
        wrapper = tk.Frame(parent, bg=self.DARK_BG)
        wrapper.pack(fill="x", pady=pady)
        header = tk.Frame(wrapper, bg=self.DARK_BG)
        header.pack(fill="x", pady=(0, 4))
        tk.Label(header, text=f"  {title}", bg=self.ACCENT, fg=self.DARK_BG,
                 font=("Courier New", 9, "bold"), padx=6, pady=2).pack(side="left")
        card = tk.Frame(wrapper, bg=self.CARD_BG, bd=0,
                        highlightbackground=self.BORDER, highlightthickness=1)
        card.pack(fill="x")
        return card

    def _path_row(self, parent, label, var, is_file=False):
        row = tk.Frame(parent, bg=self.CARD_BG)
        row.pack(fill="x", padx=10, pady=5)
        tk.Label(row, text=label, bg=self.CARD_BG, fg=self.MUTED,
                 font=("Courier New", 8), width=10, anchor="w").pack(side="left")
        entry = tk.Entry(row, textvariable=var, bg=self.PANEL_BG, fg=self.TEXT,
                         insertbackground=self.TEXT, relief="flat",
                         font=("Courier New", 9), bd=0,
                         highlightbackground=self.BORDER, highlightthickness=1)
        entry.pack(side="left", fill="x", expand=True, ipady=4, padx=(4, 4))
        btn = tk.Button(row, text="…", bg=self.BORDER, fg=self.TEXT,
                        activebackground=self.ACCENT, activeforeground=self.DARK_BG,
                        relief="flat", font=("Courier New", 9), padx=6, pady=2, cursor="hand2",
                        command=lambda: self._browse(var, is_file))
        btn.pack(side="left")

    def _browse(self, var, is_file=False):
        if is_file:
            path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        else:
            path = filedialog.askdirectory()
        if path:
            var.set(path)

    def _card_paths(self, parent):
        card = self._make_card(parent, "CHEMINS")
        self._path_row(card, "Source DOCX", self.input_dir)
        self._path_row(card, "Sortie", self.output_dir)
        tk.Frame(card, bg=self.CARD_BG, height=6).pack()

    def _card_params(self, parent):
        card = self._make_card(parent, "PARAMÈTRES")
        for label, var, from_, to in [
            ("Séances min.", self.min_sessions, 1, 20),
            ("Seuil fuzzy %", self.fuzzy_threshold, 50, 100)
        ]:
            row = tk.Frame(card, bg=self.CARD_BG)
            row.pack(fill="x", padx=10, pady=5)
            tk.Label(row, text=label, bg=self.CARD_BG, fg=self.MUTED,
                     font=("Courier New", 8), width=14, anchor="w").pack(side="left")
            spin = ttk.Spinbox(row, from_=from_, to=to, textvariable=var,
                                width=5, style="TSpinbox")
            spin.pack(side="left")

        # Slider for min_sessions
        def update_spin_from_slider(val):
            self.min_sessions.set(int(float(val)))

        slider_row = tk.Frame(card, bg=self.CARD_BG)
        slider_row.pack(fill="x", padx=10, pady=(0, 8))
        tk.Label(slider_row, text="min séances →", bg=self.CARD_BG, fg=self.BORDER,
                 font=("Courier New", 7)).pack(side="left")
        slider = tk.Scale(slider_row, from_=1, to=10, orient="horizontal",
                          variable=self.min_sessions, bg=self.CARD_BG, fg=self.MUTED,
                          troughcolor=self.BORDER, activebackground=self.ACCENT,
                          highlightthickness=0, bd=0, sliderrelief="flat",
                          font=("Courier New", 8), showvalue=False, length=150)
        slider.pack(side="left", padx=4)
        tk.Frame(card, bg=self.CARD_BG, height=4).pack()

    def _card_export(self, parent):
        card = self._make_card(parent, "EXPORT")
        for text, var in [
            ("CSV Standard (clients_fideles.csv)", self.export_standard),
            ("CSV Odoo (clients_odoo.csv)", self.export_odoo),
            ("Inclure colonne Tags Odoo", self.include_tags),
        ]:
            cb = tk.Checkbutton(card, text=text, variable=var,
                                bg=self.CARD_BG, fg=self.TEXT,
                                selectcolor=self.ACCENT, activebackground=self.CARD_BG,
                                activeforeground=self.ACCENT, relief="flat",
                                font=("Courier New", 9), cursor="hand2")
            cb.pack(anchor="w", padx=12, pady=3)
        tk.Frame(card, bg=self.CARD_BG, height=6).pack()

    def _run_button(self, parent):
        self.run_btn = tk.Button(
            parent, text="▶  LANCER L'EXTRACTION",
            bg=self.ACCENT, fg=self.DARK_BG,
            activebackground=self.ACCENT2, activeforeground=self.DARK_BG,
            relief="flat", font=("Courier New", 11, "bold"),
            pady=10, cursor="hand2",
            command=self._start_pipeline
        )
        self.run_btn.pack(fill="x", pady=(6, 0))

    def _card_log(self, parent):
        card = tk.Frame(parent, bg=self.DARK_BG)
        card.grid(row=0, column=0, sticky="nsew", pady=(0, 8))
        card.rowconfigure(1, weight=1)
        card.columnconfigure(0, weight=1)

        header = tk.Frame(card, bg=self.DARK_BG)
        header.grid(row=0, column=0, sticky="ew")
        tk.Label(header, text="  JOURNAL", bg=self.ACCENT2, fg=self.DARK_BG,
                 font=("Courier New", 9, "bold"), padx=6, pady=2).pack(side="left")
        self.clear_btn = tk.Button(header, text="effacer", bg=self.DARK_BG,
                                   fg=self.MUTED, relief="flat",
                                   font=("Courier New", 8), cursor="hand2",
                                   command=self._clear_log)
        self.clear_btn.pack(side="right")

        log_frame = tk.Frame(card, bg=self.CARD_BG,
                             highlightbackground=self.BORDER, highlightthickness=1)
        log_frame.grid(row=1, column=0, sticky="nsew")
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        self.log_text = tk.Text(
            log_frame, bg=self.CARD_BG, fg=self.TEXT,
            font=("Courier New", 9), relief="flat", bd=0,
            insertbackground=self.TEXT, wrap="word",
            state="disabled", padx=10, pady=8
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")
        self.log_text.tag_configure("ok", foreground=self.SUCCESS)
        self.log_text.tag_configure("err", foreground=self.ERROR)
        self.log_text.tag_configure("warn", foreground=self.WARNING)
        self.log_text.tag_configure("info", foreground=self.ACCENT2)
        self.log_text.tag_configure("muted", foreground=self.MUTED)

        scrollbar = tk.Scrollbar(log_frame, bg=self.PANEL_BG, troughcolor=self.BORDER,
                                  activebackground=self.ACCENT, relief="flat", bd=0)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.configure(command=self.log_text.yview)

    def _card_stats(self, parent):
        card = tk.Frame(parent, bg=self.DARK_BG)
        card.grid(row=1, column=0, sticky="nsew")
        card.columnconfigure(0, weight=1)

        tk.Label(card, text="  RÉSULTATS", bg=self.SUCCESS, fg=self.DARK_BG,
                 font=("Courier New", 9, "bold"), padx=6, pady=2).pack(anchor="w", pady=(0, 4))

        stats_frame = tk.Frame(card, bg=self.CARD_BG,
                               highlightbackground=self.BORDER, highlightthickness=1)
        stats_frame.pack(fill="both", expand=True)

        self.stat_vars = {}
        labels = [
            ("total_extracted", "Clients extraits"),
            ("after_merge", "Après déduplication"),
            ("loyal", "Clients fidèles"),
            ("with_phone", "Avec téléphone"),
            ("files_processed", "Fichiers traités"),
        ]
        for key, label in labels:
            row = tk.Frame(stats_frame, bg=self.CARD_BG)
            row.pack(fill="x", padx=12, pady=3)
            tk.Label(row, text=label, bg=self.CARD_BG, fg=self.MUTED,
                     font=("Courier New", 9), anchor="w").pack(side="left")
            var = tk.StringVar(value="—")
            self.stat_vars[key] = var
            tk.Label(row, textvariable=var, bg=self.CARD_BG, fg=self.ACCENT,
                     font=("Courier New", 10, "bold"), anchor="e").pack(side="right")

    # ── Logging helpers ──────────────────────────────────────────────────────

    def _log(self, msg: str, tag=""):
        def _do():
            self.log_text.configure(state="normal")
            timestamp = datetime.now().strftime("%H:%M:%S")
            full = f"[{timestamp}] {msg}\n"
            if not tag:
                if msg.startswith("✅") or msg.startswith("✓"):
                    t = "ok"
                elif msg.startswith("✗") or msg.startswith("❌"):
                    t = "err"
                elif msg.startswith("⚠"):
                    t = "warn"
                elif msg.startswith("📁") or msg.startswith("🔄") or msg.startswith("📊"):
                    t = "info"
                elif msg.startswith("💾"):
                    t = "ok"
                else:
                    t = ""
            else:
                t = tag
            self.log_text.insert("end", full, t)
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        self.after(0, _do)

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _update_stat(self, key, value):
        def _do():
            if key in self.stat_vars:
                self.stat_vars[key].set(str(value))
        self.after(0, _do)

    def _set_status(self, msg):
        self.after(0, lambda: self.status_var.set(msg))

    # ── Pipeline ─────────────────────────────────────────────────────────────

    def _show_deps_warning(self):
        msg = "Dépendances manquantes:\n" + "\n".join(f"  • {d}" for d in MISSING_DEPS)
        msg += "\n\nInstallez avec:\npip install " + " ".join(MISSING_DEPS)
        messagebox.showwarning("Dépendances manquantes", msg)

    def _start_pipeline(self):
        if self.is_running:
            return
        if not self.input_dir.get():
            messagebox.showerror("Erreur", "Veuillez sélectionner le dossier source DOCX.")
            return
        if MISSING_DEPS:
            messagebox.showerror("Dépendances manquantes",
                                  "Installez: pip install " + " ".join(MISSING_DEPS))
            return
        self.is_running = True
        self.run_btn.configure(state="disabled", text="⏳  EN COURS...")
        self.progress.start(12)
        self._set_status("Extraction en cours...")
        for key in self.stat_vars:
            self.stat_vars[key].set("—")
        thread = threading.Thread(target=self._run_pipeline, daemon=True)
        thread.start()

    def _run_pipeline(self):
        try:
            self._log("=" * 48, "muted")
            self._log("🌟 DÉMARRAGE DE L'EXTRACTION", "info")
            self._log("=" * 48, "muted")

            input_dir = self.input_dir.get()
            output_dir = self.output_dir.get()
            min_sessions = self.min_sessions.get()
            fuzzy = self.fuzzy_threshold.get()

            # Step 1: Extract
            all_clients = process_all_docx_files(input_dir, self._log)
            if not all_clients:
                self._log("❌ Aucun client extrait.", "err")
                self._done(success=False)
                return

            self._update_stat("total_extracted", len(all_clients))
            docx_count = len(list(Path(input_dir).glob("*.docx")))
            self._update_stat("files_processed", docx_count)

            # Step 2: Merge
            merged = merge_duplicate_clients(all_clients, fuzzy, self._log)
            self._update_stat("after_merge", len(merged))

            # Step 3: Filter
            loyal = filter_loyal_clients(merged, min_sessions, self._log)
            self._update_stat("loyal", len(loyal))
            if not loyal:
                self._log("⚠️ Aucun client fidèle trouvé avec ce critère.", "warn")
                self._done(success=False)
                return

            with_phone = sum(1 for c in loyal if c['telephone'])
            self._update_stat("with_phone", f"{with_phone}/{len(loyal)}")

            # Step 4: Export
            os.makedirs(output_dir, exist_ok=True)

            if self.export_standard.get():
                std_path = os.path.join(output_dir, "clients_fideles.csv")
                export_standard_csv(loyal, std_path, self._log)

            if self.export_odoo.get():
                odoo_path = os.path.join(output_dir, "clients_odoo.csv")
                export_odoo_csv(loyal, odoo_path, self.include_tags.get(), self._log)

            self._log("=" * 48, "muted")
            self._log(f"✅ Traitement terminé — {len(loyal)} clients fidèles exportés", "ok")
            self._log(f"📁 Dossier: {output_dir}", "info")
            self._log("=" * 48, "muted")
            self._done(success=True)

        except Exception as e:
            self._log(f"❌ Erreur inattendue: {e}", "err")
            self._done(success=False)

    def _done(self, success=True):
        def _do():
            self.is_running = False
            self.progress.stop()
            self.run_btn.configure(state="normal", text="▶  LANCER L'EXTRACTION")
            if success:
                self._set_status("✓ Extraction terminée avec succès")
                self.run_btn.configure(bg=self.SUCCESS)
                self.after(2000, lambda: self.run_btn.configure(bg=self.ACCENT))
            else:
                self._set_status("✗ Extraction terminée avec des erreurs")
        self.after(0, _do)


# ============================================================================
# ENTRY POINT
# ============================================================================

if __name__ == "__main__":
    app = RadianceCRMApp()
    app.mainloop()
