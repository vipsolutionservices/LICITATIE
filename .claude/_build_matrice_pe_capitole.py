"""
Genereaza matrice_conformitate_pe_capitole.docx (v2).

Modificari fata de v1:
 - Eliminata coloana "Cap. CdS" (redundanta cu heading)
 - Heading capitol include si NUMELE capitolului (extras din rândurile-sectiune ale tabelului sursa)
 - Coloana 3 = "Raspuns ofertant — mod indeplinire" (label conform anexa F sursa)
 - Intro extins inspirat din anexa F sursa (titlu + descriere + mod utilizare)
 - Stiluri font / dimensiuni aliniate cu sursa (16/14/11/10 pt)

Surse:
 - anexa_f_conformitate - old.docx — cerinte + nume capitole (rândurile-sectiune)
 - .claude/build_matrice.py — MAP cap -> produs (Sinteza xlsx, cu VOGO TECHNOLOGY -> <LIDER>)
"""
import re
import sys
from collections import defaultdict
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

from docx import Document
from docx.shared import Pt

ROOT = Path(r"C:\Users\adria\Documents\Claude\Projects\Licitatie 1\doc-output\oferta")
SRC_REQ = ROOT / "anexa_f_conformitate - old.docx"
OUT = ROOT / "matrice_conformitate_pe_capitole.docx"


MAP = {
    '3.4':         '<LIDER> (coordonare integrator)',
    '3.4.1.1':     '<LIDER> (arhitect sistem)',
    '3.4.1.2':     'Toate produsele C. Securitate (WAF / Honeypot / NAC / SIEM / Email / NGFW)',
    '3.4.1.3':     'Toate produsele C. Securitate + <LIDER> (conformitate legala)',
    '3.4.2.1':     'FURNIZOR LIMS COTS',
    '3.4.2.2':     'ZIPPER',
    '3.4.2.3':     'ZIPPER / VOGO Enterprise Suite',
    '3.4.2.5':     'Microsoft Power BI / SSRS / SSAS + Microsoft SSIS',
    '3.4.2.6':     'VOGO Enterprise Suite',
    '3.4.2.7':     'VOGO Enterprise Suite',
    '3.4.2.8':     'FURNIZOR GIS + VOGO Enterprise Suite',
    '3.4.2.9':     'VOGO Enterprise Suite',
    '3.4.2.10':    'VOGO Enterprise Suite',
    '3.4.2.11':    'VOGO Enterprise Suite (aplicatie mobila)',
    '3.4.2.12':    'VOGO Enterprise Suite',
    '3.4.2.14':    'Oracle Service Bus + Mirth Connect + VOGO Enterprise Suite',
    '3.4.3':       '<LIDER> (arhitect sistem)',
    '3.4.3.1':     'RHEL 9 / Oracle Linux 9 + Win Server 2022 DC (Cloud Guvernamental)',
    '3.4.3.2.1':   'NGINX Plus',
    '3.4.3.2.2':   'Microsoft IIS',
    '3.4.3.2.3':   'RHEL 9 / Oracle Linux 9 + Win Server 2022 DC',
    '3.4.3.2.4':   'Microsoft SQL Server Enterprise + Elasticsearch',
    '3.4.3.2.5':   'Microsoft SSIS',
    '3.4.3.2.6':   'Microsoft Power BI / SSRS / SSAS',
    '3.4.3.2.7':   'Keycloak Enterprise',
    '3.4.3.2.8':   'FURNIZOR GIS',
    '3.4.3.2.9':   'Oracle Service Bus',
    '3.4.3.3.1':   'ZIPPER',
    '3.4.3.3.1.1': 'ZIPPER',
    '3.4.3.3.1.2': 'ZIPPER',
    '3.4.3.3.1.3': 'ZIPPER',
    '3.4.3.3.1.4': 'ZIPPER',
    '3.4.3.3.1.5': 'ZIPPER',
    '3.4.3.3.1.6': 'ZIPPER',
    '3.4.3.3.1.7': 'ZIPPER',
    '3.4.3.3.2':   'VOGO Enterprise Suite',
    '3.4.3.3.2.1': 'VOGO Enterprise Suite (chatbot)',
    '3.4.3.3.2.2': 'VOGO Enterprise Suite (aplicatie mobila)',
    '3.4.3.3.3':   'FURNIZOR LIMS COTS',
    '3.4.3.3.3.1': 'FURNIZOR LIMS COTS',
    '3.4.3.3.3.2': 'Mirth Connect',
    '3.4.3.3.4':   '<LIDER>',
    '3.4.3.4':     'Toate produsele C. Securitate',
    '3.4.3.4.1.1': 'F5 / Imperva / FortiWeb (WAF)',
    '3.4.3.4.1.2': 'FortiDeceptor (Honeypot)',
    '3.4.3.4.1.3': 'Cisco DNA / ClearPass (NMS / NAC)',
    '3.4.3.4.1.4': 'Splunk ES / QRadar (SIEM)',
    '3.4.3.4.1.5': 'Cisco IronPort (Email Security)',
    '3.4.3.4.2.1': 'FortiGate / Palo Alto + FortiGate 100F locatii',
    '3.4.3.4.2.2': 'Echipament hardware (Switch acces) - vezi Lista_Hardware',
    '3.4.3.4.2.3': 'Echipament hardware (Switch POE) - vezi Lista_Hardware',
    '3.4.3.4.2.4': 'Echipament hardware (Access point) - vezi Lista_Hardware',
    '3.4.3.4.2.5': 'Echipament hardware (Switch agregare) - vezi Lista_Hardware',
    '3.4.3.4.2.6': 'Echipament hardware (Laptop) + MS Office H&B 2024 OEM',
    '3.4.3.4.2.7': 'Echipament hardware (Complete teren) + MS Office H&B 2024 OEM',
    '3.4.4':       '<LIDER> (PM + servicii)',
    '3.4.4.1':     '<LIDER> (Manager de proiect)',
    '3.4.4.2':     '<LIDER> (livrare/instalare/configurare)',
    '3.4.4.3':     '<LIDER> (Analist business + Arhitect)',
    '3.4.4.4':     '<LIDER> (Arhitect sistem + Team leader)',
    '3.4.4.5':     '<LIDER> (Team leader + experti dezvoltare)',
    '3.4.4.6':     '<LIDER> (Expert migrare)',
    '3.4.4.7':     '<LIDER> (echipa de implementare)',
    '3.4.4.8':     '<LIDER> (Expert testare)',
    '3.4.4.9':     '<LIDER> (Experti instruire)',
    '3.4.4.10':    '<LIDER> (echipa de punere in productie)',
    '3.4.5':       'Keycloak Enterprise + VOGO Enterprise Suite',
    '3.4.6':       'Toate produsele C. Securitate',
    '3.4.7':       'Toate produsele C. Securitate + <LIDER> (GDPR)',
    '3.4.8':       'Echipament hardware (Laptop EU Ecolabel) + furnizori ambalaje/livrare',
    '3.4.9':       '<LIDER> (Coordonator suport tehnic)',
}


def get_produs(cap):
    cap = (cap or '').strip()
    if cap in MAP:
        return MAP[cap]
    parts = cap.split('.')
    while len(parts) > 1:
        parts.pop()
        candidate = '.'.join(parts)
        if candidate in MAP:
            return MAP[candidate]
    return '<LIDER> (de validat)'


def cap_key(cap):
    parts = cap.split('.')
    out = []
    for p in parts:
        try:
            out.append(int(p))
        except ValueError:
            out.append(0)
    return tuple(out)


def normalize_cap(cap):
    m = re.search(r'(3\.4(?:\.\d+)*)', cap or '')
    if m:
        return m.group(1)
    return (cap or '').strip()


def parse_source():
    """Citeste cerinte + numele capitolelor.

    Returneaza:
     - cerinte: list[dict(nr, cap, cerinta)]
     - cap_names: dict[cap_normalizat] -> "Cap. 3.4.X — Nume Capitol" (cu whitespace normalizat)
    """
    doc = Document(str(SRC_REQ))
    t = doc.tables[0]
    cerinte = []
    cap_names = {}
    for row in t.rows:
        cells = row.cells
        texts = [c.text.strip() for c in cells]
        # Header sectiune: toate celulele au acelasi text si incepe cu "Cap."
        if len(cells) >= 2 and texts[0] == texts[1] and texts[0].startswith('Cap.'):
            full = re.sub(r'\s+', ' ', texts[0]).strip()
            # Extrag "Cap. 3.4.X.Y" + restul
            m = re.match(r'^Cap\.\s+(3\.4(?:\.\d+)*)\s*[—-]\s*(.+)$', full)
            if m:
                cap_norm = m.group(1)
                cap_names[cap_norm] = full
            continue
        # Header tabel
        if texts[0] == 'Nr.':
            continue
        # Cerinta
        if len(texts) < 3:
            continue
        nr = texts[0]
        cap = texts[1]
        cerinta = texts[2]
        if not nr.isdigit():
            continue
        cerinte.append({'nr': nr, 'cap': cap, 'cerinta': cerinta})
    return cerinte, cap_names


def set_run_default(run, size=10):
    """Setează font Calibri și dimensiunea."""
    run.font.size = Pt(size)


def add_par(doc, text, size=10, bold=False, space_before=0, space_after=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text)
    r.font.size = Pt(size)
    if bold:
        r.bold = True
    return p


def main():
    if not SRC_REQ.exists():
        raise SystemExit(f"NOT FOUND: {SRC_REQ}")
    print(f"Citesc cerinte din {SRC_REQ.name}...")
    cerinte, cap_names = parse_source()
    print(f"Cerinte gasite: {len(cerinte)}")
    print(f"Capitole cu nume identificate: {len(cap_names)}")

    by_cap = defaultdict(list)
    for c in cerinte:
        cap = normalize_cap(c['cap'])
        by_cap[cap].append(c)
    sorted_caps = sorted(by_cap.keys(), key=cap_key)
    print(f"Capitole distincte cu cerinte: {len(sorted_caps)}")

    doc = Document()

    # ============================================================
    # INTRO — inspirat din anexa F sursa
    # ============================================================
    add_par(doc, "MATRICEA DE CONFORMITATE — VARIANTĂ PE CAPITOLE",
            size=16, bold=True, space_before=0, space_after=6)
    add_par(doc, "Cerințe SIDISVA grupate pe capitole CdS, cu produsul recomandat pre-completat",
            size=14, bold=True, space_after=8)
    add_par(doc, f"SIDISVA — CN1089237 — Total cerințe: {len(cerinte)} "
                 f"(extrase automat din cap. 3.4.x.x al Caietului de Sarcini nr. 7574/CP/2025) "
                 f"— grupate pe {len(sorted_caps)} capitole distincte.",
            size=10, space_after=12)

    add_par(doc, "Mod de utilizare:", size=11, bold=True, space_before=8, space_after=4)
    add_par(doc,
        "• Pentru fiecare cerință se completează coloana \"Răspuns ofertant — mod îndeplinire\" "
        "cu descrierea concretă a modului de îndeplinire (NU \"OK\" / \"conform\" / \"soluția răspunde\"). "
        "Coloana este pre-completată automat cu produsul / componenta / echipa responsabilă "
        "din partea consorțiului condus de <LIDER>, conform mapării din "
        "Lista_Software_SIDISVA.xlsx, sheet \"Sinteza\", coloana \"Produs recomandat (oferta)\".",
        size=10, space_after=4)
    add_par(doc,
        "• Fiecare responsabil va detalia răspunsul în propunerea tehnică, indicând în clar "
        "modul în care produsul propus îndeplinește cerința (configurare, funcționalitate nativă, "
        "dezvoltare adițională etc.).",
        size=10, space_after=4)
    add_par(doc,
        "• Răspunsuri prin simpla repetare a cerinței (cu schimbarea timpului verbal) NU sunt acceptate.",
        size=10, space_after=4)
    add_par(doc,
        "• Hyperlink-uri către documentația producătorului fără citarea textului concret NU sunt acceptate.",
        size=10, space_after=4)
    add_par(doc,
        "• Lipsa răspunsurilor sau răspunsuri incomplete pot conduce la declararea ofertei ca neconformă.",
        size=10, space_after=12)

    add_par(doc,
        f"Document de uz intern, generat automat. Variantă \"pe capitole\" pentru distribuție "
        f"către responsabili (fiecare furnizor completează doar răspunsurile pentru cerințele "
        f"alocate prin coloana de produs). Pentru matricea completă cu toate cerințele într-un "
        f"singur tabel (formularul tip Anexă F), vezi anexa_f_conformitate.docx.",
        size=10, space_after=18)

    # ============================================================
    # PER CAPITOL: heading cu nume + tabel 3 coloane
    # ============================================================
    for cap in sorted_caps:
        # Heading capitol: "Capitolul X.Y — Nume"
        full_name = cap_names.get(cap)
        if full_name:
            # full_name = "Cap. 3.4.2.6 — Funcționalități..." → schimb "Cap." cu "Capitolul"
            heading_text = full_name.replace("Cap. ", "Capitolul ", 1)
        else:
            heading_text = f"Capitolul {cap}"

        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(20)
        h.paragraph_format.space_after = Pt(6)
        r = h.add_run(heading_text)
        r.bold = True
        r.font.size = Pt(13)

        # Tabel cerinte: 3 coloane (Nr / Cerinta / Raspuns)
        rows = by_cap[cap]
        t = doc.add_table(rows=1 + len(rows), cols=3)
        try:
            t.style = "Table Grid"
        except KeyError:
            pass

        hdr_titles = ["Nr.", "Cerință (citată din caietul de sarcini)",
                      "Răspuns ofertant — mod îndeplinire"]
        for i, h_text in enumerate(hdr_titles):
            cell = t.rows[0].cells[i]
            cell.text = ""
            p = cell.paragraphs[0]
            rr = p.add_run(h_text)
            rr.bold = True
            rr.font.size = Pt(10)

        for ri, c in enumerate(rows, start=1):
            cells = t.rows[ri].cells
            # Nr
            cells[0].text = ""
            r0 = cells[0].paragraphs[0].add_run(c['nr'])
            r0.font.size = Pt(9)
            # Cerinta
            cells[1].text = ""
            r1 = cells[1].paragraphs[0].add_run(c['cerinta'])
            r1.font.size = Pt(9)
            # Raspuns ofertant — mod indeplinire (= produs din Sinteza)
            produs = get_produs(normalize_cap(c['cap']))
            cells[2].text = ""
            r2 = cells[2].paragraphs[0].add_run(produs)
            r2.font.size = Pt(9)

        # Footer capitol (subțire)
        f = doc.add_paragraph()
        f.paragraph_format.space_before = Pt(4)
        f.paragraph_format.space_after = Pt(0)
        fr = f.add_run(f"({len(rows)} cerințe în acest capitol)")
        fr.italic = True
        fr.font.size = Pt(9)

    doc.save(str(OUT))
    size = OUT.stat().st_size
    print(f"\nOK | {OUT.name} salvat: {size} bytes ({size/1024:.1f} KB)")
    print(f"Total: {len(sorted_caps)} capitole, {len(cerinte)} cerinte.")
    # Verificare nume capitole
    missing_names = [c for c in sorted_caps if c not in cap_names]
    if missing_names:
        print(f"\n[INFO] Capitole fara nume identificat (folosim doar numarul):")
        for c in missing_names:
            print(f"  - {c}")


if __name__ == "__main__":
    main()
