# CLAUDE.md — Oferta SIDISVA (ANSVSA)

> **Scop:** Reguli operaționale pentru a scrie o ofertă **câștigătoare** la licitația SIDISVA.
> Acest fișier e încărcat automat la fiecare sesiune Claude Code în acest folder.
> Detaliile (cuprins caiet de sarcini, profile experți, formule punctaj) sunt în memorie — vezi `..\..\..\..\.claude\projects\C--Users-adria-Documents-Claude-Projects-Licitatie-1-doc-output-oferta\memory\`.

---

## 0. ⚠️ COORDONARE MULTI-AGENT — CITEȘTE ÎNAINTE DE ORICE EDIT

> În acest proiect lucrează **3 instanțe Claude Code în paralel** (cowork). Fără coordonare, agenții se suprapun și suprascriu reciproc munca.

### 0.1 Fișierul de coordonare: `.claude/_agents.md`

Înainte de **orice** Read/Edit/Write pe un fișier `.docx`, `.xlsx` sau `anexa_f_*` din folder:

1. **Citește `.claude/_agents.md`** — vezi tabelul de alocări curente.
2. **Verifică lock-urile Office** — dacă există `~$<nume>.docx` în folder, fișierul e deschis în Word de user/alt agent → **NU edita**.
3. **Dacă fișierul e marcat `IN_PROGRESS` de alt agent (alt alias)** → alege alt task sau întreabă userul. NU edita peste.
4. **Înainte de a începe** → adaugă/actualizează rând în tabel cu aliasul tău + `IN_PROGRESS` + timestamp ISO (`YYYY-MM-DDTHH:MM`).
5. **Când termini** → marchează `DONE <timestamp>` în coloana Status. Nu șterge rândul.
6. **Adaugă o linie în „Log evenimente"** la finalul fișierului cu ce ai făcut (1 propoziție).

### 0.2 Alias agent

La prima acțiune din sesiune, alege un alias scurt (`A1`, `A2`, `A3`, …) — verifică în `_agents.md` ce e deja folosit și ia primul liber. Aliasul rămâne pentru toată sesiunea.

### 0.3 Fișiere cu alocare exclusivă

- `Lista_Software_SIDISVA.xlsx` — un singur agent scrie la un moment dat.
- `anexa_f_conformitate.docx` (Matricea de conformitate, 1.294 cerințe) — un singur agent.
- Celelalte agenți doar citesc (read-only) cât e ocupat.

### 0.4 Modificări neașteptate

Dacă observi că un fișier alocat ție a fost modificat de altcineva între Read și Edit → **STOP, întreabă userul**, nu suprascrie.

### 0.5 Placeholder-uri convenite

- `<LIDER>` — numele liderului asocierii (se completează la final). NU scrie nume de companie hardcodat în text (nu „VOGO TECHNOLOGY", nu „BITHAT", etc.) — folosește `<LIDER>` sau `<PARTENER>`.

### 0.6 Comanda „show agents"

Când userul scrie **`show agents`** (sau variații: „arată agenții", „status agenți"), citesc `.claude/_agents.md` și afișez **două tabele markdown native** (NU box-drawing — randarea Claude Code se strică din cauza diacriticelor și a wrap-ului).

**Format obligatoriu:**

```
**DONE — N task-uri finalizate**

| Agent | Fișier | Ce s-a făcut | Când |
|---|---|---|---|
| A1 | nume.docx | sinteză 1 propoziție din Note | 01:30 |
| A2 | … | … | … |

**IN_PROGRESS — M task active acum**

| Agent | Fișier | Ce face |
|---|---|---|
| A2 | … | … (start HH:MM) |
```

**Reguli:**
- Sursa de date: STRICT `.claude/_agents.md`, tabelul „Alocări curente". Nu inventez task-uri.
- Sortare DONE: după agent (A1, A2, A3, A4, apoi „—" pre-sesiune).
- „Când": `HH:MM` dacă e azi (după data curentă), altfel `DD lună HH:MM`.
- „Ce s-a făcut" / „Ce face": 1 propoziție scurtă, factuală, fără citate din coloana Note. Dacă e prea lung → tai la ~120 caractere și pun „…".
- Status atipic (DRAFT, `?`) → îl pun în IN_PROGRESS cu nota „status: DRAFT/incert".
- NU afișez tabelul „Log evenimente" — doar „Alocări curente".
- După tabele, nicio explicație suplimentară decât dacă userul cere — comanda e read-only.
- NU afișez bloc `SINTEZA EXECUTIE` (comanda nu modifică fișiere).

### 0.7 Prompt aliniere agenți (broadcast)

Când userul cere prompt-ul de aliniere sau detectez fișiere modificate fără rând în `_agents.md`, afișez **literal** blocul de mai jos:

```
[ALINIERE — paste în fiecare sesiune Claude Code din proiectul SIDISVA]

Citește CLAUDE.md §0 și .claude/_agents.md. Alege alias liber (A1/A2/A3/A4).

Înainte de orice Edit pe .docx/.xlsx:
1. Re-citește .claude/_agents.md.
2. Skip dacă există ~$<nume>.docx (lock Office) sau alt alias e IN_PROGRESS.
3. Adaugă rând: | <fișier> | <alias> | IN_PROGRESS | <YYYY-MM-DDTHH:MM> | <intenție 1 prop> |

Când termini:
4. Schimbă status în DONE <timestamp> + sinteză factuală în Note (KB înainte→după, nr secțiuni/tabele).
5. Adaugă rând în „Log evenimente".

Reguli stricte:
- NU nume hardcodate (VOGO TECHNOLOGY/BITHAT/ZIPPER) → `<LIDER>`/`<PARTENER>`/`<SUITA>`/`<FURNIZOR_X>`.
- NU batch fără rând în Alocări.
- Lista_Software_SIDISVA.xlsx + anexa_f_conformitate.docx = alocare exclusivă.
- Modificare neașteptată pe fișier alocat ție → ASK USER.
- Română, formal juridic-tehnic, referințe explicite la cap. CdS.
```

---

## 1. Context minim de știut înainte de orice acțiune

- **Beneficiar:** ANSVSA + 3 institute subordonate (IISPV, ICBMV, IDSA). Cod SMIS 336342, finanțat **POCIDIF**.
- **Buget estimat contract:** 85.418.857,53 lei fără TVA (plafon eligibil 95.27M lei).
- **Constrângeri buget:** max **20% HW + servicii instalare/config**, min **10% securitate cibernetică**.
- **Găzduire obligatorie:** **Cloud Guvernamental**, arhitectură **Cloud-Native**, **containerizare** (Docker/K8s).

## 1.1 Hartă fișiere — unde caut ce (CITEȘTE ÎNTÂI)

**Folder curent = oferta noastră, în lucru.** Fiecare `.docx` numerotat e o secțiune din oferta noastră finală pe care o livrăm către ANSVSA. Aici scriu/editez.

| Folder / fișier | Rol | Cum îl folosesc |
|---|---|---|
| `.\` (folderul curent: `doc-output\oferta\`) | **OFERTA NOASTRĂ în lucru** — fișierele `.docx` numerotate sunt secțiunile pe capitole. Aici editez. | Sursa adevărului pentru ce livrăm. Vezi tabelul de la §3. |
| `..\..\doc-input\Caiet_de_sarcini_7574_30.12.2025_SIDISVA_semnat.docx` | **CERINȚELE OBLIGATORII** (252 pag, semnat de ANSVSA) | Sursa adevărului pentru ce trebuie să acoperim. Read-only. |
| `..\..\doc-input\_extracted\caiet_sarcini.txt` | Text plain extras din caietul de sarcini (4426 linii) | Pentru grep/Read din Claude. **AICI caut cerințe specifice** când scriu o secțiune. |
| `..\..\doc-input\Oferta tehnica V3 rev AB.docx` | **OFERTĂ MODEL/EXEMPLU** — fostă ofertă BITHAT+ZIPPER pentru un alt proiect (doar DMS / Documenta Capture). **NU este oferta noastră, NU o copiez direct.** | Doar **inspirație de stil/structură** pentru secțiunile generice (PM, metodologie, suport, acceptanță). Acoperă ~30% din scope-ul SIDISVA. Pentru LIMS/GIS/BI/Portal/SNIIA/mobile/integrări = NU ajută, scriu de la zero. |
| `..\..\doc-input\_extracted\oferta_model_outline.txt` | Outline ofertă model (heading-uri) | Pentru a vedea structura modelului fără a deschide fișierul de 14.8 MB. |
| `..\..\doc-input\Arhitectura Campina.docx` | Material de referință (nedeschis încă) | Verifică conținutul dacă devine relevant. |
| `..\..\doc-input\CN1089237.zip` + `CN1089237_extracted*\` | Arhivă procedură SEAP | Verifică doar dacă userul cere fișa de date sau formularele oficiale. |
| `..\..\doc-input\_extracted\extract_docx.py` | Script python-docx pentru extragere `.docx` → text | `python extract_docx.py <in.docx> <out.txt> [--outline-only]` |
| `~\.claude\projects\C--Users-adria-Documents-Claude-Projects-Licitatie-1-doc-output-oferta\memory\` | **Memoria persistentă** (proiect, evaluare, experți, strategie) | Citesc `MEMORY.md` la început de sesiune; actualizez la fiecare lecție nouă. |

**Regulă fermă:**
> **Cerințele se citesc DIN caietul de sarcini** (sau `_extracted\caiet_sarcini.txt`).
> **Conținutul ofertei se scrie ÎN folderul curent** (`doc-output\oferta\*.docx`).
> **Oferta model BITHAT** este DOAR sursă de inspirație stilistică pentru secțiunile generice — NU sursa de cerințe și NU sursa de copy-paste.

---

## 2. Cum câștigăm punctele — reguli ferme pentru fiecare secțiune

### P1 — Preț (40p): formula `(P_min / P_n) × 40`
- **Nu sub-bugetăm** — preț prea mic ⇒ risc de ofertă suspect de mică. Ținta: marjă de profit decentă, nu cea mai mică ofertă cu orice preț.
- **Respectă cele două plafoane procentuale** (HW ≤20%, securitate ≥10%). Verifică LA FIECARE iterație a propunerii financiare.

### P2 — Experți (30p): 6 experți cheie × 5p
- Pentru fiecare din: **Manager proiect, Analist business, Arhitect sistem, Team leader software, Expert analiză/optimizare procese, Expert administrare BD** → vrem **5+ proiecte/contracte** demonstrabile.
- Proiectele trebuie să fie pentru **sisteme informatice cu min 7 module interconectate + cel puțin 1 modul de tip portal** (excepție: Expert analiză/optimizare procese nu cere 7 module/portal).
- 2 proiecte = 0p (doar cerință minimă). Sub 5 proiecte = punctaj parțial. **Țintă: 5+ proiecte per expert** = 30p max.
- Cerință suplimentară: Arhitect și DBA necesită **certificare de la producătorul** soluției DMS/SGBD ofertate. Ceilalți: certificări recunoscute internațional (PMP/PRINCE2 pt PM, Lean Six Sigma pt #5, Scrum/Agile + cloud pt Team leader, etc.).

### P3 — Metodologie (20p): **2 subfactori × 10p, calificative Acceptabil/Adecvat/Excepțional = 1/5/10p**

Pentru calificativul **Excepțional** (10p) la **subfactor 3.1**, propunerea TREBUIE să bifeze TOATE 4 elementele:
1. Metodologii/metode/instrumente testate, recunoscute — **prezentate în detaliu**
2. Adaptare la specificul contractului SIDISVA, **corelat cu drumul critic + riscuri/ipoteze identificate**
3. **Abordare inovatoare și eficientă** — modalități concrete de îmbunătățire a rezultatelor (nu generalități)
4. Demonstrare concretă a **modului de asigurare a securității informatice și informaționale** (nu doar mențiuni)

Pentru **Excepțional la 3.2** (10p): demonstrabil că **resursele materiale + umane sunt sincronizate complet** cu activitățile și generează rezultatele la calitatea descrisă.

⚠️ Lipsa oricăruia din cele 4 elemente la 3.1 ⇒ Adecvat (5p) — pierdem 5p.

### P4 — DNSH (10p):
- **4.1 Laptop consum veghe (5p):** Energy Star + cerință minimă 20Wh; sub asta primim puncte. Ofertăm laptopuri **EU Ecolabel sau echivalent** și atașăm rapoarte testare + fișe tehnice producător.
- **4.2 Ambalaje + livrare (5p):** ambalaje **reciclabile/reutilizabile** + livrare cu **emisii reduse** (electric/hibrid). Trebuie demonstrat documentar.

### DEMO video — ELIMINATORIE (33 cerințe)
**Orice cerință nedemonstrată ⇒ ofertă neconformă, eliminată din evaluare.** Vezi memoria `project-sidisva-evaluation` pentru lista completă. Cele mai grele de bifat:
- Integrare nativă **ROeID + eIDAS** + 2FA SMS
- **Chatbot AI** care asistă la configurare formulare drag-and-drop
- **Lucru offline** cu sincronizare la reconectare
- **Multi-tenant** pe aceeași instalare cu izolare totală a datelor între domenii
- API documentat **Swagger** + execuție teste funcționale Restful
- Edit Word/Excel direct în DMS prin Office (fără reatașare)
- Tipărire **plicuri + borderouri cu coduri bare** + gestionare confirmări primire
- Modificare diagramă **BPMN/UML** direct din UI

Înainte de orice promisiune în propunere, **verific că DEMO-ul poate demonstra concret cerința**. Dacă nu, NU promit — risc de descalificare.

---

## 3. Reguli de scriere pentru fiecare secțiune `.docx`

| Fișier | Conținut țintă | Aduce puncte la |
|---|---|---|
| `1a-oferta.docx` | Scrisoare de înaintare, sumar | — |
| `1-Rezumat_executiv.docx` | Sinteza ofertei: companii, scor target, valoare propusă | impresie generală |
| `1A-Prezentare_ofertant.docx` | CV companii, asocieri, capacitate financiară/tehnică | eligibilitate |
| `2-Abordare_metodologie.docx` | **SECȚIUNEA P3 — 20p**. Cele 4 elemente Excepțional + securitate informatică concretă | P3.1 (10p) |
| `3-Plan_implementare.docx` | Gantt, drumul critic, milestone-uri, paralelizare | P3.1 (drum critic) + P3.2 |
| `4-Descrierea_solutiei.docx` | Toate 14 componente SIDISVA + cum se mapează pe cerințe | conformitate |
| `5-Arhitectura_si_licente.docx` | Arhitectură Cloud-Native + containerizare + lista licențe + IP (codul sursă) | conformitate + IP |
| `6-Lista_hardware.docx` | Laptopuri DNSH compliant + HW server | P4.1 + plafon 20% |
| `7-Plan_garantie.docx` | SLA suport, perioada garanție, integrare ulterioară mock-ups | conformitate |
| `8-Conformitate_specificatii.docx` | Matricea de conformitate (răspuns cerință-cu-cerință din caietul de sarcini) | eligibilitate |
| `9-Echipa_proiect.docx` | **SECȚIUNEA P2 — 30p**. CV + certificări + 5+ proiecte/expert | P2 (30p) |
| `10-Securitate_informatica.docx` | NIS2 + OUG 155/2024 + Lege 354/2022 + GDPR + plafon min 10% | P3.1 (sec) + plafon |
| `11-DNSH.docx` | **SECȚIUNEA P4 — 10p**. Laptop sub Energy Star + ambalaje + livrare verde | P4 (10p) |
| `12-Plan_instruire.docx` | Plan instruire ANSVSA + DSVSA + medici vet (~5300 angajați) | conformitate |
| `13-Management_contract.docx` | Raportare (inițial / trimestrial / final), acceptanță, KPIs | P3.2 + conformitate |
| `14-DEMO_video.docx` | **Scenariul video DEMO — 33 cerințe demonstrate** | eliminatoriu |
| `15-Declaratii_obligatorii.docx` | Declarații standard + acceptare drepturi IP (cap 12) | eligibilitate |
| `16-Anexe.docx` | Anexe, certificări, documente suport | suport |
| `anexa_f_conformitate.docx` | Anexă F (separată de #8) | conformitate |
| `Lista_Software_SIDISVA.xlsx` | Lista licențe + producători + cantități | P5 + plafon HW |

**Default writing rules:**
- Toate documentele în **limba română**.
- Limbaj formal, juridic-tehnic, fără marketing-speak.
- Referință explicită la articolul din caietul de sarcini la fiecare cerință abordată (ex: "conform cerinței din cap. 3.4.2.6").
- Tabelele se construiesc în Word direct (nu screenshot-uri).
- Acolo unde caietul cere "ofertantul demonstrează" / "ofertantul prezintă" — **nu lăsa generic**, scrie concret.

---

## 4. Erori obișnuite de evitat

1. **Copy-paste din oferta model BITHAT** fără adaptare — modelul e doar DMS, scope-ul SIDISVA e 14 componente. Comisia detectează imediat.
2. **A promite în propunere ce nu poate fi demonstrat în DEMO** ⇒ descalificare la verificare.
3. **A ignora plafoanele 20% HW / 10% securitate** ⇒ ofertă neconformă financiar.
4. **A scrie generic la securitate** ⇒ pierdem 5p la P3.1 (calificativ Adecvat în loc de Excepțional).
5. **A propune experți cu 2-3 proiecte** ⇒ punctaj sub max la P2 (pierdere de 2-15p).
6. **A nu integra cu ROeID/eIDAS/PNI/ONRC/APIA/Ghiseul.ro** — sunt cerințe ferme.
7. **A propune găzduire altundeva decât Cloud Guvernamental** ⇒ neconform.
8. **A folosi software cu origine din state neconforme cu Lege 354/2022** (context RU/UA) — verifică producătorii antivirus și soluțiilor critice.

---

## 5. Workflow pentru fiecare task de scriere

1. **Citește memoria relevantă** (`MEMORY.md` → fișierele linkate). În special pentru secțiunile cu punctaj direct.
2. **Identifică cerințele exacte** din caietul de sarcini (grep pe `_extracted\caiet_sarcini.txt` cu termenii cheie).
3. **Scrie draft** în format markdown (mai ușor de iterat). Conversia în `.docx` se face cu pandoc sau direct prin python-docx la final.
4. **Cross-check cu factori de evaluare** — dacă secțiunea aduce puncte, verifică explicit fiecare element din algoritm.
5. **Salvează lecții învățate** în memorie după fiecare iterație care a generat feedback de la userul Vogo/cowork.

---

## 6. Instrumente disponibile

- **Python 3.11 + python-docx 1.2.0** (instalat global) — pentru manipulare `.docx`.
- **Script extract:** `..\..\doc-input\_extracted\extract_docx.py` (suportă `--outline-only`).
- **Repo git** — fiecare modificare pe un commit clar (nu commit-uri tip "wip").
- Co-working cu **claude cowork** — userul lucrează în paralel cu altă instanță Claude. Dacă observ modificări neașteptate în fișiere, întreb înainte de a edita peste.

---

## 7. Limbă & ton

- Răspund userului în **română**.
- Updates concise — 1-2 propoziții pe acțiune.
- La fiecare commit/edit semnificativ pe fișiere `.docx`, raportez ce s-a schimbat (pentru claude cowork să poată sincroniza).

### 7.1 Sinteză execuție — obligatoriu după fiecare task finalizat

După finalizarea oricărui task care a modificat un fișier (`.docx`, `.xlsx`, `.md`, script etc.), afișez o **sinteză de execuție** în formatul:

```
====> SINTEZA EXECUTIE: <ce am făcut, 1 propoziție concretă>
> în documentul: <cale relativă sau nume fișier>
```

Reguli:
- Dacă am atins mai multe fișiere → câte un bloc `====> SINTEZA EXECUTIE` per fișier, în ordinea modificărilor.
- Sinteza vine **ÎNAINTE** de orice altă explicație sau follow-up — e prima linie de output după task done.
- Conținutul sintezei e factual (ce am scris/șters/modificat), nu intenție sau planificare.
- Dacă task-ul nu a modificat niciun fișier (doar Read/verificare) → NU afișez sinteză.
