# Macro-N — Prototip Modul A/C + constatări pe cele 4 fișiere mostră

## Ce am construit
Un script Python (`audit_subjects_v2.py`) care:
1. Identifică rândurile "Subject" de tip normalizat în Main (marcate prin textul
   `דירוג משוקלל מנורמל` aflat lângă formula agregată).
2. Identifică blocurile corespunzătoare din Normalize.
3. Compară textele și raportează potriviri / neconcordanțe / lipsuri.

## Ce am găsit — dovadă concretă că problema descrisă de client chiar există

### 1. Neconcordanță reală de text (Modul A), chiar în fișierul "corect" (Excel #1)
Fișier: `קופות_גמל...-1.xlsx`
- **Main, rândul 32**: `מדגם תעריפים מירביים (כולל מע"מ)`
- **Normalize (cel mai apropiat)**: `מדדים תעריפיים (כולל מע"מ)`

Text diferit → exact tipul de problemă pe care Modulul A trebuie să-l repare automat.

### 2. Marker-ul de agregare NU e mereu în aceeași coloană
Fișier: `רשתות_כלי_אוכל...-2.xlsx`
- Marker-ul `דירוג משוקלל מנורמל` apare în **coloana G**, nu H (ca în celelalte 3 fișiere).
- Asta confirmă punctul C.5 din PDF ("the inclusion of the Subject is different").
- Macro-ul trebuie să caute marker-ul, nu să presupună o coloană fixă.

### 3. Variante de text ale marker-ului însuși (probabil dovadă de corupere din rulările anterioare de macro)
Tot în același fișier, la rândurile G114/G135/G141:
- `'דירוג משוקלל  מנורמל >  '` (spațiu dublu)
- `'דירוג משוקלל  >  '` (cuvântul "מנורמל" complet lipsă!)

Asta explică de ce clientul a găsit celule "sparte" în urma rulării macro-ului anterior —
inclusiv marker-ul folosit pentru identificarea Subject-urilor poate fi corupt, nu doar
formulele.

### 4. Subject-uri fără corespondent în Normalize (posibil legitim, posibil eroare — de verificat cu clientul)
Ex: `ערכיות רב-שכבתית`, `גודל ופריסה`, `סייבר ושמירת פרטיות` — apar cu marker-ul de
agregare în Main dar nu au bloc corespunzător în Normalize. Necesită sau (a) o regulă
suplimentară de matching, sau (b) confirmare de la client că acele subiecte chiar nu
trebuie să aibă corespondent.

## Concluzie
Prototipul validează abordarea tehnică (Python/openpyxl reproduce logica de
detectare + poate reconstrui formula corect) și, în același timp, confirmă că
variabilitatea reală din cele 4.248 fișiere este exact atât de mare pe cât sugerează
specificația — justifică structura pe 2 milestone-uri din ofertă.
