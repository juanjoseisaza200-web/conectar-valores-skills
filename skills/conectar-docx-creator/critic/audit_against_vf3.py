"""
audit_against_vf3.py — Compara el XML de un .docx generado vs el ADN del VF3 real.

Uso:
    python audit_against_vf3.py <ruta_al_docx_a_auditar>

Salida: tabla de 8 checks críticos + score + lista de hallazgos.
"""
import sys, zipfile, re, json, os
from collections import Counter

# ADN de referencia (extraído del VF3 real)
REF_FONTS_ALLOWED = {'Arial'}
REF_COLORS_TEXT_ALLOWED = {'17375E', '5D9E9D'}  # Navy + Teal (raro)
REF_FILLS_ALLOWED = {'E7F0EF'}                    # Solo Mint
REF_BORDER_COLORS_ALLOWED = {'BFBFBF', 'C9A449', '17375E'}
REF_BORDER_SIZES_TYPICAL = {'2', '12'}            # 0.25pt y 1.5pt
REF_MARGINS_TWIPS = {'top': '1728', 'right': '1138', 'bottom': '1138', 'left': '1138'}
REF_SIZES_PT = {'18', '19', '17', '24', '20', '22', '32', '26'}  # half-points autorizados

FORBIDDEN_FILLS = {'EDEDED', '4472C4', '70AD47', 'FF0000', 'E67E22', '27AE60'}
FORBIDDEN_COLORS = {'FF0000', '4472C4', '70AD47'}


def inventory(path):
    with zipfile.ZipFile(path) as z:
        doc = z.read('word/document.xml').decode('utf-8')
        # Margins
        m = re.search(r'<w:pgMar([^/]*)/>', doc)
        margins = {}
        if m:
            for attr in ('top', 'right', 'bottom', 'left'):
                mt = re.search(rf'w:{attr}="(\d+)"', m.group(1))
                if mt: margins[attr] = mt.group(1)
        # Imágenes en media/
        media = [n for n in z.namelist() if n.startswith('word/media/')]

    # Conteo de tablas total + degeneradas
    total_tables = len(re.findall(r'<w:tbl[ >]', doc))
    # Tabla degenerada = <w:tbl>...</w:tbl> con menos de 2 <w:tc> dentro (1 celda total)
    # Aproximación: contar tablas con exactamente 1 tc
    tables_xml = re.findall(r'<w:tbl[ >].*?</w:tbl>', doc, flags=re.DOTALL)
    degenerate_tables = 0
    empty_text_tables = 0
    for t_xml in tables_xml:
        n_tc = len(re.findall(r'<w:tc[ >]', t_xml))
        # Extraer todo el texto de la tabla
        all_text = ' '.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', t_xml)).strip()
        if n_tc == 0:
            empty_text_tables += 1
        elif n_tc == 1 and not all_text:
            empty_text_tables += 1

    # Drawings (figuras inline)
    n_drawings = len(re.findall(r'<w:drawing[ >]', doc))

    return {
        'fonts': Counter(re.findall(r'w:ascii="([^"]+)"', doc)),
        'sizes': Counter(re.findall(r'<w:sz w:val="(\d+)"', doc)),
        'text_colors': Counter(re.findall(r'<w:color w:val="([0-9A-Fa-f]{6})"', doc)),
        'fills': Counter(re.findall(r'<w:shd[^>]*fill="([0-9A-Fa-f]{6})"', doc)),
        'border_colors': Counter(re.findall(r'<w:(?:top|bottom|left|right)[^/]*color="([0-9A-Fa-f]{6})"', doc)),
        'border_sizes': Counter(re.findall(r'<w:bottom[^/]*sz="(\d+)"', doc)),
        'alignments': Counter(re.findall(r'<w:jc w:val="(\w+)"', doc)),
        'margins': margins,
        'media_files': media,
        'n_drawings': n_drawings,
        'total_tables': total_tables,
        'empty_text_tables': empty_text_tables,
    }


def audit(path):
    inv = inventory(path)
    findings = []
    checks = {}

    # Check 1 — Solo Arial
    fonts_used = set(inv['fonts'].keys())
    non_arial = fonts_used - REF_FONTS_ALLOWED
    checks['1. Fonts'] = ('OK' if not non_arial else 'FAIL',
                          f"Usadas: {fonts_used}" + (f" | Prohibidas: {non_arial}" if non_arial else ""))
    if non_arial:
        findings.append(f"Fuentes no-Arial: {non_arial}")

    # Check 2 — 9pt dominante (sz=18 más frecuente entre tamaños comunes)
    sizes = inv['sizes']
    total_sizes = sum(sizes.values())
    pct_9pt = (sizes.get('18', 0) / total_sizes * 100) if total_sizes else 0
    checks['2. 9pt dominante'] = ('OK' if pct_9pt > 40 else 'FAIL',
                                   f"9pt = {pct_9pt:.0f}% de runs ({sizes.get('18',0)}/{total_sizes})")
    if pct_9pt <= 40:
        findings.append(f"9pt no dominante: {pct_9pt:.0f}% (esperado >40%)")
    # Tamaños prohibidos
    forbidden_sizes = set(sizes.keys()) - REF_SIZES_PT
    if forbidden_sizes:
        findings.append(f"Tamaños no autorizados: {forbidden_sizes}")

    # Check 3 — Colores texto solo Navy autorizados
    text_colors = set(inv['text_colors'].keys())
    forbidden_text = text_colors & FORBIDDEN_COLORS
    non_authorized = text_colors - REF_COLORS_TEXT_ALLOWED - {'000000', '4A5568'}
    checks['3. Colores texto'] = ('OK' if not forbidden_text and not non_authorized else 'FAIL',
                                   f"Usados: {dict(inv['text_colors'])}")
    if forbidden_text:
        findings.append(f"Colores texto prohibidos: {forbidden_text}")
    if non_authorized:
        findings.append(f"Colores texto no autorizados: {non_authorized}")

    # Check 4 — Fills SOLO Mint
    fills = set(inv['fills'].keys())
    forbidden_fills_used = fills & FORBIDDEN_FILLS
    extra_fills = fills - REF_FILLS_ALLOWED
    checks['4. Fills celdas'] = ('OK' if not extra_fills else 'FAIL',
                                  f"Usados: {dict(inv['fills'])}")
    if forbidden_fills_used:
        findings.append(f"Fills PROHIBIDOS encontrados: {forbidden_fills_used}")
    if extra_fills:
        findings.append(f"Fills no autorizados (solo E7F0EF Mint permitido): {extra_fills}")

    # Check 5 — Border colors
    bc = set(inv['border_colors'].keys())
    extra_bc = bc - REF_BORDER_COLORS_ALLOWED
    checks['5. Border colors'] = ('OK' if not extra_bc else 'FAIL',
                                   f"Usados: {dict(inv['border_colors'])}")
    if extra_bc:
        findings.append(f"Border colors no autorizados: {extra_bc}")

    # Check 6 — Border sizes (debe haber sz=2 y sz=12)
    bs = inv['border_sizes']
    has_hairline = bs.get('2', 0) > 0
    has_dorado = bs.get('12', 0) > 0
    # sz=4 sería ERROR (skill vieja)
    has_wrong = bs.get('4', 0) > 0
    checks['6. Border sizes'] = ('OK' if (has_hairline and has_dorado and not has_wrong)
                                  else 'FAIL',
                                  f"sz=2 (0.25pt): {bs.get('2',0)} | sz=12 (1.5pt): {bs.get('12',0)} | sz=4 (PROHIBIDO): {bs.get('4',0)}")
    if has_wrong:
        findings.append(f"Border sz=4 detectado — VF3 usa sz=2 (hairline) y sz=12 (dorado), NO sz=4")
    if not has_hairline:
        findings.append("Falta border sz=2 (hairline 0.25pt en filas cuerpo)")
    if not has_dorado:
        findings.append("Falta border sz=12 (dorado 1.5pt en header tabla)")

    # Check 7 — Center dominante (sin obligar — solo verificar que exista bastante)
    al = inv['alignments']
    center_count = al.get('center', 0)
    both_count = al.get('both', 0)
    total_al = sum(al.values())
    pct_center = (center_count / total_al * 100) if total_al else 0
    checks['7. Center dominante'] = ('OK' if pct_center > 50 else 'FAIL',
                                      f"Center: {pct_center:.0f}% ({center_count}/{total_al}) — VF3 real: 88%")
    if pct_center <= 50:
        findings.append(f"Center no dominante: {pct_center:.0f}% (VF3 real ~88%)")

    # Check 8 — Margenes exactos
    m = inv['margins']
    margins_ok = all(m.get(k) == v for k, v in REF_MARGINS_TWIPS.items())
    checks['8. Margenes'] = ('OK' if margins_ok else 'FAIL', f"Detectados: {m}")
    if not margins_ok:
        findings.append(f"Margenes incorrectos. Esperado: {REF_MARGINS_TWIPS} | Recibido: {m}")

    # Check 9 — Tablas válidas (sin tablas vacías sin texto)
    n_total = inv['total_tables']
    n_empty = inv['empty_text_tables']
    pct_empty = (n_empty / n_total * 100) if n_total else 0
    checks['9. Tablas validas'] = ('OK' if n_empty == 0 else 'FAIL',
                                    f"Total: {n_total} | Vacias sin texto: {n_empty}")
    if n_empty > 0:
        findings.append(f"Hay {n_empty} tablas sin contenido textual — usar callout o omitir")

    # Check 10 — Figuras con caption (heurístico: si hay drawings, debería haber captions)
    n_draw = inv['n_drawings']
    # No es FAIL si no hay drawings — solo informativo
    checks['10. Figuras'] = ('OK', f"Drawings inline: {n_draw} | Media files: {len(inv['media_files'])}")

    return checks, findings, inv


def print_report(checks, findings, inv, path):
    print(f"\n{'='*70}")
    print(f"AUDIT vs ADN VF3 — {os.path.basename(path)}")
    print('='*70)
    print(f"\n{'Check':40} {'Status':8} {'Detalle'}")
    print('-'*100)
    n_ok = 0
    for name, (status, detail) in checks.items():
        if status == 'OK': n_ok += 1
        print(f"{name:40} [{status:4}]  {detail}")
    print()
    print(f"SCORE: {n_ok}/{len(checks)} ({100*n_ok//len(checks)}%)")
    if findings:
        print("\nHALLAZGOS:")
        for f in findings:
            print(f"  - {f}")
    else:
        print("\nSin hallazgos. Formato VF3 reproducido fielmente.")
    print()
    if n_ok == len(checks):
        print("[PASA] Documento conforme al ADN VF3.")
    else:
        print("[FALLA] Hay desviaciones contra VF3 real.")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python audit_against_vf3.py <ruta.docx>")
        sys.exit(1)
    path = sys.argv[1]
    checks, findings, inv = audit(path)
    print_report(checks, findings, inv, path)
