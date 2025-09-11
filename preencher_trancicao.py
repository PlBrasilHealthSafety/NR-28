# -*- coding: utf-8 -*-
"""
Preenche a coluna "TRANSCRIÇÃO DO ITEM NORMATIVO" da planilha NR-28 Anexo II
lendo um PDF oficial da NR e extraindo exatamente os itens/alinéas/incisos
citados na coluna "FUNDAMENTAÇÃO LEGAL".

Como usar:
1) Ajuste CONFIG (PLANILHA_PATH, PDF_PATH, OUT_PATH, NR_NUMBER).
2) pip install: pandas openpyxl PyPDF2 pdfminer.six
3) Rode: python preencher_transcricao.py

Observações:
- O script NÃO separa por anexo; indexa o PDF inteiro e localiza
  cada item pelo número (ex.: 1.4.1, 1.5.3.2.1 etc.). Isso evita
  “perdas” por cortes de seção.
- Quando a referência tiver múltiplos itens (ex.: “1.5.3.2, alínea "b",
  1.5.4.3.1, alíneas "a", "b" e "c", e 1.5.4.3.2”), cada item é
  tratado com as suas próprias alíneas/incisos (sem “vazar”
  as letras de um item para outro).
"""

import re
import sys
import pandas as pd

# ========================== CONFIG ==========================
PLANILHA_PATH = r"C:\Users\RodrigoCinelliPLBras\Downloads\NR28_AnexoII_planilha_NR37_preenchida.xlsx"
PDF_PATH      = r"C:\Users\RodrigoCinelliPLBras\Downloads\nr-38-atualizada-2025-3.pdf"
OUT_PATH      = r"C:\Users\RodrigoCinelliPLBras\Downloads\NR28_AnexoII_planilha_NR38_preenchida.xlsx"
NR_NUMBER     = 38
# ============================================================


# ----------------------- Extração PDF -----------------------
def extrair_texto(pdf_path: str) -> str:
    """Extrai texto do PDF (pdfminer como primário, PyPDF2 como fallback)."""
    # 1) pdfminer (melhor qualidade de layout na maioria dos PDFs)
    try:
        from pdfminer.high_level import extract_text
        txt = extract_text(pdf_path)
        if txt and txt.strip():
            return txt
    except Exception:
        pass

    # 2) PyPDF2 (fallback simples e rápido)
    try:
        import PyPDF2
        with open(pdf_path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            pages = [p.extract_text() or "" for p in reader.pages]
        return "\n".join(pages)
    except Exception:
        return ""


def normalizar_texto(bruto: str) -> str:
    """Normaliza o texto sem destruir parágrafos: quebras, espaços, NBSP etc."""
    if not bruto:
        return ""
    t = bruto.replace("\r", "")
    # Quebras “estranhas”
    t = t.replace("\x0c", "\n").replace("\x0b", "\n")
    # Espaços não quebrantes e afins
    t = (t.replace("\u00A0", " ")
           .replace("\u2009", " ")
           .replace("\u2002", " ")
           .replace("\u2003", " "))
    # Colapsa espaços repetidos
    t = re.sub(r"[ \t]+", " ", t)
    # Limita blocos gigantes de linhas vazias
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t


# ------------------- Indexador de itens ---------------------
# Aceita até 7 níveis: 1, 1.2, 1.2.3, 1.2.3.4.5.6.7
_ITEM_BLOCK_RE = re.compile(
    r"(?m)^\s*(\d+(?:\.\d+){0,7})\b(.*?)(?=(?:^\s*\d+(?:\.\d+){0,7}\b)|\Z)",
    flags=re.DOTALL
)

def indexar_itens(norm: str) -> dict:
    """Indexa blocos por cabeçalhos numéricos ex.: 1, 1.4, 1.4.1, 1.5.3.2.1 etc."""
    items = {}
    for m in _ITEM_BLOCK_RE.finditer(norm):
        num = m.group(1).strip()
        block = (m.group(1) + m.group(2)).strip()
        items[num] = block
    return items


# ---------- Ajudantes para alíneas / incisos ----------
# alíneas: "a) ..." (aceita variações a )  / a) - / a) –)
_ALINEA_HEAD_RE = re.compile(
    r"(?m)^\s*([a-z])\)\s*(?:-|–)?\s+",
    flags=re.IGNORECASE
)

# incisos: "I. ..."   "I) ..."   "I - ..."   "I – ..."
_INCISO_HEAD_RE = re.compile(
    r"(?m)^\s*([IVXLCDM]+)[\.\)\-–]\s+",
    flags=re.IGNORECASE
)

def split_alineas(block: str) -> dict:
    """Retorna {'a': 'texto...', 'b': 'texto...'} a partir de 'a) ...', 'b) ...' etc."""
    alineas = {}
    pos = [(m.start(), m.group(1).lower()) for m in _ALINEA_HEAD_RE.finditer(block)]
    if not pos:
        return alineas
    pos.append((len(block), None))
    for i in range(len(pos)-1):
        start, key = pos[i]
        end, _ = pos[i+1]
        seg = block[start:end].strip()
        seg = _ALINEA_HEAD_RE.sub("", seg, count=1)
        alineas[key] = seg.strip()
    return alineas

def split_incisos(texto: str) -> dict:
    """Retorna {'I': '...', 'II': '...'} a partir de 'I. ...', 'II) ...', 'III - ...' etc."""
    incisos = {}
    pos = [(m.start(), m.group(1).upper()) for m in _INCISO_HEAD_RE.finditer(texto)]
    if not pos:
        return incisos
    pos.append((len(texto), None))
    for i in range(len(pos)-1):
        start, key = pos[i]
        end, _ = pos[i+1]
        seg = texto[start:end].strip()
        seg = _INCISO_HEAD_RE.sub("", seg, count=1)
        incisos[key] = seg.strip()
    return incisos


# --------------- Parser de referências (por item) ---------------
_NUM_PAT = r'(\d+(?:\.\d+){0,7})'

def _letters_from_blob(blob: str) -> list:
    """Extrai letras de um trecho onde podem aparecer: alínea(s) e aspas “a”, “b”, …"""
    letters_cmd = re.findall(
        r'al[ií]neas?\s+((?:"[a-zA-Z]"\s*(?:,|e)?\s*)+)|al[ií]nea\s+"?([a-zA-Z])"?',
        blob, flags=re.IGNORECASE
    )
    out = []
    if letters_cmd:
        joined = " ".join(g1 or g2 for g1, g2 in letters_cmd)
        out = re.findall(r'["“]?([a-zA-Z])["”]?', joined)
    else:
        out = re.findall(r'["“]([a-zA-Z])["”]', blob)
    return sorted(set(l.lower() for l in out))

def _romans_from_blob(blob: str) -> list:
    romans = re.findall(r'\b([IVXLCDM]+)\b', blob)
    out = []
    for r in romans:
        u = r.upper()
        if u not in out:
            out.append(u)
    return out

def parse_ref_segments(ref: str):
    """
    Divide a FUNDAMENTAÇÃO em segmentos (um por item numérico),
    trazendo as alíneas/incisos que APARECEM no mesmo segmento.
    Também suporta letras **antes** do primeiro item (serão aplicadas ao 1º).
    """
    item_pat = re.compile(rf'{_NUM_PAT}(.*?)(?={_NUM_PAT}|\Z)', re.DOTALL)
    it = list(item_pat.finditer(ref))

    # letras soltas antes do primeiro item
    carry_letters = []
    if it:
        pre = ref[:it[0].start()]
        carry_letters = _letters_from_blob(pre)

    segs = []
    for idx, m in enumerate(it):
        item = m.group(1)
        tail = m.group(2)
        letters = _letters_from_blob(tail)
        if carry_letters and not letters:
            letters = carry_letters
            carry_letters = []  # usa só no primeiro quando não houver letras no próprio segmento
        romans = _romans_from_blob(tail)
        segs.append((item, letters, romans, tail.strip()))
    return segs


def build_transcription_for_ref(ref: str, items: dict) -> str:
    """
    Monta a transcrição para *uma* referência (uma linha da planilha).
    Ex.: "NR 1 - 1.5.3.2, alínea 'b', 1.5.4.3.1, alíneas 'a', 'b' e 'c', e 1.5.4.3.2"
    """
    parts = []
    for item, letters, romans, _tail in parse_ref_segments(ref):
        block = items.get(item, "").strip()
        if not block:
            # item não existe na versão do PDF (ex.: renumeração): pula
            parts.append(f"[AVISO] Item {item} não encontrado no PDF desta versão.")
            continue

        if letters:
            alineas = split_alineas(block)
            achou_alguma = False
            for l in letters:
                seg = alineas.get(l, "")
                if seg:
                    achou_alguma = True
                    if romans:
                        incs = split_incisos(seg)
                        escolhidos = [f"{r}. {incs[r]}" for r in romans if r in incs]
                        if escolhidos:
                            parts.append(f"{item} — alínea {l})\n" + "\n".join(escolhidos))
                        else:
                            parts.append(f"{item} — alínea {l})\n{seg}")
                    else:
                        parts.append(f"{item} — alínea {l})\n{seg}")
            if not achou_alguma:
                # alíneas pedidas não estão marcadas no PDF → devolve bloco inteiro do item
                parts.append(block)
        else:
            parts.append(block)

    return "\n\n".join(p for p in parts if p).strip()


# -------------------------- Pós-processo --------------------------
_ILLEGAL_XLSX_RE = re.compile(r'[\x00-\x08\x0B\x0C\x0E-\x1F]')

def strip_heading_objetivo(txt: str) -> str:
    """Remove cabeçalhos editoriais 'Objetivo' no início do bloco, se houver."""
    if not isinstance(txt, str):
        return txt
    s = txt.strip()
    s = re.sub(r'^\s*(\d+\.\s*)?Objetivo\s*:?\s*\n+', '', s, flags=re.IGNORECASE)
    return s.strip()

def strip_carimbo_dou(txt: str) -> str:
    """Remove linhas do tipo 'Este texto não substitui o publicado no DOU'."""
    if not isinstance(txt, str):
        return txt
    s = re.sub(r'(?im)^\s*Este texto não substitui.*$', '', txt)
    s = re.sub(r'\n{3,}', '\n\n', s)
    return s.strip()

def format_alineas(txt: str) -> str:
    """Insere quebra após 'alínea x)' para legibilidade."""
    if not isinstance(txt, str):
        return txt
    return re.sub(r'(alínea\s+[a-z]\))\s+', r'\1\n', txt, flags=re.IGNORECASE)

def sanitize_for_excel(txt: str) -> str:
    """Remove caracteres ilegais para XLSX e normaliza quebras."""
    if not isinstance(txt, str):
        return txt
    s = _ILLEGAL_XLSX_RE.sub('', txt)
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r'[ \t]+\n', '\n', s)
    s = re.sub(r'\n{3,}', '\n\n', s)
    return s.strip()

def normalizar_nbsp(s: str) -> str:
    if not isinstance(s, str):
        return s
    return (s.replace("\u00A0", " ")
             .replace("\u2009", " ")
             .replace("\u2002", " ")
             .replace("\u2003", " ")
             .strip())


# ----------------------------- Main -----------------------------
def processar_planilha_para_nr(planilha_path: str, pdf_path: str, out_path: str, nr_number: int):
    # 1) PDF -> texto -> índice de itens
    bruto = extrair_texto(pdf_path)
    norm = normalizar_texto(bruto)

    if not norm.strip():
        print("[WARN] Texto do PDF veio vazio. Verifique OCR/ou permissões.")
    items = indexar_itens(norm)
    print(f"[INFO] Itens indexados a partir do PDF: {len(items)}")

    # 2) Ler planilha
    df = pd.read_excel(planilha_path)
    df.columns = [str(c).strip() for c in df.columns]

    # Assegura colunas esperadas
    if "FUNDAMENTAÇÃO LEGAL" not in df.columns:
        print("[ERRO] Coluna 'FUNDAMENTAÇÃO LEGAL' não encontrada na planilha.")
        sys.exit(2)
    if "TRANSCRIÇÃO DO ITEM NORMATIVO" not in df.columns:
        df["TRANSCRIÇÃO DO ITEM NORMATIVO"] = ""

    # dtype seguro p/ escrita de strings
    df["TRANSCRIÇÃO DO ITEM NORMATIVO"] = df["TRANSCRIÇÃO DO ITEM NORMATIVO"].astype(object)

    # Normalizar espaços especiais
    df["FUNDAMENTAÇÃO LEGAL"] = df["FUNDAMENTAÇÃO LEGAL"].apply(normalizar_nbsp)

    # 3) Filtrar linhas da NR alvo
    nr = str(nr_number)
    nr_regex = rf'^\s*NR\s*-?\s*0*{nr}\s*[—–-]\s*'
    mask = df["FUNDAMENTAÇÃO LEGAL"].fillna("").str.match(nr_regex, case=False)

    total_nr = int(mask.sum())
    print(f"[INFO] Linhas detectadas para NR {nr_number}: {total_nr}")

    # 4) Preencher as linhas dessa NR
    filled = 0
    nao_resolvidas = []

    for idx, row in df[mask].iterrows():
        ref = str(row["FUNDAMENTAÇÃO LEGAL"])
        texto = build_transcription_for_ref(ref, items)
        texto = strip_heading_objetivo(texto)
        texto = strip_carimbo_dou(texto)
        texto = format_alineas(texto)
        texto = sanitize_for_excel(texto)

        if texto:
            df.at[idx, "TRANSCRIÇÃO DO ITEM NORMATIVO"] = texto
            filled += 1
        else:
            nao_resolvidas.append(ref)

    # 5) Sanitize final e salvar
    df["TRANSCRIÇÃO DO ITEM NORMATIVO"] = df["TRANSCRIÇÃO DO ITEM NORMATIVO"].apply(sanitize_for_excel)
    df.to_excel(out_path, index=False)

    print(f"[OK] {filled} linha(s) preenchida(s) para 'NR {nr_number} —'.")
    print(f"Planilha salva em: {out_path}")

    # 6) Diagnóstico
    if nao_resolvidas:
        print("\n[DIAGNÓSTICO] Referências sem texto extraído (até 50 exemplos):")
        for s in nao_resolvidas[:50]:
            print("  •", s)

        # Diagnóstico aprofundado: quais itens não existem no PDF?
        faltantes = set()
        letras_nao_marcadas = []
        for ref in nao_resolvidas[:200]:  # limita custo
            for item, letters, romans, _tail in parse_ref_segments(ref):
                if item not in items:
                    faltantes.add(item)
                elif letters:
                    # pediu letras mas não temos marcação -> checa
                    alineas = split_alineas(items[item])
                    falt = [l for l in letters if l not in alineas]
                    if falt:
                        letras_nao_marcadas.append((item, letters, sorted(alineas.keys())))
        if faltantes:
            print("\n[DIAGNÓSTICO] Itens citados que NÃO aparecem no PDF (possível renumeração/versão):")
            print(" ", ", ".join(sorted(faltantes)) or "-")
        if letras_nao_marcadas:
            print("\n[DIAGNÓSTICO] Itens sem alíneas identificáveis no PDF (ou formatação diferente):")
            for item, letters, existentes in letras_nao_marcadas[:20]:
                print(f"  • {item}: pediu {letters} | detectadas {existentes}")


if __name__ == "__main__":
    processar_planilha_para_nr(
        planilha_path=PLANILHA_PATH,
        pdf_path=PDF_PATH,
        out_path=OUT_PATH,
        nr_number=NR_NUMBER
    )
