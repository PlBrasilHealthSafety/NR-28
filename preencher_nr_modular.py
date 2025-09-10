# -*- coding: utf-8 -*-
"""
Preenche a coluna "TRANSCRIÇÃO DO ITEM NORMATIVO" da planilha NR-28 Anexo II
a partir de um PDF oficial de NR, usando correspondências como:
"NR X — 1.4.1, alíneas "a", "b" e "c", incisos I, II..."

Como usar:
1) Ajuste as variáveis de CONFIG no topo (PLANILHA_PATH, PDF_PATH, OUT_PATH, NR_NUMBER).
2) Rode: python preencher_nr_modular.py
3) O script preenche apenas as linhas cuja FUNDAMENTAÇÃO LEGAL começa com "NR {NR_NUMBER} —"
   (aceitando também "NR {NR_NUMBER} –" e "NR {NR_NUMBER} -", com/sem zero à esquerda e com/sem hífen entre NR e número).

Requisitos:
    pip install pandas openpyxl pdfminer.six PyPDF2
"""

import re
import pandas as pd

# ========================== CONFIG ==========================
# Exemplo para NR-04 (ajuste conforme necessário)
PLANILHA_PATH = r"C:\Users\RodrigoCinelliPLBras\Downloads\NR28_AnexoII_planilha_NR4_preenchida.xlsx"
PDF_PATH      = r"C:\Users\RodrigoCinelliPLBras\Downloads\NR05atualizada2023.pdf"
OUT_PATH      = r"C:\Users\RodrigoCinelliPLBras\Downloads\NR28_AnexoII_planilha_NR5_preenchida.xlsx"

# Número da NR alvo (apenas esse prefixo será processado na planilha)
NR_NUMBER     = 5  # ex.: 1 para NR-01, 3 para NR-03, 4 para NR-04
# ===========================================================


# ----------------------- Extração PDF -----------------------
def extrair_texto(pdf_path: str) -> str:
    """Extrai texto do PDF (pdfminer como primário, PyPDF2 como fallback)."""
    try:
        from pdfminer.high_level import extract_text
        text = extract_text(pdf_path)
        if text and text.strip():
            return text
    except Exception:
        pass

    # Fallback com PyPDF2
    try:
        import PyPDF2
        with open(pdf_path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            pages = [p.extract_text() or "" for p in reader.pages]
        return "\n".join(pages)
    except Exception:
        return ""


def normalizar_texto(bruto: str) -> str:
    """Normaliza espaços e quebras sem perder parágrafos úteis."""
    if not bruto:
        return ""
    t = bruto.replace("\r", "")
    # Substitui form feed / vertical tab por quebra de linha antes de colapsar
    t = t.replace("\x0c", "\n").replace("\x0b", "\n")
    t = re.sub(r"[ \t]+", " ", t)
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t


def indexar_itens(norm: str) -> dict:
    """
    Indexa blocos por cabeçalhos numéricos: 1, 1.4, 1.4.1, 2.3.1.2, etc.
    Aceita até 5 níveis (ajuste se precisar).
    """
    item_re = re.compile(
        r"(?m)^\s*(\d+(?:\.\d+){0,5})\b(.*?)(?=(?:^\s*\d+(?:\.\d+){0,5}\b)|\Z)",
        flags=re.DOTALL
    )
    items = {}
    for m in item_re.finditer(norm):
        num = m.group(1).strip()
        block = (m.group(1) + m.group(2)).strip()
        items[num] = block
    return items


# ----------------- Alineas / Incisos helpers ----------------
_ALINEA_HEAD_RE = re.compile(r"(?m)^\s*([a-z])\)\s", flags=re.IGNORECASE)
_INCISO_HEAD_RE = re.compile(r"(?m)^\s*([IVXLCDM]+)\.\s", flags=re.IGNORECASE)

def split_alineas(block: str) -> dict:
    """Retorna {'a': 'texto...', 'b': 'texto...'} a partir de 'a) ...', 'b) ...'."""
    alineas = {}
    pos = [(m.start(), m.group(1).lower()) for m in _ALINEA_HEAD_RE.finditer(block)]
    if not pos:
        return alineas
    pos.append((len(block), None))
    for i in range(len(pos)-1):
        start, key = pos[i]
        end, _ = pos[i+1]
        seg = block[start:end].strip()
        seg = re.sub(r"^\s*[a-z]\)\s*", "", seg, count=1, flags=re.IGNORECASE)
        alineas[key] = seg.strip()
    return alineas

def split_incisos(texto: str) -> dict:
    """Retorna {'I': '...', 'II': '...'} a partir de 'I. ...', 'II. ...'."""
    incisos = {}
    pos = [(m.start(), m.group(1)) for m in _INCISO_HEAD_RE.finditer(texto)]
    if not pos:
        return incisos
    pos.append((len(texto), None))
    for i in range(len(pos)-1):
        start, key = pos[i]
        end, _ = pos[i+1]
        seg = texto[start:end].strip()
        seg = re.sub(r"^[IVXLCDM]+\.\s*", "", seg, count=1, flags=re.IGNORECASE)
        incisos[key] = seg.strip()
    return incisos


# --------------- Parsing de referências (planilha) ----------
def parse_item_numbers(ref: str) -> list:
    """
    Extrai todos os números de itens (1.4.1, 1.5.4.3.1, 2.2, 3.1...) citados na referência.
    Aceita até 5 níveis.
    """
    nums = re.findall(r'\b(\d+(?:\.\d+){0,5})\b', ref)
    # Remover algarismos romanos capturados erroneamente
    nums = [n for n in nums if not re.fullmatch(r'[IVXLCDM]+', n, flags=re.IGNORECASE)]
    # Remover duplicatas preservando ordem
    seen, out = set(), []
    for n in nums:
        if n not in seen:
            seen.add(n)
            out.append(n)
    return out

def parse_letters_list(ref: str) -> list:
    """Extrai lista de alíneas (a, b, c, ...) da referência."""
    letters = re.findall(r'al[ií]neas?\s+["“]?([a-zA-Z])["”]?', ref, flags=re.IGNORECASE)
    # Também capturar as que aparecem só entre aspas
    more = re.findall(r'["“]([a-zA-Z])["”]', ref)
    for l in more:
        if l.lower() not in [x.lower() for x in letters]:
            letters.append(l)
    return sorted(set([l.lower() for l in letters]))

def parse_romans_list(ref: str) -> list:
    """Extrai lista de incisos (I, II, III, ...) da referência."""
    romans = re.findall(r'\b([IVXLCDM]+)\b', ref)
    seen, ordered = set(), []
    for r in romans:
        up = r.upper()
        if up not in seen:
            seen.add(up)
            ordered.append(up)
    return ordered


def build_transcription_for_ref(ref: str, items: dict) -> str:
    """
    Monta a transcrição para uma referência que pode conter múltiplos itens,
    múltiplas alíneas e múltiplos incisos.
    Estratégia:
      - Se alíneas existem, extrair só essas do bloco do item; se incisos existem, filtrar também.
      - Se alíneas não existem na referência, devolver o bloco inteiro do item.
      - Fallback: se a alínea pedida não estiver marcada no PDF extraído, devolve o bloco completo do item.
    """
    item_numbers = parse_item_numbers(ref)
    if not item_numbers:
        return ""

    letters = parse_letters_list(ref)
    romans = parse_romans_list(ref)

    parts = []
    for num in item_numbers:
        block = items.get(num, "").strip()
        if not block:
            continue

        if letters:
            alineas = split_alineas(block)
            found_any = False
            for l in letters:
                seg = alineas.get(l, "")
                if seg:
                    found_any = True
                    if romans:
                        incisos = split_incisos(seg)
                        chosen = [f"{r}. {incisos[r]}" for r in romans if r in incisos]
                        if chosen:
                            parts.append(f"{num} — alínea {l})\n" + "\n".join(chosen))
                        else:
                            parts.append(f"{num} — alínea {l})\n{seg}")
                    else:
                        parts.append(f"{num} — alínea {l})\n{seg}")
            if not found_any:
                parts.append(block)  # Fallback: usar bloco do item
        else:
            parts.append(block)

    return "\n\n".join(parts).strip()


# -------------------------- Pós-processo --------------------------
_ILLEGAL_XLSX_RE = re.compile(r'[\x00-\x08\x0B\x0C\x0E-\x1F]')

def strip_heading_objetivo(txt: str) -> str:
    """Remove cabeçalhos editoriais 'Objetivo' no início do bloco, se houver."""
    if not isinstance(txt, str):
        return txt
    s = txt.strip()
    # Remover "Objetivo"/"1. Objetivo" no topo
    s = re.sub(r'^\s*(\d+\.\s*)?Objetivo\s*:?\s*\n+', '', s, flags=re.IGNORECASE)
    s = re.sub(r'^\s*(\d+\.\s*)?Objetivo\s*:?\s*\n+', '', s, flags=re.IGNORECASE)
    return s.strip()

def strip_carimbo_dou(txt: str) -> str:
    """Remove linhas do tipo 'Este texto não substitui o publicado no DOU'."""
    if not isinstance(txt, str):
        return txt
    # Remove a linha inteira, case-insensitive
    s = re.sub(r'(?im)^\s*Este texto não substitui.*$', '', txt)
    # Colapsa blank lines extra
    s = re.sub(r'\n{3,}', '\n\n', s)
    return s.strip()

def sanitize_for_excel(txt: str) -> str:
    """Remove caracteres ilegais para planilhas XLSX e normaliza quebras."""
    if not isinstance(txt, str):
        return txt
    # Tira caracteres de controle proibidos (exceto \t, \n, \r, que são permitidos)
    s = _ILLEGAL_XLSX_RE.sub('', txt)
    # Normaliza endings e espaços
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r'[ \t]+\n', '\n', s)   # tira espaços antes de quebra
    s = re.sub(r'\n{3,}', '\n\n', s)   # evita blocos gigantes de linhas vazias
    return s.strip()

def format_alineas(txt: str) -> str:
    """Insere quebra de linha após 'alínea x)' para melhor legibilidade."""
    if not isinstance(txt, str):
        return txt
    return re.sub(r'(alínea\s+[a-z]\))\s+', r'\1\n', txt, flags=re.IGNORECASE)

def normalizar_nbsp(s: str) -> str:
    """Normaliza espaços não-quebrantes e afins nas células da planilha."""
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
    texto_pdf = extrair_texto(pdf_path)
    norm = normalizar_texto(texto_pdf)

    if not norm.strip():
        print("[WARN] Texto do PDF veio vazio. Verifique se o PDF é pesquisável (não-imagem) ou rode um OCR.")
    items = indexar_itens(norm)
    print(f"[INFO] Itens indexados a partir do PDF: {len(items)}")

    # 2) Ler planilha
    df = pd.read_excel(planilha_path)
    df.columns = [str(c).strip() for c in df.columns]

    # Normalizar espaços especiais
    df["FUNDAMENTAÇÃO LEGAL"] = df["FUNDAMENTAÇÃO LEGAL"].apply(normalizar_nbsp)

    # 3) Filtrar linhas da NR alvo com regex tolerante
    nr = str(nr_number)
    # aceita "NR 4", "NR-4", "NR 04"; separadores "—" (em dash), "–" (en dash) ou "-" (hífen)
    nr_regex = rf'^\s*NR\s*-?\s*0*{nr}\s*[—–-]\s*'
    mask = df["FUNDAMENTAÇÃO LEGAL"].fillna("").str.match(nr_regex, case=False)

    total_nr = int(mask.sum())
    print(f"[INFO] Linhas detectadas para NR {nr_number}: {total_nr}")

    if total_nr == 0:
        # Diagnóstico rápido: mostra 10 exemplos de linhas que começam com "NR" para você ver o padrão
        candidatos = df["FUNDAMENTAÇÃO LEGAL"].fillna("").astype(str)
        candidatos = [c for c in candidatos if c.strip().upper().startswith("NR")]
        print("[DEBUG] Exemplos de FUNDAMENTAÇÃO que começam com 'NR':")
        for s in candidatos[:10]:
            print("  •", repr(s))

    # 4) Preencher as linhas dessa NR
    filled = 0
    for idx, row in df[mask].iterrows():
        ref = str(row["FUNDAMENTAÇÃO LEGAL"])
        texto = build_transcription_for_ref(ref, items)
        if texto:
            # Limpezas e formatações
            texto = strip_heading_objetivo(texto)
            texto = strip_carimbo_dou(texto)
            texto = format_alineas(texto)
            texto = sanitize_for_excel(texto)
            df.at[idx, "TRANSCRIÇÃO DO ITEM NORMATIVO"] = texto
            filled += 1

    # 5) Antes de salvar, sanitize geral da coluna (por segurança)
    df["TRANSCRIÇÃO DO ITEM NORMATIVO"] = df["TRANSCRIÇÃO DO ITEM NORMATIVO"].apply(sanitize_for_excel)

    # 6) Salvar
    df.to_excel(out_path, index=False)
    print(f"[OK] {filled} linha(s) preenchida(s) para 'NR {nr_number} —'.")
    print(f"Planilha salva em: {out_path}")


if __name__ == "__main__":
    processar_planilha_para_nr(
        planilha_path=PLANILHA_PATH,
        pdf_path=PDF_PATH,
        out_path=OUT_PATH,
        nr_number=NR_NUMBER
    )
