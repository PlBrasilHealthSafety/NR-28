# -*- coding: utf-8 -*-
"""
Preenche/atualiza a planilha a partir de dados MANUAIS de uma NR específica.
- 100% manual: não lê o PDF.
- Atualiza por CÓDIGO (prefixo da NR) e acrescenta códigos novos no final.
- Mantém NRs antigas intactas.

Requisitos:
    pip install pandas openpyxl
"""

import re
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Tuple

import pandas as pd

# ============== CONFIG ==============
# Base atual (já com NRs anteriores)
XLSX_IN  = Path(r"C:\Users\RodrigoCinelliPLBras\Downloads\NR28_AnexoII_planilha_PREENCHIDA.xlsx")
# Saída com a NR nova incorporada
XLSX_OUT = Path(r"C:\Users\RodrigoCinelliPLBras\Downloads\NR28_AnexoII_planilha_PREENCHIDA1.xlsx")

# NR alvo deste run
TARGET_NR = 38

# ============== DADOS MANUAIS ==============
# Cada tupla: (Item/Subitem, Código, Infração, Tipo)
MANUAL_ROWS = [
    ('NR 38 - 38.7.3.1', '138065-6', '2', 'S'),
    ('NR 38 - 38.7.3.2', '138105-9', '2', 'S'),
    ('NR 38 - 38.8.1', '138066-4', '3', 'S'),
    ('NR 38 - 38.8.1.1', '138067-2', '3', 'S'),
    ('NR 38 - 38.8.2, alíneas "a" a "d"', '138068-0', '2', 'S'),
    ('NR 38 - 38.8.2.1, alíneas "a" a "f"', '138069-9', '2', 'S'),
    ('NR 38 - 38.8.3, alíneas "a" a "d"', '138070-2', '2', 'S'),
    ('NR 38 - 38.8.3.1, alíneas "a" a "c"', '138071-0', '2', 'S'),
    ('NR 38 - 38.8.3.2 e 38.8.3.2.1', '138072-9', '2', 'S'),
    ('NR 38 - 38.8.4, alínea "a"', '138073-7', '3', 'S'),
    ('NR 38 - 38.8.4, alínea "b"', '138074-5', '2', 'S'),
    ('NR 38 - 38.8.4.1', '138075-3', '4', 'S'),
    ('NR 38 - 38.8.7', '138076-1', '2', 'S'),
    ('NR 38 - 38.8.8', '138077-0', '2', 'S'),
    ('NR 38 - 38.9.1', '138078-8', '2', 'S'),
    ('NR 38 - 38.9.10', '138090-7', '2', 'S'),
    ('NR 38 - 38.9.2', '138079-6', '2', 'S'),
    ('NR 38 - 38.9.3', '138080-0', '2', 'S'),
    ('NR 38 - 38.9.3.1, alíneas "a" a "g"', '138081-8', '2', 'S'),
    ('NR 38 - 38.9.3.2, alíneas "a" a "d"', '138082-6', '2', 'S'),
    ('NR 38 - 38.9.4', '138083-4', '2', 'S'),
    ('NR 38 - 38.9.5, alíneas "a" e "b"', '138084-2', '2', 'S'),
    ('NR 38 - 38.9.5.1', '138085-0', '2', 'S'),
    ('NR 38 - 38.9.6', '138086-9', '2', 'S'),
    ('NR 38 - 38.9.7', '138087-7', '1', 'S'),
    ('NR 38 - 38.9.8', '138088-5', '2', 'S'),
    ('NR 38 - 38.9.9', '138089-3', '2', 'S'),
    ('NR 38 - 38.10.1, alínea "a"', '138106-7', '2', 'S'),
    ('NR 38 - 38.10.1, alínea "c"', '138107-5', '2', 'S'),
    ('NR 38 - 38.10.2, alínea "a"', '138091-5', '2', 'S'),
    ('NR 38 - 38.10.2, alínea "b"', '138092-3', '2', 'S'),
    ('NR 38 - 38.10.3, alínea "a"', '138093-1', '2', 'S'),
    ('NR 38 - 38.10.3, alínea "b"', '138094-0', '2', 'S'),
    ('NR 38 - 38.10.4', '138095-8', '2', 'S'),
    ('NR 38 - 38.10.4.1', '138096-6', '1', 'S'),
    ('NR 38 - 38.10.5', '138097-4', '2', 'S'),
    ('NR 38 - 38.10.5.1, alínea "a"', '138098-2', '2', 'S'),
    ('NR 38 - 38.10.5.1, alínea "b"', '138099-0', '2', 'S'),
    ('NR 38 - 38.10.5.1, alínea "c"', '138100-8', '2', 'S'),
    ('NR 38 - 38.10.5.1.1, alíneas "a" e "b"', '138101-6', '2', 'S'),
    ('NR 38 - 38.10.6', '138102-4', '2', 'S'),
    ('NR 38 - 38.10.7, alíneas "a" e "b"', '138103-2', '2', 'S'),
]

# ============== HELPERS ==============
def _safe_write_excel(df: pd.DataFrame, out_path: Path) -> Path:
    try:
        df.to_excel(out_path, index=False)
        return out_path
    except PermissionError:
        ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        alt = out_path.parent / f"{out_path.stem}_{ts}{out_path.suffix}"
        df.to_excel(alt, index=False)
        print(f"Aviso: arquivo de saída estava em uso. Salvei como: {alt}")
        return alt

def build_df_from_manual(target_nr: int, manual_rows: List[Tuple[str, str, str, str]]) -> Tuple[pd.DataFrame, str]:
    """Constrói o DF novo e retorna também o prefixo de código (3 primeiros dígitos)."""
    rows: List[Dict] = []
    prefix = None
    for item, codigo, infracao, tipo in manual_rows:
        codigo = str(codigo).strip()
        if prefix is None and re.fullmatch(r"\d{6}-\d", codigo):
            prefix = codigo[:3]
        rows.append({
            "FUNDAMENTAÇÃO LEGAL": f"NR {target_nr} - {item}",
            "CÓDIGO": codigo,
            "INFRAÇÃO": str(infracao).strip(),
            "TIPO": str(tipo).strip().upper(),
        })
    if prefix is None and rows:
        # fallback: tenta extrair do primeiro código válido
        for r in rows:
            m = re.match(r"(\d{3})\d{3}-\d$", r["CÓDIGO"])
            if m:
                prefix = m.group(1); break
    return pd.DataFrame(rows), prefix or ""

def fill_spreadsheet_append_safe(xlsx_in: Path, xlsx_out: Path,
                                 df_new: pd.DataFrame, target_nr: int, prefix_alvo: str) -> None:
    """
    Mantém a base intacta e:
      - Atualiza SOMENTE os códigos com o prefixo alvo E cuja nova FUNDAMENTAÇÃO começa com 'NR 0*<alvo>'.
      - Acrescenta (no final) códigos novos desse prefixo.
    """
    df_new = df_new.copy()
    df_new["CÓDIGO"] = df_new["CÓDIGO"].astype(str).str.strip()

    # Garante “NR 0*<alvo>” no início
    re_nr_alvo = re.compile(rf"^NR\s*0*{target_nr}\b", re.IGNORECASE)
    df_new = df_new[df_new["FUNDAMENTAÇÃO LEGAL"].str.match(re_nr_alvo)]
    if df_new.empty:
        raise SystemExit("Lista manual está vazia ou sem 'FUNDAMENTAÇÃO LEGAL' iniciando por 'NR <alvo>'.")

    if not xlsx_in.exists():
        # gerar planilha do zero
        df_out = df_new.copy()
        # garante coluna de transcrição
        if "TRANSCRIÇÃO DO ITEM NORMATIVO" not in df_out.columns:
            df_out.insert(1, "TRANSCRIÇÃO DO ITEM NORMATIVO", "")
        saved = _safe_write_excel(df_out, xlsx_out)
        print(f"Planilha gerada do zero com {len(df_out)} linhas. Arquivo: {saved}")
        return

    df_old = pd.read_excel(xlsx_in, engine="openpyxl")
    if "CÓDIGO" not in df_old.columns:
        raise ValueError("A planilha de entrada precisa ter a coluna 'CÓDIGO'.")
    if "FUNDAMENTAÇÃO LEGAL" not in df_old.columns:
        df_old["FUNDAMENTAÇÃO LEGAL"] = ""
    if "TRANSCRIÇÃO DO ITEM NORMATIVO" not in df_old.columns:
        df_old.insert(1, "TRANSCRIÇÃO DO ITEM NORMATIVO", "")

    df_old["CÓDIGO"] = df_old["CÓDIGO"].astype(str).str.strip()

    # Filtra df_new por prefixo
    if prefix_alvo:
        df_new_target = df_new[df_new["CÓDIGO"].str.startswith(prefix_alvo)].copy()
    else:
        df_new_target = df_new.copy()

    mapa_novo = dict(zip(df_new_target["CÓDIGO"], df_new_target["FUNDAMENTAÇÃO LEGAL"]))

    # 1) Atualiza os já existentes (mesmo código)
    mask_update = df_old["CÓDIGO"].isin(df_new_target["CÓDIGO"])
    df_old.loc[mask_update, "FUNDAMENTAÇÃO LEGAL"] = df_old.loc[mask_update, "CÓDIGO"].map(mapa_novo)

    # 2) Acrescenta no final os que não existem
    codigos_existentes = set(df_old["CÓDIGO"])
    df_append = df_new_target[~df_new_target["CÓDIGO"].isin(codigos_existentes)].copy()

    # Ajusta colunas e ordem
    for col in df_old.columns:
        if col not in df_append.columns:
            df_append[col] = ""
    df_append = df_append[df_old.columns]

    df_out = pd.concat([df_old, df_append], ignore_index=True)
    saved = _safe_write_excel(df_out, xlsx_out)
    print(f"Atualizados (prefixo {prefix_alvo or '—'}): {mask_update.sum()} | Acrescentados: {len(df_append)} | Total final: {len(df_out)}")
    print(f"Planilha salva: {saved}")

# ============== RUN ==============
if __name__ == "__main__":
    df_new, prefix = build_df_from_manual(TARGET_NR, MANUAL_ROWS)
    print(f"Resumo NR {TARGET_NR}: {len(df_new)} linhas (manual) | prefixo detectado: {prefix or '—'}")
    fill_spreadsheet_append_safe(XLSX_IN, XLSX_OUT, df_new, target_nr=TARGET_NR, prefix_alvo=prefix)
