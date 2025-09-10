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
TARGET_NR = 22

# ============== DADOS MANUAIS ==============
# Cada tupla: (Item/Subitem, Código, Infração, Tipo)
MANUAL_ROWS = [
    # --- NR 22 (continuação) ---
    ('22.19.2, alínea "g"', '322359-0', '4', 'S'),
    ('22.19.2, alínea "h"', '322360-4', '4', 'S'),
    ('22.19.2, alínea "i"', '322361-2', '4', 'S'),
    ('22.19.2.1', '322362-0', '4', 'S'),
    ('22.19.3', '322363-9', '4', 'S'),
    ('22.19.4', '322364-7', '4', 'S'),
    ('22.19.5, alínea "a"', '322365-5', '2', 'S'),
    ('22.19.5, alínea "b"', '322366-3', '3', 'S'),
    ('22.19.5, alínea "c"', '322367-1', '4', 'S'),
    ('22.19.5, alínea "d"', '322368-0', '4', 'S'),
    ('22.19.5, alínea "e"', '322369-8', '4', 'S'),
    ('22.19.5, alínea "f"', '322370-1', '4', 'S'),
    ('22.19.5, alínea "g"', '322371-0', '4', 'S'),
    ('22.19.5.1, alínea "a"', '322372-8', '4', 'S'),
    ('22.19.5.1, alínea "b"', '322373-6', '4', 'S'),
    ('22.19.5.1, alínea "c"', '322374-4', '4', 'S'),
    ('22.19.5.1.1, alínea "a"', '322375-2', '4', 'S'),
    ('22.19.5.1.1, alínea "b"', '322376-0', '4', 'S'),
    ('22.19.5.1.1, alínea "c"', '322377-9', '4', 'S'),
    ('22.19.5.2, alínea "a"', '322378-7', '4', 'S'),
    ('22.19.5.2, alínea "b"', '322379-5', '4', 'S'),
    ('22.19.5.2, alínea "c"', '322380-9', '4', 'S'),
    ('22.19.5.3', '322381-7', '4', 'S'),
    ('22.19.6', '322382-5', '4', 'S'),
    ('22.19.7', '322383-3', '4', 'S'),
    ('22.19.8', '322384-1', '3', 'S'),
    ('22.19.9, alínea "a"', '322385-0', '4', 'S'),
    ('22.19.9, alínea "b"', '322386-8', '4', 'S'),
    ('22.19.9, alínea "c"', '322387-6', '4', 'S'),
    ('22.19.9, alínea "d"', '322388-4', '4', 'S'),
    ('22.19.9, alínea "e"', '322389-2', '2', 'S'),
    ('22.19.9, alínea "f"', '322390-6', '4', 'S'),
    ('22.19.9.1', '322391-4', '4', 'S'),
    ('22.19.9.2', '322392-2', '4', 'S'),
    ('22.19.10 e 22.19.10.1', '322393-0', '3', 'S'),
    ('22.19.11', '322394-9', '4', 'S'),
    ('22.19.11.1', '322395-7', '4', 'S'),
    ('22.19.12', '322396-5', '4', 'S'),
    ('22.19.12.1', '322397-3', '4', 'S'),
    ('22.19.13', '322398-1', '3', 'S'),
    ('22.19.14', '322399-0', '4', 'S'),
    ('22.19.14.1', '322400-7', '3', 'S'),
    ('22.19.15', '322401-5', '4', 'S'),
    ('22.19.16', '322402-3', '4', 'S'),
    ('22.19.17', '322403-1', '4', 'S'),
    ('22.19.18', '322404-0', '4', 'S'),
    ('22.19.19', '322405-8', '4', 'S'),
    ('22.19.20', '322406-6', '4', 'S'),
    ('22.19.21', '322407-4', '4', 'S'),
    ('22.19.22', '322408-2', '3', 'S'),
    ('22.19.23', '322409-0', '4', 'S'),
    ('22.19.24', '322410-4', '4', 'S'),
    ('22.19.25', '322411-2', '4', 'S'),
    ('22.19.26, alínea "a"', '322412-0', '4', 'S'),
    ('22.19.26, alínea "b"', '322413-9', '4', 'S'),
    ('22.19.26, alínea "c"', '322414-7', '4', 'S'),
    ('22.19.26, alínea "d"', '322415-5', '4', 'S'),
    ('22.19.26, alínea "e"', '322416-3', '4', 'S'),
    ('22.19.26, alínea "f"', '322417-1', '4', 'S'),
    ('22.19.27', '322418-0', '4', 'S'),
    ('22.19.28, alínea "a"', '322419-8', '4', 'S'),
    ('22.19.28, alínea "b"', '322420-1', '4', 'S'),
    ('22.19.28, alínea "c"', '322421-0', '4', 'S'),
    ('22.19.28, alínea "d"', '322422-8', '4', 'S'),
    ('22.19.29, alínea "a"', '322423-6', '4', 'S'),
    ('22.19.29, alínea "b"', '322424-4', '4', 'S'),
    ('22.19.29, alínea "c"', '322425-2', '4', 'S'),
    ('22.19.29, alínea "d"', '322426-0', '4', 'S'),
    ('22.19.30, alínea "a"', '322427-9', '4', 'S'),
    ('22.19.30, alínea "b"', '322428-7', '4', 'S'),
    ('22.19.30, alínea "c"', '322429-5', '4', 'S'),
    ('22.19.30.1', '322430-9', '4', 'S'),
    ('22.19.31', '322431-7', '4', 'S'),
    ('22.19.32', '322432-5', '3', 'S'),
    ('22.19.33', '322433-3', '4', 'S'),
    ('22.19.34', '322434-1', '4', 'S'),
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
