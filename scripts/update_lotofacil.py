"""
update_lotofacil.py
Lotofácil: 15 números (Bola1-Bola15), range 1-25
"""

import requests
import openpyxl
from io import BytesIO

EXCEL_URL = "https://servicebus2.caixa.gov.br/portaldeloterias/api/resultados/download?modalidade=Lotof%C3%A1cil"
OUTPUT_FILE = "historico.txt"
SEQUENCE_LENGTH = 15
MAX_NUMBER = 25

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0 Safari/537.36",
    "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*",
    "Referer": "https://loterias.caixa.gov.br/Paginas/Lotofacil.aspx",
}

def to_int(val):
    if val is None:
        return None
    try:
        return int(float(str(val).strip().replace(',', '.')))
    except (ValueError, TypeError):
        return None

def download_excel():
    print("Baixando Excel da Lotofácil...")
    resp = requests.get(EXCEL_URL, headers=HEADERS, timeout=60)
    resp.raise_for_status()
    print(f"Download OK: {len(resp.content)} bytes")
    return BytesIO(resp.content)

def parse_excel(file_bytes):
    wb = openpyxl.load_workbook(file_bytes, read_only=True, data_only=True)
    ws = wb.active

    print(f"Dimensões: max_row={ws.max_row}, max_col={ws.max_column}")

    # Lê TODAS as linhas como lista para inspeção
    all_rows = list(ws.iter_rows(values_only=True))
    print(f"Total de linhas no arquivo: {len(all_rows)}")

    # Log das primeiras 5 linhas para diagnóstico
    for i, row in enumerate(all_rows[:5]):
        print(f"  Linha {i+1}: {row[:20]}")

    # Encontra a linha do cabeçalho (que contém "Bola1")
    header_row_idx = None
    for i, row in enumerate(all_rows):
        if row and any(str(c).strip() == 'Bola1' for c in row if c is not None):
            header_row_idx = i
            print(f"Cabeçalho encontrado na linha {i+1}: {row[:20]}")
            break

    if header_row_idx is None:
        raise ValueError("Cabeçalho com 'Bola1' não encontrado!")

    # Descobre índices das colunas Bola1...Bola15
    header = all_rows[header_row_idx]
    ball_cols = []
    for i, val in enumerate(header):
        if val is not None:
            s = str(val).strip()
            if s.startswith('Bola') and s[4:].isdigit():
                ball_cols.append((int(s[4:]), i))

    ball_cols.sort()  # ordena por número da bola
    ball_cols = [i for _, i in ball_cols[:SEQUENCE_LENGTH]]
    print(f"Colunas das bolas (índices): {ball_cols}")

    if len(ball_cols) < SEQUENCE_LENGTH:
        raise ValueError(f"Esperava {SEQUENCE_LENGTH} colunas de bolas, encontrou {len(ball_cols)}")

    # Processa linhas de dados (após o cabeçalho)
    sequences = []
    for row in all_rows[header_row_idx + 1:]:
        if not row or row[0] is None:
            continue
        nums = []
        for i in ball_cols:
            n = to_int(row[i]) if i < len(row) else None
            if n is not None and 1 <= n <= MAX_NUMBER:
                nums.append(n)
        if len(nums) == SEQUENCE_LENGTH:
            sorted_nums = sorted(set(nums))
            if len(sorted_nums) == SEQUENCE_LENGTH:
                sequences.append(sorted_nums)

    print(f"Sequências válidas: {len(sequences)}")
    wb.close()
    return sequences

def write_output(sequences):
    seen = set()
    unique = []
    for s in sequences:
        key = ','.join(map(str, s))
        if key not in seen:
            seen.add(key)
            unique.append(s)
    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        for seq in unique:
            f.write(' '.join(map(str, seq)) + '\n')
    print(f"historico.txt: {len(unique)} linhas escritas")

def main():
    file_bytes = download_excel()
    sequences = parse_excel(file_bytes)
    if not sequences:
        raise ValueError("Nenhuma sequência válida encontrada!")
    write_output(sequences)
    print("✅ Lotofácil atualizado com sucesso!")

if __name__ == "__main__":
    main()
