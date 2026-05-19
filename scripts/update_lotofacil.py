"""
update_lotofacil.py
Lotofácil: 15 números (Bola1-Bola15), range 1-25
Cabeçalho na linha 1, dados a partir da linha 2, bolas nas colunas C-Q (índice 2-16)
"""

import requests
from io import BytesIO

EXCEL_URL = "https://servicebus2.caixa.gov.br/portaldeloterias/api/resultados/download?modalidade=Lotof%C3%A1cil"
OUTPUT_FILE = "historico.txt"
SEQUENCE_LENGTH = 15
MAX_NUMBER = 25
BALL_COL_START = 2   # coluna C
BALL_COL_END = 17    # coluna Q inclusive (índice 16), exclusive 17

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0 Safari/537.36",
    "Accept": "*/*",
    "Referer": "https://loterias.caixa.gov.br/Paginas/Lotofacil.aspx",
}

def to_int(val):
    if val is None or val == '':
        return None
    try:
        return int(float(str(val).strip().replace(',', '.')))
    except (ValueError, TypeError):
        return None

def download_excel():
    print("Baixando Excel da Lotofácil...")
    resp = requests.get(EXCEL_URL, headers=HEADERS, timeout=60)
    resp.raise_for_status()
    content = resp.content
    print(f"Download OK: {len(content)} bytes")
    # Detecta formato pelo magic bytes
    if content[:2] == b'PK':
        fmt = 'xlsx'
    elif content[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
        fmt = 'xls'
    else:
        fmt = 'xlsx'  # tenta xlsx por padrão
    print(f"Formato detectado: {fmt}")
    return BytesIO(content), fmt

def parse_xlsx(file_bytes):
    import openpyxl
    wb = openpyxl.load_workbook(file_bytes, data_only=True)
    ws = wb.active
    print(f"openpyxl: {ws.max_row} linhas x {ws.max_column} colunas")
    rows = []
    for row in ws.iter_rows(min_row=1, values_only=True):
        rows.append(list(row))
    wb.close()
    return rows

def parse_xls(file_bytes):
    import xlrd
    wb = xlrd.open_workbook(file_contents=file_bytes.read())
    ws = wb.sheet_by_index(0)
    print(f"xlrd: {ws.nrows} linhas x {ws.ncols} colunas")
    rows = []
    for i in range(ws.nrows):
        rows.append(ws.row_values(i))
    return rows

def extract_sequences(rows):
    print(f"Total linhas: {len(rows)}")
    # Log primeiras 3 linhas
    for i, row in enumerate(rows[:3]):
        print(f"  Linha {i+1}: {row[:20]}")

    sequences = []
    # Pula linha 0 (cabeçalho), processa a partir da linha 1
    for row in rows[1:]:
        if not row or row[0] is None or str(row[0]).strip() == '':
            continue
        nums = []
        for i in range(BALL_COL_START, BALL_COL_END):
            if i < len(row):
                n = to_int(row[i])
                if n is not None and 1 <= n <= MAX_NUMBER:
                    nums.append(n)
        if len(nums) == SEQUENCE_LENGTH:
            sorted_nums = sorted(set(nums))
            if len(sorted_nums) == SEQUENCE_LENGTH:
                sequences.append(sorted_nums)

    print(f"Sequências válidas: {len(sequences)}")
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
    file_bytes, fmt = download_excel()
    if fmt == 'xls':
        rows = parse_xls(file_bytes)
    else:
        rows = parse_xlsx(file_bytes)

    sequences = extract_sequences(rows)
    if not sequences:
        raise ValueError("Nenhuma sequência válida encontrada!")
    write_output(sequences)
    print("✅ Lotofácil atualizado com sucesso!")

if __name__ == "__main__":
    main()
