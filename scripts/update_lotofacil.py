"""
update_lotofacil.py
Lotofácil: 15 números (Bola1-Bola15), range 1-25
Cabeçalho confirmado: Concurso(0), Data Sorteio(1), Bola1(2)...Bola15(16)
"""

import requests
import openpyxl
from io import BytesIO

EXCEL_URL = "https://servicebus2.caixa.gov.br/portaldeloterias/api/resultados/download?modalidade=Lotof%C3%A1cil"
OUTPUT_FILE = "historico.txt"
SEQUENCE_LENGTH = 15
MAX_NUMBER = 25
BALL_COL_START = 2
BALL_COL_END = 17  # índices 2 a 16 inclusive

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0 Safari/537.36",
    "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*",
    "Referer": "https://loterias.caixa.gov.br/Paginas/Lotofacil.aspx",
}

def to_int(val):
    """Converte valor para inteiro — aceita int, float e string."""
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

    header = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    print(f"Cabeçalho: {header[:20]}")

    # Log primeiras 2 linhas de dados para diagnóstico
    row_count = 0
    sequences = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue

        # Log das primeiras linhas para diagnóstico
        if row_count < 2:
            print(f"Linha {row_count+2} raw: {row[:20]}")
        row_count += 1

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

    wb.close()
    print(f"Total linhas lidas: {row_count}")
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
    file_bytes = download_excel()
    sequences = parse_excel(file_bytes)
    if not sequences:
        raise ValueError("Nenhuma sequência válida encontrada!")
    write_output(sequences)
    print("✅ Lotofácil atualizado com sucesso!")

if __name__ == "__main__":
    main()
