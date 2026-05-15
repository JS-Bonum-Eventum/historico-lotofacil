"""
update_lotofacil.py
Lotofácil: 15 números, range 1-25
Colunas no Excel: Concurso(0), Data(1), Bola1(2)...Bola15(16)
"""

import requests
import openpyxl
from io import BytesIO

EXCEL_URL = "https://servicebus2.caixa.gov.br/portaldeloterias/api/resultados/download?modalidade=Lotof%C3%A1cil"
OUTPUT_FILE = "historico.txt"
SEQUENCE_LENGTH = 15
MAX_NUMBER = 25
# Bola1=col C (índice 2) até Bola15=col Q (índice 16)
BALL_COL_START = 2
BALL_COL_END = 17  # exclusive

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0 Safari/537.36",
    "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*",
    "Referer": "https://loterias.caixa.gov.br/Paginas/Lotofacil.aspx",
}

def download_excel():
    print("Baixando Excel da Lotofácil...")
    resp = requests.get(EXCEL_URL, headers=HEADERS, timeout=60)
    resp.raise_for_status()
    print(f"Download OK: {len(resp.content)} bytes")
    return BytesIO(resp.content)

def parse_excel(file_bytes):
    wb = openpyxl.load_workbook(file_bytes, read_only=True, data_only=True)
    ws = wb.active

    # Log do cabeçalho para diagnóstico
    header = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    print(f"Cabeçalho (cols 0-19): {header[:20]}")

    sequences = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        try:
            nums = []
            for i in range(BALL_COL_START, BALL_COL_END):
                if i < len(row) and row[i] is not None:
                    n = int(row[i])
                    if 1 <= n <= MAX_NUMBER:
                        nums.append(n)
            if len(nums) == SEQUENCE_LENGTH:
                sorted_nums = sorted(set(nums))
                if len(sorted_nums) == SEQUENCE_LENGTH:
                    sequences.append(sorted_nums)
        except (ValueError, TypeError):
            continue

    wb.close()
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
    print(f"historico.txt: {len(unique)} linhas")

def main():
    file_bytes = download_excel()
    sequences = parse_excel(file_bytes)
    if not sequences:
        raise ValueError("Nenhuma sequência válida encontrada!")
    write_output(sequences)
    print("✅ Lotofácil atualizado!")

if __name__ == "__main__":
    main()
