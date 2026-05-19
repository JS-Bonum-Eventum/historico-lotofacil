"""
update_lotofacil.py
Lotofácil: 15 números (Bola1-Bola15), range 1-25
"""

import requests
import pandas as pd
from io import BytesIO

EXCEL_URL = "https://servicebus2.caixa.gov.br/portaldeloterias/api/resultados/download?modalidade=Lotof%C3%A1cil"
OUTPUT_FILE = "historico.txt"
SEQUENCE_LENGTH = 15
MAX_NUMBER = 25

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0 Safari/537.36",
    "Accept": "*/*",
    "Referer": "https://loterias.caixa.gov.br/Paginas/Lotofacil.aspx",
}

def download_excel():
    print("Baixando Excel da Lotofácil...")
    resp = requests.get(EXCEL_URL, headers=HEADERS, timeout=60)
    resp.raise_for_status()
    print(f"Download OK: {len(resp.content)} bytes")
    return BytesIO(resp.content)

def parse_excel(file_bytes):
    # pandas lê todas as linhas corretamente incluindo shared strings
    df = pd.read_excel(file_bytes, engine='openpyxl', header=0)
    print(f"DataFrame: {len(df)} linhas x {len(df.columns)} colunas")
    print(f"Colunas: {list(df.columns[:20])}")

    # Identifica colunas Bola1...Bola15
    ball_cols = [c for c in df.columns if str(c).startswith('Bola') and str(c)[4:].isdigit()]
    ball_cols = sorted(ball_cols, key=lambda c: int(str(c)[4:]))[:SEQUENCE_LENGTH]
    print(f"Colunas bolas: {ball_cols}")

    if len(ball_cols) < SEQUENCE_LENGTH:
        raise ValueError(f"Esperava {SEQUENCE_LENGTH} colunas Bola, encontrou {len(ball_cols)}")

    sequences = []
    skipped = 0
    for idx, row in df.iterrows():
        try:
            nums = []
            for col in ball_cols:
                val = row[col]
                if pd.isna(val):
                    continue
                n = int(float(str(val).strip()))
                if 1 <= n <= MAX_NUMBER:
                    nums.append(n)
            if len(nums) == SEQUENCE_LENGTH:
                sorted_nums = sorted(set(nums))
                if len(sorted_nums) == SEQUENCE_LENGTH:
                    sequences.append(sorted_nums)
                else:
                    skipped += 1
            else:
                skipped += 1
                if skipped <= 3:
                    print(f"  Linha {idx+2} ignorada: {len(nums)} nums válidos de {[row[c] for c in ball_cols]}")
        except Exception as e:
            skipped += 1

    print(f"Sequências válidas: {len(sequences)}")
    print(f"Linhas ignoradas: {skipped}")
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
