"""
update_lotofacil.py
Lotofácil: 15 números, range 1-25
Baixa via Cloudflare Worker que faz proxy da Caixa
"""

import requests
import pandas as pd
import os
from io import BytesIO

# URL do seu Cloudflare Worker — substitua pelo seu subdomínio
WORKER_URL = "https://numerix-proxy.SEU-USUARIO.workers.dev/lotofacil"

# Token configurado no Worker (Settings → Variables → SECRET_TOKEN)
SECRET_TOKEN = os.environ.get("CAIXA_SECRET_TOKEN", "")

OUTPUT_FILE = "historico.txt"
SEQUENCE_LENGTH = 15
MAX_NUMBER = 25

def download_excel():
    print("Baixando Excel da Lotofácil via Worker...")
    r = requests.get(
        WORKER_URL,
        headers={"X-Auth-Token": SECRET_TOKEN},
        timeout=60,
    )
    r.raise_for_status()
    print(f"Download OK: {len(r.content)} bytes")
    return BytesIO(r.content)

def parse_excel(file_bytes):
    df = pd.read_excel(file_bytes, engine='openpyxl', header=0)
    print(f"DataFrame: {len(df)} linhas x {len(df.columns)} colunas")

    ball_cols = sorted(
        [c for c in df.columns if str(c).startswith('Bola') and str(c)[4:].isdigit()],
        key=lambda c: int(str(c)[4:])
    )[:SEQUENCE_LENGTH]
    print(f"Colunas bolas: {ball_cols}")

    if len(ball_cols) < SEQUENCE_LENGTH:
        raise ValueError(f"Esperava {SEQUENCE_LENGTH} colunas Bola, encontrou {len(ball_cols)}")

    sequences = []
    skipped = 0
    for _, row in df.iterrows():
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
        except Exception:
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
