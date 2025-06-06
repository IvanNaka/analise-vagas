import pandas as pd
import re
from collections import Counter

ARQUIVO = 'ColetaVagas.xlsx'
ARQUIVO_SAIDA = 'resultado_analise_por_formacao.txt'

def extrair_areas_formacao(texto):
    if pd.isna(texto): return []
    texto = texto.lower()
    texto = re.sub(r'\n', ' ', texto)
    return re.findall(r'(computação|engenharia|sistemas|informática|software|análise|dados|rede|segurança)', texto)

def main():
    df = pd.read_excel(ARQUIVO)

    if 'Formação Necessária ' not in df.columns:
        print("Coluna 'Formação Necessária' não encontrada.")
        return

    df['Áreas Formação'] = df['Formação Necessária '].apply(extrair_areas_formacao)
    todas_areas = df['Áreas Formação'].sum()
    contagem = Counter(todas_areas).most_common()

    with open(ARQUIVO_SAIDA, 'w', encoding='utf-8') as f:
        for area, freq in contagem:
            f.write(f"{area}: {freq} vagas\n")

    print(f"Análise por formação salva em '{ARQUIVO_SAIDA}'.")

if __name__ == "__main__":
    main()
