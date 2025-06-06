import pandas as pd
from collections import Counter
import string
import re

# Caminho do arquivo Excel
CAMINHO_ARQUIVO = 'ColetaVagas.xlsx'
ARQUIVO_SAIDA = 'resultado_analise.txt'

# Função para limpar e contar palavras
def contar_palavras(texto):
    if pd.isna(texto):
        return []
    texto = texto.lower()
    texto = re.sub(r'\n', ' ', texto)
    texto = texto.translate(str.maketrans('', '', string.punctuation))
    palavras = texto.split()
    palavras = [p for p in palavras if len(p) > 2 and not p.isdigit()]
    return palavras

def main():
    df = pd.read_excel(CAMINHO_ARQUIVO)

    df.columns = df.columns.astype(str).str.strip().str.replace('\n', '', regex=False)

    colunas_textuais = ["Hard Skills", "Soft Skills", "Formação Necessária", "Observações adicionais"]

    with open(ARQUIVO_SAIDA, 'w', encoding='utf-8') as f:
        for coluna in colunas_textuais:
            if coluna not in df.columns:
                f.write(f"\nColuna não encontrada: {coluna}\n")
                continue
            todas_palavras = df[coluna].dropna().astype(str).apply(contar_palavras).sum()
            contagem = Counter(todas_palavras)

            f.write(f"\nPalavras mais frequentes em '{coluna}':\n")
            for palavra, freq in sorted(contagem.items(), key=lambda x: (-x[1], x[0])):
                f.write(f"{palavra}: {freq}\n")

    print(f"Resultado salvo em '{ARQUIVO_SAIDA}'.")

if __name__ == "__main__":
    main()
