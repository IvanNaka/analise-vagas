import pandas as pd
from collections import Counter
import string
import re

CAMINHO_ARQUIVO = 'ColetaVagas.xlsx'
ARQUIVO_SAIDA = 'resultado_analise_por_modalidade.txt'

# Função de limpeza e contagem de palavras
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


    colunas_textuais = ["Hard Skills", "Soft Skills", "Formação Necessária", "Observações adicionais"]

    # Verificar nome correto da coluna de modalidade
    modalidade_col = [col for col in df.columns if 'modalidade' in col.lower()]
    if not modalidade_col:
        print("Coluna de modalidade não encontrada.")
        return
    coluna_modalidade = modalidade_col[0]

    # Abrir arquivo de saída
    with open(ARQUIVO_SAIDA, 'w', encoding='utf-8') as f:
        modalidades = df[coluna_modalidade].dropna().unique()

        for modalidade in modalidades:
            f.write(f"\n{'='*60}\nMODALIDADE: {modalidade}\n{'='*60}\n")
            df_modalidade = df[df[coluna_modalidade] == modalidade]

            for coluna in colunas_textuais:
                if coluna not in df.columns:
                    f.write(f"\n[Coluna não encontrada: {coluna}]\n")
                    continue
                palavras = sum(df_modalidade[coluna].dropna().astype(str).apply(contar_palavras), [])
                contagem = Counter(palavras)

                f.write(f"\nPalavras mais frequentes em '{coluna}':\n")
                for palavra, freq in sorted(contagem.items(), key=lambda x: (-x[1], x[0])):
                    f.write(f"{palavra}: {freq}\n")

    print(f"Resultado salvo em '{ARQUIVO_SAIDA}'.")

if __name__ == "__main__":
    main()
