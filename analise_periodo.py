import pandas as pd

ARQUIVO = 'ColetaVagas.xlsx'
ARQUIVO_SAIDA = 'resultado_analise_por_periodo.txt'

def categorizar_periodo(texto):
    if pd.isna(texto): return 'não mencionado'
    texto = texto.lower()
    if any(p in texto for p in ['1º', '2º']): return 'até 2º'
    elif any(p in texto for p in ['3º', '4º', '5º']): return '3º a 5º'
    elif any(p in texto for p in ['6º', '7º', '8º', '9º', '10º']): return '6º ou mais'
    elif 'qualquer' in texto or 'indiferente' in texto: return 'qualquer período'
    return 'outros'

def main():
    df = pd.read_excel(ARQUIVO)

    if 'Período Curso ' not in df.columns:
        print("Coluna 'Período Curso' não encontrada.")
        return

    df['Categoria Período'] = df['Período Curso '].apply(categorizar_periodo)
    contagem = df['Categoria Período'].value_counts()

    with open(ARQUIVO_SAIDA, 'w', encoding='utf-8') as f:
        for categoria, qtd in contagem.items():
            f.write(f"{categoria}: {qtd} vagas\n")

    print(f"Análise por período salva em '{ARQUIVO_SAIDA}'.")

if __name__ == "__main__":
    main()
