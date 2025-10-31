import pandas as pd
import os
from datetime import datetime

def conferir_acertos(jogo, resultado):
    acertos = jogo.intersection(resultado)
    return acertos, len(acertos)

# === Jogos fixos (edite aqui se quiser mudar os números) ===
jogo_jair = {3, 4, 5, 7, 8, 9, 10, 15, 16, 18, 19, 20, 21, 23, 25}
jogo_janete = {2, 3, 4, 7, 8, 10, 12, 13, 15, 16, 17, 18, 19, 20, 24}

arquivo_excel = "resultados_sorteios.xlsx"

print("=== Verificador de Acertos - Jair e Janete ===")
print(f"Jogo Jair: {sorted(jogo_jair)}")
print(f"Jogo Janete: {sorted(jogo_janete)}")

# --- Loop para registrar vários sorteios ---
while True:
    print("\n=== NOVO SORTEIO ===")
    nome_sorteio = input("Digite o nome ou número do sorteio (ex: Sorteio 1): ")

    # Gerar nome de aba com data/hora
    agora = datetime.now().strftime("%Y-%m-%d_%Hh%M")
    nome_aba = f"{nome_sorteio}_{agora}"

    resultado = input("Digite os números sorteados separados por espaço: ")
    resultado = set(map(int, resultado.split()))

    # Conferência
    acertos_jair, qtd_jair = conferir_acertos(jogo_jair, resultado)
    acertos_janete, qtd_janete = conferir_acertos(jogo_janete, resultado)

    # Exibir resultado
    print("\n=== RESULTADO ===")
    print(f"Jogo Jair acertou {qtd_jair} números: {sorted(acertos_jair)}")
    print(f"Jogo Janete acertou {qtd_janete} números: {sorted(acertos_janete)}")

    # Criar DataFrame
    dados = {
        "Jogo": ["Jair", "Janete"],
        "Números Jogados": [
            ", ".join(map(str, sorted(jogo_jair))),
            ", ".join(map(str, sorted(jogo_janete)))
        ],
        "Resultado": ", ".join(map(str, sorted(resultado))),
        "Acertos": [qtd_jair, qtd_janete],
        "Números Acertados": [
            ", ".join(map(str, sorted(acertos_jair))),
            ", ".join(map(str, sorted(acertos_janete)))
        ],
        "Data/Hora Registro": [datetime.now().strftime("%d/%m/%Y %H:%M:%S")]*2
    }

    df = pd.DataFrame(dados)

    # Salvar no Excel (cria ou adiciona aba)
    modo = "a" if os.path.exists(arquivo_excel) else "w"
    with pd.ExcelWriter(arquivo_excel, engine="openpyxl", mode=modo) as writer:
        df.to_excel(writer, sheet_name=nome_aba[:31], index=False)

    print(f"\n✅ Sorteio '{nome_sorteio}' salvo na aba '{nome_aba}' do arquivo '{arquivo_excel}'.")

    continuar = input("\nDeseja registrar outro sorteio? (s/n): ").lower()
    if continuar != "s":
        break

print("\nEncerrado. Todos os sorteios foram salvos no arquivo Excel.")
