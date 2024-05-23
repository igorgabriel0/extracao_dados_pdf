import os
import pdfplumber
import pandas as pd

# Definindo diretórios
diretorio_atual = os.getcwd()
print(diretorio_atual)

diretorio_alvo = r"C:\Users\igor.gabriel\OneDrive - JSL SA\Área de Trabalho\IGOR"
os.chdir(diretorio_alvo)

diretorio_arquivos = r"C:\Users\igor.gabriel\OneDrive - JSL SA\Área de Trabalho\NF\ARQUIVOS\NFs"

# Processando arquivos PDF
lista_arquivos = os.listdir(diretorio_arquivos)
data = []  # Lista para armazenar os dados extraídos

for arquivo in lista_arquivos:
    if ".pdf" not in arquivo:
        continue  # Ignora arquivos que não sejam PDF

    try:
        # Abrindo o PDF
        with pdfplumber.open(f"{diretorio_arquivos}/{arquivo}") as pdf:
            pagina = pdf.pages[0]
            texto = pagina.extract_text()
            linhas = texto.split("\n")
            a = 0

            # LINHA CÓDIGO
            for i, linha in enumerate(linhas):               
                if "CÓDIGO" in linha:
                    print(f"ENCONTRADO NA LINHA {i}")
                    colunas = linhas[i].split(" ")
                    valores = linhas[i + 1].split(" ")     
                    prox_linha = linhas[i + 2].split(" ")              
                    
                    #UNID
                    for j, valor in enumerate(valores):
                        if valor in ("UN", "Un", "UND", "PCS", "LT", "CDA", "PC"):
                            unid = valor
                            print(f"UNID: {unid}")
                            break    
                    if not unid:
                        unid = "NÃO ACHOU"

                    for l, linha in enumerate(linhas):
                        if "N°" in linha:
                            nf = f"{linhas[l]} {linhas[l+2]}"
                            print(f"NF: {nf}")
                            break
                        else:
                            pass

                    
                    # DESCRIÇÃO
                    descricao = []
                    ate = j - 3
                    for k in range(1, ate):
                        descricao.append(f"{valores[k].strip()}")
            
                    desc = " ".join(descricao)

                    codigo = f"{valores[0].strip()} {prox_linha[0].strip()}"
                    quant = valores[j+1].strip()
                    valor_unit = f"{valores[j+2].strip()}"
                    valor_tot = f"{valores[j+3].strip()}"

                    # DADOS ADICIONAIS
                    for y, linha in enumerate(linhas):
                        if "INFORMAÇÕES COMPLEMENTARES" in linha:
                            coluna_dados = linhas[y+1].split("\t")
                            cont = y
                        else:
                            pass

                    linha_inicial = texto.find("INFORMAÇÕES COMPLEMENTARES")
                    
                    numero_linhas = 0
                    for caractere in texto:
                      if caractere == '\n':
                        numero_linhas += 1

                    dados_adicionais = []
                    print(f"{cont} {numero_linhas}")
                    a = 0

                    
                    for g in range(cont+1, numero_linhas+1):
                        try:   
                            print(g)
                
                            coluna_dados = linhas[g].split("\t")
                            coluna_dados = linhas[g].strip()
                            dado = coluna_dados[0]
                            dados_adicionais.append(f"{coluna_dados}")
                            print(f"{coluna_dados} {cont}")
                            a += 1
                        except Exception as e:
                            print(f"O ERRO TA AQUI: {e}")
                    dados_adc = " ".join(dados_adicionais) 
                    print(dados_adc)


                    # Armazenando dados em um dicionário
                    data_linha = {
                        "ARQUIVO": arquivo,
                        "NF": nf,
                        "Código": codigo,
                        "Descrição": desc,
                        "UNID": unid,
                        "QUANT": quant,
                        "VALOR UNIT": valor_unit,
                        "VALOR TOT": valor_tot,
                        "DADOS ADICIONAIS": dados_adc,
                    }
                    data.append(data_linha)

    except Exception as e:
        print(f"ERRO no arquivo {arquivo}: {e}")

# Criando e salvando o DataFrame
df = pd.DataFrame(data)
df.to_excel(f"{diretorio_atual}\dados.xlsx", index=False)
print("\nPlanilha dados.xlsx gerada!\n")