import xmltodict
import os
import pandas as pd
import json

def pegar_infos(nome_arquivo, valores):
    print(f"Pegou as informações: {nome_arquivo}")
    with open(f"nfs/{nome_arquivo}", "rb") as arquivo_xml:
        dic_arquivo = xmltodict.parse(arquivo_xml)
        #print(json.dumps(dic_arquivo, indent=4))
        if "NFe" in dic_arquivo:
            infos_nf = dic_arquivo["NFe"]["infNFe"]
        else:
            infos_nf = dic_arquivo["nfeProc"]["NFe"]["infNFe"]
        chave_acesso = infos_nf["@Id"]
        empresa_emissora =infos_nf["emit"]["xNome"]
        numero_nota = infos_nf["ide"]["nNF"]
        serie_faturada = infos_nf["ide"]["serie"]
        data_faturamento = infos_nf["ide"]["dhEmi"]
        if "T" in data_faturamento:
            index = data_faturamento.find("T")
            data_faturamento = data_faturamento[0:index]
            data_faturamento = data_faturamento.replace("-", "/")
        nome_cliente = infos_nf["dest"]["xNome"]
        bairro_cliente = infos_nf["dest"]["enderDest"]["xBairro"]
        cidade_cliente = infos_nf["dest"]["enderDest"]["xMun"]
        valor_total_nf = infos_nf["total"]["ICMSTot"]["vNF"]
  #      valor_total_tributos = infos_nf["total"]["ICMSTot"]["vTotTrib"]
        if "vol" in infos_nf["transp"]:
            peso = infos_nf["transp"]["vol"]["pesoB"]
        else:
            peso = "Não informado"
        valores.append([chave_acesso, empresa_emissora, numero_nota, serie_faturada, nome_cliente, bairro_cliente, cidade_cliente, valor_total_nf, data_faturamento, peso])


lista_arquivos = os.listdir("nfs")

colunas = ["chave_acesso", "empresa_emissora", "numero_nota", "serie_faturada", "nome_cliente", "bairro_cliente", "cidade_cliente", "valor_total_nf", "data_faturamento", "peso"]
valores = []
for arquivo in lista_arquivos:
    pegar_infos(arquivo, valores)
 #   break 
tabela = pd.DataFrame(columns=colunas, data=valores)
tabela.to_excel("NotasFiscais.xlsx", index=False)
