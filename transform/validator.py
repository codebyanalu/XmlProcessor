import csv
import os
from config.settings import CABECALHO_CSV

def chave_produto(produto):
    return f"{produto.get('Chave_NFe','')}" f"_{produto.get('Item','')}" f"_{produto.get('cProd','')}"

def normalizar_produto(produto):
    return {col: produto.get(col, "") for col in CABECALHO_CSV}

def carregar_chaves_existentes(caminho_csv):
    chaves = set()
    if not os.path.exists(caminho_csv):
        return chaves
    try:
        with open(caminho_csv, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for linha in reader:
                chaves.add(chave_produto(linha))
    except Exception:
        pass
    return chaves

def filtrar_novos(produtos, chaves_existentes):
    novos = []
    duplicados = []
    chaves_nesta_execucao = set()
    for p in produtos:
        chave = chave_produto(p)
        produto_normalizado = normalizar_produto(p)
        if chave in chaves_existentes or chave in chaves_nesta_execucao:
            duplicados.append(produto_normalizado)
        else:
            novos.append(produto_normalizado)
            chaves_nesta_execucao.add(chave)
    return novos, duplicados
