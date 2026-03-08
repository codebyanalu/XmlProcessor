import atexit
import csv
import os
import shutil
import time
from datetime import datetime
import pandas as pd
from config.settings import CABECALHO_CSV, CSV_PRINCIPAL, CSV_TEMP, EXCEL_PRINCIPAL, EXCEL_TEMP, LOCK_FILE, LOCK_TTL_SECONDS, LOG_TEMP, SESSAO_ID, TEMP_DIR, TEMP_TTL_SECONDS, USUARIO_ID

def criar_lock():
    try:
        with open(LOCK_FILE, "w") as f:
            f.write(f"Sessao: {SESSAO_ID}\nUsuario: {USUARIO_ID}\nInicio: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
        return True
    except Exception:
        return False

def verificar_locks_ativos():
    try:
        meu_lock = os.path.basename(LOCK_FILE)
        agora = time.time()
        return [n for n in os.listdir(TEMP_DIR) if n.startswith("lock_") and n != meu_lock and (agora - os.path.getmtime(os.path.join(TEMP_DIR, n))) < LOCK_TTL_SECONDS]
    except Exception:
        return []

def inicializar_sessao():
    try:
        criar_lock()
        if os.path.exists(CSV_PRINCIPAL):
            shutil.copy2(CSV_PRINCIPAL, CSV_TEMP)
        else:
            _criar_csv_vazio(CSV_TEMP)
        sincronizar_excel_temp()
        with open(LOG_TEMP, "w", encoding="utf-8") as f:
            f.write(f"Sessao: {SESSAO_ID}\nUsuario: {USUARIO_ID}\nInicio: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
        return True
    except Exception as e:
        print(f"Erro ao inicializar sessao: {e}")
        return False

def _criar_csv_vazio(caminho):
    with open(caminho, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerow(CABECALHO_CSV)

def salvar_produtos_csv(produtos, caminho=CSV_TEMP):
    if not produtos:
        return True, "Nenhum produto para salvar"
    try:
        arquivo_existe = os.path.exists(caminho) and os.path.getsize(caminho) > 0
        with open(caminho, "a", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=CABECALHO_CSV)
            if not arquivo_existe:
                writer.writeheader()
            writer.writerows(produtos)
        return True, f"{len(produtos)} produto(s) salvos"
    except Exception as e:
        return False, f"Erro ao salvar CSV: {e}"

def total_registros(caminho=CSV_TEMP):
    try:
        if not os.path.exists(caminho):
            return 0
        with open(caminho, "r", encoding="utf-8") as f:
            return max(0, sum(1 for _ in f) - 1)
    except Exception:
        return 0

def _csv_para_dataframe(caminho):
    try:
        df = pd.read_csv(caminho, dtype=str, encoding="utf-8", on_bad_lines="skip")
    except Exception:
        linhas = []
        with open(caminho, "r", encoding="utf-8") as f:
            for i, linha in enumerate(csv.reader(f)):
                if i == 0:
                    linhas.append(CABECALHO_CSV)
                else:
                    linha = (linha + [""] * len(CABECALHO_CSV))[:len(CABECALHO_CSV)]
                    linhas.append(linha)
        df = pd.DataFrame(linhas[1:], columns=linhas[0]) if len(linhas) > 1 else pd.DataFrame(columns=CABECALHO_CSV)
    return df.reindex(columns=CABECALHO_CSV, fill_value="")

def _salvar_excel(df, caminho):
    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Produtos_NFe")

def sincronizar_excel_temp():
    if not os.path.exists(CSV_TEMP):
        return False
    try:
        _salvar_excel(_csv_para_dataframe(CSV_TEMP), EXCEL_TEMP)
        return True
    except Exception as e:
        print(f"Erro ao sincronizar Excel temp: {e}")
        return False

def atualizar_excel_principal():
    if not os.path.exists(CSV_PRINCIPAL):
        return False, "CSV principal nao encontrado"
    try:
        _salvar_excel(_csv_para_dataframe(CSV_PRINCIPAL), EXCEL_PRINCIPAL)
        return True, "Excel principal atualizado"
    except Exception as e:
        return False, f"Erro ao atualizar Excel principal: {e}"

def sincronizar_com_principal():
    try:
        if not os.path.exists(CSV_PRINCIPAL):
            _criar_csv_vazio(CSV_PRINCIPAL)
        if not os.path.exists(CSV_TEMP):
            return True, "Nenhum dado temporario para sincronizar"
        linhas_novas = []
        with open(CSV_TEMP, "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            next(reader, None)
            for linha in reader:
                if len(linha) == len(CABECALHO_CSV):
                    linhas_novas.append(linha)
        if not linhas_novas:
            return True, "Nenhum dado novo para sincronizar"
        chaves_existentes = set()
        with open(CSV_PRINCIPAL, "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            next(reader, None)
            for linha in reader:
                if len(linha) == len(CABECALHO_CSV):
                    chaves_existentes.add(f"{linha[0]}_{linha[8]}_{linha[9]}")
        backup = CSV_PRINCIPAL.replace(".csv", f"_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
        try:
            shutil.copy2(CSV_PRINCIPAL, backup)
        except Exception:
            pass
        adicionados = 0
        with open(CSV_PRINCIPAL, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            for linha in linhas_novas:
                chave = f"{linha[0]}_{linha[8]}_{linha[9]}"
                if chave not in chaves_existentes:
                    writer.writerow(linha)
                    chaves_existentes.add(chave)
                    adicionados += 1
        return True, f"{adicionados} registro(s) sincronizados"
    except Exception as e:
        return False, f"Erro na sincronizacao: {e}"

def limpar_temporarios():
    prefixos = ("temp_produtos_", "temp_excel_", "temp_log_", "lock_")
    for caminho in [CSV_TEMP, EXCEL_TEMP, LOG_TEMP, LOCK_FILE]:
        try:
            if os.path.exists(caminho):
                os.remove(caminho)
        except Exception:
            pass
    try:
        agora = time.time()
        for nome in os.listdir(TEMP_DIR):
            if any(nome.startswith(p) for p in prefixos):
                caminho = os.path.join(TEMP_DIR, nome)
                if os.path.isfile(caminho) and (agora - os.path.getmtime(caminho)) > TEMP_TTL_SECONDS:
                    try:
                        os.remove(caminho)
                    except Exception:
                        pass
    except Exception:
        pass

atexit.register(limpar_temporarios)
