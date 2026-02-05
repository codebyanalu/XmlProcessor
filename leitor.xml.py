#!/usr/bin/env python3
"""
SISTEMA MODERNO DE EXTRAÇÃO XML NF-e - GCON/SIAN
Extrai TODOS os produtos e impostos das NF-e
Sistema multiusuário com arquivos temporários
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import customtkinter as ctk
import threading
from datetime import datetime
import xml.etree.ElementTree as ET
import pandas as pd
import openpyxl
import re
import csv
import shutil
import tempfile
import uuid
import atexit
import time

# Configurar CustomTkinter
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Pasta base para arquivos compartilhados
PASTA_BASE = r"C:\Users\ana.oliveira\Downloads\Códigos\leitor.xml"
if not os.path.exists(PASTA_BASE):
    os.makedirs(PASTA_BASE, exist_ok=True)

# Pasta temporária para uso multiusuário
TEMP_DIR = os.path.join(tempfile.gettempdir(), "leitor_xml_multiusuario")
if not os.path.exists(TEMP_DIR):
    os.makedirs(TEMP_DIR, exist_ok=True)

# Gerar ID único para esta sessão
SESSAO_ID = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:8]}"
USUARIO_ID = os.environ.get('USERNAME', 'usuario_desconhecido')

# Arquivos principais (compartilhados)
CSV_PRINCIPAL = os.path.join(PASTA_BASE, "produtos_nfe.csv")
EXCEL_PRINCIPAL = os.path.join(PASTA_BASE, "produtos_nfe.xlsx")
LOG_PRINCIPAL = os.path.join(PASTA_BASE, "log_processamento.txt")

# Arquivos temporários (por usuário/sessão)
CSV_TEMP = os.path.join(TEMP_DIR, f"temp_produtos_{USUARIO_ID}_{SESSAO_ID}.csv")
EXCEL_TEMP = os.path.join(TEMP_DIR, f"temp_excel_{USUARIO_ID}_{SESSAO_ID}.xlsx")
LOG_TEMP = os.path.join(TEMP_DIR, f"temp_log_{USUARIO_ID}_{SESSAO_ID}.txt")
LOCK_FILE = os.path.join(TEMP_DIR, f"lock_{USUARIO_ID}_{SESSAO_ID}.lock")

# Função para limpar arquivos temporários ao sair
def limpar_arquivos_temporarios():
    """Remove todos os arquivos temporários desta sessão"""
    try:
        # Remover arquivos específicos da sessão
        temp_files = [CSV_TEMP, EXCEL_TEMP, LOG_TEMP, LOCK_FILE]
        for temp_file in temp_files:
            if os.path.exists(temp_file):
                try:
                    os.remove(temp_file)
                    print(f"Arquivo temporário removido: {temp_file}")
                except Exception as e:
                    print(f"Erro ao remover {temp_file}: {e}")
        
        # Limpar arquivos temporários antigos (mais de 1 hora)
        try:
            agora = time.time()
            for filename in os.listdir(TEMP_DIR):
                filepath = os.path.join(TEMP_DIR, filename)
                if os.path.isfile(filepath):
                    # Verificar se é um arquivo temporário deste sistema
                    if filename.startswith(("temp_produtos_", "temp_excel_", "temp_log_", "lock_")):
                        # Verificar idade do arquivo (mais de 1 hora)
                        idade = agora - os.path.getmtime(filepath)
                        if idade > 3600:  # 1 hora em segundos
                            try:
                                os.remove(filepath)
                                print(f"Arquivo temporário antigo removido: {filename}")
                            except:
                                pass
        except Exception as e:
            print(f"Erro ao limpar arquivos antigos: {e}")
            
    except Exception as e:
        print(f"Erro geral na limpeza de temporários: {e}")

# Registrar limpeza ao sair
atexit.register(limpar_arquivos_temporarios)

# Cabeçalho do CSV com formato ORIGINAL
CABECALHO_CSV = [
    # Dados da Nota fiscal
    'Chave_NFe', 'Numero_NFe', 'Serie_NFe', 'Data_Emissao', 
    'CNPJ_Emitente', 'Nome_Emitente', 
    'CNPJ_Destinatario', 'Nome_Destinatario',
    
    # Dados do Produto
    'Item', 'cProd', 'xProd', 'NCM', 'CFOP', 'uCom', 'qCom', 'vUnCom', 'vProd',
    
    # ICMS
    'ICMS_CST', 'ICMS_vBC', 'ICMS_pICMS', 'ICMS_vICMS',
    
    # PIS
    'PIS_CST', 'PIS_vBC', 'PIS_pPIS', 'PIS_vPIS',
    
    # COFINS
    'COFINS_CST', 'COFINS_vBC', 'COFINS_pCOFINS', 'COFINS_vCOFINS',
    
    # IBS/CBS
    'IBS_CST', 'cClassTrib', 'IBS_vBC', 'IBS_vIBS', 'CBS_vCBS',
    
    # Informações adicionais
    'Arquivo_Origem'
]

def criar_lock():
    """Cria arquivo de lock para esta sessão"""
    try:
        with open(LOCK_FILE, 'w') as f:
            f.write(f"Sessão: {SESSAO_ID}\n")
            f.write(f"Usuário: {USUARIO_ID}\n")
            f.write(f"Início: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
        return True
    except:
        return False

def verificar_lock_ativos():
    """Verifica se há outras sessões ativas"""
    try:
        if not os.path.exists(TEMP_DIR):
            return []
        
        locks_ativos = []
        for filename in os.listdir(TEMP_DIR):
            if filename.startswith("lock_") and filename != os.path.basename(LOCK_FILE):
                lock_path = os.path.join(TEMP_DIR, filename)
                # Verificar se o lock é recente (menos de 5 minutos)
                if os.path.exists(lock_path):
                    idade = time.time() - os.path.getmtime(lock_path)
                    if idade < 300:  # 5 minutos
                        locks_ativos.append(filename)
        
        return locks_ativos
    except:
        return []

def sincronizar_com_principal():
    """Sincroniza os dados temporários com o arquivo principal"""
    try:
        # Se não existir CSV principal, criar vazio
        if not os.path.exists(CSV_PRINCIPAL):
            with open(CSV_PRINCIPAL, "w", newline='', encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(CABECALHO_CSV)
        
        # Se existe CSV temporário, mesclar com principal
        if os.path.exists(CSV_TEMP):
            # Ler ambos os arquivos
            linhas_novas = []
            with open(CSV_TEMP, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                cabecalho = next(reader, None)
                for linha in reader:
                    if len(linha) == len(CABECALHO_CSV):
                        linhas_novas.append(linha)
            
            if linhas_novas:
                # Ler linhas existentes no principal
                linhas_existentes = set()
                if os.path.exists(CSV_PRINCIPAL):
                    with open(CSV_PRINCIPAL, 'r', encoding='utf-8') as f:
                        reader = csv.reader(f)
                        next(reader, None)  # Pular cabeçalho
                        for linha in reader:
                            if len(linha) == len(CABECALHO_CSV):
                                # Criar chave única para comparação
                                chave = f"{linha[0]}_{linha[8]}_{linha[9]}"  # Chave_NFe + Item + cProd
                                linhas_existentes.add(chave)
                
                # Adicionar apenas linhas novas
                with open(CSV_PRINCIPAL, 'a', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    for linha in linhas_novas:
                        chave = f"{linha[0]}_{linha[8]}_{linha[9]}"
                        if chave not in linhas_existentes:
                            writer.writerow(linha)
                
                # Criar backup do CSV principal
                backup_path = CSV_PRINCIPAL.replace('.csv', f'_backup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv')
                try:
                    shutil.copy2(CSV_PRINCIPAL, backup_path)
                except:
                    pass
                
                return True, f"{len(linhas_novas)} registros sincronizados"
        
        return True, "Nenhum dado novo para sincronizar"
        
    except Exception as e:
        return False, f"Erro na sincronização: {str(e)}"

def inicializar_arquivos_temporarios():
    """Inicializa arquivos temporários para esta sessão"""
    try:
        # Criar lock para esta sessão
        criar_lock()
        
        # Se existe CSV principal, copiar para temporário
        if os.path.exists(CSV_PRINCIPAL):
            shutil.copy2(CSV_PRINCIPAL, CSV_TEMP)
        else:
            # Criar CSV temporário vazio
            with open(CSV_TEMP, "w", newline='', encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(CABECALHO_CSV)
        
        # Criar Excel temporário sincronizado com CSV temporário
        sincronizar_excel_temp()
        
        # Criar log temporário
        with open(LOG_TEMP, "w", encoding="utf-8") as f:
            f.write(f"Sessão: {SESSAO_ID}\n")
            f.write(f"Usuário: {USUARIO_ID}\n")
            f.write(f"Início: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
        
        return True
    except Exception as e:
        print(f"Erro inicializando arquivos temporários: {e}")
        return False

def sincronizar_excel_temp():
    """Sincroniza o Excel temporário com o CSV temporário"""
    try:
        if not os.path.exists(CSV_TEMP):
            return False
        
        # Ler CSV temporário
        try:
            df = pd.read_csv(CSV_TEMP, dtype=str, encoding='utf-8', on_bad_lines='skip')
        except:
            # Tentar ler manualmente se houver erro
            linhas = []
            with open(CSV_TEMP, 'r', encoding='utf-8') as f:
                leitor = csv.reader(f)
                for i, linha in enumerate(leitor):
                    if i == 0:
                        linhas.append(CABECALHO_CSV)
                    elif len(linha) == len(CABECALHO_CSV):
                        linhas.append(linha)
                    else:
                        if len(linha) > len(CABECALHO_CSV):
                            linha = linha[:len(CABECALHO_CSV)]
                        else:
                            while len(linha) < len(CABECALHO_CSV):
                                linha.append('')
                        linhas.append(linha)
            
            df = pd.DataFrame(linhas[1:], columns=linhas[0]) if len(linhas) > 1 else pd.DataFrame(columns=CABECALHO_CSV)
        
        # Garantir colunas corretas
        df = df.reindex(columns=CABECALHO_CSV, fill_value='')
        
        # Salvar Excel temporário
        with pd.ExcelWriter(EXCEL_TEMP, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Produtos_NFe')
        
        return True
        
    except Exception as e:
        print(f"Erro sincronizando Excel temporário: {e}")
        return False

def extrair_produtos_completos(caminho_xml):
    """Função que extrai TODOS os produtos e impostos de uma NF-e"""
    produtos = []
    
    try:
        with open(caminho_xml, 'r', encoding='utf-8', errors='ignore') as f:
            conteudo = f.read()
        
        conteudo_sem_ns = conteudo
        conteudo_sem_ns = re.sub(r'xmlns(:[^=]+)?="[^"]+"', '', conteudo_sem_ns)
        conteudo_sem_ns = re.sub(r'<\?xml[^?>]*\?>', '', conteudo_sem_ns)
        
        try:
            root = ET.fromstring(conteudo_sem_ns)
        except ET.ParseError:
            conteudo_sem_ns = re.sub(r'<\?xml-stylesheet[^?>]*\?>', '', conteudo_sem_ns)
            root = ET.fromstring(conteudo_sem_ns)
        
        nfe_node = root.find('.//NFe')
        if nfe_node is None:
            if 'NFe' in root.tag:
                nfe_node = root
            else:
                return produtos, "ERRO: Tag NFe não encontrada"
        
        inf_nfe = nfe_node.find('.//infNFe')
        if inf_nfe is None:
            return produtos, "ERRO: Tag infNFe não encontrada"
        
        def extrair_texto(elemento, tag, default=""):
            if elemento is None:
                return default
            elem = elemento.find(tag)
            if elem is not None and elem.text:
                return elem.text.strip()
            return default
        
        ide = inf_nfe.find('.//ide')
        emit = inf_nfe.find('.//emit')
        dest = inf_nfe.find('.//dest')
        
        chave_nfe = inf_nfe.get('Id', '').replace('NFe', '')
        numero_nfe = extrair_texto(ide, 'nNF')
        serie_nfe = extrair_texto(ide, 'serie')
        data_emissao = extrair_texto(ide, 'dhEmi')
        cnpj_emitente = extrair_texto(emit, 'CNPJ')
        nome_emitente = extrair_texto(emit, 'xNome')
        cnpj_destinatario = extrair_texto(dest, 'CNPJ')
        nome_destinatario = extrair_texto(dest, 'xNome')
        
        lista_det = inf_nfe.findall('.//det')
        
        if not lista_det:
            return produtos, "AVISO: Nenhum produto (det) encontrado"
        
        for det_element in lista_det:
            item = det_element.get("nItem", "")
            
            prod = det_element.find('.//prod')
            if prod is None:
                continue
            
            imposto = det_element.find('.//imposto')
            
            dados_produto = {
                'Chave_NFe': chave_nfe,
                'Numero_NFe': numero_nfe,
                'Serie_NFe': serie_nfe,
                'Data_Emissao': data_emissao,
                'CNPJ_Emitente': cnpj_emitente,
                'Nome_Emitente': nome_emitente,
                'CNPJ_Destinatario': cnpj_destinatario,
                'Nome_Destinatario': nome_destinatario,
                
                'Item': item,
                'cProd': extrair_texto(prod, 'cProd'),
                'xProd': extrair_texto(prod, 'xProd'),
                'NCM': extrair_texto(prod, 'NCM'),
                'CFOP': extrair_texto(prod, 'CFOP'),
                'uCom': extrair_texto(prod, 'uCom'),
                'qCom': extrair_texto(prod, 'qCom'),
                'vUnCom': extrair_texto(prod, 'vUnCom'),
                'vProd': extrair_texto(prod, 'vProd'),
                
                'ICMS_CST': '', 'ICMS_vBC': '', 'ICMS_pICMS': '', 'ICMS_vICMS': '',
                'PIS_CST': '', 'PIS_vBC': '', 'PIS_pPIS': '', 'PIS_vPIS': '',
                'COFINS_CST': '', 'COFINS_vBC': '', 'COFINS_pCOFINS': '', 'COFINS_vCOFINS': '',
                'IBS_CST': '', 'cClassTrib': '', 'IBS_vBC': '', 'IBS_vIBS': '', 'CBS_vCBS': '',
                'Arquivo_Origem': os.path.basename(caminho_xml)
            }
            
            if imposto is not None:
                icms = imposto.find('.//ICMS')
                if icms is not None:
                    for icms_tipo in icms:
                        tipo_tag = icms_tipo.tag
                        if any(x in tipo_tag for x in ['ICMS00', 'ICMS10', 'ICMS20', 'ICMS30', 
                                                      'ICMS40', 'ICMS41', 'ICMS50', 'ICMS51', 
                                                      'ICMS60', 'ICMS70', 'ICMS90', 'ICMSSN101',
                                                      'ICMSSN102', 'ICMSSN201', 'ICMSSN202',
                                                      'ICMSSN500', 'ICMSSN900']):
                            dados_produto['ICMS_CST'] = extrair_texto(icms_tipo, 'CST')
                            dados_produto['ICMS_vBC'] = extrair_texto(icms_tipo, 'vBC')
                            dados_produto['ICMS_pICMS'] = extrair_texto(icms_tipo, 'pICMS')
                            dados_produto['ICMS_vICMS'] = extrair_texto(icms_tipo, 'vICMS')
                            break
                
                pis = imposto.find('.//PIS')
                if pis is not None:
                    for pis_tipo in pis:
                        tipo_tag = pis_tipo.tag
                        if any(x in tipo_tag for x in ['PISAliq', 'PISQtde', 'PISNT', 'PISOutr', 'PISSN']):
                            dados_produto['PIS_CST'] = extrair_texto(pis_tipo, 'CST')
                            dados_produto['PIS_vBC'] = extrair_texto(pis_tipo, 'vBC')
                            dados_produto['PIS_pPIS'] = extrair_texto(pis_tipo, 'pPIS')
                            dados_produto['PIS_vPIS'] = extrair_texto(pis_tipo, 'vPIS')
                            break
                
                cofins = imposto.find('.//COFINS')
                if cofins is not None:
                    for cofins_tipo in cofins:
                        tipo_tag = cofins_tipo.tag
                        if any(x in tipo_tag for x in ['COFINSAliq', 'COFINSQtde', 'COFINSNT', 'COFINSOutr', 'COFINSSN']):
                            dados_produto['COFINS_CST'] = extrair_texto(cofins_tipo, 'CST')
                            dados_produto['COFINS_vBC'] = extrair_texto(cofins_tipo, 'vBC')
                            dados_produto['COFINS_pCOFINS'] = extrair_texto(cofins_tipo, 'pCOFINS')
                            dados_produto['COFINS_vCOFINS'] = extrair_texto(cofins_tipo, 'vCOFINS')
                            break
                
                ibscbs = imposto.find('.//IBSCBS')
                
                if ibscbs is None:
                    ibscbs = imposto.find('.//IBS')
                
                if ibscbs is None:
                    vcbs_elem = imposto.find('.//vCBS')
                    if vcbs_elem is not None and vcbs_elem.text:
                        dados_produto['CBS_vCBS'] = vcbs_elem.text.strip()
                else:
                    dados_produto['IBS_CST'] = extrair_texto(ibscbs, 'CST')
                    dados_produto['cClassTrib'] = extrair_texto(ibscbs, 'cClassTrib')
                    
                    vbc_elem = ibscbs.find('.//vBC')
                    if vbc_elem is not None and vbc_elem.text:
                        dados_produto['IBS_vBC'] = vbc_elem.text.strip()
                    else:
                        dados_produto['IBS_vBC'] = extrair_texto(ibscbs, 'vBC')
                    
                    vibs_elem = ibscbs.find('.//vIBS')
                    if vibs_elem is not None and vibs_elem.text:
                        dados_produto['IBS_vIBS'] = vibs_elem.text.strip()
                    else:
                        dados_produto['IBS_vIBS'] = extrair_texto(ibscbs, 'vIBS')
                    
                    vcbs_elem = ibscbs.find('.//vCBS')
                    if vcbs_elem is not None and vcbs_elem.text:
                        dados_produto['CBS_vCBS'] = vcbs_elem.text.strip()
                    else:
                        for elem in ibscbs.iter():
                            if 'vCBS' in elem.tag and elem.text:
                                dados_produto['CBS_vCBS'] = elem.text.strip()
                                break
                
                if not dados_produto['CBS_vCBS']:
                    vcbs_elem = det_element.find('.//vCBS')
                    if vcbs_elem is not None and vcbs_elem.text:
                        dados_produto['CBS_vCBS'] = vcbs_elem.text.strip()
            
            produtos.append(dados_produto)
        
        if produtos:
            return produtos, f"Encontrados {len(produtos)} produto(s)"
        else:
            return produtos, "Nenhum produto encontrado"
                
    except ET.ParseError as e:
        return [], f"ERRO XML: Arquivo XML mal formado - {str(e)[:80]}"
    except Exception as e:
        return [], f"ERRO: {str(e)[:80]}"

def produto_existe_no_csv(produto, arquivo_csv_path=CSV_TEMP):
    """Verifica se um produto já existe no CSV temporário"""
    try:
        if not os.path.exists(arquivo_csv_path):
            return False
        
        chave = f"{produto.get('Chave_NFe', '')}_{produto.get('Item', '')}_{produto.get('cProd', '')}"
        if not chave or chave == "__":
            return False
        
        with open(arquivo_csv_path, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for linha in reader:
                linha_chave = f"{linha.get('Chave_NFe', '')}_{linha.get('Item', '')}_{linha.get('cProd', '')}"
                if linha_chave == chave:
                    return True
        
        return False
    except Exception as e:
        print(f"Erro verificando produto: {e}")
        return False

def adicionar_produto_ao_csv(produto, arquivo_csv_path=CSV_TEMP):
    """Adiciona um produto ao CSV temporário"""
    try:
        arquivo_existe = os.path.exists(arquivo_csv_path)
        
        with open(arquivo_csv_path, 'a', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=CABECALHO_CSV)
            
            if not arquivo_existe or f.tell() == 0:
                writer.writeheader()
            
            writer.writerow(produto)
        
        return True, "Produto adicionado com sucesso"
    except Exception as e:
        return False, f"Erro ao adicionar produto: {str(e)}"

def obter_total_registros(arquivo_csv_path=CSV_TEMP):
    """Retorna o total de registros no CSV temporário"""
    try:
        if os.path.exists(arquivo_csv_path):
            with open(arquivo_csv_path, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                return max(0, sum(1 for row in reader) - 1)
        return 0
    except:
        return 0

class AplicacaoLeitorXML:
    def __init__(self):
        self.janela = ctk.CTk()
        self.janela.title(f"Sistema de Extração NF-e - GCON/SIAN - Usuário: {USUARIO_ID}")
        self.janela.geometry("1200x850")
        self.janela.minsize(1100, 750)
        
        self.janela.protocol("WM_DELETE_WINDOW", self.fechar_aplicacao)
        
        self.janela.grid_columnconfigure(0, weight=1)
        self.janela.grid_rowconfigure(1, weight=1)
        
        self.processando = False
        self.arquivos_selecionados = []
        
        # Verificar se há outras sessões ativas
        locks_ativos = verificar_lock_ativos()
        if locks_ativos:
            resposta = messagebox.askyesno(
                "Sessões Ativas", 
                f"Existem {len(locks_ativos)} outras sessões ativas.\n"
                f"Deseja continuar mesmo assim?"
            )
            if not resposta:
                sys.exit(0)
        
        # Inicializar arquivos temporários
        if not inicializar_arquivos_temporarios():
            messagebox.showerror("Erro", "Não foi possível inicializar os arquivos temporários!")
            return
        
        self.criar_interface()
        self.log_inicial()
    
    def fechar_aplicacao(self):
        """Fecha a aplicação com limpeza de recursos"""
        if self.processando:
            if not messagebox.askyesno("Processamento em Andamento", 
                                      "Há um processamento em andamento.\n"
                                      "Deseja realmente fechar a aplicação?"):
                return
        
        # Sincronizar dados antes de fechar
        self.sincronizar_com_principal_final()
        
        # Limpar arquivos temporários
        limpar_arquivos_temporarios()
        
        self.janela.destroy()
        sys.exit(0)
    
    def sincronizar_com_principal_final(self):
        """Sincroniza dados temporários com principal ao fechar"""
        try:
            self.log("")
            self.log_divider("=")
            self.log_centered("SINCRONIZAÇÃO FINAL")
            self.log_divider("=")
            
            sucesso, mensagem = sincronizar_com_principal()
            if sucesso:
                self.log(f"Sincronização concluída: {mensagem}", "success")
                
                # Atualizar Excel principal
                if os.path.exists(CSV_PRINCIPAL):
                    try:
                        df = pd.read_csv(CSV_PRINCIPAL, dtype=str, encoding='utf-8', on_bad_lines='skip')
                        df = df.reindex(columns=CABECALHO_CSV, fill_value='')
                        with pd.ExcelWriter(EXCEL_PRINCIPAL, engine='openpyxl') as writer:
                            df.to_excel(writer, index=False, sheet_name='Produtos_NFe')
                        self.log("Excel principal atualizado", "success")
                    except Exception as e:
                        self.log(f"Erro ao atualizar Excel principal: {str(e)[:50]}", "warning")
            else:
                self.log(f"Erro na sincronização: {mensagem}", "error")
            
            self.log_divider("=")
            
        except Exception as e:
            print(f"Erro na sincronização final: {e}")
    
    def criar_interface(self):
        main_frame = ctk.CTkFrame(self.janela)
        main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        main_frame.grid_columnconfigure(0, weight=1)
        
        header_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        header_frame.grid(row=0, column=0, padx=0, pady=(0, 20), sticky="ew")
        header_frame.grid_columnconfigure(0, weight=1)
        
        title_label = ctk.CTkLabel(
            header_frame,
            text="SISTEMA DE EXTRAÇÃO COMPLETA NF-e (MULTIUSUÁRIO)",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.grid(row=0, column=0, pady=(0, 5), sticky="w")
        
        subtitle_label = ctk.CTkLabel(
            header_frame,
            text=f"GCON/SIAN - Usuário: {USUARIO_ID} - Sessão: {SESSAO_ID[:15]}...",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        subtitle_label.grid(row=1, column=0, sticky="w")
        
        botoes_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        botoes_frame.grid(row=1, column=0, padx=0, pady=(0, 20), sticky="ew")
        
        for i in range(6):
            botoes_frame.grid_columnconfigure(i, weight=1)
        
        self.btn_um_xml = ctk.CTkButton(
            botoes_frame,
            text="Selecionar 1 XML",
            command=self.selecionar_um_xml,
            height=40,
            font=ctk.CTkFont(size=14),
            corner_radius=10
        )
        self.btn_um_xml.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        self.btn_varios_xml = ctk.CTkButton(
            botoes_frame,
            text="Selecionar Vários XMLs",
            command=self.selecionar_varios_xml,
            height=40,
            font=ctk.CTkFont(size=14),
            corner_radius=10
        )
        self.btn_varios_xml.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        self.btn_visualizar_excel = ctk.CTkButton(
            botoes_frame,
            text="Visualizar Dados",
            command=self.visualizar_excel,
            height=40,
            font=ctk.CTkFont(size=14),
            corner_radius=10,
            fg_color="#2ecc71",
            hover_color="#27ae60"
        )
        self.btn_visualizar_excel.grid(row=0, column=2, padx=5, pady=5, sticky="ew")
        
        self.btn_abrir_excel = ctk.CTkButton(
            botoes_frame,
            text="Abrir Excel Local",
            command=self.abrir_excel_local,
            height=40,
            font=ctk.CTkFont(size=14),
            corner_radius=10,
            fg_color="#3498db",
            hover_color="#2980b9"
        )
        self.btn_abrir_excel.grid(row=0, column=3, padx=5, pady=5, sticky="ew")
        
        self.btn_sincronizar = ctk.CTkButton(
            botoes_frame,
            text="Sincronizar",
            command=self.sincronizar_manual,
            height=40,
            font=ctk.CTkFont(size=14),
            corner_radius=10,
            fg_color="#9b59b6",
            hover_color="#8e44ad"
        )
        self.btn_sincronizar.grid(row=0, column=4, padx=5, pady=5, sticky="ew")
        
        self.btn_fechar = ctk.CTkButton(
            botoes_frame,
            text="Fechar",
            command=self.fechar_aplicacao,
            height=40,
            font=ctk.CTkFont(size=14),
            corner_radius=10,
            fg_color="#e74c3c",
            hover_color="#c0392b"
        )
        self.btn_fechar.grid(row=0, column=5, padx=5, pady=5, sticky="ew")
        
        status_frame = ctk.CTkFrame(main_frame)
        status_frame.grid(row=2, column=0, padx=0, pady=(0, 20), sticky="ew")
        status_frame.grid_columnconfigure(0, weight=1)
        
        status_title = ctk.CTkLabel(
            status_frame,
            text="Status do Sistema",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        status_title.grid(row=0, column=0, padx=15, pady=(15, 10), sticky="w")
        
        self.lbl_status = ctk.CTkLabel(
            status_frame,
            text="PRONTO - Aguardando seleção de arquivos",
            font=ctk.CTkFont(size=14),
            text_color="green"
        )
        self.lbl_status.grid(row=1, column=0, padx=15, pady=(0, 5), sticky="w")
        
        self.lbl_contador = ctk.CTkLabel(
            status_frame,
            text="Arquivos selecionados: 0",
            font=ctk.CTkFont(size=12)
        )
        self.lbl_contador.grid(row=2, column=0, padx=15, pady=(0, 5), sticky="w")
        
        info_frame = ctk.CTkFrame(status_frame, fg_color="transparent")
        info_frame.grid(row=3, column=0, padx=15, pady=(0, 15), sticky="w")
        
        total_registros = obter_total_registros()
        self.lbl_total = ctk.CTkLabel(
            info_frame,
            text=f"Produtos nesta sessão: {total_registros}",
            font=ctk.CTkFont(size=12),
            text_color="#3498db"
        )
        self.lbl_total.pack(side="left", padx=(0, 20))
        
        self.lbl_sessao = ctk.CTkLabel(
            info_frame,
            text=f"Sessão: {SESSAO_ID[:10]}...",
            font=ctk.CTkFont(size=10),
            text_color="gray"
        )
        self.lbl_sessao.pack(side="left", padx=(0, 20))
        
        progress_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        progress_frame.grid(row=3, column=0, padx=0, pady=(0, 20), sticky="ew")
        progress_frame.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(
            progress_frame,
            text="Progresso:",
            font=ctk.CTkFont(size=14)
        ).grid(row=0, column=0, padx=(0, 10), sticky="w")
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ctk.CTkProgressBar(
            progress_frame,
            variable=self.progress_var,
            height=20,
            corner_radius=10
        )
        self.progress_bar.grid(row=0, column=1, padx=(0, 10), sticky="ew")
        self.progress_bar.set(0)
        
        self.lbl_progress = ctk.CTkLabel(
            progress_frame,
            text="0%",
            font=ctk.CTkFont(size=14, weight="bold"),
            width=50
        )
        self.lbl_progress.grid(row=0, column=2, sticky="e")
        
        log_frame = ctk.CTkFrame(main_frame)
        log_frame.grid(row=4, column=0, padx=0, pady=(0, 0), sticky="nsew")
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(1, weight=1)
        
        log_title_frame = ctk.CTkFrame(log_frame, fg_color="transparent")
        log_title_frame.grid(row=0, column=0, padx=15, pady=15, sticky="ew")
        log_title_frame.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(
            log_title_frame,
            text="Log de Processamento",
            font=ctk.CTkFont(size=16, weight="bold")
        ).grid(row=0, column=0, sticky="w")
        
        btn_limpar = ctk.CTkButton(
            log_title_frame,
            text="Limpar",
            command=self.limpar_log,
            width=80,
            height=30,
            font=ctk.CTkFont(size=12),
            corner_radius=8,
            fg_color="#e74c3c",
            hover_color="#c0392b"
        )
        btn_limpar.grid(row=0, column=1, padx=(10, 5))
        
        btn_salvar = ctk.CTkButton(
            log_title_frame,
            text="Salvar",
            command=self.salvar_log,
            width=80,
            height=30,
            font=ctk.CTkFont(size=12),
            corner_radius=8,
            fg_color="#3498db",
            hover_color="#2980b9"
        )
        btn_salvar.grid(row=0, column=2, padx=5)
        
        text_frame = ctk.CTkFrame(log_frame)
        text_frame.grid(row=1, column=0, padx=15, pady=(0, 15), sticky="nsew")
        text_frame.grid_columnconfigure(0, weight=1)
        text_frame.grid_rowconfigure(0, weight=1)
        
        self.txt_log = scrolledtext.ScrolledText(
            text_frame,
            height=20,
            wrap=tk.WORD,
            font=("Consolas", 12),
            bg='#2c3e50',
            fg='#ecf0f1',
            insertbackground='white',
            relief="flat"
        )
        self.txt_log.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        
        self.txt_log.tag_config("success", foreground="#27ae60", font=("Consolas", 12, "bold"))
        self.txt_log.tag_config("error", foreground="#e74c3c", font=("Consolas", 12, "bold"))
        self.txt_log.tag_config("warning", foreground="#f39c12", font=("Consolas", 12, "bold"))
        self.txt_log.tag_config("info", foreground="#3498db", font=("Consolas", 12))
        self.txt_log.tag_config("header", foreground="#2c3e50", background="#ecf0f1", font=("Consolas", 13, "bold"))
        self.txt_log.tag_config("border", foreground="#95a5a6", font=("Consolas", 12))
        self.txt_log.tag_config("file", foreground="#3498db", font=("Consolas", 12, "bold"))
        self.txt_log.tag_config("value", foreground="#bdc3c7", font=("Consolas", 12))
        self.txt_log.tag_config("summary", foreground="#2c3e50", background="#ecf0f1", font=("Consolas", 13, "bold"))
        self.txt_log.tag_config("timestamp", foreground="#7f8c8d", font=("Consolas", 10))
        self.txt_log.tag_config("item", foreground="#ecf0f1", font=("Consolas", 12))
    
    def sincronizar_manual(self):
        """Sincronização manual dos dados"""
        if self.processando:
            messagebox.showwarning("Aguarde", "Há um processamento em andamento!")
            return
        
        self.log("")
        self.log_divider("=")
        self.log_centered("SINCRONIZAÇÃO MANUAL")
        self.log_divider("=")
        
        sucesso, mensagem = sincronizar_com_principal()
        if sucesso:
            self.log(f"Sincronização concluída: {mensagem}", "success")
            
            # Atualizar Excel principal
            if os.path.exists(CSV_PRINCIPAL):
                try:
                    df = pd.read_csv(CSV_PRINCIPAL, dtype=str, encoding='utf-8', on_bad_lines='skip')
                    df = df.reindex(columns=CABECALHO_CSV, fill_value='')
                    with pd.ExcelWriter(EXCEL_PRINCIPAL, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name='Produtos_NFe')
                    self.log("Excel principal atualizado", "success")
                except Exception as e:
                    self.log(f"Erro ao atualizar Excel principal: {str(e)[:50]}", "warning")
            
            messagebox.showinfo("Sincronização", f"Sincronização concluída!\n{mensagem}")
        else:
            self.log(f"Erro na sincronização: {mensagem}", "error")
            messagebox.showerror("Erro", f"Falha na sincronização:\n{mensagem}")
        
        self.log_divider("=")
    
    def abrir_excel_local(self):
        """Abre o Excel temporário local"""
        try:
            if not os.path.exists(EXCEL_TEMP):
                messagebox.showwarning("Aviso", "Arquivo Excel local não encontrado!")
                return
            
            # Sincronizar antes de abrir
            if sincronizar_excel_temp():
                self.log("Excel local sincronizado", "success")
            
            os.startfile(EXCEL_TEMP)
            self.log("Excel local aberto", "info")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir Excel local:\n{str(e)}")
            self.log(f"Erro ao abrir Excel local: {str(e)}", "error")
    
    def limpar_log(self):
        self.txt_log.delete(1.0, tk.END)
        self.log("Log limpo", "info")
    
    def salvar_log(self):
        try:
            arquivo = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Arquivos de texto", "*.txt"), ("Todos os arquivos", "*.*")],
                initialfile=f"log_processamento_{USUARIO_ID}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            )
            
            if arquivo:
                with open(arquivo, 'w', encoding='utf-8') as f:
                    f.write(self.txt_log.get(1.0, tk.END))
                self.log(f"Log salvo em: {arquivo}", "success")
        except Exception as e:
            self.log(f"Erro ao salvar log: {str(e)}", "error")
    
    def log_inicial(self):
        self.txt_log.delete(1.0, tk.END)
        self.log_divider("=")
        self.log_centered("SISTEMA DE EXTRAÇÃO COMPLETA NF-e - MULTIUSUÁRIO")
        self.log_centered(f"Sessão: {SESSAO_ID}")
        self.log_centered(f"Usuário: {USUARIO_ID} - Data/Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        self.log_divider("=")
        self.log("")
        
        total_registros = obter_total_registros()
        self.log(f"Sistema pronto para processar arquivos XML", "success")
        self.log(f"Pasta base compartilhada: {PASTA_BASE}", "info")
        self.log(f"Pasta temporária: {TEMP_DIR}", "info")
        self.log(f"Produtos nesta sessão: {total_registros}", "info")
        self.log("") 
    
    def atualizar_total_registros(self):
        total_registros = obter_total_registros()
        self.lbl_total.configure(text=f"Produtos nesta sessão: {total_registros}")
    
    def log(self, mensagem, tag="info"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.txt_log.insert(tk.END, f"[{timestamp}] ", "timestamp")
        self.txt_log.insert(tk.END, f"{mensagem}\n", tag)
        self.txt_log.see(tk.END)
        self.janela.update_idletasks()
    
    def log_divider(self, char="─"):
        self.txt_log.insert(tk.END, f"{char * 100}\n", "border")
        self.txt_log.see(tk.END)
        self.janela.update_idletasks()
    
    def log_centered(self, mensagem, tag="info"):
        largura = 100
        padding = (largura - len(mensagem)) // 2
        if padding < 0:
            padding = 0
        espacada = f"{' ' * padding}{mensagem}{' ' * padding}"
        if len(mensagem) < largura:
            espacada += " "
        self.txt_log.insert(tk.END, f"{espacada}\n", tag)
        self.txt_log.see(tk.END)
        self.janela.update_idletasks()
    
    def log_arquivo(self, nome_arquivo, mensagem, tag="info"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.txt_log.insert(tk.END, f"[{timestamp}] ", "timestamp")
        self.txt_log.insert(tk.END, f"Arquivo: {nome_arquivo}\n", "file")
        self.txt_log.insert(tk.END, f"[{timestamp}]     ", "timestamp")
        
        if "ERRO" in mensagem:
            tag = "error"
        elif "AVISO" in mensagem:
            tag = "warning"
        elif "Encontrados" in mensagem:
            tag = "success"
            
        self.txt_log.insert(tk.END, f"{mensagem}\n", tag)
        self.txt_log.see(tk.END)
        self.janela.update_idletasks()
    
    def log_produto_detalhado(self, produto, indice):
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        self.txt_log.insert(tk.END, f"[{timestamp}]       ", "timestamp")
        self.txt_log.insert(tk.END, f"Produto {indice}:\n", "info")
        
        self.txt_log.insert(tk.END, f"[{timestamp}]         ", "timestamp")
        self.txt_log.insert(tk.END, f"Código: {produto.get('cProd', '')} | ", "value")
        descricao = produto.get('xProd', '')[:40]
        self.txt_log.insert(tk.END, f"Descrição: {descricao}\n", "value")
        
        self.txt_log.insert(tk.END, f"[{timestamp}]         ", "timestamp")
        self.txt_log.insert(tk.END, f"NCM: {produto.get('NCM', '')} | ", "value")
        self.txt_log.insert(tk.END, f"CFOP: {produto.get('CFOP', '')} | ", "value")
        self.txt_log.insert(tk.END, f"Qtd: {produto.get('qCom', '')} | Valor: R$ {produto.get('vProd', '0.00')}\n", "value")
        
        impostos_info = []
        
        if produto.get('ICMS_CST'):
            icms_info = f"ICMS CST: {produto.get('ICMS_CST')}"
            if produto.get('ICMS_vICMS') and produto.get('ICMS_vICMS') != '0.00':
                icms_info += f" | Valor: R$ {produto.get('ICMS_vICMS')}"
            if produto.get('ICMS_pICMS'):
                icms_info += f" | Alíquota: {produto.get('ICMS_pICMS')}%"
            impostos_info.append(icms_info)
        
        if produto.get('PIS_CST'):
            pis_info = f"PIS CST: {produto.get('PIS_CST')}"
            if produto.get('PIS_vPIS') and produto.get('PIS_vPIS') != '0.00':
                pis_info += f" | Valor: R$ {produto.get('PIS_vPIS')}"
            if produto.get('PIS_pPIS'):
                pis_info += f" | Alíquota: {produto.get('PIS_pPIS')}%"
            impostos_info.append(pis_info)
        
        if produto.get('COFINS_CST'):
            cofins_info = f"COFINS CST: {produto.get('COFINS_CST')}"
            if produto.get('COFINS_vCOFINS') and produto.get('COFINS_vCOFINS') != '0.00':
                cofins_info += f" | Valor: R$ {produto.get('COFINS_vCOFINS')}"
            if produto.get('COFINS_pCOFINS'):
                cofins_info += f" | Alíquota: {produto.get('COFINS_pCOFINS')}%"
            impostos_info.append(cofins_info)
        
        cclass = produto.get('cClassTrib', '')
        if cclass:
            impostos_info.append(f"cClassTrib: {cclass}")
        
        if produto.get('IBS_vBC'):
            ibs_info = f"IBS BC: R$ {produto.get('IBS_vBC', '0.00')}"
            if produto.get('IBS_vIBS'):
                ibs_info += f" | Valor IBS: R$ {produto.get('IBS_vIBS', '0.00')}"
            impostos_info.append(ibs_info)
        
        cbs_valor = produto.get('CBS_vCBS', '')
        if cbs_valor:
            impostos_info.append(f"CBS: R$ {cbs_valor}")
        
        if impostos_info:
            for info in impostos_info:
                self.txt_log.insert(tk.END, f"[{timestamp}]         ", "timestamp")
                self.txt_log.insert(tk.END, f"{info}\n", "value")
        
        self.txt_log.see(tk.END)
        self.janela.update_idletasks()
    
    def atualizar_status(self, texto, cor="green"):
        cores = {
            "green": "green",
            "orange": "orange",
            "red": "red",
            "blue": "#3498db"
        }
        self.lbl_status.configure(text=texto, text_color=cores.get(cor, "green"))
        self.janela.update_idletasks()
    
    def atualizar_contador(self, quantidade):
        self.lbl_contador.configure(text=f"Arquivos selecionados: {quantidade}")
    
    def atualizar_progresso(self, atual, total):
        if total > 0:
            percent = (atual / total) * 100
            self.progress_var.set(percent / 100)
            self.lbl_progress.configure(text=f"{percent:.0f}%")
            self.janela.update_idletasks()
    
    def selecionar_um_xml(self):
        if self.processando:
            messagebox.showwarning("Aguarde", "Já há um processamento em andamento!")
            return
        
        arquivo = filedialog.askopenfilename(
            title="Selecione um arquivo XML",
            filetypes=[("Arquivos XML", "*.xml"), ("Todos os arquivos", "*.*")]
        )
        
        if arquivo:
            self.arquivos_selecionados = [arquivo]
            self.atualizar_contador(1)
            self.log("")
            self.log_divider("─")
            self.log_arquivo(os.path.basename(arquivo), "Arquivo selecionado para processamento", "info")
            threading.Thread(target=self.processar_arquivos, daemon=True).start()
    
    def selecionar_varios_xml(self):
        if self.processando:
            messagebox.showwarning("Aguarde", "Já há um processamento em andamento!")
            return
        
        arquivos = filedialog.askopenfilenames(
            title="Selecione múltiplos arquivos XML",
            filetypes=[("Arquivos XML", "*.xml"), ("Todos os arquivos", "*.*")]
        )
        
        if arquivos:
            self.arquivos_selecionados = list(arquivos)
            self.atualizar_contador(len(arquivos))
            self.log("")
            self.log_divider("─")
            self.log(f"{len(arquivos)} arquivos selecionados para processamento", "info")
            threading.Thread(target=self.processar_arquivos, daemon=True).start()
    
    def processar_arquivos(self):
        self.processando = True
        self.atualizar_status("PROCESSANDO...", "orange")
        
        self.btn_um_xml.configure(state="disabled")
        self.btn_varios_xml.configure(state="disabled")
        self.btn_visualizar_excel.configure(state="disabled")
        self.btn_abrir_excel.configure(state="disabled")
        self.btn_sincronizar.configure(state="disabled")
        self.btn_fechar.configure(state="disabled")
        
        total_arquivos = len(self.arquivos_selecionados)
        arquivos_com_erro = []
        produtos_adicionados = 0
        produtos_duplicados = 0
        produtos_com_erro = 0
        total_produtos_encontrados = 0
        
        self.log("")
        self.log_divider("=")
        self.log_centered("INÍCIO DO PROCESSAMENTO")
        self.log_divider("=")
        self.log("")
        self.log(f"Arquivos a processar: {total_arquivos}", "info")
        self.log("")
        
        self.atualizar_progresso(0, total_arquivos)
        
        for i, arquivo in enumerate(self.arquivos_selecionados, 1):
            self.atualizar_progresso(i, total_arquivos)
            nome_arquivo = os.path.basename(arquivo)
            
            try:
                produtos, mensagem = extrair_produtos_completos(arquivo)
                
                if mensagem.startswith("ERRO"):
                    self.log_arquivo(nome_arquivo, f"{mensagem}", "error")
                    arquivos_com_erro.append(f"{nome_arquivo}: {mensagem[10:]}")
                    continue
                elif mensagem.startswith("AVISO"):
                    self.log_arquivo(nome_arquivo, f"{mensagem}", "warning")
                    continue
                
                total_produtos = len(produtos)
                total_produtos_encontrados += total_produtos
                
                self.log_arquivo(nome_arquivo, f"{mensagem}", "success")
                
                if produtos:
                    for j, produto in enumerate(produtos, 1):
                        try:
                            if produto_existe_no_csv(produto):
                                self.log_produto_detalhado(produto, j)
                                self.log(f"[{datetime.now().strftime('%H:%M:%S')}]         Produto já existe no sistema", "warning")
                                produtos_duplicados += 1
                            else:
                                adicionado, msg = adicionar_produto_ao_csv(produto)
                                
                                if adicionado:
                                    self.log_produto_detalhado(produto, j)
                                    self.log(f"[{datetime.now().strftime('%H:%M:%S')}]         Adicionado com sucesso", "success")
                                    produtos_adicionados += 1
                                else:
                                    self.log_produto_detalhado(produto, j)
                                    self.log(f"[{datetime.now().strftime('%H:%M:%S')}]         Erro: {msg[:50]}", "error")
                                    produtos_com_erro += 1
                                    
                        except Exception as e:
                            self.log(f"[{datetime.now().strftime('%H:%M:%S')}]         Erro ao processar produto: {str(e)[:50]}", "error")
                            produtos_com_erro += 1
                
                if produtos:
                    self.log("")
                    
            except Exception as e:
                self.log_arquivo(nome_arquivo, f"ERRO inesperado: {str(e)[:80]}", "error")
                arquivos_com_erro.append(f"{nome_arquivo}: {str(e)[:50]}")
        
        # Sincronizar Excel temporário após processamento
        try:
            if sincronizar_excel_temp():
                self.log("Excel local sincronizado", "success")
        except Exception as e:
            self.log(f"Erro ao sincronizar Excel local: {str(e)[:50]}", "warning")
        
        self.log("")
        self.log_divider("=")
        self.log_centered("RESUMO DO PROCESSAMENTO")
        self.log_divider("=")
        self.log("")
        
        self.atualizar_total_registros()
        
        self.log(f"TOTAL DE ARQUIVOS: {total_arquivos}", "info")
        self.log(f"PRODUTOS ENCONTRADOS: {total_produtos_encontrados}", "info")
        self.log("")
        
        if produtos_adicionados > 0:
            self.log(f"PRODUTOS ADICIONADOS: {produtos_adicionados}", "success")
        
        if produtos_duplicados > 0:
            self.log(f"PRODUTOS DUPLICADOS: {produtos_duplicados}", "warning")
        
        if produtos_com_erro > 0:
            self.log(f"PRODUTOS COM ERRO: {produtos_com_erro}", "error")
        
        if arquivos_com_erro:
            self.log(f"ARQUIVOS COM ERRO: {len(arquivos_com_erro)}", "error")
        
        self.log("")
        
        total_final = obter_total_registros()
        self.log(f"TOTAL NESTA SESSÃO: {total_final} produtos", "info")
        
        self.log("")
        self.log_divider("=")
        
        if arquivos_com_erro:
            self.log("")
            self.log("Detalhes dos erros:", "warning")
            for erro in arquivos_com_erro[:3]:
                self.log(f"   • {erro}", "warning")
            if len(arquivos_com_erro) > 3:
                self.log(f"   • ... e mais {len(arquivos_com_erro) - 3} outros", "warning")
            self.log("")
        
        if arquivos_com_erro:
            self.atualizar_status(f"CONCLUÍDO COM {len(arquivos_com_erro)} ERRO(S)", "red")
        elif produtos_adicionados == 0 and produtos_duplicados > 0:
            self.atualizar_status("CONCLUÍDO (todos já existiam)", "orange")
        elif produtos_adicionados > 0:
            self.atualizar_status("PROCESSAMENTO CONCLUÍDO", "green")
        else:
            self.atualizar_status("NENHUM PRODUTO ENCONTRADO", "orange")
        
        mensagem_final = ""
        if arquivos_com_erro:
            if len(arquivos_com_erro) == total_arquivos:
                mensagem_final = f"TODOS os {total_arquivos} arquivos tiveram erro!\nVerifique se são XMLs válidos."
                messagebox.showerror("Erro Grave", mensagem_final)
            else:
                mensagem_final = f"PROCESSAMENTO CONCLUÍDO COM ERROS\n\nArquivos: {total_arquivos}\nProdutos encontrados: {total_produtos_encontrados}\nAdicionados: {produtos_adicionados}\nDuplicados: {produtos_duplicados}\nErros: {len(arquivos_com_erro)}"
                messagebox.showwarning("Processamento Parcial", mensagem_final)
        elif produtos_adicionados == 0 and produtos_duplicados > 0:
            mensagem_final = f"PROCESSAMENTO CONCLUÍDO\n\nArquivos: {total_arquivos}\nProdutos encontrados: {total_produtos_encontrados}\nTodos os produtos já existiam no sistema\nNenhum novo produto adicionado."
            messagebox.showinfo("Concluído", mensagem_final)
        elif produtos_adicionados > 0:
            mensagem_final = f"PROCESSAMENTO CONCLUÍDO COM SUCESSO!\n\nArquivos: {total_arquivos}\nProdutos encontrados: {total_produtos_encontrados}\nAdicionados: {produtos_adicionados}\nDuplicados: {produtos_duplicados}\nTotal nesta sessão: {total_final}"
            messagebox.showinfo("Concluído", mensagem_final)
        else:
            mensagem_final = f"PROCESSAMENTO CONCLUÍDO\n\nArquivos: {total_arquivos}\nNenhum produto encontrado nos XMLs."
            messagebox.showinfo("Concluído", mensagem_final)
        
        self.processando = False
        self.arquivos_selecionados = []
        self.atualizar_contador(0)
        self.atualizar_progresso(0, 1)
        
        self.btn_um_xml.configure(state="normal")
        self.btn_varios_xml.configure(state="normal")
        self.btn_visualizar_excel.configure(state="normal")
        self.btn_abrir_excel.configure(state="normal")
        self.btn_sincronizar.configure(state="normal")
        self.btn_fechar.configure(state="normal")
    
    def visualizar_excel(self):
        try:
            if not os.path.exists(EXCEL_TEMP):
                messagebox.showwarning("Aviso", "Arquivo Excel local não encontrado!\nExecute algum processamento primeiro.")
                return
            
            if sincronizar_excel_temp():
                self.log("Excel local sincronizado antes da visualização", "success")
            
            df = pd.read_excel(EXCEL_TEMP, dtype=str)
            
            if df.empty or len(df) == 0:
                messagebox.showinfo("Informação", "O Excel está vazio!")
                return
            
            janela_excel = ctk.CTkToplevel(self.janela)
            janela_excel.title(f"Visualizar Dados - {len(df)} produtos (Sessão: {SESSAO_ID[:10]}...)")
            janela_excel.geometry("1400x800")
            janela_excel.transient(self.janela)
            janela_excel.grab_set()
            
            janela_excel.grid_columnconfigure(0, weight=1)
            janela_excel.grid_rowconfigure(1, weight=1)
            
            header_frame = ctk.CTkFrame(janela_excel)
            header_frame.grid(row=0, column=0, padx=20, pady=20, sticky="ew")
            header_frame.grid_columnconfigure(0, weight=1)
            
            ctk.CTkLabel(
                header_frame,
                text=f"Visualização de Dados - {len(df)} produtos (Sessão Local)",
                font=ctk.CTkFont(size=18, weight="bold")
            ).grid(row=0, column=0, pady=10, sticky="w")
            
            tree_frame = ctk.CTkFrame(janela_excel)
            tree_frame.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="nsew")
            tree_frame.grid_columnconfigure(0, weight=1)
            tree_frame.grid_rowconfigure(0, weight=1)
            
            style = ttk.Style()
            style.theme_use('default')
            style.configure("Treeview",
                background="#2a2d2e",
                foreground="white",
                rowheight=25,
                fieldbackground="#2a2d2e",
                bordercolor="#343638",
                borderwidth=0)
            style.map('Treeview', background=[('selected', '#22559b')])
            
            style.configure("Treeview.Heading",
                background="#565b5e",
                foreground="white",
                relief="flat")
            style.map("Treeview.Heading",
                background=[('active', '#3484F0')])
            
            tree = ttk.Treeview(tree_frame, style="Treeview")
            
            scroll_y = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
            scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
            tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
            
            colunas_principais = ['#', 'Chave_NFe', 'Item', 'cProd', 'xProd', 'NCM', 'CFOP', 
                                 'ICMS_CST', 'ICMS_vICMS', 
                                 'PIS_CST', 'PIS_vPIS', 
                                 'COFINS_CST', 'COFINS_vCOFINS', 
                                 'cClassTrib', 'IBS_vBC', 'IBS_vIBS', 'CBS_vCBS']
            
            tree['columns'] = colunas_principais
            tree['show'] = 'headings'
            
            larguras = {
                '#': 50, 'Chave_NFe': 150, 'Item': 60, 'cProd': 100, 'xProd': 200, 
                'NCM': 80, 'CFOP': 80,
                'ICMS_CST': 80, 'ICMS_vICMS': 100,
                'PIS_CST': 80, 'PIS_vPIS': 100,
                'COFINS_CST': 90, 'COFINS_vCOFINS': 110,
                'cClassTrib': 100, 'IBS_vBC': 100, 'IBS_vIBS': 100, 'CBS_vCBS': 100
            }
            
            for col in colunas_principais:
                tree.heading(col, text=col)
                tree.column(col, width=larguras.get(col, 100), anchor=tk.W)
            
            for i, (_, row) in enumerate(df.iterrows(), 1):
                valores = [
                    str(i),
                    row.get('Chave_NFe', '')[:20],
                    row.get('Item', ''),
                    row.get('cProd', ''),
                    str(row.get('xProd', ''))[:30],
                    row.get('NCM', ''),
                    row.get('CFOP', ''),
                    row.get('ICMS_CST', ''),
                    row.get('ICMS_vICMS', ''),
                    row.get('PIS_CST', ''),
                    row.get('PIS_vPIS', ''),
                    row.get('COFINS_CST', ''),
                    row.get('COFINS_vCOFINS', ''),
                    row.get('cClassTrib', ''),
                    row.get('IBS_vBC', ''),
                    row.get('IBS_vIBS', ''),
                    row.get('CBS_vCBS', '')
                ]
                tree.insert("", tk.END, values=valores)
            
            tree.grid(row=0, column=0, sticky="nsew")
            scroll_y.grid(row=0, column=1, sticky="ns")
            scroll_x.grid(row=1, column=0, sticky="ew")
            
            btn_frame = ctk.CTkFrame(janela_excel, fg_color="transparent")
            btn_frame.grid(row=2, column=0, padx=20, pady=(0, 20), sticky="ew")
            
            ctk.CTkButton(
                btn_frame,
                text="Copiar Seleção",
                command=lambda: self.copiar_selecao(tree),
                width=150,
                height=35,
                font=ctk.CTkFont(size=12),
                corner_radius=8
            ).pack(side="left", padx=5)
            
            ctk.CTkButton(
                btn_frame,
                text="Exportar para CSV",
                command=self.exportar_csv,
                width=150,
                height=35,
                font=ctk.CTkFont(size=12),
                corner_radius=8,
                fg_color="#2ecc71",
                hover_color="#27ae60"
            ).pack(side="left", padx=5)
            
            ctk.CTkButton(
                btn_frame,
                text="Fechar",
                command=janela_excel.destroy,
                width=150,
                height=35,
                font=ctk.CTkFont(size=12),
                corner_radius=8,
                fg_color="#e74c3c",
                hover_color="#c0392b"
            ).pack(side="left", padx=5)
            
            self.log(f"Visualizando dados: {len(df)} produtos", "info")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao visualizar dados:\n{str(e)}")
    
    def copiar_selecao(self, tree):
        try:
            selecionados = tree.selection()
            if not selecionados:
                messagebox.showinfo("Informação", "Nenhum item selecionado!")
                return
            
            textos = []
            for item in selecionados:
                valores = tree.item(item, 'values')
                textos.append("\t".join(str(v) for v in valores))
            
            self.janela.clipboard_clear()
            self.janela.clipboard_append("\n".join(textos))
            self.log("Seleção copiada para a área de transferência", "success")
        except Exception as e:
            self.log(f"Erro ao copiar seleção: {str(e)}", "error")
    
    def exportar_csv(self):
        try:
            arquivo = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("Arquivos CSV", "*.csv"), ("Todos os arquivos", "*.*")],
                initialfile=f"produtos_{USUARIO_ID}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            )
            
            if arquivo:
                df = pd.read_excel(EXCEL_TEMP, dtype=str)
                df.to_csv(arquivo, index=False, encoding='utf-8-sig')
                self.log(f"Dados exportados para: {arquivo}", "success")
                messagebox.showinfo("Sucesso", f"Dados exportados com sucesso!\n{arquivo}")
        except Exception as e:
            self.log(f"Erro ao exportar CSV: {str(e)}", "error")
    
    def run(self):
        self.janela.mainloop()

def main():
    try:
        import pandas as pd
        import openpyxl
        
        try:
            import customtkinter as ctk
        except ImportError:
            print("Instalando CustomTkinter...")
            import subprocess
            subprocess.check_call([sys.executable, "-m", "pip", "install", "customtkinter"])
            import customtkinter as ctk
        
        app = AplicacaoLeitorXML()
        app.run()
        
    except ImportError as e:
        print(f"Erro de dependência: {e}")
        print("\nInstale as dependências com:")
        print("pip install pandas openpyxl customtkinter")
        input("\nPressione Enter para sair...")
    except Exception as e:
        print(f"Erro inesperado: {e}")
        import traceback
        traceback.print_exc()
        input("\nPressione Enter para sair...")

if __name__ == "__main__":
    main()