"""
diagnostico.py — rode ENQUANTO o main.py estiver aberto, ou logo após processar.
Execute: python diagnostico.py
"""
import sys, os, hashlib, glob

print("=" * 65)
print("DIAGNÓSTICO GCON/SIAN")
print("=" * 65)

base = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, base)

# 1. Versão dos arquivos
esperados = {
    'ui/main_window.py':     ('69bb7dcd', 53645),
    'load/storage.py':       ('5f8a36d3', 11830),
    'extract/xml_reader.py': ('6ef3a2b1',  0),   # qualquer tamanho, só confere conteúdo chave
    'extract/nfse_reader.py':('0beea239', 12090),
    'config/settings.py':    ('d65da8a9',  3839),
}
print("\n1. ARQUIVOS DO PROJETO:")
for rel in esperados:
    caminho = os.path.join(base, rel)
    if not os.path.exists(caminho):
        print(f"   ✗ FALTANDO: {rel}")
        continue
    tam = os.path.getsize(caminho)
    h   = hashlib.md5(open(caminho,'rb').read()).hexdigest()[:8]
    # Verifica conteúdo chave independente de hash
    conteudo = open(caminho, encoding='utf-8', errors='ignore').read()
    if rel == 'extract/xml_reader.py':
        tem_ipi   = '_extrair_ipi' in conteudo
        tem_ibscbs= '_extrair_ibscbs' in conteudo
        ok = tem_ipi and tem_ibscbs
        detalhe = f"IPI={'OK' if tem_ipi else 'FALTA'}  IBS/CBS={'OK' if tem_ibscbs else 'FALTA'}"
    elif rel == 'ui/main_window.py':
        ok = '_ler_csv' in conteudo and 'BarChart' in conteudo
        detalhe = f"leitura CSV={'OK' if '_ler_csv' in conteudo else 'FALTA'}  gráficos={'OK' if 'BarChart' in conteudo else 'FALTA'}"
    elif rel == 'load/storage.py':
        ok = 'salvar_nfse_csv' in conteudo and '_aplicar_formatacao_excel' in conteudo
        detalhe = f"NFS-e={'OK' if 'salvar_nfse_csv' in conteudo else 'FALTA'}  formatação={'OK' if '_aplicar_formatacao_excel' in conteudo else 'FALTA'}"
    else:
        ok = True
        detalhe = f"{tam:,} bytes"
    print(f"   {'✓' if ok else '✗'} {rel:<35} {detalhe}")
    if not ok:
        print(f"          *** DESATUALIZADO — substitua pelo arquivo do Claude ***")

# 2. Arquivos temporários — mostra TODOS os existentes
print("\n2. ARQUIVOS TEMPORÁRIOS (todos na pasta):")
try:
    from config.settings import TEMP_DIR
    print(f"   Pasta: {TEMP_DIR}")
    if not os.path.exists(TEMP_DIR):
        print("   Pasta não existe ainda (normal antes de abrir o sistema)")
    else:
        csvs_nfe   = sorted(glob.glob(os.path.join(TEMP_DIR, 'temp_produtos_*.csv')))
        csvs_nfse  = sorted(glob.glob(os.path.join(TEMP_DIR, 'temp_nfse_*.csv')))
        excels_nfe = sorted(glob.glob(os.path.join(TEMP_DIR, 'temp_excel_ana*.xlsx')))
        excels_nfse= sorted(glob.glob(os.path.join(TEMP_DIR, 'temp_excel_nfse_*.xlsx')))

        def mostra(titulo, lista):
            if not lista:
                print(f"   ✗ {titulo}: nenhum encontrado")
                return
            for arq in lista[-3:]:  # últimos 3
                tam = os.path.getsize(arq)
                linhas = 0
                if arq.endswith('.csv'):
                    try:
                        linhas = sum(1 for _ in open(arq, encoding='utf-8', errors='ignore')) - 1
                    except: pass
                    print(f"   ✓ {titulo}: {os.path.basename(arq)}")
                    print(f"              {tam:,} bytes  |  {linhas} linhas de dados")
                    if linhas == 0 and tam < 500:
                        print(f"              *** CSV VAZIO — processe XMLs primeiro! ***")
                else:
                    print(f"   ✓ {titulo}: {os.path.basename(arq)}  ({tam:,} bytes)")

        mostra("CSV NF-e  ",   csvs_nfe)
        mostra("CSV NFS-e ",   csvs_nfse)
        mostra("Excel NF-e",   excels_nfe)
        mostra("Excel NFS-e",  excels_nfse)

        # Arquivo principal
        from config.settings import CSV_PRINCIPAL, CSV_NFSE_PRINCIPAL
        for nome, caminho in [("CSV NF-e PRINCIPAL", CSV_PRINCIPAL),
                               ("CSV NFS-e PRINCIPAL", CSV_NFSE_PRINCIPAL)]:
            if os.path.exists(caminho):
                tam = os.path.getsize(caminho)
                linhas = sum(1 for _ in open(caminho, encoding='utf-8', errors='ignore')) - 1
                print(f"   ✓ {nome}: {linhas} linhas  ({tam:,} bytes)")
            else:
                print(f"   ○ {nome}: ainda não criado (normal na 1ª sessão)")

except Exception as e:
    print(f"   ERRO: {e}")

# 3. Dependências
print("\n3. DEPENDÊNCIAS:")
for lib in ['pandas','openpyxl','customtkinter']:
    try:
        m = __import__(lib)
        print(f"   ✓ {lib} {getattr(m,'__version__','?')}")
    except ImportError:
        print(f"   ✗ {lib} NÃO INSTALADO — pip install {lib}")

# 4. Teste de extração com XMLs da pasta base
print("\n4. TESTE DE EXTRAÇÃO:")
try:
    from config.settings import PASTA_BASE
    print(f"   Pasta base: {PASTA_BASE}")
    arqs = glob.glob(os.path.join(PASTA_BASE, '*.xml'))
    if not arqs:
        print(f"   Nenhum XML na pasta base.")
        print(f"   Dica: selecione XMLs pela interface, não precisa colocar na pasta base.")
    else:
        from extract.xml_reader import extrair_produtos
        from extract.nfse_reader import extrair_servicos
        nfe_ok = nfse_ok = 0
        for arq in arqs[:5]:
            with open(arq,'r',encoding='utf-8',errors='ignore') as f: txt = f.read(300)
            tipo = 'nfse' if any(x in txt for x in ['CompNFe','infNFSe','nNFSe']) else 'nfe'
            if tipo == 'nfse':
                regs, msg = extrair_servicos(arq)
                nfse_ok += len(regs)
            else:
                regs, msg = extrair_produtos(arq)
                nfe_ok += len(regs)
        print(f"   Testados {min(5,len(arqs))} XMLs: {nfe_ok} produtos NF-e, {nfse_ok} notas NFS-e")
except Exception as e:
    import traceback
    print(f"   ERRO: {e}")
    traceback.print_exc()

# 5. Instrução final
print("\n" + "=" * 65)
print("PRÓXIMOS PASSOS:")
print("  1. Substitua os arquivos marcados com ✗ pelos do Claude")
print("  2. Abra o sistema (python main.py)")
print("  3. Importe os XMLs clicando em 'Selecionar XMLs'")
print("  4. Após processar, clique em 'Ver NF-e' ou 'Ver NFS-e'")
print("  OBS: as planilhas só aparecem APÓS importar XMLs na sessão")
print("=" * 65)
input("\nPressione Enter para sair...")
