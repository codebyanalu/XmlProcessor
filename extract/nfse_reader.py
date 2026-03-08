"""
extract/nfse_reader.py
Extrai NFS-e em dois formatos:
  - CompNFe  (formato municipal legado: MeuResiduo, Onfly, Vogel)
  - NFSe     (padrão nacional SPED/Fazenda: MJOTA, LEXPLAN, DCN, etc.)
Retorna lista de dicts compatível com CABECALHO_NFSE.
"""

import os
import re
import xml.etree.ElementTree as ET


def _t(el, tag, default=""):
    if el is None:
        return default
    e = el.find(tag)
    return e.text.strip() if e is not None and e.text else default

def _remover_ns(conteudo):
    conteudo = re.sub(r'xmlns(:[^=]+)?="[^"]+"', "", conteudo)
    conteudo = re.sub(r"<\?xml[^?>]*\?>", "", conteudo)
    return conteudo

def _parsear(caminho):
    with open(caminho, "r", encoding="utf-8", errors="ignore") as f:
        conteudo = f.read()
    sem_ns = _remover_ns(conteudo)
    try:
        return ET.fromstring(sem_ns)
    except ET.ParseError:
        sem_ns = re.sub(r"<\?xml-stylesheet[^?>]*\?>", "", sem_ns)
        return ET.fromstring(sem_ns)


# ── Formato CompNFe (legado municipal) ────────────────────────────────────────

def _extrair_compnfe(root, arquivo):
    nfe = root.find(".//NFe")
    if nfe is None:
        return None, "ERRO: Tag NFe nao encontrada em CompNFe"

    prest  = nfe.find(".//Prestador")
    tomad  = nfe.find(".//Tomador")
    ibscbs = nfe.find(".//IBSCBS")
    valores_ibs = ibscbs.find(".//valores") if ibscbs is not None else None

    d = {
        "Tipo_Nota":         "NFS-e",
        "Formato":           "CompNFe",
        "Chave_NFSe":        _t(nfe, "CodigoVerificador"),
        "Numero_NFSe":       _t(nfe, "NumeroNFe"),
        "Serie_RPS":         _t(nfe, "SerieRPS"),
        "Data_Emissao":      _t(nfe, "DataEmissaoNFe"),
        "Data_Competencia":  _t(nfe, "DataCompetenciaNFe"),
        "Municipio_Prestacao": _t(nfe, "MunicipioPrestacao"),
        # Tributação ISS
        "Cod_Servico_Mun":   _t(nfe, "CodigoServicoMunicipal"),
        "Desc_Servico":      _t(nfe, "DescricaoServicoMunicipal"),
        "Cod_Item_Lei116":   _t(nfe, "CodigoItemLei116"),
        "Cod_NBS":           _t(nfe, "CodigoNBS"),
        "ISS_Retido":        _t(nfe, "ISSRetido"),
        "BC_ISS":            _t(nfe, "BaseCalculoISS"),
        "Aliq_ISS":          _t(nfe, "AliquotaIss"),
        "Valor_ISS":         _t(nfe, "ValorISS"),
        # CSRF
        "BC_CSRF":           _t(nfe, "BaseCalculoCSRF"),
        "Aliq_PIS":          _t(nfe, "AliquotaPIS"),
        "Valor_PIS":         _t(nfe, "ValorPIS"),
        "Aliq_COFINS":       _t(nfe, "AliquotaCOFINS"),
        "Valor_COFINS":      _t(nfe, "ValorCOFINS"),
        "Aliq_CSLL":         _t(nfe, "AliquotaCSLL"),
        "Valor_CSLL":        _t(nfe, "ValorCSLL"),
        "BC_IRRF":           _t(nfe, "BaseCalculoIRRF"),
        "Aliq_IRRF":         _t(nfe, "AliquotaIRRF"),
        "Valor_IRRF":        _t(nfe, "ValorIRRF"),
        "BC_INSS":           _t(nfe, "BaseCalculoINSS"),
        "Aliq_INSS":         _t(nfe, "AliquotaINSS"),
        "Valor_INSS":        _t(nfe, "ValorINSS"),
        # Valores
        "Valor_Bruto":       _t(nfe, "ValorNFe"),
        "Valor_Liquido":     _t(nfe, "ValorLiquidoNFe"),
        "Discriminacao":     _t(nfe, "Discriminacao")[:200].replace("\n"," "),
        # IBS/CBS
        "IBS_vBC":           _t(valores_ibs, "vBC") if valores_ibs is not None else "",
        "IBS_pIBSUF":        "",
        "IBS_pIBSMun":       "",
        "CBS_pCBS":          "",
        # Prestador
        "CNPJ_Prestador":    _t(prest, "CnpjCpf") if prest is not None else "",
        "Nome_Prestador":    _t(prest, "RazaoSocialNome") if prest is not None else "",
        "Mun_Prestador":     "",
        "UF_Prestador":      "",
        # Tomador
        "CNPJ_Tomador":      _t(tomad, "CnpjCpf") if tomad is not None else "",
        "Nome_Tomador":      _t(tomad, "RazaoSocialNome") if tomad is not None else "",
        "Mun_Tomador":       _t(tomad, "Municipio") if tomad is not None else "",
        "UF_Tomador":        _t(tomad, "UfSigla") if tomad is not None else "",
        # Tributos nacionais (DPS dentro do CompNFe, quando existir)
        "cTribNac":          "",
        "xDescServ":         "",
        "cNBS_DPS":          "",
        "pAliq_ISS":         "",
        "tpRetISSQN":        "",
        # Simples
        "Simples_Nacional":  _t(nfe, "PrestadorOptanteSimplesNacional"),
        "Arquivo_Origem":    os.path.basename(arquivo),
    }

    # IBS/CBS detalhado
    if valores_ibs is not None:
        uf  = ibscbs.find(".//uf")
        mun = ibscbs.find(".//mun")
        fed = ibscbs.find(".//fed")
        d["IBS_pIBSUF"]  = _t(uf,  "pIBSUF")  if uf  is not None else ""
        d["IBS_pIBSMun"] = _t(mun, "pIBSMun") if mun is not None else ""
        d["CBS_pCBS"]    = _t(fed, "pCBS")     if fed is not None else ""

    # DPS embutido
    dps = nfe.find(".//DPS")
    if dps is not None:
        cserv = dps.find(".//cServ")
        if cserv is not None:
            d["cTribNac"] = _t(cserv, "cTribNac")
            d["xDescServ"] = _t(cserv, "xDescServ")[:150]
            d["cNBS_DPS"]  = _t(cserv, "cNBS")
        trib_mun = dps.find(".//tribMun")
        if trib_mun is not None:
            d["pAliq_ISS"]   = _t(trib_mun, "pAliq")
            d["tpRetISSQN"]  = _t(trib_mun, "tpRetISSQN")

    return d, f"NFS-e (CompNFe) extraida: {d['Nome_Prestador'][:40]}"


# ── Formato NFSe Nacional (SPED/Fazenda) ─────────────────────────────────────

def _extrair_nfse_nacional(root, arquivo):
    inf = root.find(".//infNFSe")
    if inf is None:
        return None, "ERRO: Tag infNFSe nao encontrada"

    emit   = inf.find(".//emit")
    ender  = emit.find(".//enderNac") if emit is not None else None
    vals   = inf.find(".//valores")
    dps    = inf.find(".//DPS")
    infdps = dps.find(".//infDPS") if dps is not None else None
    toma   = infdps.find(".//toma")   if infdps is not None else None
    serv   = infdps.find(".//serv")   if infdps is not None else None
    cserv  = serv.find(".//cServ")    if serv is not None else None
    trib   = infdps.find(".//valores/trib") if infdps is not None else None
    if trib is None and infdps is not None:
        trib = infdps.find(".//trib")
    trib_mun = trib.find(".//tribMun") if trib is not None else None
    ibscbs   = infdps.find(".//IBSCBS") if infdps is not None else inf.find(".//IBSCBS")
    vals_ibs = ibscbs.find(".//valores") if ibscbs is not None else None

    # endereço tomador
    end_toma    = toma.find(".//end")    if toma is not None else None
    end_nac_t   = end_toma.find(".//endNac") if end_toma is not None else None

    d = {
        "Tipo_Nota":         "NFS-e",
        "Formato":           "NFSe Nacional",
        "Chave_NFSe":        inf.get("Id","").replace("NFS",""),
        "Numero_NFSe":       _t(inf, "nNFSe"),
        "Serie_RPS":         _t(infdps, "serie") if infdps is not None else "",
        "Data_Emissao":      _t(infdps, "dhEmi") if infdps is not None else _t(inf, "dhProc"),
        "Data_Competencia":  _t(infdps, "dCompet") if infdps is not None else "",
        "Municipio_Prestacao": _t(inf, "xLocPrestacao"),
        # Serviço
        "Cod_Servico_Mun":   "",
        "Desc_Servico":      _t(inf, "xTribNac"),
        "Cod_Item_Lei116":   "",
        "Cod_NBS":           _t(inf, "xNBS"),
        "ISS_Retido":        "",
        "BC_ISS":            _t(vals, "vBC"),
        "Aliq_ISS":          _t(vals, "pAliqAplic"),
        "Valor_ISS":         _t(vals, "vISSQN"),
        # CSRF
        "BC_CSRF":           "",
        "Aliq_PIS":          "",
        "Valor_PIS":         "",
        "Aliq_COFINS":       "",
        "Valor_COFINS":      "",
        "Aliq_CSLL":         "",
        "Valor_CSLL":        "",
        "BC_IRRF":           "",
        "Aliq_IRRF":         "",
        "Valor_IRRF":        "",
        "BC_INSS":           "",
        "Aliq_INSS":         "",
        "Valor_INSS":        "",
        # Valores
        "Valor_Bruto":       _t(infdps.find(".//vServPrest") if infdps else None, "vServ") or _t(vals, "vLiq"),
        "Valor_Liquido":     _t(vals, "vLiq"),
        "Discriminacao":     (_t(cserv, "xDescServ") if cserv is not None else "")[:200],
        # IBS/CBS
        "IBS_vBC":           _t(vals_ibs, "vBC") if vals_ibs is not None else "",
        "IBS_pIBSUF":        "",
        "IBS_pIBSMun":       "",
        "CBS_pCBS":          "",
        # Tributos municipais
        "cTribNac":          _t(cserv, "cTribNac") if cserv is not None else "",
        "xDescServ":         (_t(cserv, "xDescServ") if cserv is not None else "")[:150],
        "cNBS_DPS":          _t(cserv, "cNBS") if cserv is not None else "",
        "pAliq_ISS":         _t(trib_mun, "pAliq") if trib_mun is not None else "",
        "tpRetISSQN":        _t(trib_mun, "tpRetISSQN") if trib_mun is not None else "",
        # Prestador
        "CNPJ_Prestador":    _t(emit, "CNPJ") if emit is not None else "",
        "Nome_Prestador":    _t(emit, "xNome") if emit is not None else "",
        "Mun_Prestador":     _t(inf, "xLocEmi"),
        "UF_Prestador":      _t(ender, "UF") if ender is not None else "",
        # Tomador
        "CNPJ_Tomador":      _t(toma, "CNPJ") if toma is not None else "",
        "Nome_Tomador":      _t(toma, "xNome") if toma is not None else "",
        "Mun_Tomador":       "",
        "UF_Tomador":        "",
        # Simples
        "Simples_Nacional":  "",
        "Arquivo_Origem":    os.path.basename(arquivo),
    }

    # IBS/CBS detalhado
    if vals_ibs is not None:
        uf  = vals_ibs.find(".//uf")
        mun = vals_ibs.find(".//mun")
        fed = vals_ibs.find(".//fed")
        d["IBS_pIBSUF"]  = _t(uf,  "pIBSUF")  if uf  is not None else ""
        d["IBS_pIBSMun"] = _t(mun, "pIBSMun") if mun is not None else ""
        d["CBS_pCBS"]    = _t(fed, "pCBS")     if fed is not None else ""

    # Regime tributário do prestador
    reg = (infdps.find(".//prest/regTrib") if infdps else None)
    if reg is not None:
        d["Simples_Nacional"] = _t(reg, "opSimpNac")

    nome = d["Nome_Prestador"] or d["CNPJ_Prestador"]
    return d, f"NFS-e (Nacional) extraida: {nome[:40]}"


# ── Função pública ─────────────────────────────────────────────────────────────

def extrair_servicos(caminho_xml):
    """
    Detecta o formato da NFS-e e extrai os dados.
    Retorna (lista_de_notas, mensagem).
    Cada nota é um dict compatível com CABECALHO_NFSE.
    """
    try:
        root = _parsear(caminho_xml)
        tag_raiz = root.tag.split("}")[-1] if "}" in root.tag else root.tag

        if tag_raiz == "CompNFe" or root.find(".//CompNFe") is not None:
            dados, msg = _extrair_compnfe(root, caminho_xml)
        elif tag_raiz in ("NFSe", "nfseProc") or root.find(".//infNFSe") is not None:
            dados, msg = _extrair_nfse_nacional(root, caminho_xml)
        elif root.find(".//CompNFe") is not None:
            dados, msg = _extrair_compnfe(root.find(".//CompNFe"), caminho_xml)
        else:
            return [], f"ERRO: Formato NFS-e nao reconhecido (tag raiz: {tag_raiz})"

        if dados is None:
            return [], msg
        return [dados], msg

    except ET.ParseError as e:
        return [], f"ERRO XML: {str(e)[:80]}"
    except Exception as e:
        return [], f"ERRO: {str(e)[:80]}"
