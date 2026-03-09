"""
extract/xml_reader.py
Extrai NF-e (modelo 55/65) — ICMS, IPI, PIS, COFINS, IBS/CBS completos.
"""

import os, re
import xml.etree.ElementTree as ET

def _t(el, tag, default=""):
    if el is None: return default
    e = el.find(tag)
    return e.text.strip() if e is not None and e.text else default

def _remover_ns(c):
    c = re.sub(r'xmlns(:[^=]+)?="[^"]+"', "", c)
    c = re.sub(r"<\?xml[^?>]*\?>", "", c)
    return c

def _parsear_xml(caminho):
    with open(caminho, "r", encoding="utf-8", errors="ignore") as f:
        c = f.read()
    s = _remover_ns(c)
    try:
        return ET.fromstring(s)
    except ET.ParseError:
        s = re.sub(r"<\?xml-stylesheet[^?>]*\?>", "", s)
        return ET.fromstring(s)

_ICMS_TAGS = (
    "ICMS00","ICMS10","ICMS20","ICMS30","ICMS40","ICMS41",
    "ICMS50","ICMS51","ICMS60","ICMS70","ICMS90",
    "ICMSSN101","ICMSSN102","ICMSSN201","ICMSSN202","ICMSSN500","ICMSSN900",
)
_PIS_TAGS    = ("PISAliq","PISQtde","PISNT","PISOutr","PISSN")
_COFINS_TAGS = ("COFINSAliq","COFINSQtde","COFINSNT","COFINSOutr","COFINSSN")

def _extrair_icms(imposto):
    r = {k:"" for k in [
        "ICMS_orig","ICMS_CST","ICMS_modBC","ICMS_vBC","ICMS_pICMS","ICMS_vICMS",
        "ICMS_vBCSTRet","ICMS_pST","ICMS_vICMSSubstituto","ICMS_vICMSSTRet",
        "ICMS_pRedBCEfet","ICMS_vBCEfet","ICMS_pICMSEfet","ICMS_vICMSEfet",
    ]}
    icms = imposto.find(".//ICMS")
    if icms is None: return r
    for filho in icms:
        if any(x in filho.tag for x in _ICMS_TAGS):
            r.update({
                "ICMS_orig":            _t(filho,"orig"),
                "ICMS_CST":             _t(filho,"CST"),
                "ICMS_modBC":           _t(filho,"modBC"),
                "ICMS_vBC":             _t(filho,"vBC"),
                "ICMS_pICMS":           _t(filho,"pICMS"),
                "ICMS_vICMS":           _t(filho,"vICMS"),
                "ICMS_vBCSTRet":        _t(filho,"vBCSTRet"),
                "ICMS_pST":             _t(filho,"pST"),
                "ICMS_vICMSSubstituto": _t(filho,"vICMSSubstituto"),
                "ICMS_vICMSSTRet":      _t(filho,"vICMSSTRet"),
                "ICMS_pRedBCEfet":      _t(filho,"pRedBCEfet"),
                "ICMS_vBCEfet":         _t(filho,"vBCEfet"),
                "ICMS_pICMSEfet":       _t(filho,"pICMSEfet"),
                "ICMS_vICMSEfet":       _t(filho,"vICMSEfet"),
            })
            break
    return r

def _extrair_ipi(imposto):
    r = {"IPI_cEnq":"","IPI_CST":"","IPI_vBC":"","IPI_pIPI":"","IPI_vIPI":""}
    ipi = imposto.find(".//IPI")
    if ipi is None: return r
    r["IPI_cEnq"] = _t(ipi,"cEnq")
    # pode estar em IPITrib ou IPINT
    for sub in ("IPITrib","IPINT"):
        el = ipi.find(sub)
        if el is not None:
            r["IPI_CST"]  = _t(el,"CST")
            r["IPI_vBC"]  = _t(el,"vBC")
            r["IPI_pIPI"] = _t(el,"pIPI")
            r["IPI_vIPI"] = _t(el,"vIPI")
            break
    return r

def _extrair_pis(imposto):
    r = {"PIS_CST":"","PIS_vBC":"","PIS_pPIS":"","PIS_vPIS":""}
    pis = imposto.find(".//PIS")
    if pis is None: return r
    for filho in pis:
        if any(x in filho.tag for x in _PIS_TAGS):
            r.update({
                "PIS_CST":  _t(filho,"CST"),
                "PIS_vBC":  _t(filho,"vBC"),
                "PIS_pPIS": _t(filho,"pPIS"),
                "PIS_vPIS": _t(filho,"vPIS"),
            })
            break
    return r

def _extrair_cofins(imposto):
    r = {"COFINS_CST":"","COFINS_vBC":"","COFINS_pCOFINS":"","COFINS_vCOFINS":""}
    cofins = imposto.find(".//COFINS")
    if cofins is None: return r
    for filho in cofins:
        if any(x in filho.tag for x in _COFINS_TAGS):
            r.update({
                "COFINS_CST":     _t(filho,"CST"),
                "COFINS_vBC":     _t(filho,"vBC"),
                "COFINS_pCOFINS": _t(filho,"pCOFINS"),
                "COFINS_vCOFINS": _t(filho,"vCOFINS"),
            })
            break
    return r

def _extrair_ibscbs(imposto, det):
    r = {k:"" for k in [
        "IBS_CST","cClassTrib","IBS_vBC",
        "pIBSUF","vIBSUF","pIBSMun","vIBSMun",
        "IBS_vIBS","pCBS","CBS_vCBS",
    ]}
    ibscbs = imposto.find(".//IBSCBS") or imposto.find(".//IBS")
    if ibscbs is None:
        vcbs = imposto.find(".//vCBS") or det.find(".//vCBS")
        if vcbs is not None and vcbs.text:
            r["CBS_vCBS"] = vcbs.text.strip()
        return r

    r["IBS_CST"]    = _t(ibscbs,"CST")
    r["cClassTrib"] = _t(ibscbs,"cClassTrib")

    g    = ibscbs.find(".//gIBSCBS")
    base = g if g is not None else ibscbs

    vbc = base.find(".//vBC")
    r["IBS_vBC"] = vbc.text.strip() if vbc is not None and vbc.text else _t(base,"vBC")

    guf = base.find(".//gIBSUF")
    if guf is not None:
        r["pIBSUF"] = _t(guf,"pIBSUF")
        r["vIBSUF"] = _t(guf,"vIBSUF")

    gmun = base.find(".//gIBSMun")
    if gmun is not None:
        r["pIBSMun"] = _t(gmun,"pIBSMun")
        r["vIBSMun"] = _t(gmun,"vIBSMun")

    vibs = base.find(".//vIBS")
    r["IBS_vIBS"] = vibs.text.strip() if vibs is not None and vibs.text else _t(base,"vIBS")

    gcbs = base.find(".//gCBS")
    if gcbs is not None:
        r["pCBS"]     = _t(gcbs,"pCBS")
        r["CBS_vCBS"] = _t(gcbs,"vCBS")
    else:
        vcbs = base.find(".//vCBS") or det.find(".//vCBS")
        if vcbs is not None and vcbs.text:
            r["CBS_vCBS"] = vcbs.text.strip()

    return r


def extrair_produtos(caminho_xml):
    produtos = []
    try:
        root = _parsear_xml(caminho_xml)
        nfe  = root.find(".//NFe")
        if nfe is None:
            nfe = root if "NFe" in root.tag else None
        if nfe is None:
            return [], "ERRO: Tag NFe nao encontrada"

        inf = nfe.find(".//infNFe")
        if inf is None:
            return [], "ERRO: Tag infNFe nao encontrada"

        ide     = inf.find(".//ide")
        emit    = inf.find(".//emit")
        dest    = inf.find(".//dest")
        endemit = inf.find(".//enderEmit")
        enddest = inf.find(".//enderDest")

        chave    = inf.get("Id","").replace("NFe","")
        num_nfe  = _t(ide,"nNF")
        serie    = _t(ide,"serie")
        dh_emi   = _t(ide,"dhEmi")
        nat_op   = _t(ide,"natOp")
        mod      = _t(ide,"mod")
        tp_nf    = _t(ide,"tpNF")

        cnpj_emit = _t(emit,"CNPJ")
        nome_emit = _t(emit,"xNome")
        fant_emit = _t(emit,"xFant")
        ie_emit   = _t(emit,"IE")
        uf_emit   = _t(endemit,"UF") if endemit is not None else _t(emit,"UF")
        mun_emit  = _t(endemit,"xMun") if endemit is not None else ""

        cnpj_dest = _t(dest,"CNPJ") if dest is not None else ""
        cpf_dest  = _t(dest,"CPF")  if dest is not None else ""
        nome_dest = _t(dest,"xNome") if dest is not None else ""
        ie_dest   = _t(dest,"IE")    if dest is not None else ""
        uf_dest   = _t(enddest,"UF")   if enddest is not None else ""
        mun_dest  = _t(enddest,"xMun") if enddest is not None else ""

        lista_det = inf.findall(".//det")
        if not lista_det:
            return [], "AVISO: Nenhum produto (det) encontrado"

        for det in lista_det:
            prod    = det.find(".//prod")
            if prod is None: continue
            imposto = det.find(".//imposto")

            d = {
                "Tipo_Nota":         "NF-e",
                "Chave_NFe":         chave,
                "Numero_NFe":        num_nfe,
                "Serie_NFe":         serie,
                "Mod_NFe":           mod,
                "NatOp":             nat_op,
                "Tp_NF":             tp_nf,
                "Data_Emissao":      dh_emi,
                "CNPJ_Emitente":     cnpj_emit,
                "Nome_Emitente":     nome_emit,
                "NomeFantasia_Emit": fant_emit,
                "IE_Emitente":       ie_emit,
                "UF_Emitente":       uf_emit,
                "Mun_Emitente":      mun_emit,
                "CNPJ_Destinatario": cnpj_dest,
                "CPF_Destinatario":  cpf_dest,
                "Nome_Destinatario": nome_dest,
                "IE_Destinatario":   ie_dest,
                "UF_Destinatario":   uf_dest,
                "Mun_Destinatario":  mun_dest,
                "Item":              det.get("nItem",""),
                "cProd":             _t(prod,"cProd"),
                "cEAN":              _t(prod,"cEAN"),
                "xProd":             _t(prod,"xProd"),
                "NCM":               _t(prod,"NCM"),
                "CEST":              _t(prod,"CEST"),
                "CFOP":              _t(prod,"CFOP"),
                "uCom":              _t(prod,"uCom"),
                "qCom":              _t(prod,"qCom"),
                "vUnCom":            _t(prod,"vUnCom"),
                "vProd":             _t(prod,"vProd"),
                "indEscala":         _t(prod,"indEscala"),
                "nFCI":              _t(prod,"nFCI"),
                # impostos zerados por padrão
                "ICMS_orig":"","ICMS_CST":"","ICMS_modBC":"",
                "ICMS_vBC":"","ICMS_pICMS":"","ICMS_vICMS":"",
                "ICMS_vBCSTRet":"","ICMS_pST":"","ICMS_vICMSSubstituto":"",
                "ICMS_vICMSSTRet":"","ICMS_pRedBCEfet":"","ICMS_vBCEfet":"",
                "ICMS_pICMSEfet":"","ICMS_vICMSEfet":"",
                "IPI_cEnq":"","IPI_CST":"","IPI_vBC":"","IPI_pIPI":"","IPI_vIPI":"",
                "PIS_CST":"","PIS_vBC":"","PIS_pPIS":"","PIS_vPIS":"",
                "COFINS_CST":"","COFINS_vBC":"","COFINS_pCOFINS":"","COFINS_vCOFINS":"",
                "IBS_CST":"","cClassTrib":"","IBS_vBC":"",
                "pIBSUF":"","vIBSUF":"","pIBSMun":"","vIBSMun":"",
                "IBS_vIBS":"","pCBS":"","CBS_vCBS":"",
                "Arquivo_Origem":    os.path.basename(caminho_xml),
            }

            if imposto is not None:
                d.update(_extrair_icms(imposto))
                d.update(_extrair_ipi(imposto))
                d.update(_extrair_pis(imposto))
                d.update(_extrair_cofins(imposto))
                d.update(_extrair_ibscbs(imposto, det))

            produtos.append(d)

        msg = f"Encontrados {len(produtos)} produto(s)"
        return (produtos, msg) if produtos else ([], "Nenhum produto encontrado")

    except ET.ParseError as e:
        return [], f"ERRO XML: {str(e)[:80]}"
    except Exception as e:
        return [], f"ERRO: {str(e)[:80]}"
