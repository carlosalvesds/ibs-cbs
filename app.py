import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
from io import BytesIO

def extrair_dados_fiscais(arquivo_xml):
    tree = ET.parse(arquivo_xml)
    root = tree.getroot()

    ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

    dados_list = []

    for nfe in root.findall('.//nfe:NFe', ns):
        dados = {}
        infNFe = nfe.find('nfe:infNFe', ns)
        if infNFe is not None:
            ide = infNFe.find('nfe:ide', ns)
            if ide is not None:
                dados['UF'] = ide.findtext('nfe:cUF', '', ns)
                dados['Numero_NF'] = ide.findtext('nfe:nNF', '', ns)
                dados['Natureza_Operacao'] = ide.findtext('nfe:natOp', '', ns)
                dados['Modelo'] = ide.findtext('nfe:mod', '', ns)
                dados['Serie'] = ide.findtext('nfe:serie', '', ns)
                dados['Data_Emissao'] = ide.findtext('nfe:dhEmi', '', ns)

            emit = infNFe.find('nfe:emit', ns)
            if emit is not None:
                dados['CNPJ_Emitente'] = emit.findtext('nfe:CNPJ', '', ns)
                dados['Nome_Emitente'] = emit.findtext('nfe:xNome', '', ns)
                dados['IE_Emitente'] = emit.findtext('nfe:IE', '', ns)

            for det in infNFe.findall('nfe:det', ns):
                det_dados = dados.copy()
                prod = det.find('nfe:prod', ns)
                if prod is not None:
                    det_dados['Codigo_Produto'] = prod.findtext('nfe:cProd', '', ns)
                    det_dados['Descricao_Produto'] = prod.findtext('nfe:xProd', '', ns)
                    det_dados['NCM'] = prod.findtext('nfe:NCM', '', ns)
                    det_dados['CFOP'] = prod.findtext('nfe:CFOP', '', ns)
                    det_dados['Quantidade'] = prod.findtext('nfe:qCom', '', ns)
                    det_dados['Valor_Unitario'] = prod.findtext('nfe:vUnCom', '', ns)
                    det_dados['Valor_Produto'] = prod.findtext('nfe:vProd', '', ns)

                imposto = det.find('nfe:imposto', ns)
                if imposto is not None:
                    icms = imposto.find('.//nfe:ICMS60', ns)
                    if icms is not None:
                        det_dados['ICMS_CST'] = icms.findtext('nfe:CST', '', ns)
                    
                    pis = imposto.find('.//nfe:PISNT', ns)
                    if pis is not None:
                        det_dados['PIS_CST'] = pis.findtext('nfe:CST', '', ns)

                    cofins = imposto.find('.//nfe:COFINSNT', ns)
                    if cofins is not None:
                        det_dados['COFINS_CST'] = cofins.findtext('nfe:CST', '', ns)
                    
                    ibscbs = imposto.find('.//nfe:IBSCBS', ns)
                    if ibscbs is not None:
                        det_dados['IBSCBS_vBC'] = ibscbs.find('.//nfe:vBC', ns).text if ibscbs.find('.//nfe:vBC', ns) is not None else ''
                        det_dados['IBSCBS_pIBSUF'] = ibscbs.find('.//nfe:pIBSUF', ns).text if ibscbs.find('.//nfe:pIBSUF', ns) is not None else ''
                        det_dados['IBSCBS_vIBSUF'] = ibscbs.find('.//nfe:vIBSUF', ns).text if ibscbs.find('.//nfe:vIBSUF', ns) is not None else ''
                        det_dados['IBSCBS_pIBSMun'] = ibscbs.find('.//nfe:pIBSMun', ns).text if ibscbs.find('.//nfe:pIBSMun', ns) is not None else ''
                        det_dados['IBSCBS_vIBSMun'] = ibscbs.find('.//nfe:vIBSMun', ns).text if ibscbs.find('.//nfe:vIBSMun', ns) is not None else ''
                        det_dados['IBSCBS_vIBS'] = ibscbs.find('.//nfe:vIBS', ns).text if ibscbs.find('.//nfe:vIBS', ns) is not None else ''
                        det_dados['IBSCBS_pCBS'] = ibscbs.find('.//nfe:pCBS', ns).text if ibscbs.find('.//nfe:pCBS', ns) is not None else ''
                        det_dados['IBSCBS_vCBS'] = ibscbs.find('.//nfe:vCBS', ns).text if ibscbs.find('.//nfe:vCBS', ns) is not None else ''
                
                dados_list.append(det_dados)

            total = infNFe.find('nfe:total', ns)
            if total is not None:
                # Adiciona totais a todos os registros de detalhes para esta NFe
                for item in dados_list:
                    if item.get('Numero_NF') == dados.get('Numero_NF'):
                        icmstot = total.find('nfe:ICMSTot', ns)
                        if icmstot is not None:
                            item['Total_vBC'] = icmstot.findtext('nfe:vBC', '', ns)
                            item['Total_vICMS'] = icmstot.findtext('nfe:vICMS', '', ns)
                            item['Total_vProd'] = icmstot.findtext('nfe:vProd', '', ns)
                            item['Total_vNF'] = icmstot.findtext('nfe:vNF', '', ns)
                        
                        ibscbstot = total.find('nfe:IBSCBSTot', ns)
                        if ibscbstot is not None:
                            item['Total_vBCIBSCBS'] = ibscbstot.findtext('nfe:vBCIBSCBS', '', ns)
                            item['Total_vIBS'] = ibscbstot.find('.//nfe:gIBS/nfe:vIBS', ns).text if ibscbstot.find('.//nfe:gIBS/nfe:vIBS', ns) is not None else ''
                            item['Total_vCBS'] = ibscbstot.find('.//nfe:gCBS/nfe:vCBS', ns).text if ibscbstot.find('.//nfe:gCBS/nfe:vCBS', ns) is not None else ''

    return dados_list

st.title('Extrator de Dados Fiscais de XML')

uploaded_file = st.file_uploader("Escolha um arquivo XML", type="xml")

if uploaded_file is not None:
    dados_fiscais = extrair_dados_fiscais(uploaded_file)
    
    if dados_fiscais:
        st.success("Dados extraídos com sucesso!")
        
        df = pd.DataFrame(dados_fiscais)
        st.dataframe(df)
        
        # Criar um botão de download para o Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Dados Fiscais')
        
        st.download_button(
            label="Baixar dados em Excel",
            data=output.getvalue(),
            file_name='dados_fiscais.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    else:
        st.error("Não foi possível extrair dados do arquivo XML. Verifique o formato do arquivo.")

