import os
import tkinter.filedialog
from tkinter import *

import pandas as pd
import xmltodict
from tqdm import tqdm


def ler_xml_prod_receita(nota):
    with open(nota, 'rb') as arquivo:
        documento = xmltodict.parse(arquivo, encoding = 'ANSI')
    dic_nota = documento['NFeLog']['procNFe']['NFe']['infNFe']
    chave = dic_nota['@Id'][3:]
    nf = dic_nota['ide']['nNF']
    produtos = dic_nota['det']
    tot = dic_nota['total']['ICMSTot']
    desconto_total = float(tot['vDesc'])
    frete_total = float(tot['vFrete'])
    ipi_total = float(tot['vIPI'])

    if 'eveNFe' in documento['NFeLog']:
        try:
            dic_evento = documento['NFeLog']['eveNFe']
            cancelamento_r = 'Não houve'
            for evento in dic_evento:
                eventos = evento['evento']['infEvento']['tpEvento']
                if '110111' in eventos:
                    cancelamento_r = 'Nota Cancelada'
                else:
                    cancelamento = 'Não houve'
                if cancelamento_r != 'Não houve':
                    cancelamento = cancelamento_r
        except:
            dic_evento = documento['NFeLog']['eveNFe']
            evento_unico = dic_evento['evento']['infEvento']['tpEvento']
            if '110111' in str(evento_unico):
                cancelamento = 'Nota Cancelada'
            else:
                cancelamento = 'Não houve'
    else:
        cancelamento = 'Não houve'

    lista_produtos = []
    try:
        for i, produto in enumerate(produtos):
            item = i + 1
            nome_produto = produto['prod']['xProd']
            qtde_produto = float(produto['prod']['qCom'])
            valor_unit = float(produto['prod']['vUnCom'])
            valor_produto = float(produto['prod']['vProd'])
            if desconto_total != 0.00:
                if 'vDesc' in produto['prod']:
                    desconto = float(produto['prod']['vDesc'])
                else:
                    desconto = 0.00
            else:
                desconto = float(desconto_total)
            if frete_total != 0.00:
                if 'vFrete' in produto['prod']:
                    frete = float(produto['prod']['vFrete'])
                else:
                    frete = 0.00
            else:
                frete = float(frete_total)
            if ipi_total != 0.00:
                mod_ipi = produto['imposto']['IPI']
                if 'IPITrib' in mod_ipi:
                    ipi = float(produto['imposto']['IPI']['IPITrib']['vIPI'])
                else:
                    ipi = 0.00
            else:
                ipi = float(ipi_total)

            ncm_produto = produto['prod']['NCM']
            cfop_produto = produto['prod']['CFOP']

            if 'PISAliq' in produto['imposto']['PIS']:
                cst_pis = produto['imposto']['PIS']['PISAliq']['CST']
            elif 'PISQtde' in produto['imposto']['PIS']:
                cst_pis = produto['imposto']['PIS']['PISQtde']['CST']
            elif 'PISNT' in produto['imposto']['PIS']:
                cst_pis = produto['imposto']['PIS']['PISNT']['CST']
            else:
                cst_pis = produto['imposto']['PIS']['PISOutr']['CST']

            if cfop_produto == '6933':
                icms = 0.00
                icms_st = 0.00
                cst_icms = cfop_produto
                origem = 0

            elif cfop_produto == '5933':
                icms = 0.00
                icms_st = 0.00
                cst_icms = cfop_produto
                origem = 0

            else:
                mod_icms = produto['imposto']['ICMS']
                if 'ICMS00' in mod_icms:
                    icms = produto['imposto']['ICMS']['ICMS00']['vICMS']
                    icms_st = 0.00
                    cst_icms = str(produto['imposto']['ICMS']['ICMS00']['CST'])
                    origem = str(produto['imposto']['ICMS']['ICMS00']['orig'])

                elif 'ICMS10' in mod_icms:
                    icms = produto['imposto']['ICMS']['ICMS10']['vICMS']
                    icms_st = produto['imposto']['ICMS']['ICMS10']['vICMSST']
                    cst_icms = str(produto['imposto']['ICMS']['ICMS10']['CST'])
                    origem = str(produto['imposto']['ICMS']['ICMS10']['orig'])

                elif 'ICMS20' in mod_icms:
                    icms = produto['imposto']['ICMS']['ICMS20']['vICMS']
                    icms_st = 0.00
                    cst_icms = str(produto['imposto']['ICMS']['ICMS20']['CST'])
                    origem = str(produto['imposto']['ICMS']['ICMS20']['orig'])

                elif 'ICMS30' in mod_icms:
                    icms = 0.00
                    icms_st = produto['imposto']['ICMS']['ICMS30']['vICMSST']
                    cst_icms = str(produto['imposto']['ICMS']['ICMS30']['CST'])
                    origem = str(produto['imposto']['ICMS']['ICMS30']['orig'])

                elif 'ICMS40' in mod_icms:
                    icms = 0.00
                    icms_st = 0.00
                    cst_icms = str(produto['imposto']['ICMS']['ICMS40']['CST'])
                    origem = str(produto['imposto']['ICMS']['ICMS40']['orig'])

                elif 'ICMS51' in mod_icms:
                    if 'vICMS' in produto['imposto']['ICMS']['ICMS51']:
                        icms = produto['imposto']['ICMS']['ICMS51']['vICMS']
                    else:
                        icms = 0.00
                    icms_st = 0.00
                    cst_icms = str(produto['imposto']['ICMS']['ICMS51']['CST'])
                    origem = str(produto['imposto']['ICMS']['ICMS51']['orig'])

                elif 'ICMS60' in mod_icms:
                    icms = 0.00
                    icms_st = 0.00
                    cst_icms = str(produto['imposto']['ICMS']['ICMS60']['CST'])
                    origem = str(produto['imposto']['ICMS']['ICMS60']['orig'])

                elif 'ICMS70' in mod_icms:
                    icms = produto['imposto']['ICMS']['ICMS70']['vICMS']
                    icms_st = produto['imposto']['ICMS']['ICMS70']['vICMSST']
                    cst_icms = str(produto['imposto']['ICMS']['ICMS70']['CST'])
                    origem = str(produto['imposto']['ICMS']['ICMS70']['orig'])

                elif 'ICMS90' in mod_icms:
                    if 'vICMS' in produto['imposto']['ICMS']['ICMS90']:
                        icms = produto['imposto']['ICMS']['ICMS90']['vICMS']
                    else:
                        icms = 0.00
                    if 'vICMSST' in produto['imposto']['ICMS']['ICMS90']:
                        icms_st = produto['imposto']['ICMS']['ICMS90']['vICMSST']
                    else:
                        icms_st = 0.00
                    cst_icms = str(produto['imposto']['ICMS']['ICMS90']['CST'])
                    origem = str(produto['imposto']['ICMS']['ICMS90']['orig'])

                elif 'ICMSPart' in mod_icms:
                    if 'vICMS' in produto['imposto']['ICMS']['ICMSPart']:
                        icms = produto['imposto']['ICMS']['ICMSPart']['vICMS']
                    else:
                        icms = 0.00
                    if 'vICMSST' in produto['imposto']['ICMS']['ICMSPart']:
                        icms_st = produto['imposto']['ICMS']['ICMSPart']['vICMSST']
                    else:
                        icms_st = 0.00
                    cst_icms = str(produto['imposto']['ICMS']['ICMSPart']['CST'])
                    origem = str(produto['imposto']['ICMS']['ICMSPart']['orig'])

                elif 'ICMSST' in mod_icms:
                    icms = 0.00
                    icms_st = 0.00
                    cst_icms = str(produto['imposto']['ICMS']['ICMSST']['CST'])
                    origem = str(produto['imposto']['ICMS']['ICMSST']['orig'])

                elif 'ICMSSN101' in mod_icms:
                    icms = 0.00
                    icms_st = 0.00
                    cst_icms = str(produto['imposto']['ICMS']['ICMSSN101']['CSOSN'])
                    origem = str(produto['imposto']['ICMS']['ICMSSN101']['orig'])

                elif 'ICMSSN102' in mod_icms:
                    icms = 0.00
                    icms_st = 0.00
                    cst_icms = str(produto['imposto']['ICMS']['ICMSSN102']['CSOSN'])
                    origem = str(produto['imposto']['ICMS']['ICMSSN102']['orig'])

                elif 'ICMSSN201' in mod_icms:
                    icms = 0.00
                    icms_st = produto['imposto']['ICMS']['ICMSSN201']['vICMSST']
                    cst_icms = str(produto['imposto']['ICMS']['ICMSSN201']['CSOSN'])
                    origem = str(produto['imposto']['ICMS']['ICMSSN201']['orig'])

                elif 'ICMSSN202' in mod_icms:
                    icms = 0.00
                    icms_st = produto['imposto']['ICMS']['ICMSSN202']['vICMSST']
                    cst_icms = str(produto['imposto']['ICMS']['ICMSSN202']['CSOSN'])
                    origem = str(produto['imposto']['ICMS']['ICMSSN202']['orig'])

                elif 'ICMSSN500' in mod_icms:
                    icms = 0.00
                    icms_st = 0.00
                    cst_icms = str(produto['imposto']['ICMS']['ICMSSN500']['CSOSN'])
                    origem = str(produto['imposto']['ICMS']['ICMSSN500']['orig'])

                elif 'ICMSSN900' in mod_icms:
                    if 'vICMS' in produto['imposto']['ICMS']['ICMSSN900']:
                        icms = produto['imposto']['ICMS']['ICMSSN900']['vICMS']
                    else:
                        icms = 0.00
                    if 'vICMSST' in produto['imposto']['ICMS']['ICMSSN900']:
                        icms_st = produto['imposto']['ICMS']['ICMSSN900']['vICMSST']
                    else:
                        icms_st = 0.00
                    cst_icms = str(produto['imposto']['ICMS']['ICMSSN900']['CSOSN'])
                    origem = str(produto['imposto']['ICMS']['ICMSSN900']['orig'])

            produtos_receita = {
                'Número da NF': nf,
                'Chave de Acesso': chave,
                'Item': item,
                'Produto': nome_produto,
                'Quantidade de Produto': float(qtde_produto),
                'Valor unitário': float(valor_unit),
                'Valor do Produto': float(valor_produto),
                'Desconto': float(desconto),
                'Frete': float(frete),
                'IPI': float(ipi),
                'ICMS': float(icms),
                'ICMS ST': float(icms_st),
                'Origem': str(origem),
                'CST ICMS / CSOSN': str(cst_icms),
                'CST Pis/Cofins': str(cst_pis),
                'NCM': str(ncm_produto),
                'CFOP': str(cfop_produto),
                'Cancelamento': cancelamento,
            }
            lista_produtos.append(produtos_receita)
    except:
        item = 1
        nome_produto = produtos['prod']['xProd']
        qtde_produto = float(produtos['prod']['qCom'])
        valor_unit = float(produtos['prod']['vUnCom'])
        valor_produto = float(produtos['prod']['vProd'])
        if desconto_total != 0.00:
            desconto = float(produtos['prod']['vDesc'])
        else:
            desconto = float(desconto_total)
        if frete_total != 0.00:
            frete = float(produtos['prod']['vFrete'])
        else:
            frete = float(frete_total)
        if ipi_total != 0.00:
            mod_ipi = produtos['imposto']['IPI']
            if 'IPITrib' in mod_ipi:
                ipi = float(produtos['imposto']['IPI']['IPITrib']['vIPI'])
            else:
                ipi = 0.00
        else:
            ipi = float(ipi_total)
        ncm_produto = produtos['prod']['NCM']
        cfop_produto = produtos['prod']['CFOP']

        if 'PISAliq' in produtos['imposto']['PIS']:
            cst_pis = produtos['imposto']['PIS']['PISAliq']['CST']
        elif 'PISQtde' in produtos['imposto']['PIS']:
            cst_pis = produtos['imposto']['PIS']['PISQtde']['CST']
        elif 'PISNT' in produtos['imposto']['PIS']:
            cst_pis = produtos['imposto']['PIS']['PISNT']['CST']
        else:
            cst_pis = produtos['imposto']['PIS']['PISOutr']['CST']

        if cfop_produto == '6933':
            icms = 0.00
            icms_st = 0.00
            cst_icms = cfop_produto
            origem = 0

        elif cfop_produto == '5933':
            icms = 0.00
            icms_st = 0.00
            cst_icms = cfop_produto
            origem = 0

        else:
            mod_icms = produtos['imposto']['ICMS']
            if 'ICMS00' in mod_icms:
                icms = produtos['imposto']['ICMS']['ICMS00']['vICMS']
                icms_st = 0.00
                cst_icms = str(produtos['imposto']['ICMS']['ICMS00']['CST'])
                origem = str(produtos['imposto']['ICMS']['ICMS00']['orig'])

            elif 'ICMS10' in mod_icms:
                icms = produtos['imposto']['ICMS']['ICMS10']['vICMS']
                icms_st = produtos['imposto']['ICMS']['ICMS10']['vICMSST']
                cst_icms = str(produtos['imposto']['ICMS']['ICMS10']['CST'])
                origem = str(produtos['imposto']['ICMS']['ICMS10']['orig'])

            elif 'ICMS20' in mod_icms:
                icms = produtos['imposto']['ICMS']['ICMS20']['vICMS']
                icms_st = 0.00
                cst_icms = str(produtos['imposto']['ICMS']['ICMS20']['CST'])
                origem = str(produtos['imposto']['ICMS']['ICMS20']['orig'])

            elif 'ICMS30' in mod_icms:
                icms = 0.00
                icms_st = produtos['imposto']['ICMS']['ICMS30']['vICMSST']
                cst_icms = str(produtos['imposto']['ICMS']['ICMS30']['CST'])
                origem = str(produtos['imposto']['ICMS']['ICMS30']['orig'])

            elif 'ICMS40' in mod_icms:
                icms = 0.00
                icms_st = 0.00
                cst_icms = str(produtos['imposto']['ICMS']['ICMS40']['CST'])
                origem = str(produtos['imposto']['ICMS']['ICMS40']['orig'])

            elif 'ICMS51' in mod_icms:
                if 'vICMS' in produtos['imposto']['ICMS']['ICMS51']:
                    icms = produtos['imposto']['ICMS']['ICMS51']['vICMS']
                else:
                    icms = 0.00
                icms_st = 0.00
                cst_icms = str(produtos['imposto']['ICMS']['ICMS51']['CST'])
                origem = str(produtos['imposto']['ICMS']['ICMS51']['orig'])

            elif 'ICMS60' in mod_icms:
                icms = 0.00
                icms_st = 0.00
                cst_icms = str(produtos['imposto']['ICMS']['ICMS60']['CST'])
                origem = str(produtos['imposto']['ICMS']['ICMS60']['orig'])

            elif 'ICMS70' in mod_icms:
                icms = produtos['imposto']['ICMS']['ICMS70']['vICMS']
                icms_st = produtos['imposto']['ICMS']['ICMS70']['vICMSST']
                cst_icms = str(produtos['imposto']['ICMS']['ICMS70']['CST'])
                origem = str(produtos['imposto']['ICMS']['ICMS70']['orig'])

            elif 'ICMS90' in mod_icms:
                if 'vICMS' in produtos['imposto']['ICMS']['ICMS90']:
                    icms = produtos['imposto']['ICMS']['ICMS90']['vICMS']
                else:
                    icms = 0.00
                if 'vICMSST' in produtos['imposto']['ICMS']['ICMS90']:
                    icms_st = produtos['imposto']['ICMS']['ICMS90']['vICMSST']
                else:
                    icms_st = 0.00
                cst_icms = str(produtos['imposto']['ICMS']['ICMS90']['CST'])
                origem = str(produtos['imposto']['ICMS']['ICMS90']['orig'])

            elif 'ICMSPart' in mod_icms:
                if 'vICMS' in produtos['imposto']['ICMS']['ICMSPart']:
                    icms = produtos['imposto']['ICMS']['ICMSPart']['vICMS']
                else:
                    icms = 0.00
                if 'vICMSST' in produtos['imposto']['ICMS']['ICMSPart']:
                    icms_st = produtos['imposto']['ICMS']['ICMSPart']['vICMSST']
                else:
                    icms_st = 0.00
                cst_icms = str(produtos['imposto']['ICMS']['ICMSPart']['CST'])
                origem = str(produtos['imposto']['ICMS']['ICMSPart']['orig'])

            elif 'ICMSST' in mod_icms:
                icms = 0.00
                icms_st = 0.00
                cst_icms = str(produtos['imposto']['ICMS']['ICMSST']['CST'])
                origem = str(produtos['imposto']['ICMS']['ICMSST']['orig'])

            elif 'ICMSSN101' in mod_icms:
                icms = 0.00
                icms_st = 0.00
                cst_icms = str(produtos['imposto']['ICMS']['ICMSSN101']['CSOSN'])
                origem = str(produtos['imposto']['ICMS']['ICMSSN101']['orig'])

            elif 'ICMSSN102' in mod_icms:
                icms = 0.00
                icms_st = 0.00
                cst_icms = str(produtos['imposto']['ICMS']['ICMSSN102']['CSOSN'])
                origem = str(produtos['imposto']['ICMS']['ICMSSN102']['orig'])

            elif 'ICMSSN201' in mod_icms:
                icms = 0.00
                icms_st = produtos['imposto']['ICMS']['ICMSSN201']['vICMSST']
                cst_icms = str(produtos['imposto']['ICMS']['ICMSSN201']['CSOSN'])
                origem = str(produtos['imposto']['ICMS']['ICMSSN201']['orig'])

            elif 'ICMSSN202' in mod_icms:
                icms = 0.00
                icms_st = produtos['imposto']['ICMS']['ICMSSN202']['vICMSST']
                cst_icms = str(produtos['imposto']['ICMS']['ICMSSN202']['CSOSN'])
                origem = str(produtos['imposto']['ICMS']['ICMSSN202']['orig'])

            elif 'ICMSSN500' in mod_icms:
                icms = 0.00
                icms_st = 0.00
                cst_icms = str(produtos['imposto']['ICMS']['ICMSSN500']['CSOSN'])
                origem = str(produtos['imposto']['ICMS']['ICMSSN500']['orig'])

            elif 'ICMSSN900' in mod_icms:
                if 'vICMS' in produtos['imposto']['ICMS']['ICMSSN900']:
                    icms = produtos['imposto']['ICMS']['ICMSSN900']['vICMS']
                else:
                    icms = 0.00
                if 'vICMSST' in produtos['imposto']['ICMS']['ICMSSN900']:
                    icms_st = produtos['imposto']['ICMS']['ICMSSN900']['vICMSST']
                else:
                    icms_st = 0.00
                cst_icms = str(produtos['imposto']['ICMS']['ICMSSN900']['CSOSN'])
                origem = str(produtos['imposto']['ICMS']['ICMSSN900']['orig'])

        produtos_receita = {
            'Número da NF': nf,
            'Chave de Acesso': chave,
            'Item': item,
            'Produto': nome_produto,
            'Quantidade de Produto': float(qtde_produto),
            'Valor unitário': float(valor_unit),
            'Valor do Produto': float(valor_produto),
            'Desconto': float(desconto),
            'Frete': float(frete),
            'IPI': float(ipi),
            'ICMS': float(icms),
            'ICMS ST': float(icms_st),
            'Origem': str(origem),
            'CST ICMS / CSOSN': str(cst_icms),
            'CST Pis/Cofins': str(cst_pis),
            'NCM': str(ncm_produto),
            'CFOP': str(cfop_produto),
            'Cancelamento': cancelamento,
        }
        lista_produtos.append(produtos_receita)
    return lista_produtos


def ler_xml_prod(nota):
    with open(nota, 'rb') as arquivo:
        documento = xmltodict.parse(arquivo)
        lista_produtos = []
    if 'nfeProc' in documento:
        dic_nota = documento['nfeProc']['NFe']['infNFe']
        chave = dic_nota['@Id'][3:]
        nf = dic_nota['ide']['nNF']
        produtos = dic_nota['det']
        tot = dic_nota['total']['ICMSTot']
        desconto_total = float(tot['vDesc'])
        frete_total = float(tot['vFrete'])
        ipi_total = float(tot['vIPI'])
        cancelamento = 'Não possui info no XML'
        try:
            for i, produto in enumerate(produtos):
                item = i + 1
                nome_produto = produto['prod']['xProd']
                qtde_produto = float(produto['prod']['qCom'])
                valor_unit = float(produto['prod']['vUnCom'])
                valor_produto = float(produto['prod']['vProd'])

                if desconto_total != 0.00:
                    if 'vDesc' in produto['prod']:
                        desconto = float(produto['prod']['vDesc'])
                    else:
                        desconto = 0.00
                else:
                    desconto = float(desconto_total)
                if frete_total != 0.00:
                    if 'vFrete' in produto['prod']:
                        frete = float(produto['prod']['vFrete'])
                    else:
                        frete = 0.00
                else:
                    frete = float(frete_total)
                if ipi_total != 0.00:
                    mod_ipi = produto['imposto']['IPI']
                    if 'IPITrib' in mod_ipi:
                        ipi = float(produto['imposto']['IPI']['IPITrib']['vIPI'])
                    else:
                        ipi = 0.00
                else:
                    ipi = float(ipi_total)

                ncm_produto = produto['prod']['NCM']
                cfop_produto = produto['prod']['CFOP']

                if 'PISAliq' in produto['imposto']['PIS']:
                    cst_pis = produto['imposto']['PIS']['PISAliq']['CST']
                elif 'PISQtde' in produto['imposto']['PIS']:
                    cst_pis = produto['imposto']['PIS']['PISQtde']['CST']
                elif 'PISNT' in produto['imposto']['PIS']:
                    cst_pis = produto['imposto']['PIS']['PISNT']['CST']
                else:
                    cst_pis = produto['imposto']['PIS']['PISOutr']['CST']

                if cfop_produto == '6933':
                    icms = 0.00
                    icms_st = 0.00
                    cst_icms = cfop_produto
                    origem = 0

                elif cfop_produto == '5933':
                    icms = 0.00
                    icms_st = 0.00
                    cst_icms = cfop_produto
                    origem = 0

                else:
                    mod_icms = produto['imposto']['ICMS']
                    if 'ICMS00' in mod_icms:
                        icms = produto['imposto']['ICMS']['ICMS00']['vICMS']
                        icms_st = 0.00
                        cst_icms = str(produto['imposto']['ICMS']['ICMS00']['CST'])
                        origem = str(produto['imposto']['ICMS']['ICMS00']['orig'])

                    elif 'ICMS10' in mod_icms:
                        icms = produto['imposto']['ICMS']['ICMS10']['vICMS']
                        icms_st = produto['imposto']['ICMS']['ICMS10']['vICMSST']
                        cst_icms = str(produto['imposto']['ICMS']['ICMS10']['CST'])
                        origem = str(produto['imposto']['ICMS']['ICMS10']['orig'])

                    elif 'ICMS20' in mod_icms:
                        icms = produto['imposto']['ICMS']['ICMS20']['vICMS']
                        icms_st = 0.00
                        cst_icms = str(produto['imposto']['ICMS']['ICMS20']['CST'])
                        origem = str(produto['imposto']['ICMS']['ICMS20']['orig'])

                    elif 'ICMS30' in mod_icms:
                        icms = 0.00
                        icms_st = produto['imposto']['ICMS']['ICMS30']['vICMSST']
                        cst_icms = str(produto['imposto']['ICMS']['ICMS30']['CST'])
                        origem = str(produto['imposto']['ICMS']['ICMS30']['orig'])

                    elif 'ICMS40' in mod_icms:
                        icms = 0.00
                        icms_st = 0.00
                        cst_icms = str(produto['imposto']['ICMS']['ICMS40']['CST'])
                        origem = str(produto['imposto']['ICMS']['ICMS40']['orig'])

                    elif 'ICMS51' in mod_icms:
                        if 'vICMS' in produto['imposto']['ICMS']['ICMS51']:
                            icms = produto['imposto']['ICMS']['ICMS51']['vICMS']
                        else:
                            icms = 0.00
                        icms_st = 0.00
                        cst_icms = str(produto['imposto']['ICMS']['ICMS51']['CST'])
                        origem = str(produto['imposto']['ICMS']['ICMS51']['orig'])

                    elif 'ICMS60' in mod_icms:
                        icms = 0.00
                        icms_st = 0.00
                        cst_icms = str(produto['imposto']['ICMS']['ICMS60']['CST'])
                        origem = str(produto['imposto']['ICMS']['ICMS60']['orig'])

                    elif 'ICMS70' in mod_icms:
                        icms = produto['imposto']['ICMS']['ICMS70']['vICMS']
                        icms_st = produto['imposto']['ICMS']['ICMS70']['vICMSST']
                        cst_icms = str(produto['imposto']['ICMS']['ICMS70']['CST'])
                        origem = str(produto['imposto']['ICMS']['ICMS70']['orig'])

                    elif 'ICMS90' in mod_icms:
                        if 'vICMS' in produto['imposto']['ICMS']['ICMS90']:
                            icms = produto['imposto']['ICMS']['ICMS90']['vICMS']
                        else:
                            icms = 0.00
                        if 'vICMSST' in produto['imposto']['ICMS']['ICMS90']:
                            icms_st = produto['imposto']['ICMS']['ICMS90']['vICMSST']
                        else:
                            icms_st = 0.00
                        cst_icms = str(produto['imposto']['ICMS']['ICMS90']['CST'])
                        origem = str(produto['imposto']['ICMS']['ICMS90']['orig'])

                    elif 'ICMSPart' in mod_icms:
                        if 'vICMS' in produto['imposto']['ICMS']['ICMSPart']:
                            icms = produto['imposto']['ICMS']['ICMSPart']['vICMS']
                        else:
                            icms = 0.00
                        if 'vICMSST' in produto['imposto']['ICMS']['ICMSPart']:
                            icms_st = produto['imposto']['ICMS']['ICMSPart']['vICMSST']
                        else:
                            icms_st = 0.00
                        cst_icms = str(produto['imposto']['ICMS']['ICMSPart']['CST'])
                        origem = str(produto['imposto']['ICMS']['ICMSPart']['orig'])

                    elif 'ICMSST' in mod_icms:
                        icms = 0.00
                        icms_st = 0.00
                        cst_icms = str(produto['imposto']['ICMS']['ICMSST']['CST'])
                        origem = str(produto['imposto']['ICMS']['ICMSST']['orig'])

                    elif 'ICMSSN101' in mod_icms:
                        icms = 0.00
                        icms_st = 0.00
                        cst_icms = str(produto['imposto']['ICMS']['ICMSSN101']['CSOSN'])
                        origem = str(produto['imposto']['ICMS']['ICMSSN101']['orig'])

                    elif 'ICMSSN102' in mod_icms:
                        icms = 0.00
                        icms_st = 0.00
                        cst_icms = str(produto['imposto']['ICMS']['ICMSSN102']['CSOSN'])
                        origem = str(produto['imposto']['ICMS']['ICMSSN102']['orig'])

                    elif 'ICMSSN201' in mod_icms:
                        icms = 0.00
                        icms_st = produto['imposto']['ICMS']['ICMSSN201']['vICMSST']
                        cst_icms = str(produto['imposto']['ICMS']['ICMSSN201']['CSOSN'])
                        origem = str(produto['imposto']['ICMS']['ICMSSN201']['orig'])

                    elif 'ICMSSN202' in mod_icms:
                        icms = 0.00
                        icms_st = produto['imposto']['ICMS']['ICMSSN202']['vICMSST']
                        cst_icms = str(produto['imposto']['ICMS']['ICMSSN202']['CSOSN'])
                        origem = str(produto['imposto']['ICMS']['ICMSSN202']['orig'])

                    elif 'ICMSSN500' in mod_icms:
                        icms = 0.00
                        icms_st = 0.00
                        cst_icms = str(produto['imposto']['ICMS']['ICMSSN500']['CSOSN'])
                        origem = str(produto['imposto']['ICMS']['ICMSSN500']['orig'])

                    elif 'ICMSSN900' in mod_icms:
                        if 'vICMS' in produto['imposto']['ICMS']['ICMSSN900']:
                            icms = produto['imposto']['ICMS']['ICMSSN900']['vICMS']
                        else:
                            icms = 0.00
                        if 'vICMSST' in produto['imposto']['ICMS']['ICMSSN900']:
                            icms_st = produto['imposto']['ICMS']['ICMSSN900']['vICMSST']
                        else:
                            icms_st = 0.00
                        cst_icms = str(produto['imposto']['ICMS']['ICMSSN900']['CSOSN'])
                        origem = str(produto['imposto']['ICMS']['ICMSSN900']['orig'])

                produtos_receita = {
                    'Número da NF': nf,
                    'Chave de Acesso': chave,
                    'Item': item,
                    'Produto': nome_produto,
                    'Quantidade de Produto': float(qtde_produto),
                    'Valor unitário': float(valor_unit),
                    'Valor do Produto': float(valor_produto),
                    'Desconto': float(desconto),
                    'Frete': float(frete),
                    'IPI': float(ipi),
                    'ICMS': float(icms),
                    'ICMS ST': float(icms_st),
                    'Origem': str(origem),
                    'CST ICMS / CSOSN': str(cst_icms),
                    'CST Pis/Cofins': str(cst_pis),
                    'NCM': str(ncm_produto),
                    'CFOP': str(cfop_produto),
                    'Cancelamento': cancelamento,
                }
                lista_produtos.append(produtos_receita)
        except:
            item = 1
            nome_produto = produtos['prod']['xProd']
            qtde_produto = float(produtos['prod']['qCom'])
            valor_unit = float(produtos['prod']['vUnCom'])
            valor_produto = float(produtos['prod']['vProd'])
            if desconto_total != 0.00:
                desconto = float(produtos['prod']['vDesc'])
            else:
                desconto = desconto_total
            if frete_total != 0.00:
                frete = float(produtos['prod']['vFrete'])
            else:
                frete = frete_total
            if ipi_total != 0.00:
                mod_ipi = produtos['imposto']['IPI']
                if 'IPITrib' in mod_ipi:
                    ipi = float(produtos['imposto']['IPI']['IPITrib']['vIPI'])
                else:
                    ipi = 0.00
            else:
                ipi = ipi_total

            ncm_produto = produtos['prod']['NCM']
            cfop_produto = produtos['prod']['CFOP']

            if 'PISAliq' in produtos['imposto']['PIS']:
                cst_pis = produtos['imposto']['PIS']['PISAliq']['CST']
            elif 'PISQtde' in produtos['imposto']['PIS']:
                cst_pis = produtos['imposto']['PIS']['PISQtde']['CST']
            elif 'PISNT' in produtos['imposto']['PIS']:
                cst_pis = produtos['imposto']['PIS']['PISNT']['CST']
            else:
                cst_pis = produtos['imposto']['PIS']['PISOutr']['CST']

            if cfop_produto != '6933':
                icms = 0.00
                icms_st = 0.00
                cst_icms = cfop_produto
                origem = 0

            elif cfop_produto != '5933':
                icms = 0.00
                icms_st = 0.00
                cst_icms = cfop_produto
                origem = 0

            else:
                mod_icms = produtos['imposto']['ICMS']
                if 'ICMS00' in mod_icms:
                    icms = produtos['imposto']['ICMS']['ICMS00']['vICMS']
                    icms_st = 0.00
                    cst_icms = str(produtos['imposto']['ICMS']['ICMS00']['CST'])
                    origem = str(produtos['imposto']['ICMS']['ICMS00']['orig'])

                elif 'ICMS10' in mod_icms:
                    icms = produtos['imposto']['ICMS']['ICMS10']['vICMS']
                    icms_st = produtos['imposto']['ICMS']['ICMS10']['vICMSST']
                    cst_icms = str(produtos['imposto']['ICMS']['ICMS10']['CST'])
                    origem = str(produtos['imposto']['ICMS']['ICMS10']['orig'])

                elif 'ICMS20' in mod_icms:
                    icms = produtos['imposto']['ICMS']['ICMS20']['vICMS']
                    icms_st = 0.00
                    cst_icms = str(produtos['imposto']['ICMS']['ICMS20']['CST'])
                    origem = str(produtos['imposto']['ICMS']['ICMS20']['orig'])

                elif 'ICMS30' in mod_icms:
                    icms = 0.00
                    icms_st = produtos['imposto']['ICMS']['ICMS30']['vICMSST']
                    cst_icms = str(produtos['imposto']['ICMS']['ICMS30']['CST'])
                    origem = str(produtos['imposto']['ICMS']['ICMS30']['orig'])

                elif 'ICMS40' in mod_icms:
                    icms = 0.00
                    icms_st = 0.00
                    cst_icms = str(produtos['imposto']['ICMS']['ICMS40']['CST'])
                    origem = str(produtos['imposto']['ICMS']['ICMS40']['orig'])

                elif 'ICMS51' in mod_icms:
                    if 'vICMS' in produtos['imposto']['ICMS']['ICMS51']:
                        icms = produtos['imposto']['ICMS']['ICMS51']['vICMS']
                    else:
                        icms = 0.00
                    icms_st = 0.00
                    cst_icms = str(produtos['imposto']['ICMS']['ICMS51']['CST'])
                    origem = str(produtos['imposto']['ICMS']['ICMS51']['orig'])

                elif 'ICMS60' in mod_icms:
                    icms = 0.00
                    icms_st = 0.00
                    cst_icms = str(produtos['imposto']['ICMS']['ICMS60']['CST'])
                    origem = str(produtos['imposto']['ICMS']['ICMS60']['orig'])

                elif 'ICMS70' in mod_icms:
                    icms = produtos['imposto']['ICMS']['ICMS70']['vICMS']
                    icms_st = produtos['imposto']['ICMS']['ICMS70']['vICMSST']
                    cst_icms = str(produtos['imposto']['ICMS']['ICMS70']['CST'])
                    origem = str(produtos['imposto']['ICMS']['ICMS70']['orig'])

                elif 'ICMS90' in mod_icms:
                    if 'vICMS' in produtos['imposto']['ICMS']['ICMS90']:
                        icms = produtos['imposto']['ICMS']['ICMS90']['vICMS']
                    else:
                        icms = 0.00
                    if 'vICMSST' in produtos['imposto']['ICMS']['ICMS90']:
                        icms_st = produtos['imposto']['ICMS']['ICMS90']['vICMSST']
                    else:
                        icms_st = 0.00
                    cst_icms = str(produtos['imposto']['ICMS']['ICMS90']['CST'])
                    origem = str(produtos['imposto']['ICMS']['ICMS90']['orig'])

                elif 'ICMSPart' in mod_icms:
                    if 'vICMS' in produtos['imposto']['ICMS']['ICMSPart']:
                        icms = produtos['imposto']['ICMS']['ICMSPart']['vICMS']
                    else:
                        icms = 0.00
                    if 'vICMSST' in produtos['imposto']['ICMS']['ICMSPart']:
                        icms_st = produtos['imposto']['ICMS']['ICMSPart']['vICMSST']
                    else:
                        icms_st = 0.00
                    cst_icms = str(produtos['imposto']['ICMS']['ICMSPart']['CST'])
                    origem = str(produtos['imposto']['ICMS']['ICMSPart']['orig'])

                elif 'ICMSST' in mod_icms:
                    icms = 0.00
                    icms_st = 0.00
                    cst_icms = str(produtos['imposto']['ICMS']['ICMSST']['CST'])
                    origem = str(produtos['imposto']['ICMS']['ICMSST']['orig'])

                elif 'ICMSSN101' in mod_icms:
                    icms = 0.00
                    icms_st = 0.00
                    cst_icms = str(produtos['imposto']['ICMS']['ICMSSN101']['CSOSN'])
                    origem = str(produtos['imposto']['ICMS']['ICMSSN101']['orig'])

                elif 'ICMSSN102' in mod_icms:
                    icms = 0.00
                    icms_st = 0.00
                    cst_icms = str(produtos['imposto']['ICMS']['ICMSSN102']['CSOSN'])
                    origem = str(produtos['imposto']['ICMS']['ICMSSN102']['orig'])

                elif 'ICMSSN201' in mod_icms:
                    icms = 0.00
                    icms_st = produtos['imposto']['ICMS']['ICMSSN201']['vICMSST']
                    cst_icms = str(produtos['imposto']['ICMS']['ICMSSN201']['CSOSN'])
                    origem = str(produtos['imposto']['ICMS']['ICMSSN201']['orig'])

                elif 'ICMSSN202' in mod_icms:
                    icms = 0.00
                    icms_st = produtos['imposto']['ICMS']['ICMSSN202']['vICMSST']
                    cst_icms = str(produtos['imposto']['ICMS']['ICMSSN202']['CSOSN'])
                    origem = str(produtos['imposto']['ICMS']['ICMSSN202']['orig'])

                elif 'ICMSSN500' in mod_icms:
                    icms = 0.00
                    icms_st = 0.00
                    cst_icms = str(produtos['imposto']['ICMS']['ICMSSN500']['CSOSN'])
                    origem = str(produtos['imposto']['ICMS']['ICMSSN500']['orig'])

                elif 'ICMSSN900' in mod_icms:
                    if 'vICMS' in produtos['imposto']['ICMS']['ICMSSN900']:
                        icms = produtos['imposto']['ICMS']['ICMSSN900']['vICMS']
                    else:
                        icms = 0.00
                    if 'vICMSST' in produtos['imposto']['ICMS']['ICMSSN900']:
                        icms_st = produtos['imposto']['ICMS']['ICMSSN900']['vICMSST']
                    else:
                        icms_st = 0.00
                    cst_icms = str(produtos['imposto']['ICMS']['ICMSSN900']['CSOSN'])
                    origem = str(produtos['imposto']['ICMS']['ICMSSN900']['orig'])

            produtos_nfe = {
                'Número da NF': nf,
                'Chave de Acesso': chave,
                'Item': item,
                'Produto': nome_produto,
                'Quantidade de Produto': float(qtde_produto),
                'Valor unitário': float(valor_unit),
                'Valor do Produto': float(valor_produto),
                'Desconto': float(desconto),
                'Frete': float(frete),
                'IPI': float(ipi),
                'ICMS': float(icms),
                'ICMS ST': float(icms_st),
                'Origem': str(origem),
                'CST ICMS / CSOSN': str(cst_icms),
                'CST Pis/Cofins': str(cst_pis),
                'NCM': str(ncm_produto),
                'CFOP': str(cfop_produto),
                'Cancelamento': cancelamento,
            }
            lista_produtos.append(produtos_nfe)
    else:
        chave = documento['NFe']['infNFe']['@Id'][3:]
        nf = documento['NFe']['infNFe']['ide']['nNF']
        produtos_nfe = {
            'Número da NF': nf,
            'Chave de Acesso': chave,
            'Item': 'Verificar se não está denegada',
            'Produto': 'Verificar se não está denegada',
            'Quantidade de Produto': 'Verificar se não está denegada',
            'Valor unitário': 'Verificar se não está denegada',
            'Valor do Produto': 'Verificar se não está denegada',
            'Desconto': 'Verificar se não está denegada',
            'Frete': 'Verificar se não está denegada',
            'IPI': 'Verificar se não está denegada',
            'ICMS': 'Verificar se não está denegada',
            'ICMS ST': 'Verificar se não está denegada',
            'Origem': 'Verificar se não está denegada',
            'CST ICMS / CSOSN': 'Verificar se não está denegada',
            'CST Pis/Cofins': 'Verificar se não está denegada',
            'NCM': 'Verificar se não está denegada',
            'CFOP': 'Verificar se não está denegada',
            'Cancelamento': 'Verificar se não está denegada',
        }
        lista_produtos.append(produtos_nfe)

    return lista_produtos


def ler_xml_prod_nfce(nota):
    with open(nota, 'rb') as arquivo:
        documento = xmltodict.parse(arquivo)
    dic_nota = documento['nfeProc']['NFe']['infNFe']
    chave = dic_nota['@Id'][3:]
    nf = dic_nota['ide']['nNF']
    produtos = dic_nota['det']
    lista_produtos = []
    try:
        for i, produto in enumerate(produtos):
            item = i + 1
            nome_produto = produto['prod']['xProd']
            qtde_produto = float(produto['prod']['qCom'])
            valor_unit = float(produto['prod']['vUnCom'])
            valor_produto = float(produto['prod']['vProd'])
            ncm_produto = produto['prod']['NCM']
            cfop_produto = produto['prod']['CFOP']
            produtos_receita = {
                'Número da NF': nf,
                'Chave de Acesso': chave,
                'Item': item,
                'Produto': nome_produto,
                'Quantidade de Produto': qtde_produto,
                'Valor unitário': valor_unit,
                'Valor do Produto': valor_produto,
                'NCM': ncm_produto,
                'CFOP': cfop_produto,
            }
            lista_produtos.append(produtos_receita)
    except:
        item = 1
        nome_produto = produtos['prod']['xProd']
        qtde_produto = float(produtos['prod']['qCom'])
        valor_unit = float(produtos['prod']['vUnCom'])
        valor_produto = float(produtos['prod']['vProd'])
        ncm_produto = produtos['prod']['NCM']
        cfop_produto = produtos['prod']['CFOP']
        produtos_nfe = {
            'Número da NF': nf,
            'Chave de Acesso': chave,
            'Item': item,
            'Produto': nome_produto,
            'Quantidade de Produto': qtde_produto,
            'Valor unitário': valor_unit,
            'Valor do Produto': valor_produto,
            'NCM': ncm_produto,
            'CFOP': cfop_produto,
        }
        lista_produtos.append(produtos_nfe)
    return lista_produtos


janela = Tk()
caminho = tkinter.filedialog.askdirectory(title='Selecione a pasta com os arquivos XML')
janela.destroy()
lista_arquivos = os.listdir(caminho)
lista_arquivos_xml = []
lista_produtos_dic = []
lista_nfce_dic = []
for arquivo in lista_arquivos:
    if '.xml' in arquivo.lower():
        lista_arquivos_xml.append(arquivo)

for nota in tqdm(lista_arquivos_xml):
    try:
        with open(f'{caminho}/{nota}', 'rb') as arquivo:
            documento = xmltodict.parse(arquivo)
        if 'procInutNFe' in documento:
            pass
        elif 'ProcInutNFe' in documento:
            pass
        elif 'retInutNFe' in documento:
            pass
        elif 'inutNFe' in documento:
            pass
        elif 'procEventoNFe' in documento:
            pass
        elif 'procEventoCTe' in documento:
            pass
        elif 'retEnvEvento' in documento:
            pass
        elif 'resEvento' in documento:
            pass
        elif 'cteProc' in documento:
            pass
        elif 'NFe' in documento:
            lista_prov = ler_xml_prod(f'{caminho}/{nota}')
            for item in lista_prov:
                lista_produtos_dic.append(item)
        elif 'NFeLog' in documento:
            lista_prov = ler_xml_prod_receita(f'{caminho}/{nota}')
            for item in lista_prov:
                lista_produtos_dic.append(item)
        elif documento['nfeProc']['NFe']['infNFe']['ide']['mod'] == '55':
            lista_prov = ler_xml_prod(f'{caminho}/{nota}')
            for item in lista_prov:
                lista_produtos_dic.append(item)
        else:
            lista_prov = ler_xml_prod_nfce(f'{caminho}/{nota}')
            for item in lista_prov:
                lista_nfce_dic.append(item)
    except:
        with open(f'{caminho}/{nota}', 'rb') as arquivo:
            documento = xmltodict.parse(arquivo, encoding = 'ANSI')
        if 'procInutNFe' in documento:
            pass
        elif 'ProcInutNFe' in documento:
            pass
        elif 'retInutNFe' in documento:
            pass
        elif 'inutNFe' in documento:
            pass
        elif 'procEventoNFe' in documento:
            pass
        elif 'procEventoCTe' in documento:
            pass
        elif 'retEnvEvento' in documento:
            pass
        elif 'resEvento' in documento:
            pass
        elif 'cteProc' in documento:
            pass
        elif 'NFe' in documento:
            lista_prov = ler_xml_prod(f'{caminho}/{nota}')
            for item in lista_prov:
                lista_produtos_dic.append(item)
        elif 'NFeLog' in documento:
            lista_prov = ler_xml_prod_receita(f'{caminho}/{nota}')
            for item in lista_prov:
                lista_produtos_dic.append(item)
        elif documento['nfeProc']['NFe']['infNFe']['ide']['mod'] == '55':
            lista_prov = ler_xml_prod(f'{caminho}/{nota}')
            for item in lista_prov:
                lista_produtos_dic.append(item)
        else:
            lista_prov = ler_xml_prod_nfce(f'{caminho}/{nota}')
            for item in lista_prov:
                lista_nfce_dic.append(item)

if lista_produtos_dic != []:
    tabela_produtos = pd.DataFrame.from_dict(lista_produtos_dic)
    tabela_produtos_red = tabela_produtos[
        ['Número da NF', 'Chave de Acesso', 'CFOP', 'Valor do Produto', 'Desconto', 'Frete', 'IPI', 'ICMS', 'ICMS ST']]
    tabela_produtos_cfop = tabela_produtos_red.groupby(['Número da NF', 'Chave de Acesso', 'CFOP']).sum()
    tabela_produtos_cfop = tabela_produtos_cfop.reset_index()
    tabela_produtos.to_excel(f'{caminho}/+ Produtos NFE.xlsx', index=False)
    tabela_produtos_cfop.to_excel(f'{caminho}/+ CFOP NFE.xlsx', index=False)
if lista_nfce_dic != []:
    tabela_produtos_nfce = pd.DataFrame.from_dict(lista_nfce_dic)
    tabela_produtos_nfce.to_excel(f'{caminho}/+ Produtos NFCE.xlsx', index=False)
