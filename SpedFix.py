# Importando dependências de código
import datetime
import os, re, xlsxwriter
import xml.etree.ElementTree as ET
import PySimpleGUI as sg

def slow_print(string):
    for char in string:
        print(char, end="", flush=True)
    print("\n", end="", flush=True)        

def log(file, string, **kwargs):
    log_file = open(file, "a+")
    log_file.write(string + "\n")
    if not kwargs.get('silent'):
        slow_print(string)
    log_file.close()

def write_efd(efd_array, SPED_SAIDA):
    new_efd = ""
    for line in efd_array:
        new_line = "|"
        for column in line:
            new_line += column + "|"
        new_line += "\n"
        new_efd += new_line
    out_file = open(SPED_SAIDA, "w+", encoding="latin-1")
    out_file.write(new_efd)
    out_file.close()

def get_value(str_value):
    if not str_value:
        return 0
    return float(str_value.replace(",", "."))

def set_value(float_value):
    try:
        return str(round(float(float_value), 2)).replace(".", ",")
    except:
        return "0,00" 

def fix_removeIPI(efd_array, LOG_SAIDA):
    wannafix = False
    for line in efd_array:
        if (line[0] == "C100" and get_value(line[24]) != 0) or (line[0] == "C170" and get_value(line[23]) != 0) or (line[0] == "C190" and get_value(line[10]) != 0):
            ans = sg.popup_yes_no("Foram encontrados valores de IPI. Deseja zerá-los?", title="Atenção!")
            if ans == "Yes":
                wannafix = True
            break
    if wannafix:
        log(LOG_SAIDA, "NOTAS COM IPI REMOVIDO:", silent=True)
        current_doc = ""
        for line in efd_array:
            if line[0] == "C100":
                current_doc = line[7]
                if get_value(line[24]) != 0:
                    log(LOG_SAIDA, "NF " + current_doc + " - " +
                        line[24] + " de IPI zerado na capa da nota.")
                    line[24] = "0"
            if line[0] == "C170" and get_value(line[23]) != 0:
                log(LOG_SAIDA, "NF " + current_doc + " - " +
                    line[23] + " de IPI zerado no item " + line[2] + ".")
                line[21] = "0"
                line[22] = "0"
                line[23] = "0"
            if line[0] == "C190" and get_value(line[10]) != 0:
                log(LOG_SAIDA, "NF " + current_doc + " - " +
                    line[10] + " de IPI zerado no registro analítico " + line[1] + "|" + line[2] + ".")
                line[10] = "0"
        log(LOG_SAIDA, "\n")
    return efd_array

def fix_removeABAT(efd_array, LOG_SAIDA):
    wannafix = False
    for line in efd_array:
        if (line[0] == "C100" and get_value(line[14]) != 0) or (line[0] == "C170" and get_value(line[37]) != 0) or (line[0] == "C190" and get_value(line[10]) != 0):
            ans = sg.popup_yes_no("Foram encontrados valores de Abatimento Não Tributado. Deseja zerá-los?", title="Atenção!")
            if ans == "Yes":
                wannafix = True
            break
    if wannafix:
        log(LOG_SAIDA, "ABATIMENTOS NÃO TRIBUTADOS REMOVIDOS:", silent=True)
        current_doc = ""
        for line in efd_array:
            if line[0] == "C100":
                current_doc = line[7]
                if get_value(line[14]) != 0:
                    log(LOG_SAIDA, "NF " + current_doc + " - " +
                        line[14] + " de Abatimento Não Tributado zerado na capa da nota.")
                    line[14] = "0"
            if line[0] == "C170" and get_value(line[37]) != 0:
                log(LOG_SAIDA, "NF " + current_doc + " - " +
                    line[37] + " de Abatimento Não Tributado zerado no item " + line[2] + ".")
                line[37] = "0"
        log(LOG_SAIDA, "\n")
    return efd_array

def fix_020_RED(efd_array, LOG_SAIDA):
    wannafix = False
    for line in efd_array:
        if line[0] == "C190" and line[1][-2:] == "20":
            vl_opr = get_value(line[4])
            vl_bc = get_value(line[5])
            vl_red_bc = get_value(line[9])
            if vl_opr != (vl_bc + vl_red_bc):
                ans = sg.popup_yes_no("Foram encontrados valores incorretos de redução da base de cálculo. Deseja corrigi-los?", title="Atenção!")
                if ans == "Yes":
                    wannafix = True
                break
    if wannafix:
        log(LOG_SAIDA, "CORREÇÕES DE REDUÇÃO DA BASE DE CÁLCULO:", silent=True)
        current_doc = ""
        for line in efd_array:
            if line[0] == "C100":
                current_doc = line[7]
            if line[0] == "C190" and line[1][-2:] == "20":
                vl_opr = get_value(line[4])
                vl_bc = get_value(line[5])
                vl_red_bc = get_value(line[9])
                if vl_opr != (vl_bc + vl_red_bc):
                    line[9] = set_value(vl_opr - vl_bc)
                    log(LOG_SAIDA, "NF " + current_doc + " - Combinação CST|CFOP: " + line[1] + "|" + line[2]
                        + ". Valor RED_BC corrigido de " + str(vl_red_bc) + " para " + line[9] + ".")
        log(LOG_SAIDA, "\n")
    return efd_array

def fix_bc_greater_than_opr(efd_array, LOG_SAIDA):
    wannafix = False
    for line in efd_array:
        vl_opr = 0
        vl_bc = 0
        if line[0] == "C170":
            vl_opr = get_value(line[6])
            vl_bc = get_value(line[12])
        if line[0] == "C190":
            vl_opr = get_value(line[4])
            vl_bc = get_value(line[5])
        if vl_bc > vl_opr:
                ans = sg.popup_yes_no("Foram encontrados valores de base de cálculo maiores que valores de operação. Deseja corrigi-los?", title="Atenção!")
                if ans == "Yes":
                    wannafix = True
                break
    if wannafix:
        log(LOG_SAIDA, "BASES DE CÁLCULO MAIORES QUE VALOR DE OPERAÇÃO:", silent=True)
        current_doc = ""
        total_bc = 0
        for line in efd_array:
            if line[0] == "C100":
                current_doc = line[7]
            if line[0] == "C170":
                vl_opr = get_value(line[6])
                vl_bc = get_value(line[12])
                total_bc += vl_bc
                if vl_bc > vl_opr:
                    line[6] = set_value(vl_opr)
                    log(LOG_SAIDA, "NF " + current_doc + " - Valor da base de cálculo corrigido de " +
                        str(vl_bc) + " para " + line[6] + " no item " + line[2] + ".")
            if line[0] == "C190":
                vl_opr = get_value(line[4])
                vl_bc = get_value(line[5])
                if vl_bc > vl_opr:
                    line[5] = set_value(vl_opr)
                    log(LOG_SAIDA, "NF " + current_doc + " - Valor da base de cálculo corrigido de " +
                        str(vl_bc) + " para " + line[5] + " no registro C190 " + line[1] + "|" + line[2] + ".")
        log(LOG_SAIDA, "\n")
    return efd_array

def fix_importCST(efd_array, LOG_SAIDA, LOG_AJUSTES):
    wannafix = False
    for line in efd_array:
        if (line[0] == "C190" and line[1][0] in ['1', '6']) or (line[0] == "C170" and line[9][0] in ['1', '6']):
            ans = sg.popup_yes_no("Foram encontrados CSTs de importação (100, 160, 600, etc). Deseja corrigi-los?", title="Atenção!")
            if ans == "Yes":
                wannafix = True
            break
    if wannafix:
        log(LOG_AJUSTES, "CSTS DE IMPORTAÇÃO CORRIGIDOS:", silent=True)
        current_doc = ""
        for line in efd_array:
            if line[0] == "C100":
                current_doc = line[7]
            if line[0] == "C170" and line[9][0] in ['1', '6']:
                old_cst = line[9]
                if line[9][0] == '1':
                    line[9] = '2' + line[9][1:]
                if line[9][0] == '6':
                    line[9] = '7' + line[9][1:]
                log(LOG_SAIDA, "NF " + current_doc + " - CST do item " +
                    line[2] + " corrigido de " + old_cst + " para " + line[9] + ".")
            if line[0] == "C190" and line[1][0] in ['1', '6']:
                old_cst = line[1]
                if line[1][0] == '1':
                    line[1] = '2' + line[1][1:]
                if line[1][0] == '6':
                    line[1] = '7' + line[1][1:]
                log(LOG_SAIDA, "NF " + current_doc + " - CST " +
                    old_cst + " corrigido para " + line[1] + ".")
                log(LOG_AJUSTES, "NF " + current_doc + " - Corrigir combinação CST|CFOP de " +
                    old_cst + "|" + line[2] + " para " + line[1] + "|" + line[2] + ".", silent=True)
        log(LOG_SAIDA, "\n\n")
        log(LOG_AJUSTES, "\n\n", silent=True)
    return efd_array

def fix_removeDuplicates(efd_array, LOG_SAIDA):
    wannafix = False
    participants_array = []
    items_array = []
    inventory_array = []
    for line in efd_array:
        if (line[0] == "0150" and line[1] in participants_array) or (line[0] == "0200" and line[1] in items_array) or (line[0] == "H010" and line[1] in inventory_array):
            ans = sg.popup_yes_no("Foram encontradas duplicatas no registro de participantes, itens e/ou no inventário. Deseja corrigir?", title="Atenção!")
            if ans == "Yes":
                wannafix = True
            break
        if line[0] == "0150":
            participants_array.append(line[1])
        if line[0] == "0200":
            items_array.append(line[1])
        if line[0] == "H010":
            inventory_array.append(line[1])
    if wannafix:
        log(LOG_SAIDA, "CORREÇÕES DE DUPLICATAS DE ITENS E INVENTÁRIO:", silent=True)
        participants_array = []
        id_array = []
        inventory_array = []
        N = len(efd_array)
        new_sped = []
        i = 0
        while i < N:
            line = efd_array[i]
            if line[0] == '0150':
                participant_id = line[1]
                if participant_id not in participants_array:
                    participants_array.append(participant_id)
                    new_sped.append(line)
                else:
                    log(LOG_SAIDA, "Registro 0150: Participante " + participant_id + " removido.")
            elif line[0] == '0200':
                item_id = line[1]
                if item_id not in id_array:
                    id_array.append(item_id)
                    new_sped.append(line)
                else:
                    log(LOG_SAIDA, "Registro 0200: Item " + item_id + " removido.")
                    if efd_array[i+1][0] == "0220":
                        i += 1
                        log(LOG_SAIDA, "Registro 0220: Fator de conversão do item " +
                            item_id + " removido.")
            elif line[0] == 'H010':
                inventory_id = line[1]
                if inventory_id not in inventory_array:
                    inventory_array.append(inventory_id)
                    new_sped.append(line)
                else:
                    log(LOG_SAIDA, "Registro H010: Item " +
                        inventory_id + " removido.")
            else:
                new_sped.append(line)
            i += 1
        log(LOG_SAIDA, "\n")
        return new_sped
    else:
        return efd_array

def fix_unusedItems(efd_array, LOG_SAIDA):
    new_sped = []
    wannafix = False
    # Gerando lista de items referenciados
    referenced_participants = []
    referenced_items = []
    for line in efd_array:
        if line[0] == "0150":
            referenced_participants.append(line[1])
        if line[0] == "C170":
            referenced_items.append(line[2])
        if line[0] == "C425":
            referenced_items.append(line[1])
        if line[0] == "H010":
            referenced_items.append(line[1])
        if line[0] == "K200":
            referenced_items.append(line[2])
    for line in efd_array:
        if (line[0] == "0200" and line[1] not in referenced_items) or (line[0] == "0150" and line[1] not in referenced_participants):
            ans = sg.popup_yes_no("Foram encontrados itens ou participantes não referenciados em nenhum registro. Deseja corrigir?", title="Atenção!")
            if ans == "Yes":
                wannafix = True
            break
    # Removendo itens não utilizados
    if wannafix:
        log(LOG_SAIDA, "CORREÇÕES DE ITENS NÃO REFERENCIADOS:", silent=True)
        for i in range(0, len(efd_array)):
            if efd_array[i][0] == "0150":
                if efd_array[i][1] in referenced_participants:
                    new_sped.append(efd_array[i])
                else:
                    log(LOG_SAIDA, "O participante " + efd_array[i][1] + " foi excluído.")
                    
            if efd_array[i][0] == "0200":
                if efd_array[i][1] in referenced_items:
                    new_sped.append(efd_array[i])
                else:
                    log(LOG_SAIDA, "O item " + efd_array[i][1] + " foi excluído.")
                    if efd_array[i+1][0] == "0220":
                        i+=1
                        log(LOG_SAIDA, "O fator de conversão do item " + efd_array[i][1] + " também foi excluído.")
        #for line in efd_array:
        #    if line[0] == "0200":
        #        if line[1] in referenced_items:
        #            new_sped.append(line)
        #        else:
        #            log(LOG_SAIDA, "O item " + line[1] + " foi excluído.")
        #            continue
            else:
                new_sped.append(efd_array[i])
        log(LOG_SAIDA, "\n")
        return new_sped
    else:
        return efd_array

def fix_inventory(efd_array, LOG_SAIDA):
    wannafix = False
    old_inv_value = 0.0
    new_inv_value = 0.0
    for line in efd_array:
        if line[0] == "H005":
            old_inv_value = float(line[2].replace(",", "."))
            ans = sg.popup_yes_no("Foi encontrado o bloco H (Inventário), com valor total de " +
                        line[2] + ". Deseja corrigir o valor?", title="Atenção!")
            if ans == "Yes":
                wannafix = True
            break
    if wannafix:
        log(LOG_SAIDA, "CORREÇÕES DE INVENTÁRIO:", silent=True)
        # Validação da entrada para o novo valor de inventário
        while(True):
            input_str = sg.popup_get_text("Digite o novo valor total do inventário: ")
            try:
                new_inv_value = float(input_str.replace(",", "."))
                break
            except:
                sg.popup("VALOR INVÁLIDO. Insira novamente.\nLembre-se de digitar apenas números e uma única vírgula ou ponto para separar inteiros de valores decimais.\nExemplo: A quantia de 1.200.042,33 deverá ser inserida como 1200300.44 ou 1200300,44")
        # Proporção a ser aplicada nos itens do inventário
        mult_ratio = new_inv_value/old_inv_value
        # Atualização dos valores de inventário
        for line in efd_array:
            if line[0] == "H005":
                line[2] = str(round(new_inv_value, 2)).replace(".", ",")
                log(LOG_SAIDA, "Valor no registro H005 corrigido de " +
                    str(old_inv_value) + " para " + str(new_inv_value) + ".")
            if line[0] == "H010":
                try:
                    QTD = get_value(line[3])
                except:
                    QTD = 0.00
                    log(LOG_SAIDA, "ERRO. QTD não pôde ser lida.")
                    log(LOG_SAIDA, line[3])
                    sg.popup(
                        "Talvez seu arquivo SPED não seja válido.\nA quantidade de um item do inventário não pôde ser lida...")
                try:
                    VL_UNIT = get_value(line[4])
                except:
                    VL_UNIT = 0.00
                    log(LOG_SAIDA, "ERRO. VL_UNIT não pôde ser lida.")
                    log(LOG_SAIDA, line[4])
                    sg.popup(
                        "Talvez seu arquivo SPED não seja válido\nO valor unitário de um item do inventário não pôde ser lida...")
                # Depois de extrair os valores de QTD e VL_ITEM, recalculamos o valor do item usando a proporção anterior
                new_VL_UNIT = round(VL_UNIT * mult_ratio, 2)
                new_VL_ITEM = round(QTD * new_VL_UNIT, 2)
                line[5] = str(round(new_VL_ITEM, 2)).replace(".", ",")
                line[10] = str(round(new_VL_ITEM, 2)).replace(".", ",")
                log(LOG_SAIDA, "Valor do item " + line[1] + " atualizado de " + str(
                    round(VL_UNIT * QTD, 6)) + " para: " + str(new_VL_ITEM))
        # Depois de ajustar os valores individuais, basta apenas o ajuste residual final,
        # que é a diferença entre o valor total do inventário e as somas dos valores individuais dos itens
        # (normalmente esta diferença vem de problemas de arredondamento).
        # A correção é adicionar ou subtrair esta diferença no primeiro item do inventário.
        #
        # Primeiro calculamos a soma do inventário
        total_inventory_sum = 0.00
        for line in efd_array:
            if line[0] == "H010":
                total_inventory_sum += get_value(line[5])
        total_residual_difference = new_inv_value - total_inventory_sum
        log(LOG_SAIDA, "Ajustando diferença residual ao primeiro item do inventário...")
        log(LOG_SAIDA, "Valor do ajuste: " + str(total_residual_difference))
        for line in efd_array:
            if line[0] == "H010":
                # Esta condicional evita subtrair mais do que há no valor do item
                # Caso o item seja maior do que o módulo da diferença, é permitido a subtração.
                # Caso o valor residual não seja negativo, então é sempre possível a operação.
                if get_value(line[5]) < (total_residual_difference * -1):
                    continue
                line[5] = set_value(get_value(line[5]) +
                                    total_residual_difference)
                line[10] = line[5]
                line[4] = set_value(get_value(line[5])/get_value(line[3]))
                log(LOG_SAIDA, "Valor adicionado ao item " + line[1] + ".")
                break
        log(LOG_SAIDA, "\n")
    return efd_array

def fix_simples(efd_array, XML_ROOT_DIR, LOG_SAIDA, LOG_AJUSTES):
    nfe_key = ""
    simples_nfe_set = set()
    # Primeiro conferimos nos XMLs quais notas são de fornecedores do Simples Nacional
    for xml_file in os.listdir(XML_ROOT_DIR):
        xml_path = XML_ROOT_DIR + xml_file
        try:
            tree = ET.parse(xml_path)
            root = tree.getroot()
            namespace = "{http://www.portalfiscal.inf.br/nfe}"
            nfe_key = ""
            for aux in root.iter(namespace + "infNFe"):
                nfe_key = aux.attrib["Id"].replace("NFe", "")
                break
            for _ in root.iter(namespace + "CSOSN"):
                simples_nfe_set.add(nfe_key)
                break
        except:
            None
    # Agora verificamos se houve alguma inconsistência nos lançamentos do SPED.
    current_doc_key = ""
    wannafix = False
    for line in efd_array:
        if line[0] == "C100":
            current_doc_key = line[8]
        if line[0] == "C170":
            if current_doc_key in simples_nfe_set:
                item_is_taxed = line[9][1:] == "00"
                bc_icms = get_value(line[12])
                aliq_icms = get_value(line[13])
                vl_icms = get_value(line[14])
                aux = bc_icms + aliq_icms + vl_icms
                if item_is_taxed or aux != 0:
                    ans = sg.popup_yes_no("Foram encontradas notas do Simples Nacional com CST de tributados ou com valores de ICMS. Deseja corrigir?", title="Atenção!")
                    if ans == "Yes":
                        wannafix = True
                    break
    if wannafix:
        log(LOG_AJUSTES, "CORREÇÕES DE NOTAS DO SIMPLES NACIONAL:", silent=True)
        for line in efd_array:
            if line[0] == "C100":
                current_doc = line[7]
                current_doc_key = line[8]
                if current_doc_key in simples_nfe_set:
                    line[20] = "0,00"
                    line[21] = "0,00"
            if line[0] == "C170":
                if current_doc_key in simples_nfe_set:
                    item_id = line[2]
                    item_is_taxed = line[9][1:] == "00"
                    bc_icms = get_value(line[12])
                    aliq_icms = get_value(line[13])
                    vl_icms = get_value(line[14])
                    if item_is_taxed:
                        old_cst = line[9]
                        line[9] = line[9][0] + "90"
                        log(LOG_SAIDA, "NFe: " + current_doc_key + " - CST corrigido de " + old_cst + " para " + line[9] + " no item " + item_id + ".")
                    if bc_icms != 0:
                        line[12] = "0,00"
                        log(LOG_SAIDA, "NFe: " + current_doc_key + " - Base de cálculo de ICMS zerada no item " + item_id + ".")
                    if aliq_icms != 0:
                        line[13] = "0,00"
                        log(LOG_SAIDA, "NFe: " + current_doc_key + " - Alíquota de ICMS zerada no item " + item_id + ".")
                    if vl_icms != 0:
                        line[14] = "0,00"
                        log(LOG_SAIDA, "NFe: " + current_doc_key + " - Valor de ICMS zerado no item " + item_id + ".")
            if line[0] == "C190":
                if current_doc_key in simples_nfe_set:
                    is_taxed = line[1][1:] == "00"
                    bc_icms = get_value(line[5])
                    aliq_icms = get_value(line[3])
                    vl_icms = get_value(line[6])
                    aux_sum = bc_icms + aliq_icms + vl_icms
                    if is_taxed:
                        old_cst = line[1]
                        cfop = line[2]
                        line[1] = line[1][0] + "90"
                        line[3] = "0,00"
                        line[5] = "0,00"
                        line[6] = "0,00"
                        aux_str = " Zerar base de cálculo, alíquota e valor de ICMS." if aux_sum != 0 else ""
                        log(LOG_AJUSTES, "NF " + current_doc + " - Corrigir combinação CST|CFOP de " + old_cst + "|" + cfop + " para " + line[1] + "|" + line[2] + "." + aux_str, silent=True)
        log(LOG_AJUSTES, "\n\n")
    return efd_array

def suggest_bonifications_corrections(efd_array, XML_ROOT_DIR, LOG_SAIDA, LOG_AJUSTES):
    nfe_key = ""
    efd_bonification_set = set()
    xml_bonification_set = set()
    efd_keys_set = set()
    # Primeiro encontramos todos os itens
    # identificados como bonificação no SPED
    nfe_key = ""
    for line in efd_array:
        if line[0] == "C100":
            nfe_key = line[8]
            efd_keys_set.add(nfe_key)
        if line[0] == "C170":
            try:
                sequence_num = int(line[1])
            except:
                sg.popup("Não foi possível converter o número de sequência " + str(line[1]) + " da NF " + str(nfe_key) + ".")
            cfop = line[10]
            if cfop[1:] == "910":
                efd_bonification_set.add((nfe_key, sequence_num))
    # Agora conferimos nos XMLs quais itens
    # de fato foram emitidos como bonificação
    for xml_file in os.listdir(XML_ROOT_DIR):
        xml_path = XML_ROOT_DIR + xml_file
        try:
            tree = ET.parse(xml_path)
            root = tree.getroot()
            namespace = "{http://www.portalfiscal.inf.br/nfe}"
            nfe_key = ""
            for aux in root.iter(namespace + "infNFe"):
                nfe_key = aux.attrib["Id"].replace("NFe", "")
                break
            for nfe_item in root.iter(namespace + "det"):
                try:
                    nItem = int(nfe_item.attrib["nItem"])
                except:
                    sg.popup("Não foi possível converter o número de sequência " + str(nfe_item.attrib["nItem"]) + " da NF " + str(nfe_key))
                prod = nfe_item.find(namespace + "prod")
                cfop = prod.find(namespace + "CFOP").text
                if cfop[1:] == "910" and nfe_key in efd_keys_set:
                    xml_bonification_set.add((nfe_key, nItem))
        except:
            None
    if not xml_bonification_set:
        return
    #Agora construímos os dois conjuntos (itens presentes no SPED mas não nos XML | itens presentes nos XMLs mas não no SPED)
    #As subtrações desses conjuntos nos dão os dois resultados que queremos
    not_recorded = xml_bonification_set.difference(efd_bonification_set)
    if not_recorded:
        log(LOG_SAIDA, "Foram encontrados nos XMLs itens de bonificação que talvez NÃO TENHAM SIDO ESCRITURADOS.")
        log(LOG_SAIDA, "Os resultados detalhados estarão no log de ajustes (log_ajustes.txt).")
        log(LOG_AJUSTES, "ITENS DE BONIFICAÇÃO NÃO ESCRITURADOS:", silent=True)
        for item in not_recorded:
            log(LOG_AJUSTES, "NF: " + item[0] + " - Item " + str(item[1]) + ".")
    recorded_incorrectly = efd_bonification_set.difference(xml_bonification_set)
    if recorded_incorrectly:
        log(LOG_SAIDA, "\n")
        log(LOG_AJUSTES, "\n", silent=True)
        log(LOG_SAIDA, "Foram encontrados no SPED itens que talvez tenham sido ESCRITURADOS INCORRETAMENTE COMO BONIFICAÇÃO.")
        log(LOG_SAIDA, "Os resultados detalhados estarão no log de ajustes (log_ajustes.txt).")
        log(LOG_AJUSTES, "ITENS INCORRETAMENTE ESCRITURADOS COMO BONIFICAÇÃO:", silent=True)
        for item in recorded_incorrectly:
            log(LOG_AJUSTES, "NF: " + item[0] + " - Item " + str(item[1]) + ".")
    if not_recorded or recorded_incorrectly:
        sg.popup("Foram encontradas inconsistências nas\nNOTAS DE BONIFICAÇÃO.\nConfira o log de ajustes para mais informações!", title="Atenção!")

def suggest_fuel_corrections(efd_array, XML_ROOT_DIR, LOG_SAIDA, LOG_AJUSTES):
    nfe_key = ""
    issue_cfops = ["5910", "6910"]
    receipt_cfops = ["1910", "2910"]
    efd_bonification_set = set()
    xml_bonification_set = set()
    efd_keys_set = set()
    # Primeiro encontramos todos os itens
    # identificados como bonificação no SPED
    nfe_key = ""
    for line in efd_array:
        if line[0] == "C100":
            nfe_key = line[8]
            efd_keys_set.add(nfe_key)
        if line[0] == "C170":
            sequence_num = line[1]
            cfop = line[10]
            if cfop in receipt_cfops:
                efd_bonification_set.add((nfe_key, sequence_num))
    # Agora conferimos nos XMLs quais itens
    # de fato foram emitidos como bonificação
    for xml_file in os.listdir(XML_ROOT_DIR):
        xml_path = XML_ROOT_DIR + xml_file
        try:
            tree = ET.parse(xml_path)
            root = tree.getroot()
            namespace = "{http://www.portalfiscal.inf.br/nfe}"
            nfe_key = ""
            for aux in root.iter(namespace + "infNFe"):
                nfe_key = aux.attrib["Id"].replace("NFe", "")
                break
            for nfe_item in root.iter(namespace + "det"):
                nItem = nfe_item.attrib["nItem"]
                prod = nfe_item.find(namespace + "prod")
                cfop = prod.find(namespace + "CFOP").text
                if cfop in issue_cfops and nfe_key in efd_keys_set:
                    xml_bonification_set.add((nfe_key, nItem))
        except:
            None
    if not xml_bonification_set:
        return
    #Agora construímos os dois conjuntos (itens presentes no SPED mas não nos XML | itens presentes nos XMLs mas não no SPED)
    #As subtrações desses conjuntos nos dão os dois resultados que queremos
    not_recorded = xml_bonification_set.difference(efd_bonification_set)
    if not_recorded:
        log(LOG_SAIDA, "Foram encontrados nos XMLs itens de bonificação que talvez NÃO TENHAM SIDO ESCRITURADOS.")
        log(LOG_SAIDA, "Os resultados detalhados estarão no log de ajustes (log_ajustes.txt).")
        log(LOG_AJUSTES, "ITENS DE BONIFICAÇÃO NÃO ESCRITURADOS:", silent=True)
        for item in not_recorded:
            log(LOG_AJUSTES, "NF: " + item[0] + " - Item " + item[1] + ".")
    recorded_incorrectly = efd_bonification_set.difference(xml_bonification_set)
    if recorded_incorrectly:
        log(LOG_SAIDA, "\n")
        log(LOG_AJUSTES, "\n", silent=True)
        log(LOG_SAIDA, "Foram encontrados no SPED itens que talvez tenham sido ESCRITURADOS INCORRETAMENTE COMO BONIFICAÇÃO.")
        log(LOG_SAIDA, "Os resultados detalhados estarão no log de ajustes (log_ajustes.txt).")
        log(LOG_AJUSTES, "ITENS INCORRETAMENTE ESCRITURADOS COMO BONIFICAÇÃO:", silent=True)
        for item in recorded_incorrectly:
            log(LOG_AJUSTES, "NF: " + item[0] + " - Item " + item[1] + ".")
    if not_recorded or recorded_incorrectly:
        sg.popup("Foram encontradas inconsistências nas\nNOTAS DE BONIFICAÇÃO.\nConfira o log de ajustes para mais informações!", title="Atenção!")


def get_codItem_simples(efd_array, nfe_key, nItem):
    CREDITED_CFOPS = ["1101", "1102", "2101", "2102", ]
    current_doc = ""
    for line in efd_array:
        if line[0] == "C100":
            current_doc = line[8]
        if line[0] == "C170":
            if current_doc == nfe_key and line[1] == nItem and line[10] in CREDITED_CFOPS:
                return line[2]

def get_simples_credit(efd_array, XML_ROOT_DIR, OUTPUT_FOLDER):
    efd_adjustments = []
    # O array de campos necessários para o lançamento de crédito do Simples é:
    # Data | NotaOrigem | CodItem | Descr | V Contábil | Alíq | V Créd
    # Lista de todas as chaves presentes na EFD
    spreadsheet_header = {
            "name" : "",
            "cnpj" : "",
            "ie" : "",
        }
    CREDIT_CSOSN = ["101", "201"]
    VC_SNC = 0.00    
    efd_keys_set = set()
    for line in efd_array:
        if line[0] == "C100":
            nfe_key = line[8]
            efd_keys_set.add(nfe_key)
        if line[0] == "0000":
            spreadsheet_header['name'] = line[5]
            spreadsheet_header['cnpj'] = line[6]
            spreadsheet_header['ie'] = line[9]
    # Primeiro conferimos se o usuário deseja gerar a planilha de crédito do Simples Nacional
    wannafix = False
    for xml_file in os.listdir(XML_ROOT_DIR):
        xml_path = XML_ROOT_DIR + xml_file
        try:
            tree = ET.parse(xml_path)
            root = tree.getroot()
            namespace = "{http://www.portalfiscal.inf.br/nfe}"
            icmssn101 = root.iter(namespace + "ICMSSN101")
            icmssn201 = root.iter(namespace + "ICMSSN201")
            infAdic = root.iter(namespace + "infAdic")
            descr_complementar = ""
            for aux in infAdic:
                descr_complementar += ET.tostring(aux, encoding='unicode')
            if not re.findall(".*permite.*cr[e|é]dito.*icms.*", descr_complementar.lower()):
                continue
            if icmssn101 or icmssn201:
                ans = sg.popup_yes_no("Foram encontradas notas do Simples Nacional que podem permitir aproveitamento de crédito. Deseja gerar a planilha?", title="Atenção!")
                if ans == "Yes":
                    wannafix = True
                break
        except:
            None
    if wannafix:
        spreadsheet_rows = []
        for xml_file in os.listdir(XML_ROOT_DIR):
            try:
                xml_path = XML_ROOT_DIR + xml_file
                tree = ET.parse(xml_path)
                root = tree.getroot()
                namespace = "{http://www.portalfiscal.inf.br/nfe}"
                nfe_key = ""
                nfe_number = ""
                # Início da leitura do XML
                infNFe = root.find(namespace + "NFe").find(namespace + "infNFe")
                nfe_key = infNFe.attrib["Id"].replace("NFe", "")
                if nfe_key not in efd_keys_set:
                    continue
                # Verificando se a nota tem a descrição correta exigida pela Lei 123
                infAdic = root.iter(namespace + "infAdic")
                descr_complementar = ""
                for aux in infAdic:
                    descr_complementar += ET.tostring(aux, encoding='unicode')
                if not re.findall(".*permite.*cr[e|é]dito.*icms.*", descr_complementar.lower()):
                    continue
                # Se chegamos aqui, é porque a nota está escriturada e a descrição contém o necessário.
                emission_dh = infNFe.find(namespace + "ide").find(namespace + "dhEmi").text
                aux = emission_dh[0:10].split("-")
                emission_date = aux[2] + "/" + aux[1] + "/" + aux[0]
                nfe_number = infNFe.find(namespace + "ide").find(namespace + "nNF").text
                # Leitura dos itens
                for nfe_item in root.iter(namespace + "det"):
                    nItem = nfe_item.attrib["nItem"]
                    cod_item = get_codItem_simples(efd_array, nfe_key, nItem)
                    # Se a função que retorna o código retornar None é porque o item não foi encontrado
                    #  ou o CFOP é de uso e consumo. Nesse caso, não aproveitamos crédito.
                    if not cod_item:                        
                        continue
                    prod = nfe_item.find(namespace + "prod")
                    description = prod.find(namespace + "xProd").text
                    v_contabil = float(prod.find(namespace + "vProd").text)
                    sn101 = nfe_item.find(namespace + "imposto").iter(namespace + "ICMSSN101")
                    sn201 = nfe_item.find(namespace + "imposto").iter(namespace + "ICMSSN201")
                    # Dentro da tag ICMSSN101 vamos encontrar valores de crédito do Simples
                    for icmssn101 in sn101:
                        try:
                            csosn = icmssn101.find(namespace + "CSOSN").text
                        except:
                            print("Não foi possível encontrar o CSOSN.")
                            continue
                        try:
                            aliq = icmssn101.find(namespace + "pCredSN").text + "%"
                        except:
                            print("Não foi possível encontrar pCredSN.")
                            continue
                        try:
                            v_credit = float(icmssn101.find(namespace + "vCredICMSSN").text)
                        except:
                            print("Não foi possível encontrar vCredICMSSN.")
                            continue
                        if csosn in CREDIT_CSOSN:
                            print(emission_date, "|", nfe_number, "|", nItem, "|", cod_item, "|", description, "|", v_contabil, "|", aliq, "|", v_credit)
                            VC_SNC += float(v_credit)
                            spreadsheet_rows.append([emission_date, nfe_number, nItem, cod_item, description, v_contabil, aliq, v_credit])
                            efd_adjustments.append([nfe_key, cod_item, v_contabil, aliq.replace("%", ""), v_credit])
                        break
                    # Agora percorremos as tags de CSOSN com final 201
                    for icmssn201 in sn201:
                        try:
                            csosn = icmssn201.find(namespace + "CSOSN").text
                        except:
                            print("Não foi possível encontrar o CSOSN.")
                            continue
                        try:
                            aliq = icmssn201.find(namespace + "pCredSN").text + "%"
                        except:
                            print("Não foi possível encontrar pCredSN.")
                            continue
                        try:
                            v_credit = float(icmssn201.find(namespace + "vCredICMSSN").text)
                        except:
                            print("Não foi possível encontrar vCredICMSSN.")
                            continue
                        if csosn in CREDIT_CSOSN:
                            print(emission_date, "|", nfe_number, "|", nItem, "|", cod_item, "|", description, "|", v_contabil, "|", aliq, "|", v_credit)
                            VC_SNC += float(v_credit)
                            spreadsheet_rows.append([emission_date, nfe_number, nItem, cod_item, description, v_contabil, aliq, v_credit])
                            efd_adjustments.append([nfe_key, cod_item, v_contabil, aliq.replace("%", ""), v_credit])
                        break
            except:
                None
        # Aqui o processo todo já foi terminado e podemos escrever a planilha
        workbook = xlsxwriter.Workbook(OUTPUT_FOLDER + 'CRÉDITO SIMPLES - ' + spreadsheet_header['name'] + '.xlsx')
        # Formatações personalizadas (variáveis usadas nas funções do xlsxwriter)
        bold = workbook.add_format({'bold': True})
        cell_format = workbook.add_format({
                                          'border': 1,
                                          })
        header_format = workbook.add_format({'align': 'center',
                                          'valign': 'vcenter',
                                          'border': 1,
                                          })
        bold_header_format = workbook.add_format({'align': 'center',
                                          'valign': 'vcenter',
                                          'border': 1,
                                          'bold' : True
                                          })
        header_format.set_bg_color("#68DAFE")
        bold_header_format.set_bg_color("#68DAFE")
        # Agora podemos começar a escrever o arquivo
        worksheet = workbook.add_worksheet("CRÉDITO DO SIMPLES")
        worksheet.set_column(0, 1, 12)
        worksheet.set_column(2, 2, 6)
        worksheet.set_column(3, 3, 12)
        worksheet.set_column(4, 4, 48)
        worksheet.set_column(5, 5, 10)
        worksheet.set_column(6, 6, 8)
        worksheet.set_column(7, 7, 12)
        worksheet.set_tab_color('white')
        # O cabeçalho padrão da planilha contém informações básicas e o nome da planilha
        worksheet.merge_range('A1:H1',"", header_format)
        worksheet.write_rich_string('A1', bold, "EMPRESA: ", spreadsheet_header['name'], header_format)
        worksheet.merge_range('A2:H2', "", header_format)
        worksheet.write_rich_string('A2', bold, "CNPJ: ", spreadsheet_header['cnpj'], header_format)
        worksheet.merge_range('A3:H3', "", header_format)
        worksheet.write_rich_string('A3', bold, "IE: ", spreadsheet_header['ie'], header_format)
        worksheet.merge_range('A5:H5', "PERMISSÃO DE CRÉDITO DE ICMS POR AQUISIÇÃO DE OPTANTE DO SIMPLES NACIONAL", bold_header_format)
        # Linhas de separação
        worksheet.merge_range('A4:H4', "")
        worksheet.merge_range('A6:H6', "")
        # O cabeçalho dos dados tem todas as informações coletadas previamente nos XML
        data_headers = ["Data", "Nota Origem", "nItem", "Cod. Item", "Descrição", "V. Cont", "Aliq.", "Valor Crédito"]
        worksheet.write_row("A7", data_headers, bold_header_format)
        aux = 7
        for row in spreadsheet_rows:
            worksheet.write_row(aux, 0, row, cell_format)
            aux += 1
        worksheet.merge_range("A"+str(aux+1)+":"+"G"+str(aux+1), "")
        worksheet.write(aux, 7, VC_SNC, bold_header_format)
        workbook.close()
        print("Total de crédito do Simples: ", VC_SNC)
        sg.popup("Planilha gerada. Pressione OK para continuar...")
    return efd_adjustments

def fix_simples_adjustments(efd_array, efd_adjustments):
    if not efd_adjustments:
        return efd_array
    nfe_key_set = set()
    for adj in efd_adjustments:
        nfe_key_set.add(adj[0])
    # Agora vamos lançar os registros 0460 e C195
    adj_0460 = ['0460', 'csimp', 'Credito por aquisição de empresa optante pelo Simples Nacional conforme Art. 23 da Lei Completar 123/2006']
    i = 0
    # Registro 0460 
    for i in range(0, len(efd_array)):
        if efd_array[i][0] > "0460":
            efd_array = efd_array[:i] + [adj_0460] + efd_array[i:]
            break
    # Registros C195 - Adicionamos o registro em cada uma das notas
    adj_c195 = ['C195', 'csimp', 'Credito por aquisição de empresa optante pelo Simples Nacional conforme Art. 23 da Lei Completar 123/2006']
    current_doc = ""
    for key in nfe_key_set:
        for i in range(0, len(efd_array)):
            if current_doc == key and (efd_array[i][0] == "C100" or efd_array[i][0] > "C197"):
                efd_array = efd_array[:i] + [adj_c195] + efd_array[i:]
                break
            if efd_array[i][0] == "C100":
                current_doc = efd_array[i][8]
                continue
    # Registro C197 - Precisamos adicionar os registros C197 embaixo dos registros C195 que adicionamos previamente
    for i in range(0, len(efd_array)):
        if efd_array[i][0] == "C100":
            current_doc = efd_array[i][8]
        if efd_array[i][0] == "C195" and efd_array[i][1] == "csimp":
            adjs_c197 = []
            # Adjustments array order: nfe_key, cod_item, v_contabil, aliq, v_credit
            for adj in efd_adjustments:
                if adj[0] == current_doc:
                    aux = 'Credito por aquisição de empresa optante pelo Simples Nacional conforme Art. 23 da Lei Completar 123/2006'
                    adjs_c197.append(['C197', 'MG10990505', aux, adj[1], set_value(adj[2]), set_value(adj[3]), set_value(adj[4]), ''])
            efd_array = efd_array[:i+1] + adjs_c197 + efd_array[i+1:]
    # É preciso adicionar/atualizar o contador dos registros 0460, já que um novo registro foi adicionado.
    has_0460_counter = False
    counter_0460 = ['9900', '0460', '']
    for line in efd_array:
        if line[0] == "9900" and line[1] == "0460":
            has_0460_counter = True
    for i in range(0, len(efd_array)):
        if (not has_0460_counter) and efd_array[i][0] == '9900' and efd_array[i][1] > '0460':
            efd_array = efd_array[:i] + [counter_0460] + efd_array[i:]
            break
    # Agora verificamos se a escrituração tem os contadores 9900|C195 e 9900|C197
    # Se sim, eles serão atualizados depois. Se não, precisamos inseri-los agora.
    has_c195_counter = False
    has_c197_counter = False
    c195_counter = ['9900', 'C195', '0']
    c197_counter = ['9900', 'C197', '0']
    for line in efd_array:
        if line[0] == "9900" and line[1] == "C195":
            has_c195_counter = True
        if line[0] == "9900" and line[1] == "C197":
            has_c197_counter = True
    for i in range(0, len(efd_array)):
        if (not has_c195_counter) and efd_array[i][0] == '9900' and efd_array[i][1] > 'C195':
            efd_array = efd_array[:i] + [c195_counter] + efd_array[i:]
            break
    for i in range(0, len(efd_array)):
        if (not has_c197_counter) and efd_array[i][0] == '9900' and efd_array[i][1] > 'C197':
            efd_array = efd_array[:i] + [c197_counter] + efd_array[i:]
            break
    return efd_array

def update_counters(efd_array, LOG_SAIDA):
    # Definição dos contadores
    counter_0150 = 0
    counter_0200 = 0
    counter_0220 = 0
    counter_0460 = 0
    counter_c195 = 0
    counter_c197 = 0
    counter_C990 = 0
    counter_H010 = 0
    counter_H990 = 0
    counter_0990 = 0
    counter_9900 = 0
    counter_9990 = 0
    counter_9999 = 0
    # Contagem
    for line in efd_array:
        counter_9999 += 1
        if line[0][0] == '0':
            counter_0990 += 1
        if line[0][0] == 'C':
            counter_C990 += 1
        if line[0][0] == 'H':
            counter_H990 += 1
        if line[0][0] == '9':
            counter_9990 += 1
        if line[0] == '0150':
            counter_0150 += 1
        if line[0] == '0200':
            counter_0200 += 1
        if line[0] == '0220':
            counter_0220 += 1
        if line[0] == '0460':
            counter_0460 += 1
        if line[0] == 'C195':
            counter_c195 += 1
        if line[0] == 'C197':
            counter_c197 += 1
        if line[0] == 'H010':
            counter_H010 += 1
        if line[0] == '9900':
            counter_9900 += 1
    # Alocando os contadores
    updated = False
    for line in efd_array:
        if line[0] == '0990':
            old_counter = line[1]
            line[1] = str(counter_0990)
            if old_counter != line[1]:
                log(LOG_SAIDA, "Registro contador 0990 atualizado de " +
                    old_counter + " para " + line[1] + ".")
                updated = True
        if line[0] == 'C990':
            old_counter = line[1]
            line[1] = str(counter_C990)
            if old_counter != line[1]:
                log(LOG_SAIDA, "Registro contador C990 atualizado de " +
                    old_counter + " para " + line[1] + ".")
                updated = True
        elif line[0] == 'H990':
            old_counter = line[1]
            line[1] = str(counter_H990)
            if old_counter != line[1]:
                log(LOG_SAIDA, "Registro contador H990 atualizado de " +
                    old_counter + " para " + line[1] + ".")
                updated = True
        elif line[0] == '9990':
            old_counter = line[1]
            line[1] = str(counter_9990)
            if old_counter != line[1]:
                log(LOG_SAIDA, "Registro contador 9990 atualizado de " +
                    old_counter + " para " + line[1] + ".")
                updated = True
        elif line[0] == '9900' and line[1] == '0150':
            old_counter = line[2]
            line[2] = str(counter_0150)
            if old_counter != line[2]:
                log(LOG_SAIDA, "Registro contador 9900|0150 atualizado de " +
                    old_counter + " para " + line[2] + ".")
                updated = True
        elif line[0] == '9900' and line[1] == '0200':
            old_counter = line[2]
            line[2] = str(counter_0200)
            if old_counter != line[2]:
                log(LOG_SAIDA, "Registro contador 9900|0200 atualizado de " +
                    old_counter + " para " + line[2] + ".")
                updated = True
        elif line[0] == '9900' and line[1] == '0220':
            old_counter = line[2]
            line[2] = str(counter_0220)
            if old_counter != line[2]:
                log(LOG_SAIDA, "Registro contador 9900|0220 atualizado de " +
                    old_counter + " para " + line[2] + ".")
                updated = True
        elif line[0] == '9900' and line[1] == '0460':
            old_counter = line[2]
            line[2] = str(counter_0460)
            if old_counter != line[2]:
                log(LOG_SAIDA, "Registro contador 9900|0460 atualizado de " +
                    old_counter + " para " + line[2] + ".")
                updated = True
        elif line[0] == '9900' and line[1] == 'C195':
            old_counter = line[2]
            line[2] = str(counter_c195)
            if old_counter != line[2]:
                log(LOG_SAIDA, "Registro contador 9900|C195 atualizado de " +
                    old_counter + " para " + line[2] + ".")
                updated = True
        elif line[0] == '9900' and line[1] == 'C197':
            old_counter = line[2]
            line[2] = str(counter_c197)
            if old_counter != line[2]:
                log(LOG_SAIDA, "Registro contador 9900|C197 atualizado de " +
                    old_counter + " para " + line[2] + ".")
                updated = True
        elif line[0] == '9900' and line[1] == '9900':
            old_counter = line[2]
            line[2] = str(counter_9900)
            if old_counter != line[2]:
                log(LOG_SAIDA, "Registro contador 9900|9900 atualizado de " +
                    old_counter + " para " + line[2] + ".")
                updated = True
        elif line[0] == '9900' and line[1] == 'H010':
            old_counter = line[2]
            line[2] = str(counter_H010)
            if old_counter != line[2]:
                log(LOG_SAIDA, "Registro contador 9900|H010 atualizado de " +
                    old_counter + " para " + line[2] + ".")
                updated = True
        elif line[0] == '9999':
            old_counter = line[1]
            line[1] = str(counter_9999)
            if old_counter != line[1]:
                log(LOG_SAIDA, "Registro contador 9999 atualizado de " +
                    old_counter + " para " + line[1] + ".")
                updated = True
    if updated:
        sg.popup("Os registros contadores foram atualizados!", title="Atenção!")
    return efd_array

def main():
    sg.theme("BlueMono")
    efd_filebrowser = [
        sg.In(size=(65, 1), enable_events=True, key="EFD"),
        sg.FileBrowse(button_text="Abrir"),
    ]
    sheet_filebrowser = [
        sg.In(size=(65, 1), enable_events=True, key="XMLS"),
        sg.FolderBrowse(button_text="Abrir"),
    ]
    output_folderbrowser = [
        sg.In(size=(60, 1), enable_events=True, key="OUTPUT_FOLDER"),
        sg.FolderBrowse(button_text="Selecionar"),
    ]
    texts = [
        [sg.Text("Bem-vindo ao SpedFIX V2!\n\nSelecione seu arquivo SPED de entrada e o local onde estão os XMLs.\nInforme também onde deseja gerar o novo SPED.\n\n", justification='c')],
        [sg.Text("EFD de entrada: ")],
        [sg.Text("Pasta dos XMLs: ")],
        [sg.Text("Diretório onde deseja gerar a nova EFD e os demais arquivos: ")],
    ]
    send_button = [sg.Column([[sg.Button('FIX SPED!',key="-SEND-")]], justification='c')]
    progress_log = [sg.Output(size=(70,10), key='-OUTPUT-')]
    layout = [[sg.Column([texts[0]], justification='c')], texts[1], efd_filebrowser, texts[2], sheet_filebrowser, [sg.HorizontalSeparator()], texts[3], output_folderbrowser, send_button, progress_log] 
    window = sg.Window("SpedFIX", layout, grab_anywhere=True)
    while True:
        event, values = window.read()
        # End program if user closes window or
        # presses the OK button
        if event == "-SEND-":
            if values['EFD'] and values['XMLS'] and values['OUTPUT_FOLDER']:
                try:
                    SPED_ENTRADA = values['EFD']
                    XML_ROOT_DIR = values['XMLS'] + "/"
                    OUTPUT_FOLDER = values['OUTPUT_FOLDER'] + "/"
                    SPED_SAIDA = OUTPUT_FOLDER + "EFD Saída.txt"
                    LOG_SAIDA = OUTPUT_FOLDER + "log_saida.txt"
                    LOG_AJUSTES = OUTPUT_FOLDER + "log_ajustes.txt"
                    efd = open(SPED_ENTRADA, "r", encoding="latin-1")
                    log_completo = open(LOG_SAIDA, "w+", encoding="latin-1")
                    log_ajustes = open(LOG_AJUSTES, "w+", encoding="latin-1")
                    log_completo.write("Correções feitas em " + str(datetime.datetime.now()))
                    log_completo.write("\n\n")
                    log_ajustes.write("Correções feitas em " + str(datetime.datetime.now()))
                    log_ajustes.write("\n\n")
                    log_completo.close()
                    log_ajustes.close()
                    efd_array = []
                    # Construindo a matriz do SPED
                    for line in efd:
                        try:
                            efd_array.append(line.replace("\n", "").split("|")[1:-1])
                            if line.replace("\n", "").split("|")[1:-1][0] == "9999":
                                break
                        except Exception as e:
                            log(LOG_SAIDA, e)
                            log(LOG_SAIDA, "ERRO NA ESTRUTURA DO ARQUIVO.")
                            log(LOG_SAIDA, "O arquivo está assinado?.")
                    # Chamada de funções de correção
                    efd_array = fix_removeIPI(efd_array, LOG_SAIDA)
                    efd_array = fix_removeABAT(efd_array, LOG_SAIDA)
                    efd_array = fix_bc_greater_than_opr(efd_array, LOG_SAIDA)
                    efd_array = fix_020_RED(efd_array, LOG_SAIDA)
                    efd_array = fix_importCST(efd_array, LOG_SAIDA, LOG_AJUSTES)
                    efd_array = fix_unusedItems(efd_array, LOG_SAIDA)
                    efd_array = fix_removeDuplicates(efd_array, LOG_SAIDA)
                    efd_array = fix_inventory(efd_array, LOG_SAIDA)
                    efd_array = fix_simples(efd_array, XML_ROOT_DIR, LOG_SAIDA, LOG_AJUSTES)
                    efd_adjustments = get_simples_credit(efd_array, XML_ROOT_DIR, OUTPUT_FOLDER)
                    efd_array = fix_simples_adjustments(efd_array, efd_adjustments)
                    suggest_bonifications_corrections(efd_array, XML_ROOT_DIR, LOG_SAIDA, LOG_AJUSTES)
                    efd_array = update_counters(efd_array, LOG_SAIDA)
                    # Escrevendo o SPED no arquivo final
                    efd.close()
                    write_efd(efd_array, SPED_SAIDA)
                    sg.popup("Prontinho!\nSeu novo SPED e os demais arquivos foram gerados!", title='Fim da Execução')
                except Exception as e:
                    print(e)
                    sg.popup("Verifique se selecionou os arquivos corretamente e se tem permissão de escrita no lugar selecionado para geração da EFD de saída.", title="Erro!")
                window.Element('EFD').Update('')
                window.Element('XMLS').Update('')
                window.Element('OUTPUT_FOLDER').Update('')
                window.Element('-OUTPUT-').Update('')
            else:
                sg.popup("Selecione os locais dos arquivos e informe o local onde deseja gerar a EFD de saída...", title="Atenção")
        if event == sg.WIN_CLOSED:
            break
    window.close()
    return

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        sg.popup("ERRO NA EXECUÇÃO DO SCRIPT:", e, title="Erro!")