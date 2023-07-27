import pandas as pd 
from datetime import time
import pandas as pd
from openpyxl import load_workbook
from descricao_tra import define_descricao

estados_regioes = {
    'AC': 'Norte',
    'AL': 'Nordeste',
    'AP': 'Norte',
    'AM': 'Norte',
    'BA': 'Nordeste',
    'CE': 'Nordeste',
    'DF': 'Centro-Oeste',
    'ES': 'Sudeste',
    'GO': 'Centro-Oeste',
    'MA': 'Nordeste',
    'MT': 'Centro-Oeste',
    'MS': 'Centro-Oeste',
    'MG': 'Sudeste',
    'PA': 'Norte',
    'PB': 'Nordeste',
    'PR': 'Sul',
    'PE': 'Nordeste',
    'PI': 'Nordeste',
    'RJ': 'Sudeste',
    'RN': 'Nordeste',
    'RS': 'Sul',
    'RO': 'Norte',
    'RR': 'Norte',
    'SC': 'Sul',
    'SP': 'Sudeste',
    'SE': 'Nordeste',
    'TO': 'Norte'
}

lider = {
    'DF': 'Gutemberg',
    'GO': 'Gutemberg',
    'MT': 'Gutemberg',
    'MS': 'Gutemberg',
    'AL': 'Saulo',
    'BA': 'Douglas',
    'CE': 'Douglas',
    'MA': 'Erivaldo',
    'PB': 'Saulo',
    'PE': 'Saulo',
    'PI': 'Luiz Bolzon',
    'RN': 'Saulo',
    'SE': 'Saulo',
    'AC': 'Erivaldo',
    'AP': 'Erivaldo',
    'AM': 'Erivaldo',
    'PA': 'Luiz Bolzon',
    'RO': 'Erivaldo',
    'RR': 'Erivaldo',
    'TO': 'Luiz Bolzon',
    'ES': 'Dolôr',
    'MG': 'Dolôr',
    'RJ': 'Dolôr',
    'SP': 'Elesandro',
    'PR': 'Sérgio Mukai',
    'RS': 'Sérgio Mukai',
    'SC': 'Sérgio Mukai'
}

causa_dicionario = {
    'Compressor': 'Infraestrutura - Sistema de Climatização',
    'Ventilador da evaporadora ': 'Infraestrutura - Sistema de Climatização',
    'Ventilador da evaporadora': 'Infraestrutura - Sistema de Climatização',
    'FAN-10': 'Infraestrutura - Sistema de Climatização',
    'Módulo de expansão 2157': 'Infraestrutura - Sistema de Climatização',
    'Transformador 220 - 24 ': 'Infraestrutura - Sistema de Climatização',
    'AGST - MP5000 M00 08000004P': 'Infraestrutura - Sistema de Climatização',
    'Válvula de expansão': 'Infraestrutura - Sistema de Climatização',
    'Ventilador da condensadora ': 'Infraestrutura - Sistema de Climatização',
    ': ZIEHL - Modelo: ZEIHL ABEGG': 'Infraestrutura - Sistema de Climatização',
    'PLC  SANRIO': 'Infraestrutura - Sistema de Climatização',
    'Módulo AGST - MP5000 M02 00080800': 'Infraestrutura - Sistema de Climatização',
    'Placa de circuito interno do Ar-condicionado SANRIO': 'Infraestrutura - Sistema de Climatização',
    'Tubulação de gás para ar-condicionado.': 'Infraestrutura - Sistema de Climatização',
    'Tampa da caixa R1': 'Infraestrutura - Sistema Retificador FCC',
    'Kit Ventilado': 'Infraestrutura - Sistema Retificador FCC',
    'RAU2X6U/A26R6AUKL40179/A26': 'Rádio - Equipamento Ericsson',
    'RAU2x7/A15Ericsson': 'Rádio - Equipamento Ericsson',
    'DVR/Câmeras/Fonte12V': 'Infraestrutura - Sistema de CFTV',
    'NVR-2 camera 22': 'Infraestrutura - Sistema de CFTV',
    'NVR': 'Infraestrutura - Sistema de CFTV',
    'KoDo PRO - Modelo: KCX-5700N': 'Infraestrutura - Sistema de CFTV',
    'SPVL-4': 'DWDM - Equipamento PADTEC',
    'SPVL-4SM': 'DWDM - Equipamento PADTEC',
    "2 Supervisores SMARTPACWEB / SNMP6 UR's 50A FLATPAC21 SPVL-901 Base para DPS3 BANDEJAS para RETIFICADORES FCC ELTEK": 'DWDM - Equipamento PADTEC',
    'SPVL-90': 'DWDM - Equipamento PADTEC',
    'Subbastidor 14uTM400# sobressalente OCM  com defeito SPVL-91# com defeito SSC  sobressalente': 'DWDM - Equipamento PADTEC',
    'SPVL 90': 'DWDM - Equipamento PADTEC',
    'SPVL-91 ': 'DWDM - Equipamento PADTEC',
    'Duas unidades SCMD3S1A#': 'DWDM - Equipamento PADTEC',
    'T100DCT-4JRT2L':'DWDM - Equipamento PADTEC',
    'T100DCT-4JRT2L	':'DWDM - Equipamento PADTEC',
    'TR400C93-QBF-QBF':'DWDM - Equipamento PADTEC',
    'T100DCT-4PTT2L': 'DWDM - Equipamento PADTEC',
    'SSC-BBAA - FAN 10':'DWDM - Equipamento PADTEC',
    'SSC-10':'DWDM - Equipamento PADTEC',
    'ROA4C301AWAHA':'DWDM - Equipamento PADTEC',
    'TM400C92QBFXHACA':'DWDM - Equipamento PADTEC',
    'TM400C92-DBF-XHF-CA.':'DWDM - Equipamento PADTEC',
    'TM400-9B':'DWDM - Equipamento PADTEC',
    'T100DCT-4JT2L':'DWDM - Equipamento PADTEC',
    'TM400C92-DBF-XHF-CA':'DWDM - Equipamento PADTEC',
    'TR400-9B':'DWDM - Equipamento PADTEC',
    'SSC-10 Modelo BBAA':'DWDM - Equipamento PADTEC',
    'LOA4C211AYAHA':'DWDM - Equipamento PADTEC',
    'SCME-4DP e CVA-4SRA':'DWDM - Equipamento PADTEC',
    '- SPVL-4SM':'DWDM - Equipamento PADTEC',
    'FAN-G8':'DWDM - Equipamento PADTEC',
    'SCMD3S1A':'DWDM - Equipamento PADTEC',
    'SCMD3S1A':'DWDM - Equipamento PADTEC',
    'T100DCT-4JTMYL':'DWDM - Equipamento PADTEC',
    '- Sobressalente necessário: Placa SSC-AAAA ':'DWDM - Equipamento PADTEC',
    'DM4001 - ETH24GX+2x10GX H Series': 'IP - Equipamento Datacom',
    'DM4001 L series ': 'IP - Equipamento Datacom',
    'DM4001 L Series': 'IP - Equipamento Datacom',
    'Chassi do switch DM4001 L Series - ETH24GX L Series': 'IP - Equipamento Datacom',
    'SFP 10 Gb  Modelo: 1200-SM-LL-L': 'IP - Equipamento Datacom',
    'DM4001 eth 24 GX L series': 'IP - Equipamento Datacom',
    'DM4000 - MPU512 ': 'IP - Equipamento Datacom',
    'DM4001 serie H - DATACOM - Fonte AC/DC': 'IP - Equipamento Datacom',
    '1KVA NB HDS LM S2': 'Infraestrutura - Nobreak',
    'QCAB': 'Infraestrutura - Balizamento de Torre',
    'Bomba injetora do GMG ': 'Infraestrutura - Grupo Motor Gerador',
    'Bateria MARCA -  DISBAL ; MODELO - S 150MD': 'Infraestrutura - Grupo Motor Gerador',
    'Bateria de GMG': 'Infraestrutura - Grupo Motor Gerador',
    'Gerador completo': 'Infraestrutura - Grupo Motor Gerador',
    'bateria do gerador  Optima - Modelo: Gel - Red Top 35 / 12 volts - 44ha - 720a (-18Cº) 910a 0Cº - RC90min': 'Infraestrutura - Grupo Motor Gerador',
    'DEEP SEA MODELO:DSE-7320': 'Infraestrutura - Grupo Motor Gerador',
    'Fonte da Telemetria Siemens': 'Infraestrutura - Sistema de Alarmes',
    'FTLB in 48V / out 24 V': 'Infraestrutura - Sistema de Alarmes',
    'OM-SMR100BR-TM-N': 'Infraestrutura - Sistema Retificador FCC',
    ': Emerson - Modelo: R48-3200': 'Infraestrutura - Sistema Retificador FCC',
    'UR 37A/48Vcc/412 Modelo OM - 1S37': 'Infraestrutura - Sistema Retificador FCC',
    'EMERSON - Modelo: R48-3200': 'Infraestrutura - Sistema Retificador FCC',
    'EMERSON R48 - 3200': 'Infraestrutura - Sistema Retificador FCC',
    'R48-3200': 'Infraestrutura - Sistema Retificador FCC',
    'Omibra OM1S50N': 'Infraestrutura - Sistema Retificador FCC',
    'FLATPAC 2 - 3kW': 'Infraestrutura - Sistema Retificador FCC',
    'Bateria Estacionaria Freedom DF2000 115Ah': 'Infraestrutura - Banco de Baterias'
}

arquivo = "Lista_modelos_bilhete.xls"
abertura_data = []
abertura_hora = []
abertura_mes = []
termino_data = []
termino_hora = []
termino_mes = []
estacao = []
descricao = []
uf = []
regiao = []
tipo_site = []
lider_regiao = []
categoria = []
subcategoria = []
prioridade = []
causa = []

def dividi_data(data_entrada, array_data, array_hora, array_mes):
    if data_entrada == '-':
        array_data.append('-')
        array_hora.append('')
        array_mes.append('-')
    else:
        data_dia, hora = map(str, data_entrada.split(' '))
        h, m ,s = map(int, hora.split(':'))
        array_data.append(data_dia)
        array_hora.append(time(h, m, s))
        array_mes.append(mes_ano(data_dia))

def mes_ano(data_entrada):
    meses = {
        '01': 'janeiro',
        '02': 'fevereiro',
        '03': 'março',
        '04': 'abril',
        '05': 'maio',
        '06': 'junho',
        '07': 'julho',
        '08': 'agosto',
        '09': 'setembro',
        '10': 'outubro',
        '11': 'novembro',
        '12': 'dezembro'
    }
    mes = data_entrada[3:5]
    ano = data_entrada[8:10]
    return f'{meses.get(mes)}-{ano}'

def estacao_id(entrada, entrada_2):
    if isinstance(entrada, float):
        estacao.append(entrada_2.lstrip()[0:11])
        descricao.append(define_descricao(entrada_2.lstrip()[0:11]))
    else:
        estacao.append(entrada)
        descricao.append(define_descricao(entrada))

def estacao_info(estacao_entrada):
    uf.append(estacao_entrada[0:2])
    regiao.append(estados_regioes.get(estacao_entrada[0:2]))
    tipo_site.append(estacao_entrada[6:8])
    lider_regiao.append(lider.get(estacao_entrada[0:2]))

def define_categoria(entrada):
    if isinstance(entrada, float) or entrada == None:
        categoria.append('')
        subcategoria.append('')
    elif ' - ' in entrada:
        cat, sub = map(str, entrada.split(' - '))
        categoria.append(cat)
        subcategoria.append(sub)
    else:
        categoria.append('')
        subcategoria.append('')

def define_prioridade(entrada):
    if isinstance(entrada, float):
        prioridade.append('Baixa')
    else:
        prioridade.append(entrada)

def define_causa(entrada, entrada2):
    if '-' in str(entrada):
        causa.append(entrada)
    else:
        entrada_tratada = str(entrada2).lstrip()
        entrada_tratada = entrada_tratada.upper()
        if 'RETIFICADOR' in entrada_tratada or 'ELTEK  - MODELO: FLATPACK 2 - 3000 W' in entrada_tratada:
            causa.append('Infraestrutura - Sistema Retificador FCC')
        elif 'TSDA' in entrada_tratada:
            causa.append('Infraestrutura - Sistema de Alarmes')
        elif 'INVERSOR' in entrada_tratada:
            causa.append('Infraestrutura - Sistema de Inversores')
        elif 'NOBREAK' in entrada_tratada:
            causa.append('Infraestrutura - Nobreak')
        elif 'BATERIAS' in entrada_tratada:
            causa.append('Infraestrutura - Banco de Baterias')
        elif 'RAU' in entrada_tratada or 'RAU2' in entrada_tratada:
            causa.append('Rádio - Equipamento Ericsson')
        elif 'MOTOR' in entrada_tratada and ('MWM' in entrada_tratada or 'MWM-D-229-4' in entrada_tratada or 'GERADOR' in entrada_tratada):
            causa.append('Infraestrutura - Grupo Motor Gerador')
        elif 'KIT VENTILADOR' in entrada_tratada or 'KIT DE VENTILADOR' in entrada_tratada or 'COMPRESSOR' in entrada_tratada:
            causa.append('Infraestrutura - Sistema de Climatização')
        elif 'HD' in entrada_tratada or 'DVR' in entrada_tratada or 'INTELBRAS' in entrada_tratada or 'SEAGATE' in entrada_tratada or 'CÂMERA' in str(entrada2).upper() or 'CAMERA' in str(entrada2).upper():
            causa.append('Infraestrutura - Sistema de CFTV')
        else:
            causa.append(causa_dicionario.get(entrada2))

arquivo_pd = pd.read_excel(arquivo)

for index, row in arquivo_pd.iterrows():
    dividi_data(row['Data abertura'], abertura_data, abertura_hora, abertura_mes)
    dividi_data(row['Término'], termino_data, termino_hora, termino_mes)
    estacao_id(row['Id da Estação'], row['Nome'])
    estacao_info(estacao[index])
    define_causa(row['Causa do alerta'], row['Sobressalente a ser verificado'])
    define_categoria(causa[index])
    define_prioridade(row['Nível de prioridade da VDS'])
        
df = {
    'Id ordem serviço': arquivo_pd['Id ordem serviço'],
    'Nome': arquivo_pd['Nome'],
    'Estado': arquivo_pd['Estado'],
    'Data abertura': abertura_data,
    'Hora Abertura': abertura_hora,
    'Mês abertura': abertura_mes,
    'Término': termino_data,
    'Hora Término': termino_hora,
    'Mês término': termino_mes,
    'Data ultima transição': arquivo_pd['Data ultima transição'],
    'Estação': estacao,
    'Descrição': descricao,
    'UF': uf,
    'Região': regiao,
    'Tipo Site': tipo_site,
    'Líder de Campo': lider_regiao,
    'Sobressalente a ser verificado': arquivo_pd['Sobressalente a ser verificado'],
    'Nível de prioridade da VDS': prioridade,
    'Causa do alerta': causa,
    'Categoria': categoria,
    'Subcategoria': subcategoria
}

pd_df = pd.DataFrame(df)

print(pd_df)

with pd.ExcelWriter('saida.xlsx', engine='xlsxwriter') as writer:
    pd_df.to_excel(writer, sheet_name='BaseVDS', index=False)
    
    workbook = writer.book
    worksheet = writer.sheets['BaseVDS']
    
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#002060',
        'font_color': '#FFFFFF',
        'border': 1
    })

    worksheet.autofilter(0, 0, 0, len(pd_df.columns) - 1)

    center_format = workbook.add_format({'align': 'center'})

   #worksheet.set_column(0, len(pd_df.columns) - 1, cell_format=center_format)

    for col_num, value in enumerate(pd_df.columns.values):
        worksheet.write(0, col_num, value, header_format)

book = load_workbook('saida.xlsx')
sheet = book['BaseVDS']

for col in sheet.columns:
    max_length = 0
    column = col[0].column_letter
    for cell in col:
        if len(str(cell.value)) > max_length:
            max_length = len(str(cell.value))
    adjusted_width = (max_length + 2)
    sheet.column_dimensions[column].width = adjusted_width

book.save('saida.xlsx')