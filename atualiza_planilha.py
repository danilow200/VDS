import pandas as pd 
from datetime import time
import pandas as pd
from openpyxl import load_workbook

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
    else:
        estacao.append(entrada)

def estacao_info(estacao_entrada):
    uf.append(estacao_entrada[0:2])
    regiao.append(estados_regioes.get(estacao_entrada[0:2]))
    tipo_site.append(estacao_entrada[6:8])
    lider_regiao.append(lider.get(estacao_entrada[0:2]))

def define_categoria(entrada):
    if isinstance(entrada, float):
        categoria.append('')
        subcategoria.append('')
    elif ' - ' in entrada:
        cat, sub = map(str, entrada.split(' - '))
        categoria.append(cat)
        subcategoria.append(sub)
    else:
        categoria.append('')
        subcategoria.append('')

arquivo_pd = pd.read_excel(arquivo)

for index, row in arquivo_pd.iterrows():
    dividi_data(row['Data abertura'], abertura_data, abertura_hora, abertura_mes)
    dividi_data(row['Término'], termino_data, termino_hora, termino_mes)
    estacao_id(row['Id da Estação'], row['Nome'])
    descricao.append('')
    estacao_info(estacao[index])
    define_categoria(row['Causa do alerta'])
        
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
    'Nível de prioridade da VDS': arquivo_pd['Nível de prioridade da VDS'],
    'Causa do alerta': arquivo_pd['Causa do alerta'],
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

    worksheet.set_column(0, len(pd_df.columns) - 1, cell_format=center_format)

    worksheet.set_column(0, len(pd_df.columns) - 1, None, None, {'level': 1, 'hidden': True})

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
    adjusted_width = (max_length + 2) * 1.2
    sheet.column_dimensions[column].width = adjusted_width

book.save('saida.xlsx')