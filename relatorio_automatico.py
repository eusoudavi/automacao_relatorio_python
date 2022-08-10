import openpyxl

# FORMATAR FUNÇÃO PARA PLOTAR O GŔAFICO NO LOCAL CERTO

# planilha = openpyxl.load_workbook('HUB_CalendárioAgosto_v5.xlsx')
planilha = openpyxl.load_workbook('teste.xlsx')

aba_pesquisada = 'AGOSTO'
aba_pesquisada = aba_pesquisada.upper().strip()
# CRIAR UMA ROTINA PARA CAPTURAR O MÊS ATUAL E PERGUNTAR AO USUÁRIO SE ELE DESEJA FAZER O RELATÓRIO ATUALIZADO DO MÊS

nome_do_consultor = 'Vanderson'
nome_do_consultor = nome_do_consultor.upper().strip()
# AJUSTAR PARA CRIAR UM RELATÓRIO ÚNICO COM TODOS OS CONSULTORES

jornal_imp = {'semanas': [0, 0, 0, 0, 0], "nome": 'JORNAL IMPRESSO'}
jornal_dig = {'semanas': [0, 0, 0, 0, 0], "nome": 'JORNAL DIGITAL'}
card_dig = {'semanas': [0, 0, 0, 0, 0], "nome": 'CARD DIGITAL'}
carrosel = {'semanas': [0, 0, 0, 0, 0], "nome": 'CARROSSEL'}
tv = {'semanas': [0, 0, 0, 0, 0], "nome": 'TV'}
radio = {'semanas': [0, 0, 0, 0, 0], "nome": 'RÁDIO'}
carro_som = {'semanas': [0, 0, 0, 0, 0], "nome": 'CARRO DE SOM'}
mercham = {'semanas': [0, 0, 0, 0, 0], "nome": 'MERCHAN'}

midias = [jornal_imp, jornal_dig, card_dig, carrosel, tv, radio, carro_som, mercham]

# ----------------CRIAÇÃO DAS ABAS (WORKSHEETS)----------------
if f'RELATÓRIO_{nome_do_consultor}' in planilha.sheetnames:
    aba_relatorio = planilha[f'RELATÓRIO_{nome_do_consultor}']
    # print('a aba já existe')
else:
    modelo = planilha[f'Modelo']
    aba_relatorio = planilha.copy_worksheet(modelo)
    aba_relatorio.title = f'RELATÓRIO_{nome_do_consultor}'

# ----------------PROCESSAMENTO DO RELATÓRIO----------------
# Colocar o nome do consultor
aba_relatorio['A1'].value = nome_do_consultor.capitalize()
# Varrer as abas da planilha para encontrar a aba onde deseja obter os dados
for abas in planilha.worksheets:                # Varrer as abas da planilha e
    if aba_pesquisada in abas.title.upper():    # verificar se há uma aba onde os dados serão consultados.
        for celulas in abas["A"]:               # Varrer a primeira coluna para encontrar o nome do consultor.
            if type(celulas.value) == str:      # Verificar se os valores são strings para eliminar as células vazias
                if nome_do_consultor in celulas.value.upper():      # IMPORTANTE
                    for i in range(5):          # Alimentação do Discionário com as informações da planilha
                        coluna_B = abas[f'B{celulas.row}'].value
                        coluna_C = abas[f'C{celulas.row}'].value
                        if f'SEMANA 0{i+1}' in coluna_B and jornal_imp['nome'] in coluna_C:
                            jornal_imp['semanas'][i] += 1
                        if f'SEMANA 0{i+1}' in coluna_B and jornal_dig['nome'] in coluna_C:
                            jornal_dig['semanas'][i] += 1
                        if f'SEMANA 0{i+1}' in coluna_B and card_dig['nome'] in coluna_C:
                            card_dig['semanas'][i] += 1
                        if f'SEMANA 0{i+1}' in coluna_B and carrosel['nome'] in coluna_C:
                            carrosel['semanas'][i] += 1
                        if f'SEMANA 0{i+1}' in coluna_B and tv['nome'] in coluna_C:
                            tv['semanas'][i] += 1
                        if f'SEMANA 0{i+1}' in coluna_B and radio['nome'] in coluna_C:
                            radio['semanas'][i] += 1
                        if f'SEMANA 0{i+1}' in coluna_B and carro_som['nome'] in coluna_C:
                            carro_som['semanas'][i] += 1
                        if f'SEMANA 0{i+1}' in coluna_B and mercham['nome'] in coluna_C:
                            mercham['semanas'][i] += 1

# Varrer as abas agora para encontrar onde o relatório foi criado para plotar os dados do Discionário
for abas in planilha.worksheets:
    if aba_relatorio.title in abas.title.upper():
        for celulas in abas["A"]:
            # print(celulas)
            if type(celulas.value) == str:
                # print(celulas.value.upper(), midias[1]['nome'])
                for i in range(len(midias)):
                    if celulas.value.upper().strip() == midias[i]['nome']:
                        abas[f'B{celulas.row}'].value = midias[i]['semanas'][0]
                        abas[f'C{celulas.row}'].value = midias[i]['semanas'][1]
                        abas[f'D{celulas.row}'].value = midias[i]['semanas'][2]
                        abas[f'E{celulas.row}'].value = midias[i]['semanas'][3]
                        abas[f'F{celulas.row}'].value = midias[i]['semanas'][4]
                        # print(celulas.row, celulas.value.upper())

# for dados in midias:
#     print(dados['nome'], dados['semanas'])

planilha.save('teste.xlsx')
