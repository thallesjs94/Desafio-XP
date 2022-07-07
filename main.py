# ========================================================================
# ======================= IMPORTAÇÃO DAS LIBS== ==========================
# ========================================================================
import os
import requests
import json
from openpyxl import Workbook

# ========================================================================
# ======================== VARIÁVEIS GLOBAIS=== ==========================
# ========================================================================
rootDir = os.path.abspath(os.curdir)


# Função responsive por consultar API e retornar a resposta como dicionário ou erro
def callApi(apiUrl):
    response = requests.get(apiUrl)

    if str(response) == '<Response [200]>':
        response_dict = json.loads(response.content)
        return response_dict
    else:
        return 'Erro ao consultar API'


# Função responsável por retornar o item mais frequente dentro de uma lista
def mostFrequent(List):
    counter = 0
    num = List[0]

    for i in List:
        currFrequency = List.count(i)
        if currFrequency > counter:
            counter = currFrequency
            num = i

    return num


# Função responsável por retornar a quantidade de vezes que um valor se repete em uma lista
def getFrequency(List, Value):
    return List.count(Value)


# Função responsável por retornar a quantidade de valores entre dois números, dentro uma lista
def totalLaunchBetweenValues(List, Value1, Value2):
    totalLaunch = 0
    for i in range(Value1, Value2):
        totalLaunch += getFrequency(List, str(i))
        return totalLaunch


# Função responsável por chamar todas as outras funções e gerar a planilha final
def generateResult():
    print('Iniciando processo')
    listYears = []
    listSites = []
    url = 'https://api.spacexdata.com/v3/launches'

    apiResponse = callApi(url)

    if apiResponse != 'Erro ao consultar API':
        try:
            wb = Workbook()

            sheet = wb.worksheets[0]
            sheet['A1'] = 'Ano com mais lançamentos'
            sheet['B1'] = 'Launch Site com mais lançamentos'
            sheet['C1'] = 'Total lançamentos entre 2019 e 2021'

            for flight in apiResponse:
                year = flight["launch_year"]
                siteId = flight["launch_site"]["site_id"]
                listYears.append(year)
                listSites.append(siteId)

            yearMostFrequent = mostFrequent(listYears)
            siteMostFrequent = mostFrequent(listSites)
            totalLaunch = totalLaunchBetweenValues(listYears, 2019, 2021)

            sheet['A2'] = yearMostFrequent
            sheet['B2'] = siteMostFrequent
            sheet['C2'] = totalLaunch

            sheet.title = 'Result'
            wb.save(rootDir + r'\Resultado.xlsx')
            wb.close()

            print('Arquivo ' + rootDir + r'\Resultado.xlsx' + ' gerado com sucesso!')

            print('Finalizando processo com sucesso')
        except Exception as e:
            print('Erro durante o processamento: ' + str(e))
    else:
        print('Erro ao consultar API, finalizando processo')


# ========================================================================
# ===================== PROCESSAMENTO PRINCIPAL ==========================
# ========================================================================
if __name__ == "__main__":
    generateResult()