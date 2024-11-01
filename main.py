from functions import formatar_relatorio_abc_servicos, formatar_relatorio_abc_insumos, formatar_relatorio_analitico, formatar_relatorio_analitico_preco_unit, formatar_relatorio_analitico_somente_insumos, formatar_relatorio_resumido, formatar_relatorio_sintetico, formatar_relatorio_sintetico_mo_eq_mat
from datetime import datetime
import os
from dotenv import load_dotenv

from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from datetime import datetime
from openpyxl.utils.cell import column_index_from_string
import locale

load_dotenv()

# Importação das funções

FOLDER = os.getenv('FOLDER')
user = os.getcwd().replace('\\', '/')
user = user[0:user.find('/', user.find('/', user.find('/') + 1) + 1)]
folder = user + FOLDER
os.chdir(folder)
files = [f for f in os.listdir('.') if os.path.isfile(f)]

# Solicitar data da base de dados
mes = -1
ano = 0
while mes < 0 or mes > 12:
  mes = int(input('-> Informe o MÊS da base de dados (de 1 a 12) ou digite 0 para cancelar:\n-> '))
  if mes == 0:
    print('-> Procedimento cancelado...')
    exit()
meses = {1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril', 5: 'Maio', 6: 'Junho',
         7: 'Julho', 8: 'Agosto', 9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'}
while ano < 1:
  ano = int(input('-> Informe o ANO da base de dados:\n-> '))
base_dados = {'mes': meses[mes], 'ano': str(ano)}
print('')

# Iniciando o contador de tempo
tempo = datetime.now().timestamp()

# Ordenando os arquivos para iniciar com as curvas ABC
for x, file in enumerate(files):
  if file.endswith('.xlsx'):

    if file.find("ABC de Insumos") != -1 and not file.startswith("05 - "):
      continue
    
    if file.find("ABC de Serviços") != -1 and not file.startswith("06 - "):
      continue
    
    files.append(file)
    files.pop(x)

# Percorrendo todos os relatórios
for file in files:
  # print(file)
  if file.endswith('.xlsx'):

    if file.find("ABC de Insumos") != -1 and not file.startswith("05 - "):

      formatar_relatorio_abc_insumos(file, base_dados)

    if file.find("Resumido") != -1 and not file.startswith("02 - "):

      formatar_relatorio_resumido(file, base_dados)

    if file.find("Orçamento Sintético") != -1 and not file.startswith("03 - "):

      formatar_relatorio_sintetico(file, base_dados)

    if file.find("Orçamento Analítico") != -1 and not file.startswith("04 - "):

      formatar_relatorio_analitico(file, base_dados)

    if file.find("ABC de Serviços") != -1 and not file.startswith("06 - "):

      formatar_relatorio_abc_servicos(file, base_dados)

    if file.find("Composições Analíticas") != -1 and not file.startswith("07 - "):

      formatar_relatorio_analitico_somente_insumos(file, base_dados)

    if file.find("Mão de Obra") != -1 and not file.startswith("08 - "):

      formatar_relatorio_sintetico_mo_eq_mat(file, base_dados)

    if file.find("Preço Unitário") != -1 and not file.startswith("09 - "):

      formatar_relatorio_analitico_preco_unit(file, base_dados)

print(
  f'\n-> tempo decorrido: {round(datetime.now().timestamp() - tempo, 1)} seg\n')
