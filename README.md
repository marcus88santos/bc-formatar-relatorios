# Automação para Formatação de Relatórios em Excel

![Badge em Desenvolvimento](http://img.shields.io/static/v1?label=STATUS&message=EM%20DESENVOLVIMENTO&color=GREEN&style=for-the-badge)

<br />
Esta aplicação automatiza a formatação de relatórios em planilhas Excel gerados a partir de um orçamento. O objetivo é otimizar a visualização e organização das planilhas para facilitar a análise dos dados. Antes, cada ciclo de formatação durava cerca de **20 min**, agora pode ser realizado em **segundos**.

## :hammer: Funcionalidades

1. **Formatação dos relatórios**:
   - Exclusão de campos como assinaturas e valores irrelevantes.
   - Ajuste de larguras de colunas com base no conteúdo.
   - Fixação de cabeçalhos e congelamento de painéis.
   - Aplicação de filtros automáticos para facilitar a navegação nos dados.
2. **Customização de relatórios**:

   - Cada tipo de relatório (sintético, analítico, ABC de serviços, ABC de insumos, etc.) é tratado de forma específica, com ajustes personalizados conforme as características do arquivo.
   - Modificação de células mescladas com dados do banco de dados atualizado.

3. **Automatização de operações repetitivas**:
   - Verificação e ajuste de valores com zeros indesejados.
   - Salvar o relatório com nomeação automática baseada no tipo e no nome do empreendimento.

## :page_facing_up: Estrutura do Código

O código é dividido em várias funções para lidar com as diferentes etapas da formatação, como:

- **Funções de Formatação**:
  - `formatar_relatorio_analitico_preco_unit()`, `formatar_relatorio_sintetico_mo_eq_mat()`, `formatar_relatorio_abc_servicos()` e outras: responsáveis por formatar tipos específicos de relatórios.
- **Funções Utilitárias**:
  - `modifica_banco_dados()`: desmescla células e atualiza informações do banco de dados.
  - `salvar_relatorio()`: salva o relatório formatado com nome adequado.
  - `aumentar_largura_coluna()`: ajusta a largura das colunas com base no conteúdo.

## ✔️ Técnicas e tecnologias utilizadas

| [![My Skills](https://skillicons.dev/icons?i=py)]() |
| :-------------------------------------------------: |
|                       Python                        |

### Bibliotecas

- `openpyxl`: Utilizada para manipulação de planilhas Excel (leitura, edição e formatação).
- `datetime`: Utilizada para trabalhar com datas, como nomeação de arquivos e uso em relatórios.
- `locale`: Utilizada para configurar a localização da aplicação e formatar corretamente as datas e números para o padrão brasileiro.
- `dotenv`: Utilizada para as variáveis de ambiente

## 🛠️ Como Usar

1. **Pré-requisitos**:
   - Instalar as bibliotecas:
     ```
     pip install -r requirements.txt
     ```
   - Criar um arquivo chamado '.env' na raiz do projeto, definindo a variável 'FOLDER' como sendo o caminho, a partir da pasta 'Users' (ou 'Usuários'), para formatar os relatórios
     ```
     echo "FOLDER=\(...)\caminho-da-pasta-com-os-relatorios" > .env
     ```
2. **Execução**:

   - Execute o arquivo main.py
   - Defina o mês da database a ser impressa nos relatórios (de 1 a 12)
   - Defina o ano da database a ser impressa nos relatórios

3. **Resultado**:
   - O relatório formatado será salvo automaticamente no mesmo diretório, com o nome ajustado de acordo com as especificações.

## ✨ Observações

- As funções estão configuradas para trabalhar com relatórios específicos. É possível adaptar para outros tipos de relatórios ajustando as funções de acordo com a necessidade.
- O código atual remove linhas de grade e ajusta o layout para uma visualização mais clara dos dados.

## 🚶 Autor

| [<img loading="lazy" src="https://github.com/marcus88santos.png?size=115" width=115><br><sub>marcUs fiLLipe santos</sub>](https://github.com/marcus88santos) |
| :----------------------------------------------------------------------------------------------------------------------------------------------------------: |

<div>
<a href="https://www.linkedin.com/in/marcus88santos" target="_blank">
<img loading="lazy" src="https://img.shields.io/badge/-LinkedIn-%230077B5?style=for-the-badge&logo=linkedin&logoColor=white"></a>   
</div>
