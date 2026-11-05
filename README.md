# Automacao_Consultas

# 🤖 RPA de Sincronização de Filas (Web → Excel)

Um robô de automação desenvolvido em Python para sincronizar dados do portal público da ANVISA com planilhas locais de controle de prazos da equipe (GGTAB).

O projeto elimina o trabalho manual de verificação diária das filas de "Registro" (6001) e "Renovação" (6003). O script acessa a web, extrai a tabela atualizada de processos, compara com a base local e realiza as atualizações necessárias de forma autônoma.

## 🚀 Principais Funcionalidades

- **Web Scraping Dinâmico:** Utiliza o Selenium para navegar pelo portal da ANVISA, selecionar filtros em menus dropdown e extrair a tabela de processos pendentes.
- **Detecção de Novas Entradas:** Cruza os dados da web com a planilha atual e insere automaticamente novos processos, calculando o prazo final de análise (90 dias para Registro, 150 dias para Renovação).
- **Baixa Automática (Lógica de "Sumiço"):** Identifica processos que constam na planilha local, mas não estão mais na fila da web, e carimba automaticamente a data de saída (finalização).
- **Formatação Condicional Nativa:** Utiliza a biblioteca `openpyxl` para clonar estilos de células, inserir fórmulas (como cálculo de tempo e mediana) e pintar as datas de vermelho, amarelo ou verde dependendo da proximidade do vencimento.
- **Ordenação Automática:** Reorganiza toda a planilha cronologicamente sempre que novos dados são inseridos.

## 🛠️ Tecnologias Utilizadas

- **Python 3.8+**
- **[Selenium WebDriver](https://www.selenium.dev/):** Para automação do navegador e interação com elementos web.
- **[Pandas](https://pandas.pydata.org/):** Para cruzamento de dados e verificação de duplicidades.
- **[OpenPyXL](https://openpyxl.readthedocs.io/):** Para manipulação avançada de arquivos `.xlsx`, injetando cores e fórmulas diretamente no arquivo.

## 📦 Como Instalar e Rodar

1. Clone o repositório:
   ```bash
   git clone [https://github.com/seu-usuario/nome-do-repositorio.git](https://github.com/seu-usuario/nome-do-repositorio.git)
   cd nome-do-repositorio
   Crie um ambiente virtual e ative-o:
  python -m venv venv
  source venv/bin/activate  # Linux/Mac
  venv\Scripts\activate     # Windows
  pip install -r requirements.txt
  
  Pré-requisito do Selenium: Certifique-se de ter o Google Chrome instalado em sua máquina. O WebDriver fará a ponte automática com o navegador.
  
  Ajuste o caminho do arquivo Excel na variável CAMINHO_PLANILHA dentro do script.


  Execute a automação:
  
  python nome_do_script.py
