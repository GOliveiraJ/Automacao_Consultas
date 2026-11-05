from copy import copy
from datetime import datetime, date, timedelta
import re
import statistics
import time
import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
import pandas as pd

# Usando o Selenium Padrão 
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

# --- CONFIGURAÇÕES DE PRAZOS ---
PRAZO_RENOVACAO_DIAS = 150  # 6072
PRAZO_REGISTRO_DIAS = 90    # 6071

# --- CAMINHO DA PLANILHA ---
CAMINHO_PLANILHA = r"C:\Users\gabriel.jesus\ANVISA - Agencia Nacional de Vigilancia Sanitaria\GGTAB - Oficial - Fumigenos\2 - REGISTRO CCTAB\REGISTRO e RENOVAÇÃO_CONTROLE DAS PETIÇÕES\CONTROLE DOS PRAZOS DE FILAS_2026.xlsx"


def padronizar_dado(val):
    if pd.isna(val) or val is None: return ""
    return re.sub(r"[^\d./-]", "", str(val))


def converter_para_data(val):
    """Converte valores da célula (datetime, date ou string DD/MM/YYYY) em objeto date."""
    if val is None or str(val).strip() == "":
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    texto = str(val).strip()
    if texto.startswith("="):
        return None
    try:
        return datetime.strptime(texto[:10], "%d/%m/%Y").date()
    except:
        try:
            return datetime.strptime(texto[:10], "%Y-%m-%d").date()
        except:
            return None


def encontrar_primeira_linha_vazia(ws, linha_inicial=10):
    linha = linha_inicial
    while ws.cell(row=linha, column=3).value is not None and str(ws.cell(row=linha, column=3).strip() != ""):
        linha += 1
    return linha


def registrar_baixas_por_sumico(ws, df_web, linha_cabecalho=9):
    if df_web.empty or "Processo" not in df_web.columns:
        print("    [ALERTA] Lista da Anvisa vazia. Abortando baixas para evitar falso-positivo.")
        return 0
        
    df_web["Chave"] = df_web["Processo"].apply(padronizar_dado) + df_web["Expediente"].apply(padronizar_dado)
    chaves_na_web = set(df_web["Chave"].tolist())
    
    col_saida = None
    for c in range(1, ws.max_column + 1):
        val = str(ws.cell(row=linha_cabecalho, column=c).value).lower()
        if "saída" in val or "saida" in val: 
            col_saida = c
            break
            
    if not col_saida: return 0
        
    hoje = datetime.now().date()
    cont_baixas = 0
    
    for r in range(linha_cabecalho + 1, ws.max_row + 1):
        cel_proc = ws.cell(row=r, column=3)
        cel_exp = ws.cell(row=r, column=4)
        cel_saida = ws.cell(row=r, column=col_saida)
        
        if cel_proc.value and str(cel_proc.value).strip() != "":
            if cel_saida.value is None or str(cel_saida.value).strip() == "":
                chave_planilha = padronizar_dado(cel_proc.value) + padronizar_dado(cel_exp.value)
                
                if chave_planilha not in chaves_na_web:
                    cel_saida.value = hoje
                    cel_saida.number_format = 'dd/mm/yyyy'
                    cont_baixas += 1
                    print(f"    ✅ BAIXA AUTOMÁTICA: Exp {cel_exp.value} (Proc: {cel_proc.value}) sumiu da fila. Saída: {hoje.strftime('%d/%m/%Y')}")
                    
    return cont_baixas


def ordenar_planilha_por_data(ws, linha_cabecalho=9):
    max_row = ws.max_row
    max_col = ws.max_column
    linhas_dados = []
    
    for r in range(linha_cabecalho + 1, max_row + 1):
        if ws.cell(row=r, column=3).value: 
            linha = [ws.cell(row=r, column=c).value for c in range(1, max_col + 1)]
            linhas_dados.append(linha)
            
    def extrair_data(linha):
        dt = converter_para_data(linha[1])
        return dt if dt else date.max
            
    linhas_dados.sort(key=extrair_data)
    
    for i, linha in enumerate(linhas_dados):
        r = linha_cabecalho + 1 + i
        for c_idx, val in enumerate(linha):
            if isinstance(val, datetime): val = val.date()
            ws.cell(row=r, column=c_idx + 1).value = val


def selecionar_opcao_segura(driver, index_select, texto_procurado, timeout=20):
    tempo_inicio = time.time()
    while time.time() - tempo_inicio < timeout:
        try:
            dropdowns = driver.find_elements(By.TAG_NAME, "select")
            if len(dropdowns) > index_select:
                select_obj = Select(dropdowns[index_select])
                for opcao in select_obj.options:
                    if texto_procurado.lower().strip() in opcao.text.lower().strip():
                        select_obj.select_by_visible_text(opcao.text)
                        return True
        except: pass
        time.sleep(1)
    raise Exception(f"A opção '{texto_procurado}' não carregou.")


def clicar_botao_por_texto(driver, texto_botao):
    botoes = driver.find_elements(By.TAG_NAME, "button")
    for botao in botoes:
        if texto_botao.lower() in botao.text.lower().strip():
            try: driver.execute_script("arguments[0].click();", botao)
            except: pass
            return True
    return False


def extrair_tabela_blindada(driver):
    dados = []
    try:
        linhas = driver.find_elements(By.XPATH, "//table//tr | //tbody//tr")
        for linha in linhas:
            texto_linha = linha.text.strip()
            if not texto_linha:
                continue
                
            match_proc = re.search(r"\d{5}\.\d+/\d{4}-\d+", texto_linha)
            processo = match_proc.group(0) if match_proc else ""
            
            match_exp = re.search(r"\d{6,7}/\d{2}-\d+", texto_linha)
            expediente = match_exp.group(0) if match_exp else ""
            
            match_data = re.search(r"\d{2}/\d{2}/\d{4}", texto_linha)
            data_ent = match_data.group(0) if match_data else ""
            
            if processo or expediente:
                dados.append({"Data de entrada": data_ent, "Processo": processo, "Expediente": expediente})
    except Exception as e: print(f"    [Aviso na varredura] {e}")
    return pd.DataFrame(dados)


def clonar_formatacao(ws, linha_inicio, qtd_linhas_novas):
    if linha_inicio <= 1: return
    linha_modelo = linha_inicio - 1
    for r in range(linha_inicio, linha_inicio + qtd_linhas_novas):
        for c in range(1, ws.max_column + 1):
            cel_modelo = ws.cell(row=linha_modelo, column=c)
            cel_nova = ws.cell(row=r, column=c)
            if cel_modelo.has_style:
                cel_nova.border, cel_nova.fill = copy(cel_modelo.border), copy(cel_modelo.fill)
                cel_nova.number_format, cel_nova.alignment = copy(cel_modelo.number_format), copy(cel_modelo.alignment)


def atualizar_indicadores_e_cores(ws, nome_aba, linha_cabecalho=9):
    col_prazo, col_saida, col_entrada, col_tempo, col_mediana = None, None, None, None, None
    for c in range(1, ws.max_column + 1):
        val = str(ws.cell(row=linha_cabecalho, column=c).value).lower()
        if "prazo" in val: col_prazo = c
        elif "saída" in val or "saida" in val: col_saida = c
        elif "entrada" in val: col_entrada = c
        elif "tempo" in val: col_tempo = c
        elif "mediana" in val: col_mediana = c
            
    if not col_prazo: return
        
    hoje = datetime.now().date()
    dias_prazo_legal = PRAZO_REGISTRO_DIAS if nome_aba == "6071" else PRAZO_RENOVACAO_DIAS
    letra_mediana = get_column_letter(col_mediana) if col_mediana else 'H'
    
    meses = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez']
    ws['B2'] = f"{hoje.day:02d}/{meses[hoje.month - 1]}"
    
    tempos_calculados = []

    for r in range(linha_cabecalho + 1, ws.max_row + 1):
        # Limpa resíduos na coluna de mediana nas linhas abaixo da linha 10
        if col_mediana and r > linha_cabecalho + 1:
            ws.cell(row=r, column=col_mediana).value = None

        cel_proc = ws.cell(row=r, column=3) 
        if cel_proc.value is None or str(cel_proc.value).strip() == "":
            continue

        cel_entrada = ws.cell(row=r, column=col_entrada) if col_entrada else ws.cell(row=r, column=2)
        cel_prazo = ws.cell(row=r, column=col_prazo)
        cel_saida = ws.cell(row=r, column=col_saida) if col_saida else None
        cel_tempo = ws.cell(row=r, column=col_tempo) if col_tempo else None
        
        dt_entrada = converter_para_data(cel_entrada.value)
        dt_saida = converter_para_data(cel_saida.value) if cel_saida else None
        dt_prazo = converter_para_data(cel_prazo.value)

        # 1. Garante que o Prazo Legal seja sempre uma data estática (substitui fórmulas antigas =B10+150)
        if dt_entrada and (dt_prazo is None or str(cel_prazo.value).strip().startswith("=")):
            dt_prazo = dt_entrada + timedelta(days=dias_prazo_legal)
            cel_prazo.value = dt_prazo
            cel_prazo.number_format = 'dd/mm/yyyy'

        # 2. Calcula diretamente o Tempo de Fila como número inteiro (sem fórmula do Excel)
        if cel_tempo:
            if dt_entrada and dt_saida:
                dias_fila = (dt_saida - dt_entrada).days
                cel_tempo.value = int(dias_fila)
                cel_tempo.number_format = '0'
                tempos_calculados.append(dias_fila)
            else:
                cel_tempo.value = None

        # 3. Aplica a coloração de alerta no Prazo Legal
        if dt_prazo is not None:
            cor_prazo = "000000" 
            
            if dt_saida is not None:
                cor_prazo = "A6A6A6"
            else:
                dias_restantes = (dt_prazo - hoje).days
                
                if dias_restantes < 0: cor_prazo = "FF0000" 
                elif 0 <= dias_restantes < 7: cor_prazo = "FFC000" 
                elif 7 <= dias_restantes <= 15: cor_prazo = "A52A2A" 
                elif 15 < dias_restantes <= 30: cor_prazo = "7030A0" 
                elif 30 < dias_restantes <= 60: cor_prazo = "00B050" 
                    
            f_prazo = cel_prazo.font
            if f_prazo: cel_prazo.font = Font(name=f_prazo.name, size=f_prazo.size, bold=f_prazo.bold, color=cor_prazo)
            else: cel_prazo.font = Font(color=cor_prazo)

    # 4. Calcula a Mediana diretamente no Python e grava o valor pronto na linha 10 (H10)
    if col_mediana:
        cel_mediana = ws.cell(row=linha_cabecalho + 1, column=col_mediana)
        if tempos_calculados:
            med_val = statistics.median(tempos_calculados)
            # Se for inteiro exato (ex: 79.0), grava como int(79); caso contrário, arredonda para 1 casa
            cel_mediana.value = int(med_val) if med_val == int(med_val) else round(med_val, 1)
        else:
            cel_mediana.value = None

        limite_meta = 70 if nome_aba == "6071" else 120
        cel_regra = f"{letra_mediana}{linha_cabecalho + 1}"
        
        # Aplica a cor diretamente na fonte da célula + mantém a regra condicional
        cor_med = "FF0000" if (cel_mediana.value is not None and cel_mediana.value > limite_meta) else "00B050"
        cel_mediana.font = Font(color=cor_med, bold=True)
        ws.conditional_formatting.add(cel_regra, CellIsRule(operator='greaterThan', formula=[str(limite_meta)], font=Font(color="FF0000", bold=True)))
        ws.conditional_formatting.add(cel_regra, CellIsRule(operator='lessThanOrEqual', formula=[str(limite_meta)], font=Font(color="00B050", bold=True)))


def consultar_fila_anvisa(driver, nome_fila, nome_subfila):
    print(f"\n=======================================================\n -> Configurando filtros para a fila: {nome_fila}...\n=======================================================")
    try:
        selecionar_opcao_segura(driver, index_select=0, texto_procurado="Derivados do Tabaco", timeout=20)
        time.sleep(2)
        selecionar_opcao_segura(driver, index_select=1, texto_procurado=nome_fila, timeout=20)
        time.sleep(2)
        selecionar_opcao_segura(driver, index_select=2, texto_procurado=nome_subfila, timeout=20)
        time.sleep(2)

        print("\n" + "*" * 60)
        print(f"👉 ATENÇÃO: Vá ao Chrome agora e clique no botão\n👉 azul 'PESQUISAR' para gerar a tabela de {nome_fila}!")
        print("*" * 60)
        input(" -> Após clicar em PESQUISAR e a tabela aparecer, PRIMA [ENTER] AQUI... ")

        print(f"\n -> A aguardar 20 segundos para a tabela {nome_fila} carregar totalmente...")
        for i in range(20, 0, -5):
            print(f"    ... {i} segundos restantes ...")
            time.sleep(5)

        print(f" -> A ler os dados da tabela de {nome_fila}...")
        df_resultado = extrair_tabela_blindada(driver)
        print(f" -> Sucesso! Extraídos {len(df_resultado)} registos da web.")

        clicar_botao_por_texto(driver, "Limpar")
        time.sleep(3)
        return df_resultado
    except Exception as e:
        print(f" [Aviso] Erro: {e}")
        return pd.DataFrame()


def extrair_dados_anvisa():
    print(">>> 1. A ligar ao Chrome validado...")
    
    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    
    try:
        driver = webdriver.Chrome(options=options)
    except Exception as e:
        print("\n[ERRO CRÍTICO] O robô não encontrou o Chrome aberto na porta 9222!")
        raise e

    try:
        print("\n" + "!" * 70)
        print("🚨 VERIFICAÇÃO MANUAL DO CLOUDFLARE 🚨")
        print("Antes de prosseguir, certifique-se de que JÁ ACEDEU a")
        print("https://consultas.anvisa.gov.br/ no Chrome aberto e")
        print("JÁ PASSOU pelo CAPTCHA com sucesso.")
        print("!" * 70)
        input(" -> SE JÁ ESTÁ NA PÁGINA INICIAL DA ANVISA, PRIMA [ENTER] AQUI... ")
        
        try:
            elementos = driver.find_elements(By.XPATH, "//*[contains(text(), 'Fila de Análise') or contains(text(), 'Fila de análise')]")
            for elem in elementos:
                if elem.is_displayed():
                    elem.click()
                    break
        except: 
            driver.get("https://consultas.anvisa.gov.br/#/fila/")
        
        time.sleep(10)
        
        df_registro_web = consultar_fila_anvisa(driver, nome_fila="REGISTRO", nome_subfila="REGISTRO")
        df_renovacao_web = consultar_fila_anvisa(driver, nome_fila="RENOVAÇÃO", nome_subfila="RENOVAÇÃO")
        return df_registro_web, df_renovacao_web
    finally:
        print("\n -> Extração concluída. A desligar o robô do navegador.")


def executar_automacao_oficial():
    print("\n>>> 2. A ler as planilhas para verificar os processos...")
    df_reg_atual = pd.read_excel(CAMINHO_PLANILHA, sheet_name="6071", header=8)
    df_renov_atual = pd.read_excel(CAMINHO_PLANILHA, sheet_name="6072", header=8)

    df_reg_web, df_renov_web = extrair_dados_anvisa()
    if df_reg_web.empty and df_renov_web.empty:
        print("\n⚠️ Nenhum dado extraído. A automação foi encerrada sem alterar a planilha.")
        return

    df_reg_novos = pd.DataFrame()
    df_renov_novos = pd.DataFrame()

    # --- REGISTRO ---
    if not df_reg_web.empty and "Processo" in df_reg_web.columns:
        df_reg_atual["Chave"] = df_reg_atual["Processo"].apply(padronizar_dado) + df_reg_atual["Expediente"].apply(padronizar_dado)
        chaves_existentes_reg = set(df_reg_atual["Chave"].tolist())
        
        df_reg_web["Chave"] = df_reg_web["Processo"].apply(padronizar_dado) + df_reg_web["Expediente"].apply(padronizar_dado)
        df_reg_novos = df_reg_web[~df_reg_web["Chave"].isin(chaves_existentes_reg)].copy()
        
        if not df_reg_novos.empty:
            df_reg_novos = df_reg_novos.drop_duplicates(subset=["Chave"], keep="first")
            df_reg_novos["Data de entrada"] = pd.to_datetime(df_reg_novos["Data de entrada"], dayfirst=True, errors="coerce")
            df_reg_novos["Prazo de análise "] = df_reg_novos["Data de entrada"] + pd.Timedelta(days=PRAZO_REGISTRO_DIAS)
            df_reg_novos["Data de entrada"] = df_reg_novos["Data de entrada"].dt.date
            df_reg_novos["Prazo de análise "] = df_reg_novos["Prazo de análise "].dt.date

    # --- RENOVAÇÃO ---
    if not df_renov_web.empty and "Processo" in df_renov_web.columns:
        df_renov_atual["Chave"] = df_renov_atual["Processo"].apply(padronizar_dado) + df_renov_atual["Expediente"].apply(padronizar_dado)
        chaves_existentes_renov = set(df_renov_atual["Chave"].tolist())
        
        df_renov_web["Chave"] = df_renov_web["Processo"].apply(padronizar_dado) + df_renov_web["Expediente"].apply(padronizar_dado)
        df_renov_novos = df_renov_web[~df_renov_web["Chave"].isin(chaves_existentes_renov)].copy()
        
        if not df_renov_novos.empty:
            df_renov_novos = df_renov_novos.drop_duplicates(subset=["Chave"], keep="first")
            df_renov_novos["Data de entrada"] = pd.to_datetime(df_renov_novos["Data de entrada"], dayfirst=True, errors="coerce")
            df_renov_novos["Prazo de análise "] = df_renov_novos["Data de entrada"] + pd.Timedelta(days=PRAZO_RENOVACAO_DIAS)
            df_renov_novos["Data de entrada"] = df_renov_novos["Data de entrada"].dt.date
            df_renov_novos["Prazo de análise "] = df_renov_novos["Prazo de análise "].dt.date

    print("\n" + "=" * 70 + "\n >>> RELATÓRIO FINAL DE ATUALIZAÇÃO DA PLANILHA 2026 <<<\n" + "=" * 70)
    wb = openpyxl.load_workbook(CAMINHO_PLANILHA)

    # --- ABA REGISTRO (6071) ---
    ws_reg = wb["6071"]
    print("\n🔍 Analisando baixas por 'Sumiço' no REGISTRO (6071)...")
    baixas_reg = registrar_baixas_por_sumico(ws_reg, df_reg_web, linha_cabecalho=9)
    print(f" -> {baixas_reg} processo(s) finalizado(s) no REGISTRO.")

    if not df_reg_novos.empty:
        print(f"\n🟢 {len(df_reg_novos)} novo(s) processo(s) de REGISTRO adicionado(s):")
        linha_alvo = encontrar_primeira_linha_vazia(ws_reg, linha_inicial=10)
        linha_inicio_formato = linha_alvo
        
        for _, row in df_reg_novos.iterrows():
            ws_reg.cell(row=linha_alvo, column=2, value=row["Data de entrada"])
            ws_reg.cell(row=linha_alvo, column=3, value=row["Processo"])
            ws_reg.cell(row=linha_alvo, column=4, value=row["Expediente"])
            ws_reg.cell(row=linha_alvo, column=5, value=row["Prazo de análise "])
            print(f"   -> Processo: {row['Processo']} | Exp: {row['Expediente']} | Data Entrada: {row['Data de entrada'].strftime('%d/%m/%Y')}")
            linha_alvo += 1
            
        clonar_formatacao(ws_reg, linha_inicio_formato, len(df_reg_novos))
    else:
        print("\n🔵 REGISTRO (6071): Nenhum processo novo de entrada.")
        
    ordenar_planilha_por_data(ws_reg, linha_cabecalho=9)

    # --- ABA RENOVAÇÃO (6072) ---
    ws_renov = wb["6072"]
    print("\n🔍 Analisando baixas por 'Sumiço' na RENOVAÇÃO (6072)...")
    baixas_renov = registrar_baixas_por_sumico(ws_renov, df_renov_web, linha_cabecalho=9)
    print(f" -> {baixas_renov} processo(s) finalizado(s) na RENOVAÇÃO.")

    if not df_renov_novos.empty:
        print(f"\n🟢 {len(df_renov_novos)} novo(s) processo(s) de RENOVAÇÃO adicionado(s):")
        linha_alvo = encontrar_primeira_linha_vazia(ws_renov, linha_inicial=10)
        linha_inicio_formato = linha_alvo
        
        for _, row in df_renov_novos.iterrows():
            ws_renov.cell(row=linha_alvo, column=2, value=row["Data de entrada"])
            ws_renov.cell(row=linha_alvo, column=3, value=row["Processo"])
            ws_renov.cell(row=linha_alvo, column=4, value=row["Expediente"])
            ws_renov.cell(row=linha_alvo, column=5, value=row["Prazo de análise "])
            print(f"   -> Processo: {row['Processo']} | Exp: {row['Expediente']} | Data Entrada: {row['Data de entrada'].strftime('%d/%m/%Y')}")
            linha_alvo += 1
            
        clonar_formatacao(ws_renov, linha_inicio_formato, len(df_renov_novos))
    else:
        print("\n🔵 RENOVAÇÃO (6072): Nenhum processo novo de entrada.")
        
    ordenar_planilha_por_data(ws_renov, linha_cabecalho=9)

    print("\n    [Calculando Indicadores, Mediana, Ordenação e Cores...]")
    atualizar_indicadores_e_cores(ws_reg, nome_aba="6071", linha_cabecalho=9)
    atualizar_indicadores_e_cores(ws_renov, nome_aba="6072", linha_cabecalho=9)

    wb.save(CAMINHO_PLANILHA)
    wb.close()
    print("\n💾 Planilha oficial guardada com valores estáticos! Fluxo completo finalizado.")
    print("=" * 70)

if __name__ == "__main__":
    executar_automacao_oficial()