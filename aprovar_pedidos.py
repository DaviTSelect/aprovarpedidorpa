import os
import sys
import pandas as pd
import zipfile
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import (
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.chrome.options import Options
import time
import tkinter as tk
import queue
import threading
import tempfile
from pathlib import Path
from tkinter import filedialog, messagebox, PhotoImage, ttk
from openpyxl import load_workbook
from copy import copy

try:
    from version import __version__
except ImportError:
    __version__ = "0.0.0"

try:
    import auto_update
except ImportError:
    auto_update = None  # type: ignore

# ---------------------------------------------------------------------------
# Caminhos – ajuste conforme a instalação local
# ---------------------------------------------------------------------------
PASTA_USUARIO = Path.home()
CHROME_BINARY = PASTA_USUARIO / "Chrome" / "GoogleChrome" / "App" / "Chrome-bin" / "chrome.exe"
CHROMEDRIVER_PATH = PASTA_USUARIO / "Chrome" / "chromedriver.exe"
PASTA_PROJETO = Path(__file__).resolve().parent
CONTEXTO_TRABALHADOR = threading.local()


def _caminho_recurso(nome: str) -> Path:
    """Retorna o caminho de um recurso (ícone, etc.) dentro do .exe ou do diretório do script."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / nome
    return PASTA_PROJETO / nome
LOADER_LOCATOR = (By.ID, "imageLoader")
NUM_NAVEGADORES = 3
SALVAR_SCREENSHOTS = False  # True para debug, False para produção
HEADLESS_CHROME = True      # False para ver as janelas do Chrome

# ---------------------------------------------------------------------------
# Utilitários de interação (idênticos ao script original)
# ---------------------------------------------------------------------------

def aguardar_loader_sumir(navegador, timeout=30):
    WebDriverWait(navegador, timeout, poll_frequency=0.1).until(
        EC.invisibility_of_element_located(LOADER_LOCATOR)
    )


def interagir_com_loader(navegador, acao, timeout=30):
    aguardar_loader_sumir(navegador, timeout=timeout)
    resultado = acao()
    aguardar_loader_sumir(navegador, timeout=timeout)
    return resultado


def clicar_elemento(navegador, elemento):
    return interagir_com_loader(navegador, elemento.click)


def limpar_elemento(navegador, elemento):
    return interagir_com_loader(navegador, elemento.clear)


def digitar_elemento(navegador, elemento, *teclas):
    return interagir_com_loader(navegador, lambda: elemento.send_keys(*teclas))


def executar_script_interacao(navegador, script, *argumentos):
    return interagir_com_loader(
        navegador,
        lambda: navegador.execute_script(script, *argumentos),
    )


def aguardar_pagina_carregar(navegador, timeout=100):
    WebDriverWait(navegador, timeout).until(
        lambda driver: driver.execute_script("return document.readyState") == "complete"
    )
    aguardar_loader_sumir(navegador, timeout=timeout)


def criar_opcoes_chrome():
    options = Options()
    options.add_argument("--log-level=3")
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    if HEADLESS_CHROME:
        options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-infobars")
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--allow-running-insecure-content")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument("--start-maximized")
    options.add_argument("--window-size=1920,1080")
    perfil = getattr(CONTEXTO_TRABALHADOR, "perfil", None)
    if perfil is not None:
        options.add_argument(f"--user-data-dir={perfil}")
    if CHROME_BINARY.is_file():
        options.binary_location = str(CHROME_BINARY)
    return options


def _caminho_env():
    """Retorna o diretório onde o .env deve estar (ao lado do .exe ou do script)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent


def realizar_login():
    from conveniar import iniciar_sessao

    usuario = os.environ.get("USUARIO", "")
    senha   = os.environ.get("SENHA", "")

    env_path = _caminho_env() / ".env"
    if env_path.is_file() and (not usuario or not senha):
        try:
            from dotenv import load_dotenv
            load_dotenv(env_path, override=False)
            usuario = os.environ.get("USUARIO", "")
            senha   = os.environ.get("SENHA", "")
        except ImportError:
            pass

    if not usuario or not senha:
        raise RuntimeError(
            f"Credenciais não encontradas.\n"
            f"Crie o arquivo  {env_path}  com:\n"
            f"  USUARIO=seu.login\n"
            f"  SENHA=SuaSenha"
        )

    return iniciar_sessao(usuario, senha)


def confirmar(mensagem):
    """Exibe mensagem no console e aguarda 's' para continuar ou 'n' para pular."""
    while True:
        resposta = input(f"\n>>> {mensagem} [s/n]: ").strip().lower()
        if resposta in ("s", "n"):
            return resposta == "s"
        print("    Digite s para continuar ou n para cancelar este pedido.")


# ---------------------------------------------------------------------------
# Gravação de resultado na planilha
# ---------------------------------------------------------------------------

_lock_planilha = threading.Lock()

_CONCLUIDOS = ("Aprovado",)


def _copiar_linha_para_aba(wb, ws1, linha, col_res, resultado, nome_aba):
    """Copia uma linha de ws1 para a aba nome_aba (criando-a se necessário) e remove de ws1."""
    if nome_aba not in wb.sheetnames:
        ws_dest = wb.create_sheet(nome_aba)
        for c in range(1, ws1.max_column + 1):
            ws_dest.cell(1, c, ws1.cell(1, c).value)
    else:
        ws_dest = wb[nome_aba]
    proxima = ws_dest.max_row + 1
    for c in range(1, ws1.max_column + 1):
        ws_dest.cell(proxima, c, ws1.cell(linha, c).value)
    ws_dest.cell(proxima, col_res, resultado)
    ws1.delete_rows(linha)


def registrar_e_mover_se_concluido(caminho_arquivo, numero_pedido, resultado):
    """Registra resultado na planilha.
    - Aprovado / Pedido não encontrado → move para aba 'Processados'.
    - Qualquer erro                    → move para aba 'Erros'.
    Tenta até 10 vezes com backoff caso o arquivo esteja temporariamente bloqueado.
    """
    tentativas = 10
    for tentativa in range(1, tentativas + 1):
        try:
            with _lock_planilha:
                # Valida que é um zip válido antes de abrir
                with open(caminho_arquivo, "rb") as f_check:
                    if f_check.read(4) != b"PK\x03\x04":
                        raise zipfile.BadZipFile("Arquivo não é um xlsx válido (não começa com PK)")
                wb = load_workbook(caminho_arquivo)
                ws1 = wb.active
                _registrar_interno(wb, ws1, caminho_arquivo, numero_pedido, resultado)
            return
        except zipfile.BadZipFile as exc:
            msg = (
                f"⚠️  Pedido {numero_pedido} – tentativa {tentativa}/{tentativas}: "
                f"arquivo corrompido ou aberto no Excel ({exc})"
            )
            print(msg)
            if tentativa == tentativas:
                print(f"❌  Desistindo de salvar pedido {numero_pedido}. Feche o arquivo no Excel e tente novamente.")
                return
            time.sleep(2 * tentativa)
        except Exception as exc:
            if tentativa == tentativas:
                print(f"⚠️  Não foi possível salvar pedido {numero_pedido} após {tentativas} tentativas: {exc}")
                return
            time.sleep(1 * tentativa)


def _registrar_interno(wb, ws1, caminho_arquivo, numero_pedido, resultado):
    cabecalhos = {
        str(ws1.cell(1, c).value or "").strip(): c
        for c in range(1, ws1.max_column + 1)
    }
    col_pedido = cabecalhos.get("Nº Pedido", 1)
    col_res = cabecalhos.get("Resultado Aprovação")
    if col_res is None:
        col_res = ws1.max_column + 1
        ws1.cell(1, col_res, "Resultado Aprovação")

    linha_encontrada = None
    for row in range(2, ws1.max_row + 1):
        val = str(ws1.cell(row, col_pedido).value or "").strip()
        if val == str(numero_pedido).strip():
            linha_encontrada = row
            break

    if linha_encontrada is None:
        wb.close()
        return

    if resultado in _CONCLUIDOS or resultado == "Pedido não encontrado":
        _copiar_linha_para_aba(wb, ws1, linha_encontrada, col_res, resultado, "Processados")
    else:
        _copiar_linha_para_aba(wb, ws1, linha_encontrada, col_res, resultado, "Erros")

    wb.save(caminho_arquivo)
    wb.close()


def consolidar_planilha(caminho_arquivo, registros_status):
    """Cria a aba 'status_aprov' com o resultado de cada pedido processado."""
    origem = Path(caminho_arquivo)
    saida = origem.with_name(origem.stem + "_aprovados.xlsx")
    wb = load_workbook(origem)

    # Remove a aba se já existir (re-execução)
    if "status_aprov" in wb.sheetnames:
        del wb["status_aprov"]

    ws_status = wb.create_sheet("status_aprov")
    cabecalhos = ["Nº Pedido", "Item PT", "Favorecido", "Resultado Aprovação"]
    for col, titulo in enumerate(cabecalhos, 1):
        ws_status.cell(1, col, titulo)

    for linha_num, registro in enumerate(registros_status, 2):
        for col, valor in enumerate(registro, 1):
            ws_status.cell(linha_num, col, valor)

    # Ajusta largura das colunas
    for col in ws_status.columns:
        largura = max((len(str(c.value or "")) for c in col), default=10) + 4
        ws_status.column_dimensions[col[0].column_letter].width = min(largura, 60)

    wb.save(saida)
    wb.close()
    return saida


# ---------------------------------------------------------------------------
# Função principal de aprovação
# ---------------------------------------------------------------------------

def funcao_aprovar(navegador, pedido, numero, rubrica, data_pagamento, competencia, data_nascimento=""):
    """Busca o pedido pelo número e o aprova usando apenas a rubrica da planilha."""
    wait = WebDriverWait(navegador, 20)

    # --- PESQUISA ---
    cb_bolsas = "/html/body/div[3]/form/div[5]/div/div/div[2]/div/div[1]/fieldset/div/div/ul/li[2]/table/tbody/tr[1]/td[1]/span/input"
    checkbox = wait.until(EC.presence_of_element_located((By.XPATH, cb_bolsas)))
    executar_script_interacao(
        navegador, "arguments[0].scrollIntoView({behavior:'smooth',block:'center'});", checkbox
    )
    if not checkbox.is_selected():
        clicar_elemento(navegador, checkbox)

    cb_pf = "/html/body/div[3]/form/div[5]/div/div/div[2]/div/div[1]/fieldset/div/div/ul/li[2]/table/tbody/tr[3]/td[1]/span/input"
    checkbox_pf = wait.until(EC.presence_of_element_located((By.XPATH, cb_pf)))
    if not checkbox_pf.is_selected():
        clicar_elemento(navegador, checkbox_pf)

    cb_diaria = "/html/body/div[3]/form/div[5]/div/div/div[2]/div/div[1]/fieldset/div/div/ul/li[2]/table/tbody/tr[2]/td[2]/span/input"
    checkbox_diaria = wait.until(EC.presence_of_element_located((By.XPATH, cb_diaria)))
    if not checkbox_diaria.is_selected():
        clicar_elemento(navegador, checkbox_diaria)

    # Número do pedido
    campo_pedido = wait.until(EC.presence_of_element_located((
        By.XPATH,
        "/html/body/div[3]/form/div[5]/div/div/div[2]/div/div[1]/fieldset/div/div/ul/li[3]/input",
    )))
    limpar_elemento(navegador, campo_pedido)
    digitar_elemento(navegador, campo_pedido, pedido)

    # Limpa campos de data
    limpar_elemento(navegador, wait.until(EC.presence_of_element_located((
        By.XPATH,
        "/html/body/div[3]/form/div[5]/div/div/div[2]/div/div[1]/fieldset/div/div/ul/li[8]/input[1]",
    ))))
    for xpath_data in (
        "/html/body/div[3]/form/div[5]/div/div/div[2]/div/div[1]/fieldset/div/div/ul/li[9]/input[1]",
        "/html/body/div[3]/form/div[5]/div/div/div[2]/div/div[1]/fieldset/div/div/ul/li[9]/input[2]",
    ):
        el = wait.until(EC.presence_of_element_located((By.XPATH, xpath_data)))
        executar_script_interacao(navegador, "arguments[0].value = '';", el)

    # Status "Enviado" (16)
    select_element = wait.until(EC.presence_of_element_located((
        By.XPATH,
        "/html/body/div[3]/form/div[5]/div/div/div[2]/div/div[1]/fieldset/div/div/ul/li[14]/select",
    )))
    sel = Select(select_element)
    interagir_com_loader(navegador, lambda: sel.select_by_value("16"))

    # Executar pesquisa
    btn_consultar = wait.until(EC.element_to_be_clickable((
        By.ID,
        "ctl00_ContentPlaceHolder1_ConsultaPedidoPagamentoUserControl2_btnConsultar",
    )))
    executar_script_interacao(navegador, "arguments[0].click();", btn_consultar)
    aguardar_loader_sumir(navegador)

    # Aguarda os links de edição ou ausência de registros
    WebDriverWait(navegador, 15).until(
        lambda driver: (
            driver.find_elements(By.XPATH, "//a[@title='Editar registro']")
            or driver.find_elements(
                By.XPATH,
                "//td[normalize-space()='Nenhum registro encontrado.']",
            )
        )
    )

    edit_links = navegador.find_elements(By.XPATH, "//a[@title='Editar registro']")

    if not edit_links:
        print(f"❌ Pedido {pedido} não encontrado na busca")
        return "Pedido não encontrado"

    print(f"\n{'='*60}")
    print(f"PEDIDO: {pedido}  |  RUBRICA: {rubrica}")
    print(f"{'='*60}")
    print(f"   {len(edit_links)} resultado(s) encontrado(s) na busca.")

    aba_consulta = navegador.current_window_handle

    for link in edit_links:
        try:
            executar_script_interacao(
                navegador,
                "arguments[0].scrollIntoView({behavior:'smooth',block:'center'});",
                link,
            )

            # Abre em nova aba
            abas_antes = set(navegador.window_handles)
            executar_script_interacao(navegador, "arguments[0].click();", link)
            novas_abas = WebDriverWait(navegador, 25).until(
                lambda driver: list(set(driver.window_handles) - abas_antes) or False
            )
            navegador.switch_to.window(novas_abas[0])
            aguardar_loader_sumir(navegador)
            print(f"   Aba do pedido aberta. URL: {navegador.current_url}")

            # -----------------------------------------------------------------
            # Helpers locais para esta aba
            # -----------------------------------------------------------------
            pasta_screenshots = PASTA_PROJETO / "screenshots"
            pasta_screenshots.mkdir(exist_ok=True)
            _step_num = [0]

            def ss(nome_etapa):
                _step_num[0] += 1
                if not SALVAR_SCREENSHOTS:
                    return
                caminho_ss = pasta_screenshots / f"{pedido.replace('/', '-')}_{_step_num[0]:02d}_{nome_etapa}.png"
                try:
                    navegador.save_screenshot(str(caminho_ss))
                except Exception:
                    pass

            def js(script, *args):
                return navegador.execute_script(script, *args)

            def scroll_js_clicar(elem):
                js("arguments[0].scrollIntoView({block:'center'});", elem)
                js("arguments[0].click();", elem)
                aguardar_loader_sumir(navegador)

            def eid(id_elem, timeout=25):
                return WebDriverWait(navegador, timeout, poll_frequency=0.1).until(
                    EC.presence_of_element_located((By.ID, id_elem))
                )

            def rolar(elem):
                """Rola até o elemento via JS antes de qualquer interação."""
                js("arguments[0].scrollIntoView({block:'center'});", elem)

            def ok_dialog(timeout=15):
                """Clica OK no diálogo jQuery UI e aguarda loader."""
                btn_ok = WebDriverWait(navegador, timeout).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, ".ui-dialog-buttonset button"))
                )
                js("arguments[0].scrollIntoView({block:'center'});", btn_ok)
                js("arguments[0].click();", btn_ok)
                aguardar_loader_sumir(navegador)

            def fechar_dialog(timeout=15):
                """Fecha o diálogo jQuery UI via botão X e aguarda loader."""
                btn_x = WebDriverWait(navegador, timeout).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, ".ui-dialog-titlebar-close"))
                )
                js("arguments[0].click();", btn_x)
                aguardar_loader_sumir(navegador)

            def ok_dialog_com_sim(timeout_sim=4):
                """Clica 'sim' se aparecer diálogo de confirmação antes do OK de sucesso."""
                aguardar_loader_sumir(navegador)
                try:
                    btn_sim = WebDriverWait(navegador, timeout_sim).until(
                        EC.element_to_be_clickable((
                            By.XPATH,
                            "//div[contains(@class,'ui-dialog-buttonset')]"
                            "//button/span[normalize-space()='sim']/..",
                        ))
                    )
                    js("arguments[0].scrollIntoView({block:'center'});", btn_sim)
                    js("arguments[0].click();", btn_sim)
                    aguardar_loader_sumir(navegador)
                    print("      (diálogo 'sim/não' confirmado com sim)")
                except TimeoutException:
                    pass
                ok_dialog()

            # -----------------------------------------------------------------
            # Passo 2b – Atualizar Dados Cadastrais + Data de Nascimento
            # -----------------------------------------------------------------
            print("   [2b] Clicando em Atualizar Dados Cadastrais...")
            btn_atualizar = eid("ctl00_ContentPlaceHolder1_PagamentoPessoaFisicaUserControl_btnAtualizar")
            scroll_js_clicar(btn_atualizar)
            ss("atualizar_dados")

            if data_nascimento:
                campo_nasc = eid("ctl00_ContentPlaceHolder1_PessoaFisicaEnderecoPersonalizadoUserControl1_txtDataNascimento")
                rolar(campo_nasc)
                valor_atual = (campo_nasc.get_attribute("value") or "").strip()
                if not valor_atual:
                    print(f"   [2b] Campo vazio no site — preenchendo: {data_nascimento}")
                    js("arguments[0].value = arguments[1];", campo_nasc, data_nascimento)
                    js("arguments[0].dispatchEvent(new Event('blur'));", campo_nasc)
                    aguardar_loader_sumir(navegador)
                    ss("data_nascimento_preenchida")
                else:
                    print(f"   [2b] Data de nascimento já preenchida no site ({valor_atual}) — ignorando")

            print("   [2b] Clicando em OK (dados cadastrais)...")
            btn_ok_nasc = eid("ctl00_ContentPlaceHolder1_PessoaFisicaEnderecoPersonalizadoUserControl1_btnOK")
            scroll_js_clicar(btn_ok_nasc)
            ss("ok_dados_cadastrais")

            # -----------------------------------------------------------------
            # Passo 3 – Procurar conta bancária
            # -----------------------------------------------------------------
            print("   [3] Clicando em Procurar conta bancária...")
            btn_proc = eid("ctl00_ContentPlaceHolder1_PagamentoPessoaFisicaUserControl_btnConsultarConta")
            scroll_js_clicar(btn_proc)
            ss("procurar_conta")

            # -----------------------------------------------------------------
            # Passo 4 – Selecionar conta principal no grid bancário
            # -----------------------------------------------------------------
            print("   [4] Selecionando conta principal no grid bancário...")
            grid_banco = WebDriverWait(navegador, 20).until(
                EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_WucBuscaCadastraPessoaContaBancaria1_gdwConsultaBanco"))
            )
            conta_ok = False
            for linha_banco in grid_banco.find_elements(By.CSS_SELECTOR, "tr.gridRow"):
                cols_banco = linha_banco.find_elements(By.TAG_NAME, "td")
                if cols_banco and cols_banco[-1].text.strip() == "Sim":
                    link_sel = linha_banco.find_element(By.CSS_SELECTOR, "a[id*='lbtSelecionar']")
                    scroll_js_clicar(link_sel)
                    conta_ok = True
                    break
            if not conta_ok:
                raise RuntimeError("Conta principal não encontrada no grid bancário")
            ss("conta_selecionada")

            # -----------------------------------------------------------------
            # Passo 5 – Ler finalidade (antes de salvar)
            # -----------------------------------------------------------------
            print("   [5] Lendo finalidade do campo Finalidade...")
            campo_finalidade = eid("ctl00_ContentPlaceHolder1_PagamentoPessoaFisicaUserControl_txtJustificativa")
            rolar(campo_finalidade)
            finalidade = campo_finalidade.get_attribute("value").strip()
            print(f"      Finalidade: {finalidade[:100]}")
            ss("finalidade_lida")

            # -----------------------------------------------------------------
            # Passo 6 – Confirmar Conta Bancária (checkbox)
            # -----------------------------------------------------------------
            print("   [6] Marcando 'Confirmar Conta Bancária'...")
            cb_conf = eid("ctl00_ContentPlaceHolder1_PagamentoPessoaFisicaUserControl_cbxConfContaBancaria")
            js("arguments[0].scrollIntoView({block:'center'});", cb_conf)
            if not cb_conf.is_selected():
                interagir_com_loader(navegador, lambda: js("arguments[0].click();", cb_conf))
            ss("checkbox_conta")

            # -----------------------------------------------------------------
            # Passo 7 – Salvar
            # -----------------------------------------------------------------
            print("   [7] Clicando em Salvar...")
            btn_salvar = eid("ctl00_ContentPlaceHolder1_btnSalvar")
            scroll_js_clicar(btn_salvar)
            ss("salvar")

            # -----------------------------------------------------------------
            # Passo 8 – Confirmar diálogo OK após Salvar
            # -----------------------------------------------------------------
            print("   [8] Confirmando diálogo OK após Salvar...")
            ok_dialog()
            ss("ok_salvar")

            # -----------------------------------------------------------------
            # Passo 9 – Clicar em Aprovar (abre o formulário de aprovação)
            # -----------------------------------------------------------------
            print("   [9] Clicando em Aprovar (abre formulário de aprovação)...")
            btn_aprovar_1 = eid("ctl00_ContentPlaceHolder1_btnAprovar")
            scroll_js_clicar(btn_aprovar_1)
            ss("aprovar_abrir_form")

            # -----------------------------------------------------------------
            # Passo 10 – Rubrica = 23 (Serviços de Terceiros Pessoa Física)
            # -----------------------------------------------------------------
            print("   [10] Selecionando Rubrica = Serviços de Terceiros Pessoa Física...")
            ddl_rub = eid("ctl00_ContentPlaceHolder1_PagamentoPessoaFisicaAprovarUserControl_ddlRubrica")
            js("arguments[0].scrollIntoView({block:'center'});", ddl_rub)
            interagir_com_loader(navegador, lambda: Select(ddl_rub).select_by_value("23"))
            ss("rubrica_23")

            # -----------------------------------------------------------------
            # Passo 11 – Histórico = finalidade
            # -----------------------------------------------------------------
            print("   [11] Preenchendo Histórico com a finalidade...")
            campo_hist = eid("ctl00_ContentPlaceHolder1_PagamentoPessoaFisicaAprovarUserControl_txtHistorico")
            rolar(campo_hist)
            limpar_elemento(navegador, campo_hist)
            digitar_elemento(navegador, campo_hist, finalidade)
            ss("historico_preenchido")

            # -----------------------------------------------------------------
            # Passo 12 – Data de Pagamento (TAB dispara onchange / postback)
            # -----------------------------------------------------------------
            print(f"   [12] Preenchendo Data de Pagamento: {data_pagamento}...")
            campo_dp = eid("ctl00_ContentPlaceHolder1_PagamentoPessoaFisicaAprovarUserControl_txtDataPagamento")
            rolar(campo_dp)
            limpar_elemento(navegador, campo_dp)
            digitar_elemento(navegador, campo_dp, data_pagamento)
            interagir_com_loader(navegador, lambda: campo_dp.send_keys(Keys.TAB))
            ss("data_pagamento")

            # -----------------------------------------------------------------
            # Passo 13 – Mês de Competência
            # -----------------------------------------------------------------
            print(f"   [13] Preenchendo Mês de Competência: {competencia}...")
            campo_comp = eid("ctl00_ContentPlaceHolder1_PagamentoPessoaFisicaAprovarUserControl_txtDataCompetencia")
            rolar(campo_comp)
            limpar_elemento(navegador, campo_comp)
            digitar_elemento(navegador, campo_comp, competencia)
            ss("competencia")

            # -----------------------------------------------------------------
            # Passo 14 – Horas Trabalhadas = 000:00
            # -----------------------------------------------------------------
            print("   [14] Preenchendo Horas Trabalhadas = 000:00...")
            campo_horas = eid("txtHorasTrabalhadas")
            rolar(campo_horas)
            limpar_elemento(navegador, campo_horas)
            digitar_elemento(navegador, campo_horas, "000:00")
            ss("horas_trabalhadas")

            # -----------------------------------------------------------------
            # Passo 15 – Calcular Imposto
            # -----------------------------------------------------------------
            print("   [15] Calculando impostos...")
            btn_calc = eid("ctl00_ContentPlaceHolder1_PagamentoPessoaFisicaAprovarUserControl_btnCalcularImposto")
            scroll_js_clicar(btn_calc)
            ss("calcular_imposto")

            # -----------------------------------------------------------------
            # Passo 16 – Marcar "Plano de trabalho para Impostos"
            # -----------------------------------------------------------------
            print("   [16] Marcando 'Plano de trabalho para Impostos'...")
            chk_plano_imp = eid("ctl00_ContentPlaceHolder1_PagamentoPessoaFisicaAprovarUserControl_chkImpostosPlano")
            js("arguments[0].scrollIntoView({block:'center'});", chk_plano_imp)
            if not chk_plano_imp.is_selected():
                interagir_com_loader(navegador, lambda: js("arguments[0].click();", chk_plano_imp))
            ss("chk_plano_impostos")

            # -----------------------------------------------------------------
            # Passo 17 – Marcar TODOS os impostos (header checkbox)
            # -----------------------------------------------------------------
            print("   [17] Marcando todos os impostos via header checkbox...")
            chk_hdr = eid("ctl00_ContentPlaceHolder1_PagamentoPessoaFisicaAprovarUserControl_gdwImposCalcImposto_3_ctl01_chkHeader")
            js("arguments[0].scrollIntoView({block:'center'});", chk_hdr)
            if not chk_hdr.is_selected():
                interagir_com_loader(navegador, lambda: js("arguments[0].click();", chk_hdr))
            ss("todos_impostos_marcados")

            # -----------------------------------------------------------------
            # Passo 18 – Desmarcar INSS Patronal no grid de impostos
            # -----------------------------------------------------------------
            print("   [18] Desmarcando INSS Patronal do plano de impostos...")
            grid_imp = eid("ctl00_ContentPlaceHolder1_PagamentoPessoaFisicaAprovarUserControl_gdwImposCalcImposto_3")
            tem_inss = False
            for linha_imp in grid_imp.find_elements(By.CSS_SELECTOR, "tr.gridRow, tr.gridAlternateRow"):
                tds_imp = linha_imp.find_elements(By.TAG_NAME, "td")
                if any("INSS Patronal" in td.text for td in tds_imp):
                    cbs_imp = linha_imp.find_elements(By.CSS_SELECTOR, "input[type='checkbox']")
                    if cbs_imp:
                        cb_inss = cbs_imp[0]
                        js("arguments[0].scrollIntoView({block:'center'});", cb_inss)
                        if cb_inss.is_selected():
                            interagir_com_loader(navegador, lambda el=cb_inss: js("arguments[0].click();", el))
                        tem_inss = True
                    break
            if not tem_inss:
                print("   ⚠ INSS Patronal não encontrado no grid de impostos")
            ss("desmarcar_inss_patronal")

            # -----------------------------------------------------------------
            # Passo 19 – Definir plano de trabalho (itens não-INSS)
            # -----------------------------------------------------------------
            print("   [19] Clicando em Definir plano de trabalho...")
            btn_def_pt = eid("ctl00_ContentPlaceHolder1_PagamentoPessoaFisicaAprovarUserControl_btnDefinirPlanoTrabalhoLote")
            scroll_js_clicar(btn_def_pt)
            ss("definir_plano_trabalho")

            # -----------------------------------------------------------------
            # Passo 20 – Marcar a linha da rubrica no grid do plano de trabalho
            # Usa JS para localizar o checkbox e evitar StaleElementReferenceException
            # -----------------------------------------------------------------
            print(f"   [20] Procurando rubrica '{rubrica}' no plano de trabalho...")
            cb_id_pt = js("""
                var grid = document.getElementById(
                    'ctl00_ContentPlaceHolder1_wucConvenioPlanoTrabalhoAprovarUserControl1_gdwPlanoTrabalho'
                );
                if (!grid) return null;
                var rows = grid.querySelectorAll('tr[data-grupo-rubrica]');
                var termo = arguments[0].toLowerCase();
                for (var i = 0; i < rows.length; i++) {
                    var tds = rows[i].querySelectorAll('td');
                    if (tds.length >= 4 && tds[3].textContent.trim().toLowerCase().indexOf(termo) !== -1) {
                        var cb = rows[i].querySelector('input[type=checkbox]');
                        if (cb) return cb.id;
                    }
                }
                return null;
            """, rubrica)
            if not cb_id_pt:
                raise RuntimeError(f"Rubrica '{rubrica}' não encontrada no plano de trabalho")
            cb_pt = eid(cb_id_pt)
            rolar(cb_pt)
            print(f"      ✓ Encontrada (id={cb_id_pt})")
            if not cb_pt.is_selected():
                interagir_com_loader(navegador, lambda el=cb_pt: js("arguments[0].click();", el))
            ss("rubrica_no_plano")

            # -----------------------------------------------------------------
            # Passo 21 – OK no plano de trabalho
            # -----------------------------------------------------------------
            print("   [21/22] OK no plano + diálogo de confirmação...")
            btn_ok_pt = eid("ctl00_ContentPlaceHolder1_wucConvenioPlanoTrabalhoAprovarUserControl1_btnDefinir")
            scroll_js_clicar(btn_ok_pt)
            ok_dialog_com_sim()
            ss("ok_plano_e_dialogo")

            # -----------------------------------------------------------------
            # Passo 23 – Fechar diálogo do plano de trabalho
            # -----------------------------------------------------------------
            print("   [23] Fechando diálogo do plano de trabalho...")
            fechar_dialog()
            ss("fechar_plano")

            # -----------------------------------------------------------------
            # Passo 24 – Marcar INSS Patronal no grid de impostos
            # -----------------------------------------------------------------
            print("   [24] Marcando INSS Patronal no grid de impostos...")
            grid_imp2 = eid("ctl00_ContentPlaceHolder1_PagamentoPessoaFisicaAprovarUserControl_gdwImposCalcImposto_3")
            for linha_imp2 in grid_imp2.find_elements(By.CSS_SELECTOR, "tr.gridRow, tr.gridAlternateRow"):
                tds_imp2 = linha_imp2.find_elements(By.TAG_NAME, "td")
                if any("INSS Patronal" in td.text for td in tds_imp2):
                    cbs_imp2 = linha_imp2.find_elements(By.CSS_SELECTOR, "input[type='checkbox']")
                    if cbs_imp2:
                        cb_inss2 = cbs_imp2[0]
                        js("arguments[0].scrollIntoView({block:'center'});", cb_inss2)
                        if not cb_inss2.is_selected():
                            interagir_com_loader(navegador, lambda el=cb_inss2: js("arguments[0].click();", el))
                    break
            ss("marcar_inss_patronal")

            # -----------------------------------------------------------------
            # Passo 25 – Definir plano de trabalho para INSS Patronal
            # -----------------------------------------------------------------
            print("   [25] Definindo plano de trabalho para INSS Patronal...")
            btn_def_pt2 = eid("ctl00_ContentPlaceHolder1_PagamentoPessoaFisicaAprovarUserControl_btnDefinirPlanoTrabalhoLote")
            scroll_js_clicar(btn_def_pt2)
            ss("definir_plano_inss")

            # -----------------------------------------------------------------
            # Passo 26 – Marcar linha INSS Patronal no plano
            # (linha especial com lbtDefinir em vez de checkbox)
            # -----------------------------------------------------------------
            print("   [26] Marcando INSS Patronal no plano de trabalho...")
            # JS localiza checkbox ou link lbtDefinir da linha INSS Patronal
            inss_info = js("""
                var grid = document.getElementById(
                    'ctl00_ContentPlaceHolder1_wucConvenioPlanoTrabalhoAprovarUserControl1_gdwPlanoTrabalho'
                );
                if (!grid) return null;
                var rows = grid.querySelectorAll('tr');
                for (var i = 0; i < rows.length; i++) {
                    if (rows[i].textContent.indexOf('INSS Patronal') !== -1) {
                        var cb = rows[i].querySelector('input[type=checkbox]');
                        if (cb) return {tipo: 'checkbox', id: cb.id};
                        var lnk = rows[i].querySelector('a[title="Definir pagamento"]');
                        if (lnk) return {tipo: 'link', id: lnk.id};
                    }
                }
                return null;
            """)
            if not inss_info:
                raise RuntimeError("Linha INSS Patronal não encontrada no plano de trabalho")
            elem_inss = eid(inss_info["id"])
            rolar(elem_inss)
            if inss_info["tipo"] == "checkbox":
                if not elem_inss.is_selected():
                    interagir_com_loader(navegador, lambda el=elem_inss: js("arguments[0].click();", el))
            else:
                interagir_com_loader(navegador, lambda el=elem_inss: js("arguments[0].click();", el))
            ss("inss_no_plano")

            # -----------------------------------------------------------------
            # Passo 27/28 – OK no plano INSS + diálogo de confirmação
            # -----------------------------------------------------------------
            print("   [27/28] OK no plano INSS + diálogo de confirmação...")
            btn_ok_pt2 = eid("ctl00_ContentPlaceHolder1_wucConvenioPlanoTrabalhoAprovarUserControl1_btnDefinir")
            scroll_js_clicar(btn_ok_pt2)
            ok_dialog_com_sim()
            ss("ok_plano_inss_e_dialogo")

            # -----------------------------------------------------------------
            # Passo 29 – Fechar diálogo do plano INSS Patronal
            # -----------------------------------------------------------------
            print("   [29] Fechando diálogo do plano INSS Patronal...")
            fechar_dialog()
            ss("fechar_plano_inss")

            # -----------------------------------------------------------------
            # Passo 30 – Aprovação final
            # -----------------------------------------------------------------
            print("   [30] Clicando em Aprovar (aprovação final)...")
            btn_aprovar_final = eid("ctl00_ContentPlaceHolder1_btnAprovar")
            scroll_js_clicar(btn_aprovar_final)
            ss("aprovar_final_clicado")

            # Pode aparecer diálogo "condições identificadas" sim/não → confirmar
            aguardar_loader_sumir(navegador)
            try:
                btn_sim_final = WebDriverWait(navegador, 5).until(
                    EC.element_to_be_clickable((
                        By.XPATH,
                        "//div[contains(@class,'ui-dialog-buttonset')]"
                        "//button/span[normalize-space()='sim']/..",
                    ))
                )
                js("arguments[0].scrollIntoView({block:'center'});", btn_sim_final)
                js("arguments[0].click();", btn_sim_final)
                aguardar_loader_sumir(navegador)
                print("      (diálogo de condições confirmado com sim)")
            except TimeoutException:
                pass

            # Aguardar diálogo de sucesso "Pedido Nº XXXXX aprovado."
            print("   [30] Aguardando mensagem de sucesso...")
            dlg_ok = WebDriverWait(navegador, 30).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, ".ui-dialog-buttonset button"))
            )
            # Lê texto pelo label específico do diálogo de aprovação
            texto_dialogo = js("""
                var lbl = document.getElementById(
                    'ctl00_ContentPlaceHolder1_DialogUserControl1_lblMensagem'
                );
                if (lbl) return lbl.textContent.trim();
                var dialogs = document.querySelectorAll('.ui-dialog');
                for (var i = 0; i < dialogs.length; i++) {
                    var d = dialogs[i];
                    if (window.getComputedStyle(d).display !== 'none') {
                        var c = d.querySelector('.ui-dialog-content');
                        if (c) return c.textContent.trim();
                    }
                }
                return '';
            """) or ""

            if "aprovado" not in texto_dialogo.lower():
                raise RuntimeError(
                    f"Diálogo final inesperado (esperado 'aprovado'): {texto_dialogo!r}"
                )

            js("arguments[0].scrollIntoView({block:'center'});", dlg_ok)
            js("arguments[0].click();", dlg_ok)
            aguardar_loader_sumir(navegador)
            ss("aprovado_final_ok")

            print(f"   ✅ Pedido {pedido} aprovado com sucesso! Mensagem: {texto_dialogo.strip()!r}")
            navegador.close()
            navegador.switch_to.window(aba_consulta)
            return "Aprovado"

        except Exception as exc:
            print(f"❌ Erro ao processar {pedido}: {exc}")
            try:
                handles = navegador.window_handles
                if len(handles) > 1:
                    navegador.switch_to.window(handles[-1])
                    navegador.close()
                if navegador.window_handles:
                    navegador.switch_to.window(navegador.window_handles[0])
            except (WebDriverException, IndexError):
                pass
            return f"Falha: {exc}"

    return "Pedido não encontrado"


# ---------------------------------------------------------------------------
# Worker por thread / navegador
# ---------------------------------------------------------------------------

eventos_painel: queue.Queue = queue.Queue()


def executar_parte(parte, caminho_arquivo, indice_trabalhador, data_pagamento, competencia):
    navegador = None
    resultados = []        # (posicao, pedido, resultado)
    registros_status = []  # (pedido, item_pt, favorecido, resultado) → aba status_aprov
    try:
        with tempfile.TemporaryDirectory(prefix=f"aprovar-{indice_trabalhador}-") as pasta:
            CONTEXTO_TRABALHADOR.perfil = pasta
            eventos_painel.put(("status", f"Navegador {indice_trabalhador}: iniciando sessão..."))
            navegador = realizar_login()
            if navegador is None:
                raise RuntimeError("Login não retornou um navegador Selenium.")

            for linha_indice, linha in parte.iterrows():
                pedido = str(linha["Nº Pedido"]).strip()
                rubrica = str(linha.get("Item PT", "")).strip()
                favorecido = str(linha.get("Favorecido", "")).strip()
                data_nascimento = str(linha.get("Data Nascimento", "")).strip()
                if data_nascimento in ("nan", "None", "NaT"):
                    data_nascimento = ""
                posicao = int(linha_indice)
                eventos_painel.put(("pedido", posicao, pedido, indice_trabalhador))
                try:
                    resultado = funcao_aprovar(
                        navegador, pedido, posicao, rubrica, data_pagamento, competencia,
                        data_nascimento=data_nascimento,
                    )
                    resultado = resultado or "Resultado incerto"
                except Exception as exc:
                    resultado = f"Falha: {exc}"

                resultados.append((posicao, pedido, resultado))
                registros_status.append((pedido, rubrica, favorecido, resultado))
                eventos_painel.put(("concluido", posicao, pedido, resultado, indice_trabalhador))

                # Grava imediatamente; move para aba 2 se aprovado ou não encontrado
                try:
                    registrar_e_mover_se_concluido(caminho_arquivo, pedido, resultado)
                except Exception as exc_salvar:
                    eventos_painel.put(("status", f"Aviso: não foi possível salvar na planilha: {exc_salvar}"))


    except Exception as exc:
        eventos_painel.put(("status", f"Navegador {indice_trabalhador}: falha na sessão ({exc})"))
        for posicao, linha in parte.iterrows():
            if not any(r[0] == posicao for r in resultados):
                res = f"Falha na sessão do navegador {indice_trabalhador}: {exc}"
                pedido = str(linha["Nº Pedido"]).strip()
                rubrica = str(linha.get("Item PT", "")).strip()
                favorecido = str(linha.get("Favorecido", "")).strip()
                resultados.append((int(posicao), pedido, res))
                registros_status.append((pedido, rubrica, favorecido, res))
                eventos_painel.put((
                    "concluido", int(posicao), pedido, res, indice_trabalhador,
                ))
    finally:
        if navegador is not None:
            try:
                navegador.quit()
            except WebDriverException:
                pass
        CONTEXTO_TRABALHADOR.perfil = None

    return resultados, registros_status


# ---------------------------------------------------------------------------
# Coordenador de planilha
# ---------------------------------------------------------------------------

def processar_planilha(caminho_arquivo, data_pagamento, competencia):
    try:
        from concurrent.futures import ThreadPoolExecutor, as_completed

        if Path(caminho_arquivo).suffix.lower() != ".xlsx":
            raise ValueError("Selecione uma planilha .xlsx (converta .xls antes de iniciar).")

        df = pd.read_excel(caminho_arquivo, dtype={"Nº Pedido": str})
        df.columns = [str(c).strip() for c in df.columns]
        # Normaliza variações de nome da coluna data de nascimento
        _NASC_ALIASES = {"data_nascimento", "datanascim", "data nascimento", "dt_nascimento", "nascimento"}
        df.columns = [
            "Data Nascimento" if c.lower().replace(" ", "_").strip() in _NASC_ALIASES else c
            for c in df.columns
        ]

        if "Nº Pedido" not in df.columns:
            raise ValueError("A planilha precisa ter a coluna 'Nº Pedido'.")
        if "Item PT" not in df.columns:
            raise ValueError("A planilha precisa ter a coluna 'Item PT'.")

        df = df[df["Nº Pedido"].notna()].copy()
        df["Nº Pedido"] = df["Nº Pedido"].astype(str).str.strip()
        df = df[(df["Nº Pedido"] != "") & (df["Nº Pedido"].str.casefold() != "nan")].copy()

        # Processa apenas linhas com Status = "ok" (ou todas se não houver coluna Status)
        if "Status" in df.columns:
            df = df[df["Status"].astype(str).str.strip().str.lower() == "ok"].copy()

        total = len(df)
        if not total:
            eventos_painel.put(("aviso", "Nenhum pedido válido (Status=ok) encontrado na planilha."))
            return

        df = df.reset_index(drop=True)
        n = min(NUM_NAVEGADORES, len(df))
        partes = [df.iloc[i::n].copy() for i in range(n) if len(df.iloc[i::n]) > 0]
        eventos_painel.put(("total", total))
        eventos_painel.put(("pedidos", [(int(i), str(row["Nº Pedido"])) for i, row in df.iterrows()]))
        eventos_painel.put(("status", f"Executando {len(partes)} navegadores simultâneos..."))

        todos_resultados = []
        todos_registros_status = []
        with ThreadPoolExecutor(max_workers=NUM_NAVEGADORES, thread_name_prefix="aprovar") as executor:
            futuros = [
                executor.submit(executar_parte, parte, caminho_arquivo, n, data_pagamento, competencia)
                for n, parte in enumerate(partes, 1)
            ]
            for futuro in as_completed(futuros):
                lote, registros = futuro.result()
                todos_resultados.extend(lote)
                todos_registros_status.extend(registros)

        saida = consolidar_planilha(caminho_arquivo, todos_registros_status)
        aprovados = sum(1 for _, _, s in todos_resultados if s == "Aprovado")
        nao_encontrados = sum(1 for _, _, s in todos_resultados if "não encontrado" in s)
        falhas = sum(1 for _, _, s in todos_resultados if s.startswith("Falha"))

        msg = (
            f"Concluído: {aprovados} aprovados, {nao_encontrados} não encontrados, "
            f"{falhas} falhas.\nArquivo: {saida}"
        )
        eventos_painel.put(("fim", msg))

    except Exception as exc:
        eventos_painel.put(("erro", str(exc)))


# ---------------------------------------------------------------------------
# Interface gráfica
# ---------------------------------------------------------------------------
processamento_ativo = False
concluidos_painel = 0
total_painel = 0

# ---------------------------------------------------------------------------
# Cores — inspiradas no padrão visual da empresa (Central de Passagens)
# ---------------------------------------------------------------------------
BG         = "#203447"
CARD_BG    = "#f8f7f3"
CARD_FT    = "#f1f1ed"
TEXT       = "#efede5"
TEXT_MUTED = "#8a9bab"
LABEL_FG   = "#344a5b"
LABEL_MUT  = "#7a838a"
ACTION     = "#e94c1f"
ACTION_HV  = "#c73e18"
STEP_BG    = "#e9eef0"
BORDER     = "#dde1df"
DIVIDER    = "#415364"

# ---------------------------------------------------------------------------
# Janela raiz
# ---------------------------------------------------------------------------
janela = tk.Tk()
janela.title(f"Aprovação de Pedidos  v{__version__}")
janela.geometry("1020x750")
janela.minsize(820, 580)
janela.configure(bg=BG)

try:
    janela.iconbitmap(str(_caminho_recurso("logo.ico")))
except Exception:
    pass

# ---------------------------------------------------------------------------
# Cabeçalho
# ---------------------------------------------------------------------------
frame_cabecalho = tk.Frame(janela, bg=BG)
frame_cabecalho.pack(fill="x", padx=48, pady=(22, 0))

cab_esq = tk.Frame(frame_cabecalho, bg=BG)
cab_esq.pack(side="left", fill="y")
tk.Label(cab_esq, text="A U T O M A Ç Ã O   O P E R A C I O N A L",
         font=("Segoe UI", 8, "bold"), bg=BG, fg="#85929c").pack(anchor="w")
tk.Label(cab_esq, text="Aprovar Pedidos",
         font=("Segoe UI", 15, "bold"), bg=BG, fg=TEXT).pack(anchor="w", pady=(4, 0))

cab_dir = tk.Frame(frame_cabecalho, bg=BG)
cab_dir.pack(side="right", anchor="e")
_lbl_update_dot = tk.Label(cab_dir, text="●", font=("Segoe UI", 10), bg=BG,
                            fg=auto_update.COR_UNKNOWN if auto_update else "#85929c")
_lbl_update_dot.pack(side="left", padx=(0, 4))
_lbl_update_txt = tk.Label(cab_dir, text="Verificando...", font=("Segoe UI", 9),
                            bg=BG, fg=TEXT_MUTED)
_lbl_update_txt.pack(side="left")

tk.Frame(janela, bg=DIVIDER, height=1).pack(fill="x", pady=(14, 0))

# ---------------------------------------------------------------------------
# Corpo: painel esquerdo + cartão direito
# ---------------------------------------------------------------------------
corpo = tk.Frame(janela, bg=BG)
corpo.pack(fill="both", expand=True, padx=48, pady=24)

# --- painel esquerdo ---
painel_esq = tk.Frame(corpo, bg=BG, width=250)
painel_esq.pack(side="left", fill="y", padx=(0, 44))
painel_esq.pack_propagate(False)

tk.Label(painel_esq, text="A P R O V A Ç Ã O",
         font=("Segoe UI", 8, "bold"), bg=BG, fg=ACTION).pack(anchor="w", pady=(10, 8))
tk.Label(painel_esq, text="Execute aprovações\ncom segurança.",
         font=("Segoe UI", 17, "bold"), bg=BG, fg=TEXT, justify="left").pack(anchor="w")
tk.Label(painel_esq,
         text="Selecione a planilha com os pedidos e preencha as datas. O robô abrirá o Conveniar e aprovará cada pedido automaticamente.",
         font=("Segoe UI", 10), bg=BG, fg="#a5afb6",
         wraplength=220, justify="left").pack(anchor="w", pady=(14, 0))

tk.Frame(painel_esq, bg=DIVIDER, height=1).pack(fill="x", pady=20)

rodape_esq = tk.Frame(painel_esq, bg=BG)
rodape_esq.pack(anchor="w")
tk.Label(rodape_esq, text="🔒", font=("Segoe UI", 15), bg=BG, fg="#304658").pack(side="left", padx=(0, 8), anchor="n")
tk.Label(rodape_esq,
         text="Seus arquivos são processados localmente e não são compartilhados.",
         font=("Segoe UI", 9), bg=BG, fg="#8998a3",
         wraplength=185, justify="left").pack(side="left")

# --- cartão direito ---
frame_cartao = tk.Frame(corpo, bg=CARD_BG)
frame_cartao.pack(side="left", fill="both", expand=True)

# título do cartão
frame_titulo = tk.Frame(frame_cartao, bg=CARD_BG)
frame_titulo.pack(fill="x", padx=28, pady=(18, 0))

rotulo_contagem = tk.Label(frame_titulo, text="0 de 3",
                            font=("Segoe UI", 9, "bold"), bg=CARD_FT, fg=LABEL_MUT,
                            relief="solid", bd=1, padx=10, pady=5)
rotulo_contagem.pack(side="right", anchor="n")

tk.Label(frame_titulo, text="Configurar aprovação",
         font=("Segoe UI", 13, "bold"), bg=CARD_BG, fg=LABEL_FG).pack(anchor="w")
tk.Label(frame_titulo, text="Preencha as três etapas para continuar",
         font=("Segoe UI", 9), bg=CARD_BG, fg=LABEL_MUT).pack(anchor="w", pady=(3, 10))

tk.Frame(frame_cartao, bg=BORDER, height=1).pack(fill="x")


def _etapa(parent, numero, titulo, descricao, criar_widget=None):
    """Linha de etapa: ícone  |  textos  |  widget opcional à direita."""
    f = tk.Frame(parent, bg=CARD_BG)
    f.pack(fill="x")
    inner = tk.Frame(f, bg=CARD_BG)
    inner.pack(fill="x", padx=16, pady=16)

    icone_f = tk.Frame(inner, bg=STEP_BG, width=42, height=42)
    icone_f.pack(side="left", padx=(0, 12))
    icone_f.pack_propagate(False)
    tk.Label(icone_f, text=numero, font=("Segoe UI", 12, "bold"),
             bg=STEP_BG, fg=LABEL_FG).pack(expand=True)

    widget = None
    if criar_widget:
        widget = criar_widget(inner)
        widget.pack(side="right", padx=(12, 0), anchor="center")

    textos = tk.Frame(inner, bg=CARD_BG)
    textos.pack(side="left", fill="both", expand=True)
    tk.Label(textos, text="ETAPA " + numero,
             font=("Segoe UI", 8, "bold"), bg=CARD_BG, fg=LABEL_MUT).pack(anchor="w")
    tk.Label(textos, text=titulo,
             font=("Segoe UI", 11, "bold"), bg=CARD_BG, fg=LABEL_FG).pack(anchor="w")
    tk.Label(textos, text=descricao,
             font=("Segoe UI", 9), bg=CARD_BG, fg="#6d7780",
             wraplength=260, justify="left").pack(anchor="w", pady=(2, 0))
    return f, widget


# etapa 1 — planilha
_caminho_planilha = {"valor": ""}

def _btn_linha(parent, texto, cmd):
    return tk.Button(parent, text=texto, command=cmd,
                     font=("Segoe UI", 10, "bold"),
                     bg="white", fg=LABEL_FG, relief="solid", bd=1,
                     padx=14, pady=7, cursor="hand2",
                     activebackground="#f1f3f1", activeforeground=LABEL_FG)

btn_planilha = None

def _selecionar_planilha_ui():
    caminho = filedialog.askopenfilename(
        title="Selecione a planilha de pedidos",
        filetypes=[("Arquivos Excel", "*.xlsx")],
    )
    if caminho:
        _caminho_planilha["valor"] = caminho
        nome = Path(caminho).name
        rotulo_planilha.configure(text=nome)
        btn_planilha.configure(text="Alterar")
        _verificar_pronto()

_, _w1 = _etapa(frame_cartao, "01", "Planilha de pedidos",
                "Arquivo .xlsx com os pedidos a aprovar.",
                lambda p: _btn_linha(p, "Selecionar planilha", _selecionar_planilha_ui))
btn_planilha = _w1

rotulo_planilha = tk.Label(frame_cartao, text="Nenhuma selecionada",
                            font=("Segoe UI", 9, "italic"), bg=CARD_BG, fg=LABEL_MUT)
rotulo_planilha.pack(anchor="e", padx=20, pady=(0, 6))

tk.Frame(frame_cartao, bg=BORDER, height=1).pack(fill="x")

# etapa 2 — data de pagamento
entrada_data = None

def _criar_entry_data(parent):
    global entrada_data
    entrada_data = tk.Entry(parent, font=("Segoe UI", 11), width=13,
                            relief="solid", bd=1, fg=LABEL_FG)
    entrada_data.bind("<KeyRelease>", lambda e: _verificar_pronto())
    return entrada_data

_etapa(frame_cartao, "02", "Data de Pagamento",
       "Formato DD/MM/AAAA.", _criar_entry_data)

tk.Frame(frame_cartao, bg=BORDER, height=1).pack(fill="x")

# etapa 3 — competência
entrada_competencia = None

def _criar_entry_comp(parent):
    global entrada_competencia
    entrada_competencia = tk.Entry(parent, font=("Segoe UI", 11), width=13,
                                   relief="solid", bd=1, fg=LABEL_FG)
    entrada_competencia.bind("<KeyRelease>", lambda e: _verificar_pronto())
    return entrada_competencia

_etapa(frame_cartao, "03", "Competência",
       "Formato MM/AAAA.", _criar_entry_comp)

tk.Frame(frame_cartao, bg=BORDER, height=1).pack(fill="x")

# --- Ocultar janela (checkbox headless) ---
frame_headless = tk.Frame(frame_cartao, bg=CARD_BG)
frame_headless.pack(fill="x", padx=20, pady=10)

var_headless = tk.BooleanVar(value=HEADLESS_CHROME)

def _ao_toggle_headless():
    global HEADLESS_CHROME
    HEADLESS_CHROME = var_headless.get()

tk.Checkbutton(
    frame_headless, text="Ocultar janela",
    variable=var_headless, command=_ao_toggle_headless,
    font=("Segoe UI", 10), bg=CARD_BG, fg=LABEL_FG,
    activebackground=CARD_BG, selectcolor="white",
).pack(side="left")
tk.Label(frame_headless, text="(marque para rodar em segundo plano, sem mostrar o Chrome)",
         font=("Segoe UI", 8), bg=CARD_BG, fg=LABEL_MUT).pack(side="left", padx=4)

# --- Rodapé do cartão ---
frame_rodape_cartao = tk.Frame(frame_cartao, bg=CARD_FT)
frame_rodape_cartao.pack(fill="x", side="bottom")

inner_rodape = tk.Frame(frame_rodape_cartao, bg=CARD_FT)
inner_rodape.pack(fill="x", padx=28, pady=18)

textos_rod = tk.Frame(inner_rodape, bg=CARD_FT)
textos_rod.pack(side="left", fill="both", expand=True)
rotulo_pronto_titulo = tk.Label(textos_rod, text="Aguardando configurações",
                                 font=("Segoe UI", 10, "bold"), bg=CARD_FT, fg=LABEL_FG)
rotulo_pronto_titulo.pack(anchor="w")
rotulo_pronto_dica = tk.Label(textos_rod, text="Complete as etapas acima para continuar.",
                               font=("Segoe UI", 9), bg=CARD_FT, fg=LABEL_MUT, wraplength=320)
rotulo_pronto_dica.pack(anchor="w")

botao_acao = tk.Button(
    inner_rodape, text="INICIAR APROVAÇÃO  →",
    font=("Segoe UI", 10, "bold"), bg="#d8dad9", fg="#92999b",
    relief="flat", padx=18, pady=12, cursor="hand2",
    state="disabled",
)
botao_acao.pack(side="right")

# ---------------------------------------------------------------------------
# Área de progresso (aparece ao processar, abaixo do cabeçalho)
# ---------------------------------------------------------------------------
frame_progresso = tk.Frame(janela, bg=BG)

rotulo_status = tk.Label(frame_progresso,
                          text="Processando...",
                          font=("Segoe UI", 10), bg=BG, fg=TEXT, wraplength=600)
rotulo_status.pack(pady=(4, 2), padx=48, anchor="w")

progresso = tk.IntVar(value=0)
barra_progresso = ttk.Progressbar(frame_progresso, variable=progresso,
                                   maximum=100, length=600, mode="determinate")
barra_progresso.pack(pady=4, padx=48)

rotulo_progresso = tk.Label(frame_progresso, text="0/0 pedidos",
                              font=("Segoe UI", 9), bg=BG, fg=TEXT_MUTED)
rotulo_progresso.pack(pady=(0, 4), padx=48, anchor="w")

tk.Label(frame_progresso, text="PEDIDOS", font=("Segoe UI", 9, "bold"),
         bg=BG, fg=TEXT_MUTED).pack(anchor="w", padx=48, pady=(4, 2))

quadro_lista = tk.Frame(frame_progresso, bg=BG)
quadro_lista.pack(fill="both", expand=True, padx=48, pady=(0, 8))

_sty = ttk.Style()
_sty.theme_use("default")
_sty.configure("Dark.Treeview", background="#2a4055", foreground=TEXT,
               fieldbackground="#2a4055", rowheight=24, font=("Segoe UI", 9))
_sty.configure("Dark.Treeview.Heading", background="#1a3047", foreground=TEXT_MUTED,
               font=("Segoe UI", 9, "bold"), relief="flat")
_sty.map("Dark.Treeview", background=[("selected", "#3a5a75")])

lista_pedidos = ttk.Treeview(quadro_lista, columns=("pedido", "status"),
                              show="headings", height=8, style="Dark.Treeview")
lista_pedidos.heading("pedido", text="Número do Pedido")
lista_pedidos.heading("status", text="Status")
lista_pedidos.column("pedido", width=300, minwidth=150, anchor="w")
lista_pedidos.column("status", width=320, minwidth=160, anchor="w")

scroll_pedidos = ttk.Scrollbar(quadro_lista, orient="vertical", command=lista_pedidos.yview)
lista_pedidos.configure(yscrollcommand=scroll_pedidos.set)
lista_pedidos.pack(side="left", fill="both", expand=True)
scroll_pedidos.pack(side="right", fill="y")

# ---------------------------------------------------------------------------
# Rodapé da janela
# ---------------------------------------------------------------------------
tk.Frame(janela, bg=DIVIDER, height=1).pack(fill="x")
rodape_janela = tk.Frame(janela, bg=BG)
rodape_janela.pack(fill="x", padx=48, pady=12)
tk.Label(rodape_janela, text="Aprovar Pedidos RPA",
         font=("Segoe UI", 9), bg=BG, fg="#83919c").pack(side="left")
tk.Label(rodape_janela, text="Ambiente seguro",
         font=("Segoe UI", 9), bg=BG, fg="#83919c").pack(side="right")

# ---------------------------------------------------------------------------
# Lógica do painel
# ---------------------------------------------------------------------------

def _verificar_pronto():
    """Habilita o botão de início somente quando os 3 campos estiverem preenchidos."""
    planilha_ok = bool(_caminho_planilha["valor"])
    data_ok = bool(entrada_data and entrada_data.get().strip())
    comp_ok = bool(entrada_competencia and entrada_competencia.get().strip())
    count = sum([planilha_ok, data_ok, comp_ok])
    rotulo_contagem.configure(text=f"{count} de 3")
    if count == 3:
        rotulo_pronto_titulo.configure(text="Tudo pronto para começar")
        rotulo_pronto_dica.configure(text="Revise as escolhas e inicie a aprovação.")
        botao_acao.configure(state="normal", bg=ACTION, fg="white",
                             activebackground=ACTION_HV, activeforeground="white")
    else:
        rotulo_pronto_titulo.configure(text="Aguardando configurações")
        rotulo_pronto_dica.configure(text="Complete as etapas acima para continuar.")
        botao_acao.configure(state="disabled", bg="#d8dad9", fg="#92999b")


def iniciar_processamento():
    global processamento_ativo, concluidos_painel, total_painel

    caminho = _caminho_planilha["valor"]
    data_pagamento = entrada_data.get().strip()
    competencia = entrada_competencia.get().strip()

    if not caminho:
        messagebox.showwarning("Campo obrigatório", "Selecione a planilha.", parent=janela)
        return
    if not data_pagamento:
        messagebox.showwarning("Campo obrigatório", "Preencha a Data de Pagamento (DD/MM/AAAA).", parent=janela)
        return
    if not competencia:
        messagebox.showwarning("Campo obrigatório", "Preencha a Competência de Pagamento (MM/AAAA).", parent=janela)
        return

    concluidos_painel = 0
    total_painel = 0
    processamento_ativo = True
    botao_acao.configure(state="disabled", bg="#d8dad9", fg="#92999b", text="◌  Processando...")
    rotulo_pronto_titulo.configure(text="Aprovação em andamento")
    rotulo_pronto_dica.configure(text="Aguarde: os pedidos estão sendo processados no Conveniar.")

    # mostra área de progresso e reduz janela de configuração
    corpo.pack_forget()
    frame_progresso.pack(fill="both", expand=True, pady=(0, 0))

    threading.Thread(
        target=processar_planilha,
        args=(caminho, data_pagamento, competencia),
        daemon=True,
    ).start()


def atualizar_painel():
    global processamento_ativo, concluidos_painel, total_painel
    try:
        while True:
            tipo, *dados = eventos_painel.get_nowait()
            if tipo == "total":
                total_painel = dados[0]
                rotulo_progresso.configure(text=f"0/{total_painel} pedidos")
            elif tipo == "pedidos":
                lista_pedidos.delete(*lista_pedidos.get_children())
                for indice, pedido in dados[0]:
                    lista_pedidos.insert("", "end", iid=str(indice), values=(pedido, "Pendente"))
            elif tipo == "status":
                rotulo_status.configure(text=dados[0])
            elif tipo == "pedido":
                indice, pedido, nav_id = dados
                rotulo_status.configure(text=f"Navegador {nav_id}: {pedido}")
                if lista_pedidos.exists(str(indice)):
                    lista_pedidos.set(str(indice), "status", f"Processando (N{nav_id})")
                    lista_pedidos.see(str(indice))
            elif tipo == "concluido":
                indice, pedido, estado, nav_id = dados
                concluidos_painel += 1
                progresso.set(round(concluidos_painel * 100 / max(total_painel, 1)))
                rotulo_progresso.configure(text=f"{concluidos_painel}/{total_painel} pedidos")
                if lista_pedidos.exists(str(indice)):
                    lista_pedidos.set(str(indice), "status", estado)
            elif tipo == "aviso":
                rotulo_status.configure(text=dados[0])
                messagebox.showwarning("Planilha vazia", dados[0], parent=janela)
                processamento_ativo = False
                botao_acao.configure(state="normal", bg=ACTION, fg="white", text="INICIAR APROVAÇÃO  →")
            elif tipo in ("fim", "erro"):
                rotulo_status.configure(text=dados[0])
                processamento_ativo = False
                rotulo_pronto_titulo.configure(text="Concluído")
                rotulo_pronto_dica.configure(text="Selecione nova planilha para processar novamente.")
                botao_acao.configure(state="normal", bg=ACTION, fg="white",
                                     text="NOVA APROVAÇÃO  →",
                                     command=_reiniciar_interface)
                if tipo == "fim":
                    messagebox.showinfo("Finalizado", dados[0], parent=janela)
                else:
                    messagebox.showerror("Erro", dados[0], parent=janela)
    except queue.Empty:
        pass
    janela.after(100, atualizar_painel)


def _reiniciar_interface():
    """Volta para a tela de configuração após o processamento."""
    global processamento_ativo
    processamento_ativo = False
    _caminho_planilha["valor"] = ""
    rotulo_planilha.configure(text="Nenhuma selecionada")
    btn_planilha.configure(text="Selecionar planilha")
    progresso.set(0)
    lista_pedidos.delete(*lista_pedidos.get_children())
    frame_progresso.pack_forget()
    corpo.pack(fill="both", expand=True, padx=48, pady=24)
    botao_acao.configure(text="INICIAR APROVAÇÃO  →", command=iniciar_processamento)
    _verificar_pronto()


def ao_fechar():
    if processamento_ativo:
        messagebox.showwarning(
            "Processamento em andamento",
            "Aguarde o término antes de fechar o painel.",
            parent=janela,
        )
        return
    janela.destroy()


botao_acao.configure(command=iniciar_processamento)
janela.protocol("WM_DELETE_WINDOW", ao_fechar)

# ---------------------------------------------------------------------------
# Verificação de atualização em background
# ---------------------------------------------------------------------------

def _verificar_atualizacao():
    if not auto_update:
        return
    if not auto_update.REPO_GITHUB:
        janela.after(0, lambda: (
            _lbl_update_dot.configure(fg=auto_update.COR_UNKNOWN),
            _lbl_update_txt.configure(text=""),
        ))
        return
    try:
        release = auto_update.buscar_release(auto_update.REPO_GITHUB)
        if auto_update.tem_versao_nova(release, __version__):
            nova = release.get("tag_name", "?")

            def _iniciar_update():
                _lbl_update_dot.configure(fg=auto_update.COR_BUSY)
                _lbl_update_txt.configure(text=f"Atualizando para {nova}...")
                if auto_update.pode_instalar():
                    threading.Thread(target=_baixar_e_instalar,
                                     args=(release,), daemon=True).start()

            janela.after(0, _iniciar_update)
        else:
            janela.after(0, lambda: (
                _lbl_update_dot.configure(fg=auto_update.COR_OK),
                _lbl_update_txt.configure(text="Atualizado"),
            ))
    except Exception:
        janela.after(0, lambda: (
            _lbl_update_dot.configure(fg=auto_update.COR_UNKNOWN),
            _lbl_update_txt.configure(text=""),
        ))


def _baixar_e_instalar(release):
    try:
        janela.after(0, lambda: _lbl_update_txt.configure(text="Baixando atualização..."))
        auto_update.instalar(release, auto_update.NOME_EXE)
        janela.after(0, lambda: (
            _lbl_update_txt.configure(text="Reiniciando..."),
            janela.after(1500, janela.destroy),
        ))
    except Exception as e:
        janela.after(0, lambda: (
            _lbl_update_dot.configure(fg=auto_update.COR_UNKNOWN),
            _lbl_update_txt.configure(text="Falha na atualização"),
            messagebox.showerror("Erro na atualização", str(e), parent=janela),
        ))


threading.Thread(target=_verificar_atualizacao, daemon=True).start()

if __name__ == "__main__":
    atualizar_painel()
    janela.mainloop()
