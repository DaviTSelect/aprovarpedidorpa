"""Sessões independentes e interações verificadas com o Conveniar."""
import os
import tempfile
import time
from pathlib import Path

from selenium import webdriver
from selenium.common.exceptions import (
    ElementClickInterceptedException, ElementNotInteractableException,
    NoSuchElementException, StaleElementReferenceException,
)
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select, WebDriverWait

URL_CONSULTA = "https://facto.conveniar.com.br/Fundacao/Forms/Convenio/ConsultaPedidoPagamento.aspx"
FILTROS = "/html/body/div[3]/form/div[5]/div/div/div[2]/div/div[1]/fieldset/div/div/ul"
TIMEOUT = float(os.getenv("CONVENIAR_TIMEOUT", "30"))
if TIMEOUT <= 0:
    raise ValueError("CONVENIAR_TIMEOUT deve ser maior que zero.")


def xpath(texto):
    return By.XPATH, texto


def ident(nome):
    return By.ID, nome


class Tela:
    def __init__(self, driver, timeout=TIMEOUT):
        self.driver = driver
        self.timeout = timeout

    def esperar(self, condicao, mensagem="Condição esperada não apareceu"):
        return WebDriverWait(
            self.driver, self.timeout, poll_frequency=0.15,
            ignored_exceptions=(NoSuchElementException, StaleElementReferenceException),
        ).until(condicao, mensagem)

    def pronta(self):
        """Exige um intervalo ocioso, incluindo postbacks ASP.NET atrasados."""
        inicio_ocioso = None

        def ociosa(driver):
            nonlocal inicio_ocioso
            pronto = driver.execute_script("""
                const visivel = e => !!(e.getClientRects().length) &&
                    getComputedStyle(e).visibility !== 'hidden' &&
                    getComputedStyle(e).display !== 'none';
                const pm = window.Sys && Sys.WebForms &&
                    Sys.WebForms.PageRequestManager.getInstance();
                return document.readyState === 'complete' &&
                    !(pm && pm.get_isInAsyncPostBack()) &&
                    !(window.jQuery && jQuery.active > 0) &&
                    !Array.from(document.querySelectorAll('#imageLoader')).some(visivel);
            """)
            if not pronto:
                inicio_ocioso = None
                return False
            if inicio_ocioso is None:
                inicio_ocioso = time.monotonic()
            return time.monotonic() - inicio_ocioso >= 0.35

        self.esperar(ociosa, "Página ou loader não terminou de carregar")

    def visiveis(self, localizador):
        return [e for e in self.driver.find_elements(*localizador) if e.is_displayed()]

    def localizar(self, localizador, editavel=False):
        elementos = self.visiveis(localizador)
        if len(elementos) > 1:
            raise RuntimeError(f"Seletor ambíguo: {localizador}")
        if not elementos:
            return False
        elemento = elementos[0]
        if editavel and (not elemento.is_enabled() or elemento.get_attribute("readonly")):
            return False
        return elemento

    def elemento(self, localizador, editavel=False):
        return self.esperar(lambda d: self.localizar(localizador, editavel),
                            f"Elemento indisponível: {localizador}")

    def agir(self, localizador, acao):
        """Relocaliza apenas se a interação não foi enviada ao elemento."""
        self.pronta()

        def tentar(driver):
            try:
                elemento = self.localizar(localizador, editavel=True)
                if not elemento:
                    return False
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elemento)
                acao(elemento)
                return True
            except (StaleElementReferenceException, ElementClickInterceptedException,
                    ElementNotInteractableException):
                return False

        self.esperar(tentar, f"Não foi possível interagir: {localizador}")
        # Fora da repetição: timeout após envio nunca repete a ação de negócio.
        self.pronta()

    def clicar(self, localizador):
        self.agir(localizador, lambda e: e.click())

    def preencher(self, localizador, valor, tab=True):
        # Cada comando tem sua própria relocalização. Um TAB pode gerar postback.
        self.agir(localizador, lambda e: e.send_keys(Keys.CONTROL, "a"))
        self.agir(localizador, lambda e: e.send_keys(Keys.BACKSPACE))
        if valor:
            self.agir(localizador, lambda e: e.send_keys(str(valor)))
        if tab:
            self.agir(localizador, lambda e: e.send_keys(Keys.TAB))
        self.esperar(lambda d: self.elemento(localizador).get_attribute("value") == str(valor),
                     f"Campo não preservou o valor esperado: {localizador}")

    def marcar(self, localizador, marcado=True):
        self.pronta()
        if self.elemento(localizador, editavel=True).is_selected() != marcado:
            self.clicar(localizador)
        self.esperar(lambda d: self.elemento(localizador).is_selected() == marcado,
                     f"Checkbox não preservou a seleção: {localizador}")

    def selecionar(self, localizador, valor):
        self.agir(localizador, lambda e: Select(e).select_by_value(valor))
        self.esperar(lambda d: self.elemento(localizador).get_attribute("value") == valor)

    def desaparecer(self, localizador):
        self.esperar(lambda d: not self.visiveis(localizador), f"Painel permaneceu aberto: {localizador}")


def criar_opcoes_chrome(perfil):
    options = Options()
    if os.getenv("CONVENIAR_HEADLESS", "0") == "1":
        options.add_argument("--headless=new")
    options.add_argument("--window-size=1440,1000")
    options.add_argument("--disable-extensions")
    options.add_argument("--log-level=3")
    options.add_argument(f"--user-data-dir={perfil}")
    chrome = Path.home() / "Chrome/GoogleChrome/App/Chrome-bin/chrome.exe"
    if chrome.is_file():
        options.binary_location = str(chrome)
    return options


def encerrar_sessao(driver):
    perfil = getattr(driver, "_perfil_conveniar", None)
    try:
        driver.quit()
    finally:
        if perfil:
            perfil.cleanup()


def iniciar_sessao(usuario, senha):
    perfil = tempfile.TemporaryDirectory(prefix="conveniar-")
    driver = None
    try:
        caminho = Path.home() / "Chrome/chromedriver.exe"
        service = Service(executable_path=str(caminho)) if caminho.is_file() else Service()
        driver = webdriver.Chrome(service=service, options=criar_opcoes_chrome(perfil.name))
        driver._perfil_conveniar = perfil
        driver.set_page_load_timeout(TIMEOUT * 2)
        driver.set_script_timeout(TIMEOUT)
        driver.implicitly_wait(0)
        driver.get(URL_CONSULTA)
        tela = Tela(driver)
        tela.preencher((By.CSS_SELECTOR, "input[type='text']"), usuario, tab=False)
        tela.preencher((By.CSS_SELECTOR, "input[type='password']"), senha, tab=False)
        tela.clicar(xpath("/html/body/form/main/div/div[2]/div/table/tbody/tr/td/div/div[3]/input"))
        tela.elemento(xpath(FILTROS + "/li[3]/input"), editavel=True)
        tela.pronta()
        return driver
    except Exception:
        if driver is not None:
            try:
                encerrar_sessao(driver)
            except Exception:
                pass  # Preserva a exceção original do login.
        else:
            perfil.cleanup()
        raise
