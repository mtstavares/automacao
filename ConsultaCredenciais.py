# === IMPORTACOES ===
import json
import logging
import math
import os
import re
import subprocess
import sys
import tempfile
import time
import unicodedata
import urllib3
from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from logging.handlers import RotatingFileHandler
from urllib.parse import urlparse

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from selenium import webdriver
from selenium.common.exceptions import (
    NoSuchElementException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

script_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
arquivo_excel = os.path.join(script_dir, "Credenciais.xlsx")

LOGIN_URL_MS = "http://ms.policiamilitar.sp.gov.br/login.aspx"
LOGIN_URL_AD = "http://www.rh.intranet.policiamilitar.sp.gov.br"
SUCCESS_URL_AD = "https://sgp-prod.intranet.policiamilitar.sp.gov.br/Portal/Portal"

COL_EMAIL_INFORMADO = "A"
COL_SENHA = "B"
COL_DOCUMENTO_INFORMADO = "C"
COL_CPF_RESOLVIDO = "D"
COL_NOME = "E"
COL_EMAIL_FUNCIONAL = "F"
COL_RESULTADO_MS = "G"
COL_MSG_MS = "H"
COL_RESULTADO_AD = "I"
COL_MSG_AD = "J"
COL_ORIGEM_CPF = "K"
COL_STATUS_PROCESSAMENTO = "L"
COL_DATA_TESTE = "M"
COL_URL_VAZAMENTO = "N"

HEADER_JA_IDENTIFICADO_MES_ANTIGO = "ja identificado neste mes"
HEADER_CONTAGEM_IDENTIFICACOES = "Identificado quantas vezes"

HEADERS = [
    "E-mail informado",
    "Senha de entrada",
    "CPF ou RE informado",
    "CPF resolvido",
    "Nome completo",
    "E-mail funcional",
    "Resultado MS",
    "Mensagem MS",
    "Resultado AD",
    "Mensagem AD",
    "Origem do CPF",
    "Status do processamento",
    "Data e hora do teste",
    "URL de vazamento",
]

HEADERS_SAIDA = [
    "Nome",
    "E-mail",
    "CPF",
    "Senha",
    "MS",
    "AD",
    "Mensagem do MS",
    "Mensagem do AD",
    "URL de vazamento",
]

HEADERS_MENSAL = [
    "Nome",
    "E-mail",
    "CPF",
    "Senha",
    "MS",
    "AD",
    "Mensagem do MS",
    "Mensagem do AD",
    HEADER_CONTAGEM_IDENTIFICACOES,
    "Primeira identificação",
    "Data desta identificação",
    "Dia da semana",
    "URL de vazamento",
]

MESES_PT = {
    1: "JAN",
    2: "FEV",
    3: "MAR",
    4: "ABR",
    5: "MAI",
    6: "JUN",
    7: "JUL",
    8: "AGO",
    9: "SET",
    10: "OUT",
    11: "NOV",
    12: "DEZ",
}

DIAS_SEMANA_PT = {
    0: "Segunda-feira",
    1: "Terça-feira",
    2: "Quarta-feira",
    3: "Quinta-feira",
    4: "Sexta-feira",
    5: "Sábado",
    6: "Domingo",
}

RESULTADO_SIM = "SIM"
RESULTADO_NAO = "NÃO"
RESULTADO_ERRO = "ERRO"
RESULTADO_NT = "N/T"
RESULTADO_INCONCLUSIVO = "INCONCLUSIVO"

SUCESSO = "SUCESSO"
CREDENCIAL_INVALIDA = "CREDENCIAL_INVALIDA"
CONTA_BLOQUEADA = "CONTA_BLOQUEADA"
CONTA_DESATIVADA = "CONTA_DESATIVADA"
SENHA_EXPIRADA = "SENHA_EXPIRADA"
ALTERACAO_DE_SENHA_OBRIGATORIA = "ALTERACAO_DE_SENHA_OBRIGATORIA"
ACESSO_NEGADO = "ACESSO_NEGADO"
USUARIO_NAO_ENCONTRADO = "USUARIO_NAO_ENCONTRADO"
SISTEMA_INDISPONIVEL = "SISTEMA_INDISPONIVEL"
ERRO_DE_REDE = "ERRO_DE_REDE"
TIMEOUT = "TIMEOUT"
ELEMENTO_NAO_ENCONTRADO = "ELEMENTO_NAO_ENCONTRADO"
PAGINA_INESPERADA = "PAGINA_INESPERADA"
ERRO_SELENIUM = "ERRO_SELENIUM"
ERRO_INTERNO = "ERRO_INTERNO"
DADOS_AUSENTES = "DADOS_AUSENTES"
NAO_TESTADO = "NAO_TESTADO"
INCONCLUSIVO = "INCONCLUSIVO"

CPF_INFORMADO = "CPF_INFORMADO"
CPF_RECONSTRUIDO_ZERO_INICIAL = "CPF_RECONSTRUIDO_ZERO_INICIAL"
RE_API = "RE_API"
RE_RECONSTRUIDO_ZERO_INICIAL = "RE_RECONSTRUIDO_ZERO_INICIAL"
EMAIL_DSQUERY = "EMAIL_DSQUERY"
NAO_ENCONTRADO = "NAO_ENCONTRADO"
AMBIGUO = "AMBIGUO"

CONCLUIDO = "CONCLUIDO"
CONCLUIDO_COM_RESSALVA = "CONCLUIDO_COM_RESSALVA"
ERRO_DE_ENTRADA = "ERRO_DE_ENTRADA"
ERRO_TECNICO = "ERRO_TECNICO"

CLASSIFICACOES_FUNCIONAIS = {
    CREDENCIAL_INVALIDA,
    CONTA_BLOQUEADA,
    CONTA_DESATIVADA,
    SENHA_EXPIRADA,
    ALTERACAO_DE_SENHA_OBRIGATORIA,
    ACESSO_NEGADO,
    USUARIO_NAO_ENCONTRADO,
}

CLASSIFICACOES_TECNICAS = {
    SISTEMA_INDISPONIVEL,
    ERRO_DE_REDE,
    TIMEOUT,
    ELEMENTO_NAO_ENCONTRADO,
    PAGINA_INESPERADA,
    ERRO_SELENIUM,
    ERRO_INTERNO,
}


class SeletoresMS:
    CPF = [(By.NAME, "vUSRNUMCPFAUX")]
    SENHA = [(By.NAME, "vSENHA")]
    BOTAO_LOGIN = [(By.NAME, "BTN_LOGIN")]
    MENSAGEM_ERRO = [(By.CSS_SELECTOR, "span#gxErrorViewer > div")]


class SeletoresAD:
    DROPDOWN_LOGIN = [(By.CLASS_NAME, "dropdown-toggle")]
    CPF = [(By.NAME, "CPF")]
    SENHA = [(By.CLASS_NAME, "password")]
    BOTAO_AUTENTICAR = [(By.ID, "btnAutenticar")]
    MENSAGEM_AUTENTICACAO = [(By.ID, "msgAutenticacao")]


@dataclass
class ResultadoOperacao:
    status: str
    mensagem: str
    valor: object = None
    detalhe_tecnico: str = ""


@dataclass
class ResultadoAuth:
    classificacao: str
    resultado_planilha: str
    mensagem: str
    detalhe_tecnico: str = ""


@dataclass
class IdentificadorNormalizado:
    tipo: str
    valor: str
    origem: str
    mensagem: str = ""


class ContextFilter(logging.Filter):
    def __init__(self, execution_id):
        super().__init__()
        self.execution_id = execution_id

    def filter(self, record):
        defaults = {
            "execution_id": self.execution_id,
            "fase": "-",
            "linha": "-",
            "sistema": "-",
            "cpf": "-",
            "operacao": "-",
            "status": "-",
            "duracao": "-",
        }
        for key, value in defaults.items():
            if not hasattr(record, key):
                setattr(record, key, value)
        return True


def configurar_logger(nome, caminho, execution_id):
    logger = logging.getLogger(nome)
    logger.setLevel(logging.INFO)
    for handler in logger.handlers:
        handler.close()
    logger.handlers.clear()
    handler = logging.NullHandler()
    logger.addHandler(handler)
    logger.propagate = False
    return logger


def log_event(logger, level, fase, linha, sistema, cpf, operacao, status, duracao, mensagem, exc_info=False):
    extra = {
        "fase": fase,
        "linha": linha if linha is not None else "-",
        "sistema": sistema or "-",
        "cpf": mascarar_cpf(cpf),
        "operacao": operacao,
        "status": status,
        "duracao": f"{duracao:.3f}s" if isinstance(duracao, (int, float)) else "-",
    }
    logger.log(level, mensagem, extra=extra, exc_info=exc_info)


def adicionar_erro(resumo, chave, linha, mensagem):
    resumo.setdefault(chave, [])
    resumo[chave].append((linha, mensagem))


def texto_sem_acentos(valor):
    texto = "" if valor is None else str(valor)
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(ch for ch in texto if not unicodedata.combining(ch))


def normalizar_header(valor):
    return re.sub(r"[^a-z0-9]+", "", texto_sem_acentos(valor).lower())


def mascarar_cpf(cpf):
    digitos = re.sub(r"\D", "", str(cpf or ""))
    if len(digitos) != 11:
        return "-"
    return f"***.***.{digitos[6:9]}-{digitos[9:]}"


def limpar_texto_identificador(valor):
    if valor is None:
        return ""
    if isinstance(valor, int):
        return str(valor)
    if isinstance(valor, float):
        if math.isnan(valor) or math.isinf(valor):
            return ""
        if valor.is_integer() and abs(valor) < 10**15:
            return str(int(valor))
        return ""
    texto = str(valor).strip()
    texto = texto.replace("\u200b", "").replace("\ufeff", "").strip()
    if texto.startswith("'"):
        texto = texto[1:].strip()
    if texto.startswith("="):
        return ""
    if re.fullmatch(r"[+-]?\d+(\.0+)?", texto):
        return texto.split(".")[0].lstrip("+")
    if re.fullmatch(r"[+-]?\d+(\.\d+)?[eE][+-]?\d+", texto):
        try:
            decimal = Decimal(texto)
            if decimal == decimal.to_integral_value():
                return str(int(decimal))
        except (InvalidOperation, ValueError):
            return ""
    return re.sub(r"\D", "", texto)


def cpf_valido(cpf):
    cpf = re.sub(r"\D", "", str(cpf or ""))
    if len(cpf) != 11:
        return False
    if cpf == cpf[0] * 11:
        return False
    soma = sum(int(cpf[i]) * (10 - i) for i in range(9))
    dig1 = (soma * 10) % 11
    dig1 = 0 if dig1 == 10 else dig1
    soma = sum(int(cpf[i]) * (11 - i) for i in range(10))
    dig2 = (soma * 10) % 11
    dig2 = 0 if dig2 == 10 else dig2
    return dig1 == int(cpf[9]) and dig2 == int(cpf[10])


def normalizar_cpf(valor):
    digitos = limpar_texto_identificador(valor)
    if not digitos:
        return ResultadoOperacao(DADOS_AUSENTES, "CPF vazio.")
    if len(digitos) == 11 and cpf_valido(digitos):
        return ResultadoOperacao(SUCESSO, "CPF valido.", digitos)
    if len(digitos) == 10:
        reconstruido = "0" + digitos
        if cpf_valido(reconstruido):
            return ResultadoOperacao(SUCESSO, "CPF reconstruido com zero inicial.", reconstruido, "zero_inicial")
    if len(digitos) == 11:
        return ResultadoOperacao(ERRO_DE_ENTRADA, "CPF com digitos verificadores invalidos.")
    return ResultadoOperacao(ERRO_DE_ENTRADA, "CPF com quantidade incorreta de digitos.")


def normalizar_re(valor):
    digitos = limpar_texto_identificador(valor)
    if not digitos:
        return ResultadoOperacao(DADOS_AUSENTES, "RE vazio.")
    if len(digitos) == 6:
        return ResultadoOperacao(SUCESSO, "RE valido.", digitos)
    if len(digitos) == 5:
        return ResultadoOperacao(SUCESSO, "RE reconstruido com zero inicial.", "0" + digitos, "zero_inicial")
    return ResultadoOperacao(ERRO_DE_ENTRADA, "RE com quantidade incorreta de digitos.")


def normalizar_identificador(valor):
    digitos = limpar_texto_identificador(valor)
    if not digitos:
        return ResultadoOperacao(DADOS_AUSENTES, "CPF/RE ausente.")
    if len(digitos) in (10, 11):
        resultado_cpf = normalizar_cpf(valor)
        if resultado_cpf.status == SUCESSO:
            origem = CPF_RECONSTRUIDO_ZERO_INICIAL if resultado_cpf.detalhe_tecnico == "zero_inicial" else CPF_INFORMADO
            return ResultadoOperacao(SUCESSO, resultado_cpf.mensagem, IdentificadorNormalizado("CPF", resultado_cpf.valor, origem))
        return resultado_cpf
    if len(digitos) in (5, 6):
        resultado_re = normalizar_re(valor)
        if resultado_re.status == SUCESSO:
            origem = RE_RECONSTRUIDO_ZERO_INICIAL if resultado_re.detalhe_tecnico == "zero_inicial" else RE_API
            return ResultadoOperacao(SUCESSO, resultado_re.mensagem, IdentificadorNormalizado("RE", resultado_re.valor, origem))
        return resultado_re
    return ResultadoOperacao(ERRO_DE_ENTRADA, "Identificador nao e CPF nem RE reconhecido.")


class ClienteAPIsInternas:
    def __init__(self, logger):
        self.session = requests.Session()
        self.logger = logger
        self.timeout = (5, 15)
        self.max_tentativas = 3

    def fechar(self):
        self.session.close()

    def _get_json(self, url, operacao, linha=None, cpf=None):
        inicio = time.perf_counter()
        for tentativa in range(1, self.max_tentativas + 1):
            try:
                response = self.session.get(url, verify=False, timeout=self.timeout)
                if response.status_code in (429, 500, 502, 503, 504):
                    if tentativa < self.max_tentativas:
                        time.sleep(0.5 * tentativa)
                        continue
                    return ResultadoOperacao(SISTEMA_INDISPONIVEL, "API indisponivel.", detalhe_tecnico=f"HTTP {response.status_code}")
                if response.status_code in (401, 403):
                    return ResultadoOperacao(ACESSO_NEGADO, "Acesso negado pela API.", detalhe_tecnico=f"HTTP {response.status_code}")
                if response.status_code == 404:
                    return ResultadoOperacao(NAO_ENCONTRADO, "Registro nao encontrado pela API.", detalhe_tecnico="HTTP 404")
                response.raise_for_status()
                if not response.content:
                    return ResultadoOperacao(INCONCLUSIVO, "Resposta vazia da API.")
                try:
                    return ResultadoOperacao(SUCESSO, "Consulta realizada.", response.json())
                except ValueError as exc:
                    return ResultadoOperacao(ERRO_INTERNO, "JSON invalido retornado pela API.", detalhe_tecnico=str(exc))
            except requests.Timeout as exc:
                if tentativa < self.max_tentativas:
                    time.sleep(0.5 * tentativa)
                    continue
                return ResultadoOperacao(TIMEOUT, "Tempo limite excedido consultando API.", detalhe_tecnico=str(exc))
            except requests.ConnectionError as exc:
                if tentativa < self.max_tentativas:
                    time.sleep(0.5 * tentativa)
                    continue
                return ResultadoOperacao(ERRO_DE_REDE, "Erro de rede consultando API.", detalhe_tecnico=str(exc))
            except requests.RequestException as exc:
                return ResultadoOperacao(ERRO_DE_REDE, "Erro HTTP consultando API.", detalhe_tecnico=str(exc))
            finally:
                duracao = time.perf_counter() - inicio
                log_event(self.logger, logging.INFO, "API", linha, "API", cpf, operacao, "tentativa", duracao, f"Tentativa {tentativa}")
        return ResultadoOperacao(ERRO_INTERNO, "Erro inesperado na consulta da API.")

    def buscar_cpf_por_re(self, re_militar, linha=None):
        url = f"https://webservices.intranet.policiamilitar.sp.gov.br/pmesp.cdpm/api/v1/PolicialMilitar/re/{re_militar}/dadosResumidos"
        resposta = self._get_json(url, "buscar_cpf_por_re", linha=linha)
        if resposta.status != SUCESSO:
            return resposta
        dados = resposta.valor.get("dados") if isinstance(resposta.valor, dict) else None
        if not isinstance(dados, list) or not dados:
            return ResultadoOperacao(NAO_ENCONTRADO, "CPF nao encontrado pelo RE.")
        cpf = dados[0].get("cpf", {}).get("cpfComDigito")
        cpf_normalizado = normalizar_cpf(cpf)
        if cpf_normalizado.status != SUCESSO:
            return ResultadoOperacao(INCONCLUSIVO, "API retornou CPF invalido para o RE.", detalhe_tecnico=cpf_normalizado.mensagem)
        return ResultadoOperacao(SUCESSO, "CPF encontrado pelo RE.", cpf_normalizado.valor)

    def buscar_nome_por_cpf(self, cpf, linha=None):
        url = f"https://webservices.intranet.policiamilitar.sp.gov.br/pmesp.cdpm/api/v1/PolicialMilitar/cpf/{cpf}/dadosResumidos"
        resposta = self._get_json(url, "buscar_nome_por_cpf", linha=linha, cpf=cpf)
        if resposta.status != SUCESSO:
            return resposta
        dados = resposta.valor.get("dados") if isinstance(resposta.valor, dict) else None
        if not isinstance(dados, list) or not dados:
            return ResultadoOperacao(NAO_ENCONTRADO, "Nome nao encontrado pelo CPF.")
        nome = dados[0].get("nomeCompleto")
        if not nome:
            return ResultadoOperacao(NAO_ENCONTRADO, "Nome ausente na resposta da API.")
        return ResultadoOperacao(SUCESSO, "Nome encontrado.", str(nome).strip())

    def buscar_email_por_cpf(self, cpf, linha=None):
        url = f"https://webservices.intranet.policiamilitar.sp.gov.br/pmesp.cdpm/api/v1/PolicialMilitar/cpf/{cpf}/informacaoContato"
        resposta = self._get_json(url, "buscar_email_por_cpf", linha=linha, cpf=cpf)
        if resposta.status != SUCESSO:
            return resposta
        dados = resposta.valor.get("dados") if isinstance(resposta.valor, dict) else None
        if not isinstance(dados, list) or not dados:
            return ResultadoOperacao(NAO_ENCONTRADO, "E-mail funcional nao encontrado pelo CPF.")
        for email_info in dados[0].get("emails", []):
            if email_info.get("tipoContato", {}).get("identificador") == 4:
                endereco = email_info.get("endereco")
                if endereco:
                    return ResultadoOperacao(SUCESSO, "E-mail funcional encontrado.", str(endereco).strip())
        return ResultadoOperacao(NAO_ENCONTRADO, "E-mail funcional nao encontrado pelo CPF.")


def buscar_re_por_email(email, logger, linha=None):
    inicio = time.perf_counter()
    dsquery_cmd = ["dsquery", "*", "-filter", f"(mail={email.strip()})", "-attr", "employeeNumber"]
    try:
        cmd = subprocess.run(dsquery_cmd, capture_output=True, text=True, timeout=20)
        match = re.search(r"employeeNumber\s+(\d+)", cmd.stdout)
        if match:
            log_event(logger, logging.INFO, "IDENTIDADE", linha, "DSQUERY", None, "pesquisa_email", SUCESSO, time.perf_counter() - inicio, "RE encontrado por e-mail.")
            return ResultadoOperacao(SUCESSO, "RE encontrado por e-mail.", match.group(1))
        return ResultadoOperacao(NAO_ENCONTRADO, "RE nao encontrado pelo e-mail.")
    except FileNotFoundError as exc:
        return ResultadoOperacao(ERRO_DE_REDE, "dsquery nao encontrado neste ambiente.", detalhe_tecnico=str(exc))
    except subprocess.TimeoutExpired as exc:
        return ResultadoOperacao(TIMEOUT, "Tempo limite excedido consultando dsquery.", detalhe_tecnico=str(exc))
    except Exception as exc:
        return ResultadoOperacao(ERRO_INTERNO, "Erro ao consultar dsquery.", detalhe_tecnico=str(exc))


def validar_arquivo_entrada(caminho):
    if not os.path.exists(caminho):
        raise FileNotFoundError(f"Arquivo de entrada nao encontrado: {caminho}")
    if not caminho.lower().endswith(".xlsx"):
        raise ValueError("Arquivo de entrada precisa ter extensao .xlsx.")
    if os.path.getsize(caminho) == 0:
        raise ValueError("Arquivo de entrada esta vazio.")
    diretorio = os.path.dirname(caminho)
    if not os.access(caminho, os.R_OK):
        raise PermissionError("Sem permissao de leitura no arquivo de entrada.")
    if not os.access(diretorio, os.W_OK):
        raise PermissionError("Sem permissao de escrita no diretorio de saida.")
    teste_path = None
    try:
        fd, teste_path = tempfile.mkstemp(prefix=".credenciais_write_test_", suffix=".tmp", dir=diretorio)
        os.close(fd)
    finally:
        if teste_path and os.path.exists(teste_path):
            os.remove(teste_path)
    wb = load_workbook(caminho)
    ws = wb.active
    if ws.max_row < 2:
        raise ValueError("Planilha sem linhas de dados.")
    if ws.max_column < 3:
        raise ValueError("Planilha precisa ter ao menos as colunas A, B e C.")
    headers = [normalizar_header(ws.cell(1, col).value) for col in range(1, min(ws.max_column, len(HEADERS)) + 1)]
    layout_antigo = headers[:3] == ["email", "senha", "cpf"]
    headers_esperados = [normalizar_header(h) for h in HEADERS]
    layout_novo = headers[:len(HEADERS)] == headers_esperados or headers[:len(HEADERS) - 1] == headers_esperados[:-1]
    if not layout_antigo and not layout_novo:
        raise ValueError("Layout da planilha nao e compativel com o esperado.")
    return wb, ws, layout_novo


def localizar_coluna_url_vazamento(ws):
    header_url = normalizar_header("URL de vazamento")
    for col in range(1, ws.max_column + 1):
        if normalizar_header(ws.cell(1, col).value) == header_url:
            return col
    return None


def preparar_layout(ws, layout_novo):
    if layout_novo:
        if normalizar_header(ws[f"{COL_URL_VAZAMENTO}1"].value) != normalizar_header("URL de vazamento"):
            ws[f"{COL_URL_VAZAMENTO}1"] = "URL de vazamento"
        return
    coluna_url = localizar_coluna_url_vazamento(ws)
    linhas = []
    for row in range(2, ws.max_row + 1):
        linhas.append((
            ws[f"A{row}"].value,
            ws[f"B{row}"].value,
            ws[f"C{row}"].value,
            ws.cell(row, coluna_url).value if coluna_url else None,
        ))
    for col, header in enumerate(HEADERS, start=1):
        cell = ws.cell(1, col)
        cell.value = header
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
    for offset, (email, senha, documento, url_vazamento) in enumerate(linhas, start=2):
        ws[f"{COL_EMAIL_INFORMADO}{offset}"] = email
        ws[f"{COL_SENHA}{offset}"] = senha
        ws[f"{COL_DOCUMENTO_INFORMADO}{offset}"] = documento
        ws[f"{COL_URL_VAZAMENTO}{offset}"] = url_vazamento
        for col in range(4, 14):
            ws.cell(offset, col).value = None


def aplicar_estilos(ws):
    success_fill = PatternFill("solid", fgColor="C6EFCE")
    failure_fill = PatternFill("solid", fgColor="FFC7CE")
    error_fill = PatternFill("solid", fgColor="FFEB9C")
    inconclusive_fill = PatternFill("solid", fgColor="D9EAF7")
    nt_fill = PatternFill("solid", fgColor="D9D9D9")
    fills = {
        RESULTADO_SIM: success_fill,
        RESULTADO_NAO: failure_fill,
        RESULTADO_ERRO: error_fill,
        RESULTADO_INCONCLUSIVO: inconclusive_fill,
        RESULTADO_NT: nt_fill,
    }
    for row in range(2, ws.max_row + 1):
        for col in [COL_RESULTADO_MS, COL_RESULTADO_AD]:
            cell = ws[f"{col}{row}"]
            if cell.value in fills:
                cell.fill = fills[cell.value]
            cell.alignment = Alignment(horizontal="center")


def aplicar_estilos_saida(ws):
    success_fill = PatternFill("solid", fgColor="C6EFCE")
    failure_fill = PatternFill("solid", fgColor="FFC7CE")
    for col in range(1, len(HEADERS_SAIDA) + 1):
        cell = ws.cell(1, col)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
    for row in range(2, ws.max_row + 1):
        for col in (5, 6):
            cell = ws.cell(row, col)
            if cell.value == RESULTADO_SIM:
                cell.fill = success_fill
            elif cell.value == RESULTADO_NAO:
                cell.fill = failure_fill
            cell.alignment = Alignment(horizontal="center")
    larguras = {
        "A": 34,
        "B": 38,
        "C": 16,
        "D": 28,
        "E": 12,
        "F": 12,
        "G": 58,
        "H": 58,
        "I": 58,
    }
    for coluna, largura in larguras.items():
        ws.column_dimensions[coluna].width = largura


def criar_workbook_saida(ws_processado):
    wb_saida = Workbook()
    ws_saida = wb_saida.active
    ws_saida.title = "Resultado"
    for col, header in enumerate(HEADERS_SAIDA, start=1):
        ws_saida.cell(1, col).value = header
    for row in range(2, ws_processado.max_row + 1):
        linha_saida = row
        ws_saida.cell(linha_saida, 1).value = ws_processado[f"{COL_NOME}{row}"].value
        ws_saida.cell(linha_saida, 2).value = ws_processado[f"{COL_EMAIL_FUNCIONAL}{row}"].value
        ws_saida.cell(linha_saida, 3).value = ws_processado[f"{COL_CPF_RESOLVIDO}{row}"].value
        ws_saida.cell(linha_saida, 4).value = ws_processado[f"{COL_SENHA}{row}"].value
        ws_saida.cell(linha_saida, 5).value = ws_processado[f"{COL_RESULTADO_MS}{row}"].value
        ws_saida.cell(linha_saida, 6).value = ws_processado[f"{COL_RESULTADO_AD}{row}"].value
        ws_saida.cell(linha_saida, 7).value = ws_processado[f"{COL_MSG_MS}{row}"].value
        ws_saida.cell(linha_saida, 8).value = ws_processado[f"{COL_MSG_AD}{row}"].value
        ws_saida.cell(linha_saida, 9).value = ws_processado[f"{COL_URL_VAZAMENTO}{row}"].value
    aplicar_estilos_saida(ws_saida)
    return wb_saida


def resultado_planilha_por_classificacao(classificacao):
    if classificacao == SUCESSO:
        return RESULTADO_SIM
    if classificacao in CLASSIFICACOES_FUNCIONAIS:
        return RESULTADO_NAO
    if classificacao in CLASSIFICACOES_TECNICAS:
        return RESULTADO_ERRO
    if classificacao in (DADOS_AUSENTES, NAO_TESTADO):
        return RESULTADO_NT
    return RESULTADO_INCONCLUSIVO


def classificar_mensagem_autenticacao(mensagem):
    texto = texto_sem_acentos(mensagem).lower()
    if not texto:
        return INCONCLUSIVO
    if "senha invalida" in texto or "usuario ou senha incorreta" in texto or "usuario e senha incorreta" in texto:
        return CREDENCIAL_INVALIDA
    if "bloquead" in texto:
        return CONTA_BLOQUEADA
    if "inativo" in texto or "desativad" in texto:
        return CONTA_DESATIVADA
    if "senha expirada" in texto:
        return SENHA_EXPIRADA
    if "alteracao de senha" in texto or "troca de senha" in texto:
        return ALTERACAO_DE_SENHA_OBRIGATORIA
    if "nao cadastrado" in texto or "cpf invalido" in texto or "usuario nao encontrado" in texto:
        return USUARIO_NAO_ENCONTRADO
    if "acesso negado" in texto:
        return ACESSO_NEGADO
    return INCONCLUSIVO


def mensagem_para_planilha(classificacao, mensagem_original=""):
    if mensagem_original and classificacao in CLASSIFICACOES_FUNCIONAIS:
        return str(mensagem_original).strip()
    mensagens = {
        SUCESSO: "Autenticacao realizada com sucesso.",
        CREDENCIAL_INVALIDA: "Credencial invalida confirmada pela mensagem da pagina.",
        CONTA_BLOQUEADA: "Conta bloqueada.",
        CONTA_DESATIVADA: "Conta desativada ou inativa.",
        SENHA_EXPIRADA: "Senha expirada.",
        ALTERACAO_DE_SENHA_OBRIGATORIA: "Alteracao de senha obrigatoria.",
        ACESSO_NEGADO: "Acesso negado.",
        USUARIO_NAO_ENCONTRADO: "Usuario nao encontrado ou CPF invalido.",
        SISTEMA_INDISPONIVEL: "Sistema indisponivel.",
        ERRO_DE_REDE: "Erro de rede.",
        TIMEOUT: "Tempo limite excedido aguardando resposta.",
        ELEMENTO_NAO_ENCONTRADO: "Elemento esperado nao localizado.",
        PAGINA_INESPERADA: "Pagina retornada nao corresponde ao fluxo esperado.",
        ERRO_SELENIUM: "Erro tecnico do Selenium.",
        ERRO_INTERNO: "Erro interno durante o processamento.",
        DADOS_AUSENTES: "CPF ou senha ausente.",
        NAO_TESTADO: "Nao testado.",
        INCONCLUSIVO: "Nao foi possivel confirmar sucesso ou falha.",
    }
    return mensagens.get(classificacao, mensagens[INCONCLUSIVO])


def aguardar_primeiro_elemento_valido(driver, seletores, timeout=10, clicavel=False):
    ultimo_erro = None
    for seletor in seletores:
        try:
            condicao = EC.element_to_be_clickable(seletor) if clicavel else EC.presence_of_element_located(seletor)
            return WebDriverWait(driver, timeout).until(condicao)
        except TimeoutException as exc:
            ultimo_erro = exc
    raise TimeoutException(str(ultimo_erro) if ultimo_erro else "Elemento nao localizado.")


def elemento_existe(driver, seletores):
    for by, value in seletores:
        try:
            if driver.find_elements(by, value):
                return True
        except WebDriverException:
            return False
    return False


def texto_visivel_primeiro_elemento(driver, seletores):
    for by, value in seletores:
        try:
            for elemento in driver.find_elements(by, value):
                if elemento.is_displayed():
                    texto = elemento.text.strip()
                    if texto:
                        return texto
        except WebDriverException:
            return ""
    return ""


def selecionar_janela_com_url(driver, trecho_url):
    for handle in driver.window_handles:
        driver.switch_to.window(handle)
        if trecho_url.lower() in (driver.current_url or "").lower():
            return True
    return False


def normalizar_url(url):
    parsed = urlparse(url or "")
    hostname = (parsed.hostname or "").lower()
    path = (parsed.path or "").rstrip("/").lower()
    return hostname, path


def url_corresponde(url_atual, url_esperada):
    host_atual, path_atual = normalizar_url(url_atual)
    host_esperado, path_esperado = normalizar_url(url_esperada)
    return host_atual == host_esperado and path_atual == path_esperado


def criar_driver():
    options = webdriver.ChromeOptions()
    chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    if os.path.exists(chrome_path):
        options.binary_location = chrome_path
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(30)
    return driver


def encerrar_driver(driver):
    if driver:
        try:
            driver.quit()
        except WebDriverException:
            pass


def validar_sessao_limpa(driver):
    try:
        cookies_limpos = len(driver.get_cookies()) == 0
        storage_limpo = driver.execute_script("return window.localStorage.length + window.sessionStorage.length;") == 0
        uma_janela = len(driver.window_handles) == 1
        return cookies_limpos and storage_limpo and uma_janela
    except WebDriverException:
        return False


def limpar_sessao_driver(driver):
    try:
        handles = driver.window_handles
        if not handles:
            return False
        principal = handles[0]
        for handle in handles[1:]:
            driver.switch_to.window(handle)
            driver.close()
        driver.switch_to.window(principal)
        driver.delete_all_cookies()
        driver.execute_script("window.localStorage.clear();")
        driver.execute_script("window.sessionStorage.clear();")
        driver.get("about:blank")
        driver.delete_all_cookies()
        return validar_sessao_limpa(driver)
    except WebDriverException:
        return False


def recriar_driver(driver):
    encerrar_driver(driver)
    return criar_driver()


def validar_login_ms(driver, cpf, senha):
    try:
        driver.get(LOGIN_URL_MS)
        cpf_input = aguardar_primeiro_elemento_valido(driver, SeletoresMS.CPF, timeout=10)
        senha_input = aguardar_primeiro_elemento_valido(driver, SeletoresMS.SENHA, timeout=10)
        botao = aguardar_primeiro_elemento_valido(driver, SeletoresMS.BOTAO_LOGIN, timeout=10, clicavel=True)
        cpf_input.clear()
        senha_input.clear()
        cpf_input.send_keys(cpf)
        senha_input.send_keys(senha)
        botao.click()

        try:
            WebDriverWait(driver, 10).until(
                lambda d: selecionar_janela_com_url(d, "home.aspx") or elemento_existe(d, SeletoresMS.MENSAGEM_ERRO)
            )
        except TimeoutException:
            return ResultadoAuth(TIMEOUT, resultado_planilha_por_classificacao(TIMEOUT), mensagem_para_planilha(TIMEOUT))

        mensagem = ""
        try:
            mensagem = aguardar_primeiro_elemento_valido(driver, SeletoresMS.MENSAGEM_ERRO, timeout=2).text.strip()
        except TimeoutException:
            mensagem = ""

        if mensagem:
            classificacao = classificar_mensagem_autenticacao(mensagem)
            if classificacao == INCONCLUSIVO:
                return ResultadoAuth(INCONCLUSIVO, RESULTADO_INCONCLUSIVO, mensagem_para_planilha(INCONCLUSIVO), mensagem)
            return ResultadoAuth(classificacao, resultado_planilha_por_classificacao(classificacao), mensagem_para_planilha(classificacao, mensagem), mensagem)

        url_home = "home.aspx" in driver.current_url.lower()
        formulario_ausente = not elemento_existe(driver, SeletoresMS.CPF) and not elemento_existe(driver, SeletoresMS.SENHA)
        if url_home and formulario_ausente:
            return ResultadoAuth(SUCESSO, RESULTADO_SIM, mensagem_para_planilha(SUCESSO))
        if url_home:
            return ResultadoAuth(INCONCLUSIVO, RESULTADO_INCONCLUSIVO, mensagem_para_planilha(INCONCLUSIVO), "URL de home sem segundo indicador comprovado.")
        return ResultadoAuth(PAGINA_INESPERADA, RESULTADO_ERRO, mensagem_para_planilha(PAGINA_INESPERADA), driver.current_url)
    except TimeoutException as exc:
        return ResultadoAuth(TIMEOUT, RESULTADO_ERRO, mensagem_para_planilha(TIMEOUT), str(exc))
    except NoSuchElementException as exc:
        return ResultadoAuth(ELEMENTO_NAO_ENCONTRADO, RESULTADO_ERRO, mensagem_para_planilha(ELEMENTO_NAO_ENCONTRADO), str(exc))
    except WebDriverException as exc:
        return ResultadoAuth(ERRO_SELENIUM, RESULTADO_ERRO, mensagem_para_planilha(ERRO_SELENIUM), str(exc))
    except Exception as exc:
        return ResultadoAuth(ERRO_INTERNO, RESULTADO_ERRO, mensagem_para_planilha(ERRO_INTERNO), str(exc))


def validar_login_ad(driver, cpf, senha):
    try:
        driver.get(LOGIN_URL_AD)
        dropdown = aguardar_primeiro_elemento_valido(driver, SeletoresAD.DROPDOWN_LOGIN, timeout=10, clicavel=True)
        dropdown.click()
        cpf_input = aguardar_primeiro_elemento_valido(driver, SeletoresAD.CPF, timeout=10)
        senha_input = aguardar_primeiro_elemento_valido(driver, SeletoresAD.SENHA, timeout=10)
        botao = aguardar_primeiro_elemento_valido(driver, SeletoresAD.BOTAO_AUTENTICAR, timeout=10, clicavel=True)
        cpf_input.clear()
        senha_input.clear()
        cpf_input.send_keys(cpf)
        senha_input.send_keys(senha)
        botao.click()

        try:
            WebDriverWait(driver, 10).until(
                lambda d: url_corresponde(d.current_url, SUCCESS_URL_AD) or bool(texto_visivel_primeiro_elemento(d, SeletoresAD.MENSAGEM_AUTENTICACAO))
            )
        except TimeoutException:
            return ResultadoAuth(TIMEOUT, RESULTADO_ERRO, mensagem_para_planilha(TIMEOUT), driver.current_url)

        mensagem = ""
        try:
            mensagem = texto_visivel_primeiro_elemento(driver, SeletoresAD.MENSAGEM_AUTENTICACAO)
        except WebDriverException:
            mensagem = ""

        if mensagem:
            classificacao = classificar_mensagem_autenticacao(mensagem)
            if classificacao == INCONCLUSIVO:
                return ResultadoAuth(INCONCLUSIVO, RESULTADO_INCONCLUSIVO, mensagem_para_planilha(INCONCLUSIVO), mensagem)
            return ResultadoAuth(classificacao, resultado_planilha_por_classificacao(classificacao), mensagem_para_planilha(classificacao, mensagem), mensagem)

        url_sucesso = url_corresponde(driver.current_url, SUCCESS_URL_AD)
        formulario_ausente = not elemento_existe(driver, SeletoresAD.CPF) and not elemento_existe(driver, SeletoresAD.SENHA)
        if url_sucesso and formulario_ausente:
            return ResultadoAuth(SUCESSO, RESULTADO_SIM, mensagem_para_planilha(SUCESSO))
        if url_sucesso:
            return ResultadoAuth(INCONCLUSIVO, RESULTADO_INCONCLUSIVO, mensagem_para_planilha(INCONCLUSIVO), "URL de portal sem segundo indicador comprovado.")
        return ResultadoAuth(PAGINA_INESPERADA, RESULTADO_ERRO, mensagem_para_planilha(PAGINA_INESPERADA), f"URL final: {driver.current_url}")
    except TimeoutException as exc:
        return ResultadoAuth(TIMEOUT, RESULTADO_ERRO, mensagem_para_planilha(TIMEOUT), str(exc))
    except NoSuchElementException as exc:
        return ResultadoAuth(ELEMENTO_NAO_ENCONTRADO, RESULTADO_ERRO, mensagem_para_planilha(ELEMENTO_NAO_ENCONTRADO), str(exc))
    except WebDriverException as exc:
        return ResultadoAuth(ERRO_SELENIUM, RESULTADO_ERRO, mensagem_para_planilha(ERRO_SELENIUM), str(exc))
    except Exception as exc:
        return ResultadoAuth(ERRO_INTERNO, RESULTADO_ERRO, mensagem_para_planilha(ERRO_INTERNO), str(exc))


def atualizar_contador_auth(resumo, sistema, classificacao):
    chave = f"{sistema}_{classificacao}"
    resumo[chave] += 1


def processar_identidade(ws, cliente_api, logger, resumo):
    print(f"🔎 Iniciando resolução de identidade")
    print(f"📄 Entrada: {arquivo_excel}\n")
    for row in range(2, ws.max_row + 1):
        inicio = time.perf_counter()
        email_informado = str(ws[f"{COL_EMAIL_INFORMADO}{row}"].value or "").strip()
        documento_original = ws[f"{COL_DOCUMENTO_INFORMADO}{row}"].value
        ws[f"{COL_STATUS_PROCESSAMENTO}{row}"] = None
        ws[f"{COL_ORIGEM_CPF}{row}"] = None

        if not documento_original and not email_informado:
            ws[f"{COL_RESULTADO_MS}{row}"] = RESULTADO_NT
            ws[f"{COL_RESULTADO_AD}{row}"] = RESULTADO_NT
            ws[f"{COL_MSG_MS}{row}"] = "CPF/RE e e-mail ausentes."
            ws[f"{COL_MSG_AD}{row}"] = "CPF/RE e e-mail ausentes."
            ws[f"{COL_STATUS_PROCESSAMENTO}{row}"] = ERRO_DE_ENTRADA
            ws[f"{COL_ORIGEM_CPF}{row}"] = NAO_ENCONTRADO
            resumo["linhas_ignoradas"] += 1
            adicionar_erro(resumo, "erros_identidade", row, "Dados ausentes nas colunas A e C")
            print(f"  ❌ Linha {row}: dados ausentes")
            continue

        cpf = None
        origem = NAO_ENCONTRADO
        normalizado = normalizar_identificador(documento_original)
        if normalizado.status == SUCESSO and normalizado.valor.tipo == "CPF":
            cpf = normalizado.valor.valor
            origem = normalizado.valor.origem
            resumo["cpfs_validos_informados"] += 1
            if origem == CPF_RECONSTRUIDO_ZERO_INICIAL:
                resumo["cpfs_reconstruidos_zero_inicial"] += 1
        elif normalizado.status == SUCESSO and normalizado.valor.tipo == "RE":
            resposta_cpf = cliente_api.buscar_cpf_por_re(normalizado.valor.valor, linha=row)
            if resposta_cpf.status == SUCESSO:
                cpf = resposta_cpf.valor
                origem = normalizado.valor.origem
                resumo["cpfs_encontrados_por_re"] += 1
                if origem == RE_RECONSTRUIDO_ZERO_INICIAL:
                    resumo["res_reconstruidos_zero_inicial"] += 1
            else:
                resumo["erros_api"] += 1 if resposta_cpf.status in CLASSIFICACOES_TECNICAS else 0
                ws[f"{COL_STATUS_PROCESSAMENTO}{row}"] = ERRO_TECNICO if resposta_cpf.status in CLASSIFICACOES_TECNICAS else ERRO_DE_ENTRADA
                adicionar_erro(resumo, "erros_identidade", row, f"Erro ao buscar CPF pelo RE: {resposta_cpf.mensagem}")
                log_event(logger, logging.WARNING, "IDENTIDADE", row, "API", cpf, "buscar_cpf_por_re", resposta_cpf.status, time.perf_counter() - inicio, resposta_cpf.mensagem)
        elif normalizado.status not in (DADOS_AUSENTES,):
            ws[f"{COL_STATUS_PROCESSAMENTO}{row}"] = ERRO_DE_ENTRADA

        if not cpf and email_informado:
            resposta_re = buscar_re_por_email(email_informado, logger, linha=row)
            if resposta_re.status == SUCESSO:
                resposta_cpf = cliente_api.buscar_cpf_por_re(resposta_re.valor, linha=row)
                if resposta_cpf.status == SUCESSO:
                    cpf = resposta_cpf.valor
                    origem = EMAIL_DSQUERY
                    resumo["cpfs_encontrados_por_email"] += 1
                else:
                    resumo["erros_api"] += 1 if resposta_cpf.status in CLASSIFICACOES_TECNICAS else 0
                    adicionar_erro(resumo, "erros_identidade", row, f"Erro ao buscar CPF pelo e-mail: {resposta_cpf.mensagem}")
                    log_event(logger, logging.WARNING, "IDENTIDADE", row, "API", cpf, "buscar_cpf_por_re_email", resposta_cpf.status, time.perf_counter() - inicio, resposta_cpf.mensagem)
            elif resposta_re.status in CLASSIFICACOES_TECNICAS:
                resumo["erros_dsquery"] += 1
                adicionar_erro(resumo, "erros_identidade", row, resposta_re.mensagem)
                log_event(logger, logging.WARNING, "IDENTIDADE", row, "DSQUERY", cpf, "pesquisa_email", resposta_re.status, time.perf_counter() - inicio, resposta_re.mensagem)

        if not cpf:
            ws[f"{COL_CPF_RESOLVIDO}{row}"] = None
            ws[f"{COL_ORIGEM_CPF}{row}"] = NAO_ENCONTRADO
            ws[f"{COL_STATUS_PROCESSAMENTO}{row}"] = ws[f"{COL_STATUS_PROCESSAMENTO}{row}"].value or ERRO_DE_ENTRADA
            ws[f"{COL_RESULTADO_MS}{row}"] = RESULTADO_NT
            ws[f"{COL_RESULTADO_AD}{row}"] = RESULTADO_NT
            ws[f"{COL_MSG_MS}{row}"] = "CPF nao encontrado."
            ws[f"{COL_MSG_AD}{row}"] = "CPF nao encontrado."
            resumo["cpfs_nao_encontrados"] += 1
            adicionar_erro(resumo, "erros_identidade", row, "CPF não encontrado")
            print(f"  ❌ Linha {row}: CPF não encontrado")
            continue

        ws[f"{COL_CPF_RESOLVIDO}{row}"] = cpf
        ws[f"{COL_ORIGEM_CPF}{row}"] = origem
        nome = cliente_api.buscar_nome_por_cpf(cpf, linha=row)
        email = cliente_api.buscar_email_por_cpf(cpf, linha=row)
        if nome.status == SUCESSO:
            ws[f"{COL_NOME}{row}"] = nome.valor
        else:
            ws[f"{COL_NOME}{row}"] = "Nome nao encontrado pelo CPF"
            resumo["erros_api"] += 1 if nome.status in CLASSIFICACOES_TECNICAS else 0
            adicionar_erro(resumo, "erros_identidade", row, f"Erro ao consultar nome pelo CPF: {nome.mensagem}")
        if email.status == SUCESSO:
            ws[f"{COL_EMAIL_FUNCIONAL}{row}"] = email.valor
        else:
            ws[f"{COL_EMAIL_FUNCIONAL}{row}"] = "E-mail funcional nao encontrado pelo CPF"
            resumo["erros_api"] += 1 if email.status in CLASSIFICACOES_TECNICAS else 0
            adicionar_erro(resumo, "erros_identidade", row, f"Erro ao consultar e-mail funcional pelo CPF: {email.mensagem}")

        if nome.status == SUCESSO and email.status == SUCESSO:
            ws[f"{COL_STATUS_PROCESSAMENTO}{row}"] = CONCLUIDO
        else:
            ws[f"{COL_STATUS_PROCESSAMENTO}{row}"] = CONCLUIDO_COM_RESSALVA
        resumo["linhas_processadas"] += 1
        log_event(logger, logging.INFO, "IDENTIDADE", row, "API", cpf, "resolver_identidade", ws[f"{COL_STATUS_PROCESSAMENTO}{row}"].value, time.perf_counter() - inicio, "Identidade processada.")
        print(f"  ✅ Linha {row}: CPF {mascarar_cpf(cpf)} resolvido ({origem})")


def processar_autenticacoes(ws, logger, resumo):
    print("\n🔐 Iniciando testes de autenticação\n")
    driver = None
    linhas_com_driver = 0
    try:
        for row in range(2, ws.max_row + 1):
            cpf = str(ws[f"{COL_CPF_RESOLVIDO}{row}"].value or "").strip()
            senha = str(ws[f"{COL_SENHA}{row}"].value or "").strip()
            ws[f"{COL_DATA_TESTE}{row}"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            cpf_ok = normalizar_cpf(cpf)
            if cpf_ok.status != SUCESSO or not senha:
                classificacao = DADOS_AUSENTES
                msg = "CPF ausente ou invalido." if cpf_ok.status != SUCESSO else "Senha ausente."
                ws[f"{COL_RESULTADO_MS}{row}"] = RESULTADO_NT
                ws[f"{COL_MSG_MS}{row}"] = msg
                ws[f"{COL_RESULTADO_AD}{row}"] = RESULTADO_NT
                ws[f"{COL_MSG_AD}{row}"] = msg
                atualizar_contador_auth(resumo, "MS", classificacao)
                atualizar_contador_auth(resumo, "AD", classificacao)
                adicionar_erro(resumo, "erros_autenticacao", row, f"MS/AD - {msg}")
                print(f"  ⏭️ Linha {row}: não testada ({msg})")
                continue

            if not driver:
                driver = criar_driver()
                linhas_com_driver = 0
            elif linhas_com_driver >= 8:
                driver = recriar_driver(driver)
                linhas_com_driver = 0

            if not limpar_sessao_driver(driver):
                driver = recriar_driver(driver)

            print(f"  🧪 Linha {row}: testando MS | CPF {mascarar_cpf(cpf)}")
            inicio_ms = time.perf_counter()
            resultado_ms = validar_login_ms(driver, cpf, senha)
            ws[f"{COL_RESULTADO_MS}{row}"] = resultado_ms.resultado_planilha
            ws[f"{COL_MSG_MS}{row}"] = resultado_ms.mensagem
            atualizar_contador_auth(resumo, "MS", resultado_ms.classificacao)
            mensagem_log_ms = resultado_ms.mensagem
            if resultado_ms.detalhe_tecnico:
                mensagem_log_ms = f"{mensagem_log_ms} | detalhe={resultado_ms.detalhe_tecnico}"
            log_event(logger, logging.INFO, "AUTENTICACAO", row, "MS", cpf, "validar_login_ms", resultado_ms.classificacao, time.perf_counter() - inicio_ms, mensagem_log_ms)
            if resultado_ms.classificacao != SUCESSO:
                adicionar_erro(resumo, "erros_autenticacao", row, f"MS - {resultado_ms.mensagem}")
                print(f"     ❌ MS: {resultado_ms.resultado_planilha} | {resultado_ms.mensagem}")
            else:
                print(f"     ✅ MS: SIM")

            if resultado_ms.classificacao in CLASSIFICACOES_TECNICAS or not limpar_sessao_driver(driver):
                driver = recriar_driver(driver)

            print(f"  🧪 Linha {row}: testando AD | CPF {mascarar_cpf(cpf)}")
            inicio_ad = time.perf_counter()
            resultado_ad = validar_login_ad(driver, cpf, senha)
            ws[f"{COL_RESULTADO_AD}{row}"] = resultado_ad.resultado_planilha
            ws[f"{COL_MSG_AD}{row}"] = resultado_ad.mensagem
            atualizar_contador_auth(resumo, "AD", resultado_ad.classificacao)
            mensagem_log_ad = resultado_ad.mensagem
            if resultado_ad.detalhe_tecnico:
                mensagem_log_ad = f"{mensagem_log_ad} | detalhe={resultado_ad.detalhe_tecnico}"
            log_event(logger, logging.INFO, "AUTENTICACAO", row, "AD", cpf, "validar_login_ad", resultado_ad.classificacao, time.perf_counter() - inicio_ad, mensagem_log_ad)
            if resultado_ad.classificacao != SUCESSO:
                adicionar_erro(resumo, "erros_autenticacao", row, f"AD - {resultado_ad.mensagem}")
                print(f"     ❌ AD: {resultado_ad.resultado_planilha} | {resultado_ad.mensagem}")
            else:
                print(f"     ✅ AD: SIM")

            if resultado_ad.classificacao in CLASSIFICACOES_TECNICAS or not limpar_sessao_driver(driver):
                driver = recriar_driver(driver)
            linhas_com_driver += 1
    finally:
        encerrar_driver(driver)


def remover_senhas(ws):
    for row in range(2, ws.max_row + 1):
        if ws[f"{COL_SENHA}{row}"].value:
            ws[f"{COL_SENHA}{row}"] = "REMOVIDA_APOS_PROCESSAMENTO"


def criar_aba_resumo(wb, resumo):
    if "Resumo" in wb.sheetnames:
        del wb["Resumo"]
    ws = wb.create_sheet("Resumo")
    ws["A1"] = "Campo"
    ws["B1"] = "Valor"
    ws["A1"].font = Font(bold=True)
    ws["B1"].font = Font(bold=True)
    linhas = [
        ("ID da execucao", resumo["execution_id"]),
        ("Arquivo de entrada", resumo["arquivo_entrada"]),
        ("Arquivo final", resumo["arquivo_final"]),
        ("Inicio", resumo["inicio"].strftime("%Y-%m-%d %H:%M:%S")),
        ("Termino", resumo["termino"].strftime("%Y-%m-%d %H:%M:%S")),
        ("Duracao total (s)", round((resumo["termino"] - resumo["inicio"]).total_seconds(), 2)),
        ("Total de linhas", resumo["total_linhas"]),
        ("Linhas processadas", resumo["linhas_processadas"]),
        ("Linhas ignoradas", resumo["linhas_ignoradas"]),
        ("CPFs validos informados", resumo["cpfs_validos_informados"]),
        ("CPFs reconstruidos com zero inicial", resumo["cpfs_reconstruidos_zero_inicial"]),
        ("CPFs encontrados por RE", resumo["cpfs_encontrados_por_re"]),
        ("CPFs encontrados por e-mail", resumo["cpfs_encontrados_por_email"]),
        ("CPFs nao encontrados", resumo["cpfs_nao_encontrados"]),
        ("Erros de API", resumo["erros_api"]),
        ("Erros de dsquery", resumo["erros_dsquery"]),
    ]
    for sistema in ("MS", "AD"):
        for classificacao in [SUCESSO, CREDENCIAL_INVALIDA, CONTA_BLOQUEADA, CONTA_DESATIVADA, SENHA_EXPIRADA, NAO_TESTADO, INCONCLUSIVO, TIMEOUT, SISTEMA_INDISPONIVEL, ERRO_SELENIUM, ERRO_INTERNO, DADOS_AUSENTES]:
            linhas.append((f"{sistema} - {classificacao}", resumo.get(f"{sistema}_{classificacao}", 0)))
    for idx, (campo, valor) in enumerate(linhas, start=2):
        ws[f"A{idx}"] = campo
        ws[f"B{idx}"] = valor
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 80


def salvar_atomicamente(wb, destino):
    diretorio = os.path.dirname(destino)
    temp_path = None
    try:
        fd, temp_path = tempfile.mkstemp(prefix=".credenciais_", suffix=".xlsx", dir=diretorio)
        os.close(fd)
        wb.save(temp_path)
        if not os.path.exists(temp_path) or os.path.getsize(temp_path) == 0:
            raise RuntimeError("Arquivo temporario nao foi criado corretamente.")
        teste = load_workbook(temp_path, read_only=True)
        teste.close()
        os.replace(temp_path, destino)
    except Exception:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass
        raise


def obter_competencia_atual(data_execucao):
    mes = MESES_PT[data_execucao.month]
    ano = f"{data_execucao.year % 100:02d}"
    return f"{mes}{ano}"


def obter_dia_semana_portugues(data_execucao):
    return DIAS_SEMANA_PT[data_execucao.weekday()]


def normalizar_cpf_para_comparacao(valor):
    resultado = normalizar_cpf(valor)
    if resultado.status == SUCESSO:
        return resultado.valor
    return None


def resultado_eh_sim(valor):
    return str(valor or "").strip().upper() == RESULTADO_SIM


def extrair_registros_mensais(ws_processado, sistema):
    col_resultado = COL_RESULTADO_AD if sistema == "AD" else COL_RESULTADO_MS
    registros = []
    for row in range(2, ws_processado.max_row + 1):
        if not resultado_eh_sim(ws_processado[f"{col_resultado}{row}"].value):
            continue
        registros.append({
            "nome": ws_processado[f"{COL_NOME}{row}"].value,
            "email": ws_processado[f"{COL_EMAIL_FUNCIONAL}{row}"].value,
            "cpf": ws_processado[f"{COL_CPF_RESOLVIDO}{row}"].value,
            "senha": ws_processado[f"{COL_SENHA}{row}"].value,
            "ms": ws_processado[f"{COL_RESULTADO_MS}{row}"].value,
            "ad": ws_processado[f"{COL_RESULTADO_AD}{row}"].value,
            "msg_ms": ws_processado[f"{COL_MSG_MS}{row}"].value,
            "msg_ad": ws_processado[f"{COL_MSG_AD}{row}"].value,
            "url_vazamento": ws_processado[f"{COL_URL_VAZAMENTO}{row}"].value,
        })
    return registros


def criar_workbook_mensal():
    wb = Workbook()
    ws = wb.active
    ws.title = "Resultado"
    for col, header in enumerate(HEADERS_MENSAL, start=1):
        cell = ws.cell(1, col)
        cell.value = header
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    aplicar_formatacao_mensal(ws)
    return wb


def validar_layout_mensal(ws):
    headers_existentes = [ws.cell(1, col).value for col in range(1, ws.max_column + 1)]
    while headers_existentes and headers_existentes[-1] is None:
        headers_existentes.pop()
    headers_migrados = list(headers_existentes)
    if len(headers_migrados) >= 9:
        header_coluna_contagem = str(headers_migrados[8] or "").lower()
        if headers_migrados[8] != HEADER_CONTAGEM_IDENTIFICACOES and "identificado" in header_coluna_contagem:
            headers_migrados[8] = HEADER_CONTAGEM_IDENTIFICACOES
    if headers_migrados == HEADERS_MENSAL:
        for col, header in enumerate(HEADERS_MENSAL, start=1):
            ws.cell(1, col).value = header
        return
    if headers_migrados == HEADERS_MENSAL[:len(headers_migrados)]:
        for col in range(1, len(HEADERS_MENSAL) + 1):
            cell = ws.cell(1, col)
            cell.value = HEADERS_MENSAL[col - 1]
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        return
    raise ValueError("Layout mensal existente incompat?vel; arquivo preservado sem altera??o.")

def carregar_ou_criar_arquivo_mensal(caminho):
    if os.path.exists(caminho):
        wb = load_workbook(caminho)
        ws = wb.active
        validar_layout_mensal(ws)
        return wb, ws, False
    wb = criar_workbook_mensal()
    return wb, wb.active, True


def obter_contagem_identificacoes(valor):
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        try:
            return max(int(valor), 1)
        except (TypeError, ValueError):
            return 1
    texto = str(valor or "").strip()
    if texto.isdigit():
        return max(int(texto), 1)
    return 1


def indexar_cpfs_existentes(ws):
    indice = {}
    for row in range(2, ws.max_row + 1):
        cpf = normalizar_cpf_para_comparacao(ws.cell(row, 3).value)
        if not cpf:
            continue
        primeira = ws.cell(row, 10).value
        if not primeira:
            primeira = ws.cell(row, 11).value
        contagem = obter_contagem_identificacoes(ws.cell(row, 9).value)
        if cpf not in indice:
            indice[cpf] = {
                "row": row,
                "primeira": primeira,
                "contagem": contagem,
            }
        else:
            indice[cpf]["contagem"] += contagem
            if not indice[cpf]["primeira"] and primeira:
                indice[cpf]["primeira"] = primeira
    return indice


def aplicar_formatacao_mensal(ws):
    success_fill = PatternFill("solid", fgColor="C6EFCE")
    failure_fill = PatternFill("solid", fgColor="FFC7CE")
    error_fill = PatternFill("solid", fgColor="F4B183")
    inconclusive_fill = PatternFill("solid", fgColor="FFEB9C")
    nt_fill = PatternFill("solid", fgColor="D9D9D9")
    resultado_fills = {
        RESULTADO_SIM: success_fill,
        RESULTADO_NAO: failure_fill,
        RESULTADO_ERRO: error_fill,
        RESULTADO_INCONCLUSIVO: inconclusive_fill,
        RESULTADO_NT: nt_fill,
    }
    for col in range(1, len(HEADERS_MENSAL) + 1):
        cell = ws.cell(1, col)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row in range(2, ws.max_row + 1):
        for col in (5, 6):
            cell = ws.cell(row, col)
            if cell.value in resultado_fills:
                cell.fill = resultado_fills[cell.value]
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for col in (3, 9, 10, 11, 12):
            ws.cell(row, col).alignment = Alignment(horizontal="center", vertical="center")
        for col in (7, 8, 13):
            ws.cell(row, col).alignment = Alignment(wrap_text=True, vertical="top")
        for col in (10, 11):
            ws.cell(row, col).number_format = "dd/mm/yyyy hh:mm:ss"
    larguras = {
        "A": 34,
        "B": 38,
        "C": 16,
        "D": 28,
        "E": 12,
        "F": 12,
        "G": 58,
        "H": 58,
        "I": 24,
        "J": 24,
        "K": 24,
        "L": 18,
        "M": 58,
    }
    for coluna, largura in larguras.items():
        ws.column_dimensions[coluna].width = largura
    ws.auto_filter.ref = f"A1:M{max(ws.max_row, 1)}"
    ws.freeze_panes = "A2"


def adicionar_registros_mensais(ws, registros, data_execucao):
    indice_cpfs = indexar_cpfs_existentes(ws)
    adicionados = 0
    duplicados = 0
    dia_semana = obter_dia_semana_portugues(data_execucao)
    for registro in registros:
        cpf_normalizado = normalizar_cpf_para_comparacao(registro["cpf"])
        primeira_identificacao = data_execucao
        contagem_identificacoes = 1
        linha_nova = True
        if cpf_normalizado and cpf_normalizado in indice_cpfs:
            duplicados += 1
            info_cpf = indice_cpfs[cpf_normalizado]
            row = info_cpf["row"]
            primeira_identificacao = info_cpf["primeira"] or data_execucao
            contagem_identificacoes = info_cpf["contagem"] + 1
            linha_nova = False
        else:
            row = ws.max_row + 1
        valores = [
            registro["nome"],
            registro["email"],
            registro["cpf"],
            registro["senha"],
            registro["ms"],
            registro["ad"],
            registro["msg_ms"],
            registro["msg_ad"],
            contagem_identificacoes,
            primeira_identificacao,
            data_execucao,
            dia_semana,
            registro["url_vazamento"],
        ]
        for col, valor in enumerate(valores, start=1):
            ws.cell(row, col).value = valor
        if cpf_normalizado:
            indice_cpfs[cpf_normalizado] = {
                "row": row,
                "primeira": primeira_identificacao,
                "contagem": contagem_identificacoes,
            }
        if linha_nova:
            adicionados += 1
    aplicar_formatacao_mensal(ws)
    return adicionados, duplicados


def salvar_workbook_mensal_atomicamente(wb, destino, linhas_antes, novas_linhas):
    diretorio = os.path.dirname(destino)
    temp_path = None
    try:
        fd, temp_path = tempfile.mkstemp(prefix=".mensal_", suffix=".xlsx", dir=diretorio)
        os.close(fd)
        wb.save(temp_path)
        teste = load_workbook(temp_path, read_only=True)
        ws_teste = teste.active
        headers = [ws_teste.cell(1, col).value for col in range(1, len(HEADERS_MENSAL) + 1)]
        if headers != HEADERS_MENSAL:
            teste.close()
            raise RuntimeError("Cabeçalhos do arquivo mensal ficaram inválidos.")
        if ws_teste.max_column != len(HEADERS_MENSAL):
            teste.close()
            raise RuntimeError("Quantidade de colunas do arquivo mensal ficou inválida.")
        if ws_teste.max_row != linhas_antes + novas_linhas:
            teste.close()
            raise RuntimeError("Quantidade de linhas do arquivo mensal não confere.")
        teste.close()
        os.replace(temp_path, destino)
    except PermissionError:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass
        raise PermissionError(f"Não foi possível atualizar {os.path.basename(destino)} porque o arquivo está aberto ou sem permissão de escrita.")
    except Exception:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass
        raise


def atualizar_arquivo_mensal(sistema, registros, data_execucao):
    competencia = obter_competencia_atual(data_execucao)
    nome_arquivo = f"Credenciais_{sistema}_{competencia}.xlsx"
    caminho = os.path.join(script_dir, nome_arquivo)
    resultado = {
        "sistema": sistema,
        "arquivo": caminho,
        "encontrados": len(registros),
        "adicionados": 0,
        "duplicados": 0,
        "erro": "",
    }
    if not registros:
        resultado["arquivo"] = caminho if os.path.exists(caminho) else ""
        return resultado
    try:
        wb, ws, _criado = carregar_ou_criar_arquivo_mensal(caminho)
        linhas_antes = ws.max_row
        adicionados, duplicados = adicionar_registros_mensais(ws, registros, data_execucao)
        salvar_workbook_mensal_atomicamente(wb, caminho, linhas_antes, adicionados)
        resultado["adicionados"] = adicionados
        resultado["duplicados"] = duplicados
    except Exception as exc:
        resultado["erro"] = str(exc)
    return resultado


def atualizar_consolidacoes_mensais(ws_processado, data_execucao):
    registros_ad = extrair_registros_mensais(ws_processado, "AD")
    registros_ms = extrair_registros_mensais(ws_processado, "MS")
    return {
        "AD": atualizar_arquivo_mensal("AD", registros_ad, data_execucao),
        "MS": atualizar_arquivo_mensal("MS", registros_ms, data_execucao),
    }


def escrever_log_resumido(caminho, titulo_erros, erros, linhas_sucesso, linhas_falha, resultado_final, consolidacao_mensal=None):
    with open(caminho, "w", encoding="utf-8") as log:
        log.write(f"=== NOVA EXECUÇÃO: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===\n")
        log.write(f"{titulo_erros}\n")
        if erros:
            for linha, erro in erros:
                log.write(f"Linha {linha} - {erro}\n")
        else:
            log.write("Nenhum erro encontrado.\n")
        log.write(f"\n✔️ Registros com sucesso: {linhas_sucesso}\n")
        log.write(f"❌ Registros com falha: {linhas_falha}\n")
        log.write(f"📁 Planilha gerada: {resultado_final}\n")
        if consolidacao_mensal:
            log.write(f"\nArquivo individual gerado:\n{resultado_final}\n")
            for sistema in ("AD", "MS"):
                info = consolidacao_mensal.get(sistema, {})
                if info.get("erro"):
                    log.write(f"\nFalha ao atualizar arquivo mensal do {sistema}: {info['erro']}\n")
                    continue
                if info.get("encontrados", 0) == 0:
                    log.write(f"\nNenhum resultado SIM encontrado para o {sistema} nesta execução.\n")
                    continue
                log.write(f"\nArquivo mensal {sistema}:\n{info.get('arquivo', '')}\n")
                log.write(f"Novas linhas adicionadas ao arquivo {sistema}:\n{info.get('adicionados', 0)}\n")
                log.write(f"Registros atualizados por identificacao repetida no {sistema}:\n{info.get('duplicados', 0)}\n")


def escrever_logs_finais(resumo, resultado_final, consolidacao_mensal=None):
    total_linhas = resumo["total_linhas"]
    sucesso_identidade = resumo["linhas_processadas"]
    falha_identidade = total_linhas - sucesso_identidade
    erros_identidade = resumo.get("erros_identidade", [])
    escrever_log_resumido(
        os.path.join(script_dir, "Logs_BuscarNome.txt"),
        "-- ERROS NA EXECUÇÃO --",
        erros_identidade,
        sucesso_identidade,
        falha_identidade,
        resultado_final,
        consolidacao_mensal,
    )

    sucesso_auth = resumo.get("MS_SUCESSO", 0) + resumo.get("AD_SUCESSO", 0)
    total_testes_auth = total_linhas * 2
    falha_auth = total_testes_auth - sucesso_auth
    erros_auth = resumo.get("erros_autenticacao", [])
    escrever_log_resumido(
        os.path.join(script_dir, "Logs_testarCredenciais.txt"),
        "-- ERROS NA EXECUÇÃO --",
        erros_auth,
        sucesso_auth,
        falha_auth,
        resultado_final,
        consolidacao_mensal,
    )


def main():
    execution_id = datetime.now().strftime("%Y%m%d%H%M%S")
    inicio = datetime.now()
    logger_identidade = configurar_logger("buscar_nome", os.path.join(script_dir, "Logs_BuscarNome.txt"), execution_id)
    logger_auth = configurar_logger("testar_credenciais", os.path.join(script_dir, "Logs_testarCredenciais.txt"), execution_id)
    resumo = Counter()
    resumo["execution_id"] = execution_id
    resumo["arquivo_entrada"] = arquivo_excel
    resumo["inicio"] = inicio
    cliente_api = None
    try:
        wb, ws, layout_novo = validar_arquivo_entrada(arquivo_excel)
        preparar_layout(ws, layout_novo)
        resumo["total_linhas"] = ws.max_row - 1
        cliente_api = ClienteAPIsInternas(logger_identidade)
        processar_identidade(ws, cliente_api, logger_identidade, resumo)
        processar_autenticacoes(ws, logger_auth, resumo)
        nome_final = f"Credenciais_{datetime.now().strftime('%d-%m-%Y_%H-%M')}.xlsx"
        resultado_final = os.path.join(script_dir, nome_final)
        resumo["arquivo_final"] = resultado_final
        resumo["termino"] = datetime.now()
        wb_saida = criar_workbook_saida(ws)
        salvar_atomicamente(wb_saida, resultado_final)
        consolidacao_mensal = atualizar_consolidacoes_mensais(ws, resumo["termino"])
        escrever_logs_finais(resumo, resultado_final, consolidacao_mensal)
        for sistema in ("AD", "MS"):
            info = consolidacao_mensal.get(sistema, {})
            print(f"\n{sistema}:")
            print(f"Resultados SIM encontrados nesta execução: {info.get('encontrados', 0)}")
            print(f"Novas linhas adicionadas: {info.get('adicionados', 0)}")
            print(f"Registros atualizados por identificacao repetida neste mes: {info.get('duplicados', 0)}")
            if info.get("erro"):
                print(f"Falha ao atualizar arquivo mensal: {info['erro']}")
            elif info.get("arquivo"):
                print(f"Arquivo mensal: {info['arquivo']}")
            else:
                print("Arquivo mensal: não criado nesta execução")
        print("\n📊 Resumo final")
        print(f"📁 Arquivo final: {resultado_final}")
        print(f"✅ Identidades resolvidas: {resumo['linhas_processadas']}")
        print(f"⚠️ Linhas ignoradas: {resumo['linhas_ignoradas']}")
        print(f"🟢 MS sucesso: {resumo.get('MS_SUCESSO', 0)}")
        print(f"🟢 AD sucesso: {resumo.get('AD_SUCESSO', 0)}")
        print("🏁 Execução concluída.")
        return 0
    except Exception as exc:
        print(f"💥 Erro fatal: {exc}")
        return 1
    finally:
        if cliente_api:
            cliente_api.fechar()


if __name__ == "__main__":
    sys.exit(main())
