"""
Monitor de Produtividade - TimeSheet (versão integrada)
=======================================================
Esta versão substitui o envio de e-mail pelo registro DIRETO no TimeSheet
(WebAppSystems), via API REST com autenticação por usuário/senha do sistema.

Mudanças em relação ao app original:
  1. Login com usuário e senha do TimeSheet (em vez de digitar o e-mail).
  2. Ao finalizar a atividade, o usuário seleciona Cliente, Área, Tipo e
     Solicitante (listas carregadas do próprio TimeSheet) e confirma o horário.
  3. O lançamento é gravado direto no TimeSheet (em vez de enviar e-mail).

O motor de monitoramento de janelas (win32gui / bandeja) é o mesmo do original.

Dependências: requests, pywin32, psutil, pystray, pillow
    pip install requests pywin32 psutil pystray pillow
"""

import time
import tkinter as tk
from tkinter import ttk, messagebox
import win32gui
import win32process
import win32api
import win32con
import psutil
import json
import os
import sys
import math
import unicodedata
import ctypes
import ctypes.wintypes
import re
from datetime import datetime, timedelta
import threading
import pystray
from PIL import Image, ImageDraw, ImageTk
import requests

# ============================================================
# CONFIGURAÇÃO
# ============================================================
# Versão do app (usada no título e no futuro aviso de atualização)
APP_VERSAO = "1.0.2"

# URL do servidor, escolhida automaticamente:
#   - Executável instalado (.exe do PyInstaller) -> PRODUÇÃO
#   - Rodando o .py (desenvolvimento)             -> localhost
# Para forçar outra URL (ex.: testar o .exe contra o localhost), defina a variável
# de ambiente TIMESHEET_API_URL antes de abrir o app.
URL_PRODUCAO = "https://ecadvogados.azurewebsites.net"
URL_DESENVOLVIMENTO = "https://localhost:5095"
API_BASE_URL = os.environ.get("TIMESHEET_API_URL") or (
    URL_PRODUCAO if getattr(sys, "frozen", False) else URL_DESENVOLVIMENTO
)

# Em desenvolvimento (localhost) o certificado é autoassinado; desabilita a
# verificação SSL só nesse caso. Em produção (Azure), mantém a verificação.
VERIFY_SSL = not API_BASE_URL.startswith("https://localhost")
if not VERIFY_SSL:
    # Evita o aviso de "InsecureRequestWarning" no console durante o dev
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def caminho_asset(nome):
    """Resolve o caminho de um arquivo em 'assets', funcionando tanto ao rodar o .py
    quanto no .exe empacotado (PyInstaller extrai os dados em sys._MEIPASS)."""
    import sys
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, "assets", nome)


def ja_esta_rodando():
    """Garante instância única usando um mutex nomeado do Windows.
    Retorna True se JÁ existe outra instância do app em execução."""
    try:
        import win32event
        import win32api
        import winerror
        # Nome único do mutex para este app
        ja_esta_rodando._mutex = win32event.CreateMutex(None, False, "JRW_MonitorTimeSheet_SingleInstance")
        return win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS
    except Exception:
        # Em caso de falha no mutex, não bloqueia a execução
        return False

# Paleta de cores (alinhada ao visual da web - roxo/gradiente #667eea -> #764ba2)
COR_FUNDO = "#f4f6f9"          # fundo claro como a web
COR_CARD = "#ffffff"           # cartões brancos
COR_PRIMARIA = "#667eea"       # roxo/azul da web
COR_PRIMARIA_ESCURA = "#764ba2"  # roxo do gradiente
COR_PRIMARIA_HOVER = "#5568d3"
COR_TEXTO = "#2d3748"          # texto escuro sobre fundo claro
COR_TEXTO_SUAVE = "#718096"
COR_TEXTO_HEADER = "#ffffff"   # texto sobre o cabeçalho roxo
COR_SUCESSO = "#2f855a"
COR_ERRO = "#c53030"
COR_INPUT_BG = "#f7fafc"       # input claro
COR_BORDA = "#e2e8f0"

# --- CONFIGURAÇÃO DE PASTAS ---
pasta_usuario = os.path.expanduser("~")
pasta_documentos = os.path.join(pasta_usuario, "Documents")
arquivo_config = os.path.join(pasta_documentos, "config_timesheet_escritorio.json")

# --- VARIÁVEIS GLOBAIS ---
cronometro_rodando = False
monitoramento_pausado = False
tempo_inicio = 0
hora_inicio_texto = ""
hwnd_monitorado = None
titulo_monitorado = ""
caminho_arquivo_monitorado = ""
janelas_ja_perguntadas = set()
icone_bandeja = None

# Cronômetro flutuante (widget no canto inferior direito)
cronometro_widget = None
cronometro_label = None
cronometro_after_id = None
cronometro_tooltip = None
cronometro_partes = []    # (cartão, ícone, texto) do cronômetro flutuante, para trocar as cores

# Pausa do cronômetro (tecla de atalho)
cronometro_pausado = False
pausa_desde = 0.0         # instante em que a pausa atual começou
pausa_total = 0.0         # segundos já pausados nesta atividade (pausas anteriores)
inicio_dt = None          # data/hora reais do início da atividade
ATALHO_PADRAO = "Ctrl+Alt+P"
atalho_pausa = ATALHO_PADRAO

# Borda vermelha ao redor da janela monitorada
borda_janelas = []        # as 4 faixas (topo, base, esq, dir)
borda_after_id = None
BORDA_COR = "#e53e3e"
BORDA_COR_PAUSA = "#d69e2e"   # âmbar enquanto o cronômetro está pausado
BORDA_ESPESSURA = 3

# Sessão do usuário autenticado no TimeSheet
sessao = {
    "token": None,
    "userId": None,
    "userName": None,
    "departmentId": None
}

# Catálogo de programas que PODEM ser monitorados (rótulo amigável -> termo no título da janela).
# O usuário escolhe quais ativar na tela de Configurações (bandeja).
PROGRAMAS_DISPONIVEIS = [
    ("Microsoft Word", "word"),
    ("Microsoft Excel", "excel"),
    ("Microsoft PowerPoint", "powerpoint"),
    ("Outlook (e-mail)", "outlook"),
    ("Arquivos PDF", "pdf"),
    ("Google Chrome", "chrome"),
    ("Microsoft Edge", "edge"),
    ("Mozilla Firefox", "firefox"),
    ("eproc", "eproc"),
    ("Projudi", "projudi"),
]

# Lista efetivamente monitorada (preenchida a partir da config salva)
palavras_alvo = []
termos_ignorar = ["lista de atalhos", "menu", "iniciar", "pesquisa", "configurações"]

# Padrão quando ainda não há config salva: monitora tudo do catálogo
PALAVRAS_ALVO_PADRAO = [termo for _, termo in PROGRAMAS_DISPONIVEIS]


def carregar_palavras_alvo():
    """Carrega os termos monitorados do arquivo de config; usa o padrão se não houver."""
    global palavras_alvo
    cfg = carregar_configuracoes()
    monitorados = cfg.get("programas_monitorados")
    if monitorados and isinstance(monitorados, list):
        palavras_alvo = monitorados
    else:
        palavras_alvo = list(PALAVRAS_ALVO_PADRAO)


def salvar_palavras_alvo(lista):
    """Salva a lista de termos monitorados, preservando o restante da config."""
    global palavras_alvo
    cfg = carregar_configuracoes()
    cfg["programas_monitorados"] = lista
    try:
        with open(arquivo_config, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        palavras_alvo = lista
        return True
    except Exception:
        return False


# Preferência: exibir o cronômetro flutuante no canto da tela? (o tempo é contado de qualquer forma)
exibir_cronometro = True


def carregar_preferencias():
    """Carrega as preferências de exibição do arquivo de config (padrão: cronômetro visível)."""
    global exibir_cronometro
    global aviso_retomada_ativo
    cfg = carregar_configuracoes()
    exibir_cronometro = bool(cfg.get("mostrar_cronometro", True))
    aviso_retomada_ativo = bool(cfg.get("aviso_retomada", True))


def definir_exibir_cronometro(valor):
    """Salva a preferência e já aplica, inclusive em um cronômetro que esteja rodando agora."""
    global exibir_cronometro
    valor = bool(valor)
    cfg = carregar_configuracoes()
    cfg["mostrar_cronometro"] = valor
    try:
        with open(arquivo_config, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        return False
    exibir_cronometro = valor
    if cronometro_rodando:
        if valor:
            mostrar_cronometro_flutuante()
        else:
            esconder_cronometro_flutuante()
    return True

# Regras para montar a descrição automática a partir do título da janela.
# Cada regra: (palavras que identificam o programa, modelo de frase).
# {arquivo} é substituído pelo nome do arquivo detectado (sem extensão), quando houver.
REGRAS_DESCRICAO = [
    (["eproc", "projudi"], "Acompanhamento processual no sistema {programa}"),
    (["outlook"], "Tratativas por e-mail"),
    (["word", "winword"], "Elaboração/revisão do documento {arquivo}"),
    (["excel"], "Trabalho na planilha {arquivo}"),
    (["powerpoint"], "Elaboração da apresentação {arquivo}"),
    (["pdf", "foxit", "acrobat"], "Análise do documento {arquivo}"),
    (["chrome", "edge", "firefox"], "Pesquisa/consulta: {arquivo}"),
]


def _nome_programa(titulo_min):
    """Identifica um rótulo amigável do programa a partir do título em minúsculas."""
    mapa = {
        "eproc": "eproc", "projudi": "Projudi", "outlook": "Outlook",
        "winword": "Word", "word": "Word", "excel": "Excel",
        "powerpoint": "PowerPoint", "foxit": "PDF", "acrobat": "PDF", "pdf": "PDF",
        "chrome": "navegador", "edge": "navegador", "firefox": "navegador",
    }
    for chave, nome in mapa.items():
        if chave in titulo_min:
            return nome
    return "sistema"


def _extrair_nome_arquivo(titulo):
    """Extrai um nome de arquivo/documento legível do título da janela."""
    # Títulos costumam ser "arquivo.ext - Programa" ou "Programa - arquivo.ext"
    partes = [p.strip() for p in titulo.split(" - ") if p.strip()]
    candidato = None
    # Prefere a parte que tenha cara de arquivo (com extensão)
    for p in partes:
        if any(p.lower().endswith(ext) for ext in
               (".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".pdf", ".txt")):
            candidato = p
            break
    if candidato is None and partes:
        # Senão, usa a primeira parte (geralmente é o conteúdo/arquivo)
        candidato = partes[0]
    if not candidato:
        return ""
    # Remove a extensão
    for ext in (".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt", ".pdf", ".txt"):
        if candidato.lower().endswith(ext):
            candidato = candidato[: -len(ext)]
            break
    return candidato.strip()


LIMITE_DESCRICAO = 200  # mesmo limite do TimeSheet (banco e cadastro web)


def gerar_descricao_automatica(titulo):
    """Monta uma sugestão de descrição a partir do título da janela (sem IA)."""
    if not titulo:
        return ""
    titulo_min = titulo.lower()
    arquivo = _extrair_nome_arquivo(titulo)
    programa = _nome_programa(titulo_min)

    for palavras, modelo in REGRAS_DESCRICAO:
        if any(p in titulo_min for p in palavras):
            frase = modelo.format(arquivo=arquivo or "o arquivo", programa=programa)
            # Evita "o arquivo" solto quando não há nome
            return frase.replace("  ", " ").strip()
    # Fallback genérico
    if arquivo:
        return f"Atividade relacionada a {arquivo}"
    return ""


# ============================================================
# LOGIN GUARDADO ("Manter conectado") E INÍCIO COM O WINDOWS
# ============================================================
# O usuário e a senha ficam só neste computador, protegidos pela criptografia do
# próprio Windows (DPAPI): só o mesmo usuário do Windows, nesta máquina, consegue ler.
import hashlib

pasta_dados_app = os.path.join(os.environ.get("APPDATA") or pasta_usuario, "JRW", "MonitorTimeSheet")
# Um arquivo por servidor (produção x testes), para um não sobrescrever o outro
arquivo_login = os.path.join(
    pasta_dados_app, "login_" + hashlib.sha1(API_BASE_URL.encode("utf-8")).hexdigest()[:8] + ".dat")
_ENTROPIA_LOGIN = b"JRW.MonitorTimeSheet.v1"


def salvar_login_guardado(login, senha):
    """Grava o login criptografado com o Windows. Retorna True se gravou."""
    try:
        import win32crypt
        os.makedirs(pasta_dados_app, exist_ok=True)
        bruto = json.dumps({"login": login, "senha": senha}).encode("utf-8")
        cifrado = win32crypt.CryptProtectData(bruto, "MonitorTimeSheet", _ENTROPIA_LOGIN, None, None, 0)
        with open(arquivo_login, "wb") as f:
            f.write(cifrado)
        return True
    except Exception:
        return False


def carregar_login_guardado():
    """Retorna (login, senha) guardados ou None (sem arquivo, ilegível ou de outro usuário do Windows)."""
    try:
        import win32crypt
        if not os.path.exists(arquivo_login):
            return None
        with open(arquivo_login, "rb") as f:
            cifrado = f.read()
        _, bruto = win32crypt.CryptUnprotectData(cifrado, _ENTROPIA_LOGIN, None, None, 0)
        dados = json.loads(bruto.decode("utf-8"))
        if dados.get("login") and dados.get("senha"):
            return dados["login"], dados["senha"]
    except Exception:
        pass
    return None


def apagar_login_guardado():
    try:
        if os.path.exists(arquivo_login):
            os.remove(arquivo_login)
    except Exception:
        pass


def gravar_config(**valores):
    """Atualiza chaves do arquivo de configuração sem perder as demais. Retorna True se gravou."""
    cfg = carregar_configuracoes()
    cfg.update(valores)
    try:
        with open(arquivo_config, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


# Mesmo nome de valor que o instalador (Inno Setup) grava, para os dois controlarem a mesma entrada
_NOME_INICIALIZACAO = "MonitorTimeSheet"
_CHAVE_RUN = r"Software\Microsoft\Windows\CurrentVersion\Run"


def _comando_inicializacao():
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}"'
    pythonw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    exe = pythonw if os.path.exists(pythonw) else sys.executable
    return f'"{exe}" "{os.path.abspath(__file__)}"'


def iniciar_com_windows_ativo():
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _CHAVE_RUN) as k:
            winreg.QueryValueEx(k, _NOME_INICIALIZACAO)
            return True
    except Exception:
        return False


def definir_iniciar_com_windows(ativar):
    """Liga/desliga a abertura do Monitor junto com o Windows (só para o usuário atual)."""
    try:
        import winreg
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, _CHAVE_RUN) as k:
            if ativar:
                winreg.SetValueEx(k, _NOME_INICIALIZACAO, 0, winreg.REG_SZ, _comando_inicializacao())
            else:
                try:
                    winreg.DeleteValue(k, _NOME_INICIALIZACAO)
                except FileNotFoundError:
                    pass
        gravar_config(iniciar_com_windows=bool(ativar))
        return True
    except Exception:
        return False


def ativar_inicio_automatico_na_primeira_vez():
    """Ao marcar 'Manter conectado' pela primeira vez no app instalado, liga também a
    abertura com o Windows (o usuário pode desligar em Configurações). Só no .exe instalado,
    para não registrar o script de desenvolvimento no Windows."""
    if not getattr(sys, "frozen", False):
        return
    if "iniciar_com_windows" not in carregar_configuracoes():
        definir_iniciar_com_windows(True)


def aplicar_sessao(dados):
    sessao["token"] = dados.get("token")
    sessao["userId"] = dados.get("userId")
    sessao["userName"] = dados.get("userName")
    sessao["departmentId"] = dados.get("departmentId")


# ============================================================
# CLIENTE DA API DO TIMESHEET
# ============================================================
class TimeSheetAPI:
    """Encapsula as chamadas à API de integração do WebAppSystems."""

    def __init__(self, base_url):
        self.base_url = base_url.rstrip("/")
        self.token = None
        # Credenciais só em memória, para renovar a sessão quando o token vence (8 h)
        self._login = None
        self._senha = None
        # Motivo da última falha de login: "credenciais" (usuário/senha recusados),
        # "temporario" (sem rede, servidor fora, limite de tentativas) ou "outro"
        self.ultimo_erro = None
        self._trava = threading.Lock()
        # Chamado (em qualquer thread) quando o servidor recusa o login guardado
        self.ao_credenciais_recusadas = None

    def definir_credenciais(self, login, senha):
        self._login, self._senha = login, senha

    def limpar_sessao(self):
        self.token = None
        self._login = self._senha = None
        sessao.update({"token": None, "userId": None, "userName": None, "departmentId": None})

    def renovar_sessao(self):
        """Entra de novo com as credenciais em memória. Retorna True se renovou."""
        if not self._login:
            return False
        with self._trava:
            ok, _ = self.login(self._login, self._senha)
        if not ok and self.ultimo_erro == "credenciais":
            # Senha alterada (ou usuário desativado): o login guardado não serve mais
            login_antigo = self._login
            apagar_login_guardado()
            self.limpar_sessao()
            if self.ao_credenciais_recusadas:
                try:
                    self.ao_credenciais_recusadas(login_antigo)
                except Exception:
                    pass
        return ok

    def _requisitar(self, metodo, caminho, timeout=20, **kw):
        """Chamada autenticada. Se não há token ou o servidor responde 401, renova a sessão
        com o login guardado e repete a chamada uma vez."""
        if not self.token and self._login:
            self.renovar_sessao()
        url = f"{self.base_url}{caminho}"
        r = requests.request(metodo, url, headers=self._headers(), verify=VERIFY_SSL,
                             timeout=timeout, **kw)
        if r.status_code == 401 and self._login and self.renovar_sessao():
            r = requests.request(metodo, url, headers=self._headers(), verify=VERIFY_SSL,
                                 timeout=timeout, **kw)
        return r

    def _headers(self):
        h = {"Content-Type": "application/json"}
        if self.token:
            h["Authorization"] = f"Bearer {self.token}"
        return h

    def login(self, login, senha):
        """Autentica e guarda o token. Retorna (ok, dados_ou_mensagem)."""
        try:
            r = requests.post(
                f"{self.base_url}/api/LoginApi/authenticate",
                json={"Login": login, "Senha": senha},
                verify=VERIFY_SSL,
                timeout=20
            )
            if r.status_code == 200:
                data = r.json()
                self.token = data.get("token")
                self._login, self._senha = login, senha
                self.ultimo_erro = None
                aplicar_sessao(data)
                return True, data
            elif r.status_code == 401:
                self.ultimo_erro = "credenciais"
                return False, "Usuário ou senha incorretos. Verifique e tente novamente."
            elif r.status_code == 429:
                self.ultimo_erro = "temporario"
                return False, "Muitas tentativas de login. Aguarde 1 minuto e tente de novo."
            else:
                self.ultimo_erro = "temporario" if r.status_code >= 500 else "outro"
                try:
                    msg = r.json().get("message", "Não foi possível entrar. Tente novamente.")
                except Exception:
                    msg = "Não foi possível entrar. Tente novamente."
                return False, msg
        except requests.exceptions.ConnectionError:
            self.ultimo_erro = "temporario"
            return False, ("Não foi possível conectar ao TimeSheet.\n"
                          "Verifique se o sistema está no ar e tente novamente.")
        except requests.exceptions.Timeout:
            self.ultimo_erro = "temporario"
            return False, "O servidor demorou para responder. Verifique sua conexão e tente novamente."
        except requests.exceptions.RequestException:
            self.ultimo_erro = "temporario"
            return False, "Falha de conexão com o TimeSheet. Tente novamente em instantes."

    def get_clientes(self):
        r = self._requisitar("GET", "/api/TimesheetApi/clientes")
        r.raise_for_status()
        return r.json()

    def get_areas(self):
        r = self._requisitar("GET", "/api/TimesheetApi/areas")
        r.raise_for_status()
        return r.json()

    def get_tipos(self):
        r = self._requisitar("GET", "/api/TimesheetApi/tipos")
        r.raise_for_status()
        return r.json()

    def get_versao(self):
        r = self._requisitar("GET", "/api/TimesheetApi/versao")
        r.raise_for_status()
        return r.json()

    def baixar_instalador(self, destino, ao_progresso=None):
        """Baixa o instalador para 'destino', chamando ao_progresso(baixado, total)."""
        r = self._requisitar("GET", "/api/TimesheetApi/instalador", timeout=60, stream=True)
        r.raise_for_status()
        total = int(r.headers.get("Content-Length") or 0)
        baixado = 0
        with open(destino, "wb") as f:
            for bloco in r.iter_content(chunk_size=256 * 1024):
                if bloco:
                    f.write(bloco)
                    baixado += len(bloco)
                    if ao_progresso:
                        ao_progresso(baixado, total)
        return baixado

    def get_solicitantes(self, cliente_id):
        r = self._requisitar("GET", f"/api/TimesheetApi/clientes/{cliente_id}/solicitantes")
        r.raise_for_status()
        return r.json()

    def registrar_lancamento(self, payload):
        """Grava o lançamento. Retorna (ok, mensagem)."""
        try:
            r = self._requisitar("POST", "/api/TimesheetApi/lancamentos", timeout=30, json=payload)
            if r.status_code == 200:
                return True, r.json().get("message", "Lançamento registrado.")
            elif r.status_code == 401:
                return False, ("Sua sessão expirou e não foi possível renovar. Use 'Trocar de usuário' "
                              "no ícone da bandeja para entrar de novo. O tempo desta atividade "
                              "continua nesta tela.")
            elif r.status_code == 404:
                return False, ("O TimeSheet ainda não está atualizado para o Monitor. "
                              "Avise o suporte (JRW).")
            else:
                try:
                    return False, r.json().get("message", "Não foi possível registrar. Tente novamente.")
                except Exception:
                    return False, "Não foi possível registrar. Tente novamente."
        except requests.exceptions.RequestException as e:
            return False, f"Falha de conexão ao registrar.\n({e})"


api = TimeSheetAPI(API_BASE_URL)


# ============================================================
# HELPERS DE UI (estilização)
# ============================================================
def centralizar(janela, largura, altura):
    janela.update_idletasks()
    x = (janela.winfo_screenwidth() // 2) - (largura // 2)
    y = (janela.winfo_screenheight() // 2) - (altura // 2)
    janela.geometry(f"{largura}x{altura}+{x}+{y}")


def criar_botao(parent, texto, comando, primario=True):
    bg = COR_PRIMARIA if primario else "#edf2f7"
    fg = "#ffffff" if primario else COR_TEXTO_SUAVE
    btn = tk.Button(
        parent, text=texto, command=comando,
        bg=bg, fg=fg, activebackground=COR_PRIMARIA_HOVER, activeforeground="#ffffff",
        relief="flat", bd=0, font=("Segoe UI", 10, "bold"),
        cursor="hand2", padx=20, pady=10
    )
    hover = COR_PRIMARIA_HOVER if primario else "#e2e8f0"
    btn.bind("<Enter>", lambda e: btn.config(bg=hover))
    btn.bind("<Leave>", lambda e: btn.config(bg=bg))
    return btn


def _hex_para_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _desenhar_gradiente(canvas, largura, altura, cor1, cor2):
    """Desenha um gradiente diagonal suave (cor1 -> cor2) usando PIL (sem faixas)."""
    r1, g1, b1 = _hex_para_rgb(cor1)
    r2, g2, b2 = _hex_para_rgb(cor2)

    img = Image.new("RGB", (largura, altura))
    px = img.load()
    # gradiente diagonal (135deg, como o CSS da web): mistura x e y
    for y in range(altura):
        for x in range(largura):
            t = ((x / largura) + (y / altura)) / 2
            r = int(r1 + (r2 - r1) * t)
            g = int(g1 + (g2 - g1) * t)
            b = int(b1 + (b2 - b1) * t)
            px[x, y] = (r, g, b)

    foto = ImageTk.PhotoImage(img)
    canvas.create_image(0, 0, anchor="nw", image=foto)
    # guarda referência para o garbage collector não apagar a imagem
    if not hasattr(canvas, "_imgs_ref"):
        canvas._imgs_ref = []
    canvas._imgs_ref.append(foto)


def _retangulo_arredondado(canvas, x1, y1, x2, y2, raio, **kwargs):
    """Desenha um retângulo com cantos arredondados no canvas."""
    pontos = [
        x1 + raio, y1, x2 - raio, y1, x2, y1, x2, y1 + raio,
        x2, y2 - raio, x2, y2, x2 - raio, y2, x1 + raio, y2,
        x1, y2, x1, y2 - raio, x1, y1 + raio, x1, y1
    ]
    return canvas.create_polygon(pontos, smooth=True, **kwargs)


def _gerar_retangulo_gradiente(largura, altura, raio, cor1, cor2):
    """Gera (PIL) um retângulo arredondado preenchido com gradiente diagonal, suave."""
    esc = 3
    w, h = largura * esc, altura * esc
    grad = Image.new("RGB", (w, h))
    gpx = grad.load()
    r1, g1, b1 = _hex_para_rgb(cor1)
    r2, g2, b2 = _hex_para_rgb(cor2)
    for y in range(h):
        for x in range(w):
            t = ((x / w) + (y / h)) / 2
            gpx[x, y] = (int(r1 + (r2 - r1) * t), int(g1 + (g2 - g1) * t), int(b1 + (b2 - b1) * t))
    mask = Image.new("L", (w, h), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, w - 1, h - 1], radius=raio * esc, fill=255)
    out = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    out.paste(grad, (0, 0), mask)
    return out.resize((largura, altura), Image.LANCZOS)


def _retangulo_arredondado_img(canvas, largura, altura, raio, cor1, cor2):
    """Coloca no canvas um retângulo arredondado com gradiente (imagem PIL suave)."""
    img = _gerar_retangulo_gradiente(largura, altura, raio, cor1, cor2)
    foto = ImageTk.PhotoImage(img)
    canvas.create_image(0, 0, anchor="nw", image=foto)
    if not hasattr(canvas, "_imgs_ref"):
        canvas._imgs_ref = []
    canvas._imgs_ref.append(foto)


def _desenhar_relogio_simples(canvas, cx, cy, raio):
    """Relógio branco (contorno + ponteiros) para dentro do badge, estilo da web."""
    import math as _m
    canvas.create_oval(cx - raio, cy - raio, cx + raio, cy + raio, outline="#ffffff", width=2)
    # marcações de minutos (tracinhos finos)
    for k in range(12):
        ang = _m.radians(k * 30)
        x_ini = cx + (raio - 3) * _m.sin(ang)
        y_ini = cy - (raio - 3) * _m.cos(ang)
        x_fim = cx + raio * _m.sin(ang)
        y_fim = cy - raio * _m.cos(ang)
        canvas.create_line(x_ini, y_ini, x_fim, y_fim, fill="#ffffff", width=1)
    # ponteiros
    canvas.create_line(cx, cy, cx, cy - raio + 5, fill="#ffffff", width=2)
    canvas.create_line(cx, cy, cx + raio - 7, cy, fill="#ffffff", width=2)


_icone_cache = {}


def _garantir_ico():
    """Gera (uma vez) um arquivo .ico do logo em várias resoluções e retorna o caminho.
    O .ico é o formato que o Windows usa melhor para barra de título e de tarefas."""
    if "ico_path" in _icone_cache:
        return _icone_cache["ico_path"]
    try:
        caminho = os.path.join(pasta_documentos, "timesheet_icon.ico")
        logo = _gerar_logo_relogio(256)
        # .ico com múltiplos tamanhos (o Windows escolhe o melhor para cada lugar)
        logo.save(caminho, format="ICO",
                 sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
        _icone_cache["ico_path"] = caminho
        return caminho
    except Exception:
        return None


def aplicar_icone_janela(janela):
    """Define o ícone do app (logo do relógio) na janela, substituindo a pena padrão."""
    ico = _garantir_ico()
    if ico:
        try:
            janela.iconbitmap(ico)
            return
        except Exception:
            pass
    # Fallback: iconphoto com fundo branco
    try:
        if "foto" not in _icone_cache:
            logo = _gerar_logo_relogio(32)
            fundo = Image.new("RGBA", logo.size, (255, 255, 255, 255))
            fundo.alpha_composite(logo)
            _icone_cache["foto"] = ImageTk.PhotoImage(fundo.convert("RGB"))
        janela.iconphoto(True, _icone_cache["foto"])
    except Exception:
        pass


def _botao_gradiente(parent, texto, comando, largura=380, altura=50):
    """Cria um 'botão' com fundo em gradiente roxo (Canvas clicável), estilo da web."""
    cv = tk.Canvas(parent, width=largura, height=altura, highlightthickness=0, bd=0, cursor="hand2")
    img_normal = _gerar_retangulo_gradiente(largura, altura, 10, "#667eea", "#764ba2")
    img_hover = _gerar_retangulo_gradiente(largura, altura, 10, "#5a6fd6", "#6a4296")
    foto_n = ImageTk.PhotoImage(img_normal)
    foto_h = ImageTk.PhotoImage(img_hover)
    cv._imgs_ref = [foto_n, foto_h]
    bg_item = cv.create_image(0, 0, anchor="nw", image=foto_n)
    txt_item = cv.create_text(largura // 2, altura // 2, text=texto, fill="#ffffff",
                             font=("Segoe UI", 12, "bold"))

    estado = {"ativo": True}

    def definir(texto=None, ativo=None):
        if texto is not None:
            cv.itemconfig(txt_item, text=texto)
        if ativo is not None:
            estado["ativo"] = ativo

    cv.bind("<Enter>", lambda e: cv.itemconfig(bg_item, image=foto_h))
    cv.bind("<Leave>", lambda e: cv.itemconfig(bg_item, image=foto_n))
    cv.bind("<Button-1>", lambda e: comando() if estado["ativo"] else None)
    # método auxiliar para o chamador alterar texto/estado sem mexer no .config nativo
    cv.definir = definir
    return cv


def _gerar_logo_relogio(tamanho=96):
    """Reproduz o logo oficial do sistema (favicon.svg) com PIL, em alta resolução
    e com antialiasing (desenha 4x maior e reduz) para ficar suave, sem serrilhado.
    Retorna um objeto PIL.Image (RGBA)."""
    esc = 4
    s = tamanho * esc
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    def p(v):  # converte coordenada do viewBox 0..100 para pixels
        return v / 100 * s

    # Círculo de fundo com gradiente roxo (aproxima o linearGradient do SVG)
    raio_fundo = p(48)
    cx = cy = s / 2
    # desenha o gradiente num quadrado e recorta em círculo via máscara
    grad = Image.new("RGB", (s, s))
    gpx = grad.load()
    r1, g1, b1 = _hex_para_rgb("#667eea")
    r2, g2, b2 = _hex_para_rgb("#764ba2")
    for y in range(s):
        for x in range(s):
            t = ((x / s) + (y / s)) / 2
            gpx[x, y] = (int(r1 + (r2 - r1) * t), int(g1 + (g2 - g1) * t), int(b1 + (b2 - b1) * t))
    mask = Image.new("L", (s, s), 0)
    ImageDraw.Draw(mask).ellipse([cx - raio_fundo, cy - raio_fundo, cx + raio_fundo, cy + raio_fundo], fill=255)
    img.paste(grad, (0, 0), mask)

    # Mostrador branco
    rf = p(35)
    d.ellipse([cx - rf, cy - rf, cx + rf, cy + rf], fill=(255, 255, 255, 242))

    roxo1 = _hex_para_rgb("#667eea")
    roxo2 = _hex_para_rgb("#764ba2")

    # Marcadores das horas (12, 3, 6, 9)
    rm = p(2)
    for (mx, my) in [(50, 20), (80, 50), (50, 80), (20, 50)]:
        d.ellipse([p(mx) - rm, p(my) - rm, p(mx) + rm, p(my) + rm], fill=roxo1)

    # Ponteiro das horas
    d.line([p(50), p(50), p(50), p(35)], fill=roxo2, width=int(p(3)))
    # Ponteiro dos minutos
    d.line([p(50), p(50), p(65), p(50)], fill=roxo1, width=int(p(2.5)))
    # Centro
    rc = p(4)
    d.ellipse([cx - rc, cy - rc, cx + rc, cy + rc], fill=roxo2)
    # Pequeno "play" no canto inferior direito
    d.polygon([(p(72), p(72)), (p(82), p(77)), (p(72), p(82))], fill=roxo1)

    # Reduz para o tamanho final com antialiasing (LANCZOS) -> fica suave
    return img.resize((tamanho, tamanho), Image.LANCZOS)


def _desenhar_logo_timetracker(canvas, cx, cy, tamanho=96):
    """Coloca o logo oficial (relógio) renderizado com PIL no canvas."""
    img = _gerar_logo_relogio(tamanho)
    foto = ImageTk.PhotoImage(img)
    canvas.create_image(cx, cy, image=foto)
    if not hasattr(canvas, "_imgs_ref"):
        canvas._imgs_ref = []
    canvas._imgs_ref.append(foto)


def estilizar_combo(root):
    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure("TS.TCombobox",
                    fieldbackground="#ffffff", background="#ffffff",
                    foreground=COR_TEXTO, arrowcolor=COR_PRIMARIA,
                    bordercolor=COR_BORDA, lightcolor=COR_BORDA, darkcolor=COR_BORDA,
                    borderwidth=1, padding=7)
    style.map("TS.TCombobox",
             fieldbackground=[("readonly", "#ffffff")],
             bordercolor=[("focus", COR_PRIMARIA), ("active", COR_PRIMARIA)],
             selectbackground=[("readonly", "#ffffff")],
             selectforeground=[("readonly", COR_TEXTO)])


# ============================================================
# ÍCONE DE RELÓGIO PARA A BANDEJA (igual ao original)
# ============================================================
def criar_imagem_bandeja():
    # Usa o mesmo logo oficial (relógio com gradiente) renderizado com PIL
    return _gerar_logo_relogio(64).convert("RGBA")


# ============================================================
# TELA DE LOGIN
# ============================================================
def tela_login(root, aviso=None, login_inicial=""):
    """Exibe a janela de login como Toplevel da raiz única. Retorna True se autenticou.
    'aviso' mostra uma mensagem no topo (ex.: senha alterada); 'login_inicial' preenche o usuário."""
    login_ok = {"valor": False}

    LARG, ALT = 500, 720
    win = tk.Toplevel(root)
    win.title("TimeSheet - Login")
    win.configure(bg="#ffffff")
    win.resizable(False, False)
    aplicar_icone_janela(win)
    centralizar(win, LARG, ALT)

    # ---------- Topo: badge do relógio + títulos (fundo branco, igual à web) ----------
    topo = tk.Frame(win, bg="#ffffff")
    topo.pack(fill="x", pady=(44, 0))

    # Badge roxo arredondado com o relógio dentro
    badge = tk.Canvas(topo, width=72, height=72, bg="#ffffff", highlightthickness=0, bd=0)
    _retangulo_arredondado_img(badge, 72, 72, 18, "#667eea", "#764ba2")
    _desenhar_relogio_simples(badge, 36, 36, 17)
    badge.pack()

    tk.Label(topo, text="Time Tracker", bg="#ffffff", fg="#1a202c",
            font=("Segoe UI", 24, "bold")).pack(pady=(16, 2))
    tk.Label(topo, text="Entre na sua conta", bg="#ffffff", fg="#718096",
            font=("Segoe UI", 11)).pack()

    # ---------- Formulário ----------
    form = tk.Frame(win, bg="#ffffff")
    form.pack(fill="x", padx=60, pady=(34, 0))

    def campo_com_icone(parent, rotulo, icone_txt, senha=False):
        tk.Label(parent, text=rotulo, bg="#ffffff", fg="#4a5568",
                font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 6))
        wrap = tk.Frame(parent, bg="#f7fafc", highlightbackground="#e2e8f0",
                       highlightthickness=1)
        wrap.pack(fill="x")
        tk.Label(wrap, text=icone_txt, bg="#f7fafc", fg="#a0aec0",
                font=("Segoe UI", 12)).pack(side="left", padx=(12, 6), pady=10)
        ent = tk.Entry(wrap, bg="#f7fafc", fg="#2d3748", insertbackground="#2d3748",
                      relief="flat", font=("Segoe UI", 11),
                      show="•" if senha else "")
        ent.pack(side="left", fill="x", expand=True, ipady=6, padx=(0, 8))
        # efeito de foco (borda roxa)
        ent.bind("<FocusIn>", lambda e: wrap.config(highlightbackground="#667eea", highlightcolor="#667eea"))
        ent.bind("<FocusOut>", lambda e: wrap.config(highlightbackground="#e2e8f0"))
        return ent, wrap

    entry_login, _ = campo_com_icone(form, "Usuário", "\U0001F464")  # 👤
    tk.Frame(form, bg="#ffffff", height=16).pack()

    entry_senha, wrap_senha = campo_com_icone(form, "Senha", "\U0001F512", senha=True)  # 🔒
    # botão "olho" para mostrar/ocultar senha
    mostrar = {"on": False}
    def toggle_senha():
        mostrar["on"] = not mostrar["on"]
        entry_senha.config(show="" if mostrar["on"] else "•")
        btn_olho.config(text="\U0001F441" if mostrar["on"] else "\U0001F576")  # 👁 / 🕶
    btn_olho = tk.Label(wrap_senha, text="\U0001F576", bg="#f7fafc", fg="#a0aec0",
                       font=("Segoe UI", 11), cursor="hand2")
    btn_olho.pack(side="right", padx=(0, 12))
    btn_olho.bind("<Button-1>", lambda e: toggle_senha())

    var_manter = tk.BooleanVar(value=True)
    tk.Checkbutton(
        form, text="  Manter conectado neste computador", variable=var_manter,
        bg="#ffffff", fg="#4a5568", selectcolor="#ffffff", activebackground="#ffffff",
        activeforeground="#667eea", font=("Segoe UI", 10), anchor="w", cursor="hand2",
        highlightthickness=0, bd=0
    ).pack(anchor="w", pady=(14, 0))

    lbl_erro = tk.Label(form, text=aviso or "", bg="#ffffff", fg=COR_ERRO if not aviso else "#b7791f",
                       font=("Segoe UI", 9), wraplength=380, justify="left")
    lbl_erro.pack(anchor="w", pady=(10, 0))

    def fazer_login():
        login = entry_login.get().strip()
        senha = entry_senha.get()
        if not login or not senha:
            lbl_erro.config(text="Informe usuário e senha.")
            return
        btn_entrar.definir(texto="Entrando...", ativo=False)
        win.update()
        ok, resultado = api.login(login, senha)
        if ok:
            # (a sessão já foi preenchida por api.login)
            if var_manter.get():
                if salvar_login_guardado(login, senha):
                    ativar_inicio_automatico_na_primeira_vez()
            else:
                apagar_login_guardado()
            login_ok["valor"] = True
            win.destroy()
        else:
            lbl_erro.config(fg=COR_ERRO, text=str(resultado))
            btn_entrar.definir(texto="Entrar  →", ativo=True)

    # Botão "Entrar" com gradiente roxo (imagem)
    btn_entrar = _botao_gradiente(form, "Entrar  →", fazer_login, largura=380, altura=50)
    btn_entrar.pack(fill="x", pady=(22, 0))

    # ---------- Rodapé ----------
    rodape = tk.Frame(win, bg="#ffffff")
    rodape.pack(side="bottom", fill="x", pady=(0, 20))
    tk.Frame(rodape, bg="#edf2f7", height=1).pack(fill="x", padx=60, pady=(0, 14))
    tk.Label(rodape, text="JRWIT.COM", bg="#ffffff", fg="#2d3748",
            font=("Segoe UI", 11, "bold")).pack()
    tk.Label(rodape, text="SMART SOLUTIONS FOR YOUR BUSINESS", bg="#ffffff", fg="#a0aec0",
            font=("Segoe UI", 8)).pack(pady=(2, 0))
    tk.Label(rodape, text=f"Versão {APP_VERSAO}", bg="#ffffff", fg="#cbd5e0",
            font=("Segoe UI", 8)).pack(pady=(4, 0))
    if API_BASE_URL.startswith("https://localhost"):
        tk.Label(rodape, text="⚠ Ambiente de testes", bg="#ffffff", fg="#d69e2e",
                font=("Segoe UI", 8)).pack(pady=(6, 0))

    entry_senha.bind("<Return>", lambda e: fazer_login())
    if login_inicial:
        entry_login.insert(0, login_inicial)
        entry_senha.focus()
    else:
        entry_login.focus()
    win.grab_set()

    win.protocol("WM_DELETE_WINDOW", lambda: win.destroy())
    win.wait_window()  # bloqueia até a janela de login fechar (sem criar outro mainloop)
    return login_ok["valor"]


# ============================================================
# GERENCIAMENTO DE CONFIGURAÇÕES (mantém o último login)
# ============================================================
def carregar_configuracoes():
    if os.path.exists(arquivo_config):
        try:
            with open(arquivo_config, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


# ============================================================
# CAPTURAR O CAMINHO DO ARQUIVO (REDE U:\) - igual ao original
# ============================================================
def obter_caminho_do_processo(hwnd):
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        processo = psutil.Process(pid)
        for arquivo in processo.open_files():
            caminho = arquivo.path
            if caminho and ("U:\\" in caminho or "u:\\" in caminho):
                return caminho
        cmdline = processo.cmdline()
        for arg in cmdline:
            if arg and ("U:\\" in arg or "u:\\" in arg) and os.path.exists(arg):
                return arg
    except Exception:
        pass
    return ""


# ============================================================
# LÓGICA DE MONITORAMENTO (igual ao original)
# ============================================================
def verificar_janelas():
    global cronometro_rodando, monitoramento_pausado, hwnd_monitorado
    global titulo_monitorado, caminho_arquivo_monitorado, janelas_ja_perguntadas

    if not monitoramento_pausado:
        try:
            hwnd_atual = win32gui.GetForegroundWindow()
            if cronometro_rodando and hwnd_monitorado is not None:
                if not win32gui.IsWindow(hwnd_monitorado) or not win32gui.IsWindowVisible(hwnd_monitorado):
                    parar_cronometro_automatico()
            elif not cronometro_rodando:
                if win32gui.IsWindowVisible(hwnd_atual):
                    titulo_atual = win32gui.GetWindowText(hwnd_atual).strip()
                    if titulo_atual:
                        titulo_minusculo = titulo_atual.lower()
                        eh_menu_ou_atalho = any(termo in titulo_minusculo for termo in termos_ignorar)
                        if not eh_menu_ou_atalho:
                            matches = [p for p in palavras_alvo if p in titulo_minusculo]
                            if matches and hwnd_atual not in janelas_ja_perguntadas:
                                janelas_ja_perguntadas.add(hwnd_atual)
                                perguntar_iniciar_cronometro(hwnd_atual, titulo_atual)
        except Exception:
            pass

    janela_principal.after(4000, verificar_janelas)


def perguntar_iniciar_cronometro(hwnd, titulo_janela):
    """Janela customizada perguntando se deseja monitorar (substitui o messagebox padrão)."""
    win = tk.Toplevel(janela_principal)
    win.title("Monitoramento")
    win.configure(bg=COR_FUNDO)
    win.resizable(False, False)
    win.attributes("-topmost", True)
    centralizar(win, 420, 300)

    resposta = {"valor": False}

    # Cabeçalho com gradiente roxo
    header = tk.Canvas(win, width=420, height=64, highlightthickness=0, bd=0)
    header.pack(fill="x")
    _desenhar_gradiente(header, 420, 64, "#667eea", "#764ba2")
    header.create_text(22, 32, anchor="w", text="⏱  Atividade detectada",
                      fill=COR_TEXTO_HEADER, font=("Segoe UI", 13, "bold"))

    body = tk.Frame(win, bg=COR_FUNDO)
    body.pack(fill="both", expand=True, padx=28, pady=22)

    tk.Label(body, text="Detectamos que você abriu:", bg=COR_FUNDO, fg=COR_TEXTO_SUAVE,
            font=("Segoe UI", 10)).pack(anchor="w")

    # Nome da janela em destaque, dentro de um "chip"
    chip = tk.Frame(body, bg="#eef1fb")
    chip.pack(fill="x", pady=(8, 16))
    nome = titulo_janela if len(titulo_janela) <= 55 else titulo_janela[:55] + "..."
    tk.Label(chip, text=nome, bg="#eef1fb", fg=COR_PRIMARIA_ESCURA,
            font=("Segoe UI", 11, "bold"), wraplength=330, justify="left").pack(anchor="w", padx=14, pady=10)

    tk.Label(body, text="Deseja iniciar o cronômetro para registrar o tempo desta atividade?",
            bg=COR_FUNDO, fg=COR_TEXTO, font=("Segoe UI", 10), wraplength=360,
            justify="left").pack(anchor="w")

    def confirmar(valor):
        resposta["valor"] = valor
        win.destroy()

    frame_btns = tk.Frame(win, bg=COR_FUNDO)
    frame_btns.pack(fill="x", padx=28, pady=(0, 20))
    criar_botao(frame_btns, "Sim, cronometrar", lambda: confirmar(True), primario=True).pack(side="right")
    criar_botao(frame_btns, "Agora não", lambda: confirmar(False), primario=False).pack(side="right", padx=(0, 10))

    win.protocol("WM_DELETE_WINDOW", lambda: confirmar(False))
    win.grab_set()
    win.wait_window()

    if resposta["valor"]:
        iniciar_cronometro(hwnd, titulo_janela)


# ============================================================
# CRONÔMETRO FLUTUANTE (canto inferior direito)
# ============================================================
def mostrar_cronometro_flutuante():
    """Cria uma janelinha sem bordas no canto inferior direito com o tempo correndo."""
    global cronometro_widget, cronometro_label

    if cronometro_widget is not None:
        return  # já existe

    w = tk.Toplevel(janela_principal)
    w.overrideredirect(True)          # sem barra de título/bordas
    w.attributes("-topmost", True)    # sempre visível
    try:
        w.attributes("-alpha", 0.96)
    except Exception:
        pass

    larg, alt = 150, 48
    w.update_idletasks()
    sw = w.winfo_screenwidth()
    sh = w.winfo_screenheight()
    # canto inferior direito, com margem (acima da barra de tarefas)
    x = sw - larg - 20
    y = sh - alt - 70
    w.geometry(f"{larg}x{alt}+{x}+{y}")

    # cartão roxo arredondado (simulado com frame)
    card = tk.Frame(w, bg=COR_PRIMARIA, cursor="fleur")
    card.pack(fill="both", expand=True)

    icone = tk.Label(card, text="⏱", bg=COR_PRIMARIA, fg="#ffffff", font=("Segoe UI", 15), cursor="fleur")
    icone.pack(side="left", padx=(12, 6))

    lbl = tk.Label(card, text="00:00", bg=COR_PRIMARIA, fg="#ffffff", font=("Consolas", 15, "bold"), cursor="fleur")
    lbl.pack(side="left")

    cronometro_widget = w
    cronometro_label = lbl
    cronometro_partes[:] = [card, icone, lbl]
    _aplicar_visual_pausa()

    # --- Arrastar o cronômetro para qualquer posição ---
    arrasto = {"x": 0, "y": 0, "movendo": False}

    def iniciar_arrasto(e):
        arrasto["x"] = e.x
        arrasto["y"] = e.y
        arrasto["movendo"] = False
        _esconder_tooltip_cronometro()

    def mover(e):
        arrasto["movendo"] = True
        nx = w.winfo_x() + (e.x - arrasto["x"])
        ny = w.winfo_y() + (e.y - arrasto["y"])
        w.geometry(f"+{nx}+{ny}")

    # Tooltip ao passar o mouse + bindings de arraste
    for widget in (card, icone, lbl):
        widget.bind("<Enter>", lambda e: (None if arrasto["movendo"] else _mostrar_tooltip_cronometro()))
        widget.bind("<Leave>", lambda e: _esconder_tooltip_cronometro())
        widget.bind("<Button-1>", iniciar_arrasto)
        widget.bind("<B1-Motion>", mover)

    _tick_cronometro()


def _tick_cronometro():
    """Atualiza o tempo exibido a cada segundo."""
    global cronometro_after_id
    if cronometro_widget is None or not cronometro_rodando:
        return
    decorrido = tempo_trabalhado_seg()
    h = decorrido // 3600
    m = (decorrido % 3600) // 60
    s = decorrido % 60
    if h > 0:
        cronometro_label.config(text=f"{h:02d}:{m:02d}:{s:02d}")
    else:
        cronometro_label.config(text=f"{m:02d}:{s:02d}")
    cronometro_after_id = janela_principal.after(1000, _tick_cronometro)


def _mostrar_tooltip_cronometro():
    global cronometro_tooltip
    if cronometro_widget is None or cronometro_tooltip is not None:
        return
    texto = titulo_monitorado if titulo_monitorado else "Monitorando atividade"
    if len(texto) > 55:
        texto = texto[:55] + "..."
    if cronometro_pausado:
        texto += f"\n⏸ Pausado — {atalho_pausa} para retomar"
    else:
        texto += f"\n{atalho_pausa} para pausar"

    tip = tk.Toplevel(janela_principal)
    tip.overrideredirect(True)
    tip.attributes("-topmost", True)
    lbl = tk.Label(tip, text=texto, bg="#2d3748", fg="#ffffff",
                  font=("Segoe UI", 9), padx=10, pady=6, justify="left", wraplength=260)
    lbl.pack()
    tip.update_idletasks()
    # posiciona o tooltip logo acima do cronômetro
    cx = cronometro_widget.winfo_x()
    cy = cronometro_widget.winfo_y()
    tw = tip.winfo_width()
    tip.geometry(f"+{cx + 150 - tw}+{cy - tip.winfo_height() - 8}")
    cronometro_tooltip = tip


def _esconder_tooltip_cronometro():
    global cronometro_tooltip
    if cronometro_tooltip is not None:
        cronometro_tooltip.destroy()
        cronometro_tooltip = None


def esconder_cronometro_flutuante():
    """Remove o cronômetro flutuante (quando o cronômetro para)."""
    global cronometro_widget, cronometro_label, cronometro_after_id
    _esconder_tooltip_cronometro()
    if cronometro_after_id is not None:
        try:
            janela_principal.after_cancel(cronometro_after_id)
        except Exception:
            pass
        cronometro_after_id = None
    if cronometro_widget is not None:
        cronometro_widget.destroy()
        cronometro_widget = None
        cronometro_label = None


# ============================================================
# BORDA VERMELHA ao redor da janela monitorada
# ============================================================
def _hwnd_da_faixa(f):
    """HWND real (janela de nível superior) de uma faixa do tkinter."""
    h = win32gui.GetParent(f.winfo_id())
    return h or f.winfo_id()


def _criar_faixa():
    """Cria uma faixa vermelha: janela sem borda, click-through, que não rouba foco e não
    aparece na barra de tarefas. NÃO é 'sempre no topo': a ordem em relação às demais janelas
    é ajustada a cada ciclo por _posicionar_borda_acima_da_janela()."""
    f = tk.Toplevel(janela_principal)
    f.overrideredirect(True)
    f.configure(bg=BORDA_COR_PAUSA if cronometro_pausado else BORDA_COR)
    try:
        f.update_idletasks()
        hwnd = _hwnd_da_faixa(f)
        estilo = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        win32gui.SetWindowLong(
            hwnd, win32con.GWL_EXSTYLE,
            estilo | win32con.WS_EX_LAYERED | win32con.WS_EX_TRANSPARENT   # clique atravessa
                   | win32con.WS_EX_NOACTIVATE | win32con.WS_EX_TOOLWINDOW  # sem foco / sem barra de tarefas
        )
    except Exception:
        pass
    return f


def _posicionar_borda_acima_da_janela(alvo):
    """Coloca as 4 faixas LOGO ACIMA da janela monitorada na pilha de janelas do Windows.

    Assim a borda se comporta como a moldura do próprio app: continua visível mesmo que
    outro app esteja em foco (ex.: em outro monitor), e qualquer janela que COBRIR o app
    cobre a borda também, em vez de a borda ficar flutuando por cima de tudo."""
    faixas = {_hwnd_da_faixa(f) for f in borda_janelas}

    def acima_de(h):
        try:
            return win32gui.GetWindow(h, win32con.GW_HWNDPREV)
        except Exception:
            return 0

    def eh_topmost(h):
        return bool(win32gui.GetWindowLong(h, win32con.GWL_EXSTYLE) & win32con.WS_EX_TOPMOST)

    anterior = acima_de(alvo)                 # janela imediatamente acima do app
    while anterior and anterior in faixas:    # ignora as nossas próprias faixas
        anterior = acima_de(anterior)

    if not anterior or (eh_topmost(anterior) and not eh_topmost(alvo)):
        # nada (ou só janelas "sempre no topo") acima do app: topo da pilha normal
        depois_de = win32con.HWND_TOP
    else:
        depois_de = anterior                  # fica logo abaixo dessa janela = logo acima do app

    flags = win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE
    for f in borda_janelas:
        win32gui.SetWindowPos(_hwnd_da_faixa(f), depois_de, 0, 0, 0, 0, flags)


def mostrar_borda():
    """Cria as 4 faixas da borda e inicia o acompanhamento da janela monitorada."""
    global borda_janelas
    if borda_janelas:
        return
    borda_janelas = [_criar_faixa() for _ in range(4)]  # topo, base, esq, dir
    _atualizar_borda()


def _atualizar_borda():
    """Reposiciona as 4 faixas ao redor da janela monitorada. Roda em loop."""
    global borda_after_id
    if not borda_janelas or not cronometro_rodando or hwnd_monitorado is None:
        return
    try:
        # A borda acompanha o app enquanto ele estiver visível, com ou sem foco (importante
        # com 2 monitores). Some só se o app fechou/minimizou. Se outra janela o cobrir, a
        # borda fica POR BAIXO dela (ver _posicionar_borda_acima_da_janela).
        if (not win32gui.IsWindow(hwnd_monitorado)
                or not win32gui.IsWindowVisible(hwnd_monitorado)
                or win32gui.IsIconic(hwnd_monitorado)):
            for f in borda_janelas:
                f.withdraw()
        else:
            l, t, r, b = win32gui.GetWindowRect(hwnd_monitorado)

            # Se a janela estiver MAXIMIZADA, o GetWindowRect vaza ~8px além do
            # monitor. Nesse caso usamos a área de trabalho do monitor onde a
            # janela está (multi-monitor), para a borda ficar certinha nas beiradas.
            try:
                placement = win32gui.GetWindowPlacement(hwnd_monitorado)
                maximizada = placement[1] == win32con.SW_SHOWMAXIMIZED
            except Exception:
                maximizada = False

            if maximizada:
                try:
                    monitor = win32api.MonitorFromWindow(hwnd_monitorado, 2)  # MONITOR_DEFAULTTONEAREST
                    info = win32api.GetMonitorInfo(monitor)
                    l, t, r, b = info["Work"]  # área de trabalho (desconta a barra de tarefas)
                except Exception:
                    pass

            largura = r - l
            altura = b - t
            if largura < 10 or altura < 10:
                for f in borda_janelas:
                    f.withdraw()
            else:
                e = BORDA_ESPESSURA
                topo, base, esq, dir_ = borda_janelas
                topo.geometry(f"{largura}x{e}+{l}+{t}")
                base.geometry(f"{largura}x{e}+{l}+{b - e}")
                esq.geometry(f"{e}x{altura}+{l}+{t}")
                dir_.geometry(f"{e}x{altura}+{r - e}+{t}")
                for f in borda_janelas:
                    # deiconify() só quando estava escondida (no Windows ele também traz
                    # a janela para frente e ativa; fazer isso a cada ciclo bagunçaria a pilha)
                    if f.state() == "withdrawn":
                        f.deiconify()
                # O Tk só aplica geometry() quando fica ocioso. Se o SetWindowPos da pilha de
                # janelas rodar antes disso, o evento dele sobrescreve a posição pendente com
                # a antiga (0,0). Por isso aplicamos as posições ANTES de ajustar a pilha.
                janela_principal.update_idletasks()
                _posicionar_borda_acima_da_janela(hwnd_monitorado)
    except Exception:
        pass
    borda_after_id = janela_principal.after(120, _atualizar_borda)


def esconder_borda():
    """Remove as faixas da borda e para o acompanhamento."""
    global borda_janelas, borda_after_id
    if borda_after_id is not None:
        try:
            janela_principal.after_cancel(borda_after_id)
        except Exception:
            pass
        borda_after_id = None
    for f in borda_janelas:
        try:
            f.destroy()
        except Exception:
            pass
    borda_janelas = []


# ============================================================
# PAUSA DO CRONÔMETRO (tecla de atalho global)
# ============================================================
def tempo_trabalhado_seg():
    """Segundos efetivamente trabalhados: tempo total menos as pausas (inclusive a atual)."""
    agora = time.time()
    pausado_agora = (agora - pausa_desde) if cronometro_pausado else 0
    return max(int(agora - tempo_inicio - pausa_total - pausado_agora), 0)


def _aplicar_visual_pausa():
    """Cores do cronômetro flutuante e da borda: roxo/vermelho rodando, âmbar pausado."""
    cor_cartao = "#d69e2e" if cronometro_pausado else COR_PRIMARIA
    try:
        if len(cronometro_partes) == 3:
            card, icone, lbl = cronometro_partes
            card.config(bg=cor_cartao)
            icone.config(bg=cor_cartao, text="⏸" if cronometro_pausado else "⏱")
            lbl.config(bg=cor_cartao)
    except Exception:
        pass
    for f in borda_janelas:
        try:
            f.configure(bg=BORDA_COR_PAUSA if cronometro_pausado else BORDA_COR)
        except Exception:
            pass


def _aviso_rapido(texto, ms=2200):
    """Aviso discreto no canto da tela (não rouba o foco), para confirmar pausa/retomada."""
    try:
        w = tk.Toplevel(janela_principal)
        w.overrideredirect(True)
        w.attributes("-topmost", True)
        tk.Label(w, text=texto, bg="#2d3748", fg="#ffffff", font=("Segoe UI", 10, "bold"),
                 padx=16, pady=10).pack()
        w.update_idletasks()
        x = w.winfo_screenwidth() - w.winfo_width() - 20
        y = w.winfo_screenheight() - w.winfo_height() - 130
        w.geometry(f"+{x}+{y}")
        try:
            hwnd = win32gui.GetParent(w.winfo_id()) or w.winfo_id()
            estilo = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
            win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE,
                                   estilo | win32con.WS_EX_NOACTIVATE | win32con.WS_EX_TOOLWINDOW)
        except Exception:
            pass
        w.after(ms, w.destroy)
    except Exception:
        pass


def alternar_pausa_cronometro():
    """Pausa ou retoma o cronômetro da atividade em andamento (atalho ou bandeja)."""
    global cronometro_pausado, pausa_desde, pausa_total
    if not cronometro_rodando:
        _aviso_rapido("Nenhuma atividade em andamento")
        return
    agora = time.time()
    if cronometro_pausado:
        pausa_total += agora - pausa_desde
        cronometro_pausado = False
        fechar_aviso_retomada()
        _aviso_rapido("▶ Cronômetro retomado")
    else:
        pausa_desde = agora
        cronometro_pausado = True
        _reiniciar_deteccao_retomada(carencia=RETOMADA_CARENCIA_SEG)
        _aviso_rapido("⏸ Cronômetro pausado")
    _aplicar_visual_pausa()
    atualizar_menu_bandeja()


# ============================================================
# AVISO DE RETOMADA (Melhoria 2b)
# ============================================================
# Durante uma pausa, se a pessoa voltar a trabalhar (janela em foco + uso de mouse/teclado,
# sem ler o que é digitado), o Monitor pergunta, sem roubar o foco, se deseja retomar.
# Se ela abrir OUTRO programa monitorado, nunca inicia uma segunda contagem sozinho:
# pergunta se continua a atividade pausada, encerra e inicia uma nova, ou mantém a pausa.
RETOMADA_INTERVALO_MS = 2000      # frequência da checagem
RETOMADA_CARENCIA_SEG = 60        # não pergunta no 1º minuto da pausa
RETOMADA_ATIVIDADE_SEG = 8        # uso contínuo necessário antes de perguntar
RETOMADA_OCIOSO_MS = 2500         # sem mouse/teclado há mais que isso = não está usando
RETOMADA_SILENCIO_SEG = 600       # depois de "continuar pausado", não pergunta por 10 min
RETOMADA_AVISO_EXPIRA_MS = 45000  # aviso some sozinho (vale como "manter pausa")

aviso_retomada_ativo = True       # preferência do usuário (Configurações)
_retomada = {"alvo": None, "seg": 0, "silencio_ate": 0.0, "janela": None}


class _LASTINPUTINFO(ctypes.Structure):
    _fields_ = [("cbSize", ctypes.c_uint), ("dwTime", ctypes.c_uint)]


def _ms_sem_uso():
    """Milissegundos desde o último uso de mouse/teclado (só o horário, não o conteúdo)."""
    try:
        info = _LASTINPUTINFO()
        info.cbSize = ctypes.sizeof(info)
        ctypes.windll.user32.GetLastInputInfo(ctypes.byref(info))
        return (ctypes.windll.kernel32.GetTickCount() - info.dwTime) & 0xFFFFFFFF
    except Exception:
        return 10 ** 9


def _titulo_monitoravel(hwnd):
    """Título da janela se ela for de um programa monitorado (mesma regra da detecção), senão None."""
    try:
        if not hwnd or not win32gui.IsWindowVisible(hwnd):
            return None
        titulo = win32gui.GetWindowText(hwnd).strip()
        t = titulo.lower()
        if not titulo or any(termo in t for termo in termos_ignorar):
            return None
        return titulo if any(p in t for p in palavras_alvo) else None
    except Exception:
        return None


def _janela_do_proprio_app(hwnd):
    try:
        return win32process.GetWindowThreadProcessId(hwnd)[1] == os.getpid()
    except Exception:
        return False


def _reiniciar_deteccao_retomada(carencia=0):
    _retomada["alvo"] = None
    _retomada["seg"] = 0
    _retomada["silencio_ate"] = time.time() + carencia


def fechar_aviso_retomada():
    w = _retomada.get("janela")
    _retomada["janela"] = None
    if w is not None:
        try:
            w.destroy()
        except Exception:
            pass


def verificar_retomada():
    """Roda sempre (a cada 2 s); só age com o cronômetro pausado e o aviso ligado."""
    _passo_retomada()
    janela_principal.after(RETOMADA_INTERVALO_MS, verificar_retomada)


def _passo_retomada():
    try:
        if (cronometro_rodando and cronometro_pausado and aviso_retomada_ativo
                and _retomada["janela"] is None and time.time() >= _retomada["silencio_ate"]):
            fg = win32gui.GetForegroundWindow()
            alvo = None
            if fg and not _janela_do_proprio_app(fg) and _ms_sem_uso() < RETOMADA_OCIOSO_MS:
                if fg == hwnd_monitorado:
                    alvo = ("mesma", fg, titulo_monitorado)
                else:
                    titulo = _titulo_monitoravel(fg)
                    if titulo:
                        alvo = ("outra", fg, titulo)
            if alvo and _retomada["alvo"] and alvo[:2] == _retomada["alvo"][:2]:
                _retomada["seg"] += RETOMADA_INTERVALO_MS / 1000
            else:
                _retomada["seg"] = RETOMADA_INTERVALO_MS / 1000 if alvo else 0
            _retomada["alvo"] = alvo
            if alvo and _retomada["seg"] >= RETOMADA_ATIVIDADE_SEG:
                mostrar_aviso_retomada(*alvo)
    except Exception:
        pass


def _encurtar(txt, n=48):
    txt = txt or ""
    return txt if len(txt) <= n else txt[:n - 3] + "..."


def mostrar_aviso_retomada(tipo, hwnd_novo, titulo_novo):
    """Cartão no canto da tela, sem roubar o foco de quem está digitando."""
    fechar_aviso_retomada()
    w = tk.Toplevel(janela_principal)
    w.overrideredirect(True)
    w.attributes("-topmost", True)
    _retomada["janela"] = w
    cartao = tk.Frame(w, bg="#ffffff", highlightbackground="#d69e2e", highlightthickness=2)
    cartao.pack(fill="both", expand=True)
    corpo = tk.Frame(cartao, bg="#ffffff")
    corpo.pack(fill="both", expand=True, padx=14, pady=12)

    if tipo == "mesma":
        titulo_aviso = "⏸ O cronômetro está pausado"
        texto = f"Parece que você voltou a trabalhar em:\n{_encurtar(titulo_monitorado)}\n\nRetomar o cronômetro?"
    else:
        titulo_aviso = "⏸ Há uma atividade pausada"
        texto = (f"Pausada: {_encurtar(titulo_monitorado)}\n"
                 f"Agora em: {_encurtar(titulo_novo)}\n\nO que deseja fazer?")
    tk.Label(corpo, text=titulo_aviso, bg="#ffffff", fg="#744210",
             font=("Segoe UI", 10, "bold")).pack(anchor="w")
    tk.Label(corpo, text=texto, bg="#ffffff", fg=COR_TEXTO, font=("Segoe UI", 9),
             justify="left", wraplength=320).pack(anchor="w", pady=(4, 10))

    botoes = tk.Frame(corpo, bg="#ffffff")
    botoes.pack(fill="x")

    def botao(txt, cmd, primario):
        b = criar_botao(botoes, txt, cmd, primario=primario)
        b.config(padx=10, pady=5, font=("Segoe UI", 9, "bold"))
        b.pack(fill="x", pady=(0, 5))

    def continuar_pausado():
        fechar_aviso_retomada()
        _reiniciar_deteccao_retomada(carencia=RETOMADA_SILENCIO_SEG)

    def retomar():
        fechar_aviso_retomada()
        if cronometro_rodando and cronometro_pausado:
            alternar_pausa_cronometro()

    def encerrar_e_iniciar():
        fechar_aviso_retomada()
        if not cronometro_rodando:
            return
        parar_cronometro_automatico()           # abre a tela de lançamento da pausada
        if win32gui.IsWindow(hwnd_novo):
            janelas_ja_perguntadas.add(hwnd_novo)
            iniciar_cronometro(hwnd_novo, titulo_novo)

    if tipo == "mesma":
        botao("Retomar o cronômetro", retomar, True)
        botao("Continuar pausado", continuar_pausado, False)
    else:
        botao("Continuar a atividade pausada", retomar, True)
        botao("Encerrar a pausada e iniciar nova", encerrar_e_iniciar, False)
        botao("Manter a pausa", continuar_pausado, False)

    w.update_idletasks()
    larg = max(w.winfo_reqwidth(), 340)
    alt = w.winfo_reqheight()
    x = w.winfo_screenwidth() - larg - 20
    y = w.winfo_screenheight() - alt - 130
    w.geometry(f"{larg}x{alt}+{x}+{y}")
    try:
        h = win32gui.GetParent(w.winfo_id()) or w.winfo_id()
        estilo = win32gui.GetWindowLong(h, win32con.GWL_EXSTYLE)
        win32gui.SetWindowLong(h, win32con.GWL_EXSTYLE,
                               estilo | win32con.WS_EX_NOACTIVATE | win32con.WS_EX_TOOLWINDOW)
    except Exception:
        pass

    def expirar():
        if _retomada.get("janela") is w:
            continuar_pausado()
    w.after(RETOMADA_AVISO_EXPIRA_MS, expirar)


def definir_aviso_retomada(valor):
    global aviso_retomada_ativo
    aviso_retomada_ativo = bool(valor)
    if not aviso_retomada_ativo:
        fechar_aviso_retomada()
    return gravar_config(aviso_retomada=aviso_retomada_ativo)


_MODIFICADORES_ATALHO = {"ctrl": 0x0002, "alt": 0x0001, "shift": 0x0004}


def interpretar_atalho(texto):
    """'Ctrl+Alt+P' -> (modificadores, código da tecla) ou None se inválido."""
    try:
        partes = [p.strip() for p in str(texto).split("+") if p.strip()]
        if not partes:
            return None
        *mods_txt, tecla = partes
        mods = 0
        for m in mods_txt:
            bit = _MODIFICADORES_ATALHO.get(m.lower())
            if bit is None:
                return None
            mods |= bit
        t = tecla.upper()
        if len(t) == 1 and (t.isalpha() or t.isdigit()):
            vk = ord(t)
        elif t.startswith("F") and t[1:].isdigit() and 1 <= int(t[1:]) <= 12:
            vk = 0x70 + int(t[1:]) - 1
        else:
            return None
        # Letras/números soltos atrapalhariam a digitação: exige ao menos Ctrl ou Alt
        if len(t) == 1 and not (mods & 0x0003):
            return None
        return mods, vk
    except Exception:
        return None


_hotkey = {"tid": None, "thread": None}


def parar_atalho_pausa():
    tid, th = _hotkey["tid"], _hotkey["thread"]
    _hotkey["tid"] = _hotkey["thread"] = None
    if tid:
        try:
            ctypes.windll.user32.PostThreadMessageW(tid, 0x0012, 0, 0)  # WM_QUIT
        except Exception:
            pass
    if th:
        th.join(1.5)


def registrar_atalho_pausa(texto):
    """Registra o atalho global de pausa. Retorna (ok, mensagem de erro)."""
    parar_atalho_pausa()
    interpretado = interpretar_atalho(texto)
    if not interpretado:
        return False, "Atalho inválido. Use, por exemplo, Ctrl+Alt+P."
    mods, vk = interpretado
    estado = {}
    pronto = threading.Event()

    def laco():
        u32 = ctypes.windll.user32
        estado["tid"] = ctypes.windll.kernel32.GetCurrentThreadId()
        msg = ctypes.wintypes.MSG()
        u32.PeekMessageW(ctypes.byref(msg), None, 0, 0, 0)   # cria a fila de mensagens da thread
        estado["ok"] = bool(u32.RegisterHotKey(None, 1, mods | 0x4000, vk))  # MOD_NOREPEAT
        pronto.set()
        if not estado["ok"]:
            return
        while u32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
            if msg.message == 0x0312:   # WM_HOTKEY
                try:
                    janela_principal.after(0, alternar_pausa_cronometro)
                except Exception:
                    pass
        u32.UnregisterHotKey(None, 1)

    th = threading.Thread(target=laco, daemon=True)
    th.start()
    pronto.wait(3)
    if estado.get("ok"):
        _hotkey["tid"], _hotkey["thread"] = estado["tid"], th
        return True, ""
    return False, (f"A combinação {texto} já está em uso por outro programa. "
                   "Escolha outra em Configurações.")


def definir_atalho_pausa(novo):
    """Guarda o atalho escolhido (por usuário do Windows) e atualiza o menu da bandeja."""
    global atalho_pausa
    atalho_pausa = novo
    gravar_config(atalho_pausa=novo)
    atualizar_menu_bandeja()


def iniciar_cronometro(hwnd, titulo_janela):
    global cronometro_rodando, tempo_inicio, hwnd_monitorado
    global titulo_monitorado, caminho_arquivo_monitorado, hora_inicio_texto
    global cronometro_pausado, pausa_desde, pausa_total, inicio_dt
    cronometro_rodando = True
    hwnd_monitorado = hwnd
    titulo_monitorado = titulo_janela
    caminho_arquivo_monitorado = obter_caminho_do_processo(hwnd)
    tempo_inicio = time.time()
    inicio_dt = datetime.now()
    cronometro_pausado, pausa_desde, pausa_total = False, 0.0, 0.0
    hora_inicio_texto = inicio_dt.strftime("%H:%M")
    if exibir_cronometro:           # preferência do usuário; a contagem do tempo é a mesma
        mostrar_cronometro_flutuante()
    mostrar_borda()
    atualizar_menu_bandeja()


def parar_cronometro_automatico():
    global cronometro_rodando, hwnd_monitorado, titulo_monitorado
    global caminho_arquivo_monitorado, hora_inicio_texto

    global cronometro_pausado, pausa_desde, pausa_total

    # Se terminou durante uma pausa, ela vale até este instante (não conta como trabalho)
    agora = time.time()
    if cronometro_pausado:
        pausa_total += agora - pausa_desde
        cronometro_pausado = False
    trabalhado = (agora - tempo_inicio) - pausa_total
    pausa_min = int(round(pausa_total / 60))

    cronometro_rodando = False
    fechar_aviso_retomada()
    esconder_cronometro_flutuante()
    esconder_borda()

    minutos_gastos = math.ceil(max(trabalhado, 0) / 60)
    if minutos_gastos == 0:
        minutos_gastos = 1

    if pausa_min >= 1 and inicio_dt is not None:
        # Mantém o início real e ajusta o fim: início + tempo efetivamente trabalhado
        hora_fim_texto = (inicio_dt + timedelta(minutes=minutos_gastos)).strftime("%H:%M")
    else:
        pausa_min = 0
        hora_fim_texto = datetime.now().strftime("%H:%M")


    # Abre a tela de lançamento (substitui o envio de e-mail)
    abrir_tela_lancamento(
        titulo=titulo_monitorado,
        hora_inicial=hora_inicio_texto,
        hora_final=hora_fim_texto,
        minutos=minutos_gastos,
        pausa_min=pausa_min
    )
    pausa_total = 0.0

    hwnd_monitorado = None
    titulo_monitorado = ""
    caminho_arquivo_monitorado = ""
    atualizar_menu_bandeja()


# ============================================================
# CAMPO DE SELEÇÃO COM BUSCA (para listas longas, como a de clientes)
# ============================================================
def _norm_busca(texto):
    """Minúsculas e sem acentos, para comparar com o que o usuário digita."""
    s = unicodedata.normalize("NFD", str(texto or ""))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower().strip()


class ComboBusca:
    """Seleção com busca digitada, no estilo do Select2 da web.

    Ao digitar, a lista é filtrada por qualquer trecho do nome (sem diferenciar acentos
    ou maiúsculas), com os que COMEÇAM pelo texto primeiro. Setas + Enter, Esc e clique
    do mouse funcionam. Só aceita itens da lista (texto solto é descartado ao sair).

    API usada pela tela: pack(), set_valores(lista), current() -> índice ou -1,
    limpar() e o atributo ao_selecionar (função chamada quando o usuário escolhe um item).
    """

    def __init__(self, parent, janela, placeholder="Digite para buscar...", linhas=8):
        self.janela = janela
        self.placeholder = placeholder
        self.linhas = linhas
        self.valores = []
        self._visiveis = []          # índices (em self.valores) exibidos na lista
        self._indice = -1            # item escolhido (-1 = nenhum)
        self._ultimo_texto = ""
        self._placeholder_ativo = False
        self._aberta = False         # lista de sugestões aberta?
        self.ao_selecionar = None

        # Caixa do campo (visual parecido com os demais combos)
        self.caixa = tk.Frame(parent, bg="#ffffff", highlightbackground=COR_BORDA, highlightthickness=1)
        self.entry = tk.Entry(self.caixa, bg="#ffffff", fg=COR_TEXTO, insertbackground=COR_TEXTO,
                              relief="flat", bd=0, highlightthickness=0, font=("Segoe UI", 10))
        self.entry.pack(side="left", fill="x", expand=True, ipady=7, padx=(10, 0))
        self.seta = tk.Label(self.caixa, text="▾", bg="#ffffff", fg=COR_PRIMARIA,
                             font=("Segoe UI", 12), cursor="hand2")
        self.seta.pack(side="right", padx=(0, 10))

        # Lista de sugestões: filha da JANELA (não do corpo) para flutuar sobre o formulário
        self.lista_box = tk.Frame(janela, bg=COR_PRIMARIA)
        self.listbox = tk.Listbox(self.lista_box, height=linhas, activestyle="none",
                                  exportselection=False, takefocus=0, relief="flat", bd=0,
                                  highlightthickness=0, bg="#ffffff", fg=COR_TEXTO,
                                  selectbackground=COR_PRIMARIA, selectforeground="#ffffff",
                                  font=("Segoe UI", 10))
        self.scroll = tk.Scrollbar(self.lista_box, orient="vertical", command=self.listbox.yview,
                                   takefocus=0)
        self.listbox.configure(yscrollcommand=self.scroll.set)

        self.entry.bind("<FocusIn>", self._ao_focar)
        self.entry.bind("<FocusOut>", self._agendar_perda_foco)
        self.entry.bind("<KeyRelease>", self._ao_digitar)
        self.entry.bind("<Down>", self._seta_baixo)
        self.entry.bind("<Up>", self._seta_cima)
        self.entry.bind("<Return>", self._enter)
        self.entry.bind("<Escape>", self._esc)
        self.seta.bind("<Button-1>", self._alternar)
        self.listbox.bind("<Button-1>", self._clique_lista)   # "break": não rouba o foco do campo
        self.listbox.bind("<Motion>", self._mouse_sobre_item)
        janela.bind("<Button-1>", self._clique_fora, add="+")

        self._mostrar_placeholder()

    # ---------- API pública ----------
    def pack(self, **kw):
        self.caixa.pack(**kw)

    def set_valores(self, valores):
        self.valores = list(valores)
        self.limpar()

    def current(self):
        return self._indice

    def limpar(self):
        self._indice = -1
        self._esconder_lista()
        self._definir_texto("")

    # ---------- texto / placeholder ----------
    def _tem_foco(self):
        try:
            return self.janela.focus_get() is self.entry
        except Exception:
            return False

    def _texto(self):
        return "" if self._placeholder_ativo else self.entry.get()

    def _mostrar_placeholder(self):
        if not self.entry.get():
            self._placeholder_ativo = True
            self.entry.insert(0, self.placeholder)
            self.entry.config(fg="#a0aec0")

    def _tirar_placeholder(self):
        if self._placeholder_ativo:
            self.entry.delete(0, "end")
            self.entry.config(fg=COR_TEXTO)
            self._placeholder_ativo = False

    def _definir_texto(self, texto):
        self._tirar_placeholder()
        self.entry.delete(0, "end")
        self.entry.insert(0, texto)
        self._ultimo_texto = texto
        if not texto and not self._tem_foco():
            self._mostrar_placeholder()

    # ---------- lista ----------
    def _lista_visivel(self):
        # Marca própria: winfo_ismapped() só muda no próximo ciclo de eventos
        return self._aberta

    def _esconder_lista(self):
        self._aberta = False
        self.lista_box.place_forget()

    def _atualizar_lista(self, texto, tudo=False):
        alvo = "" if tudo else _norm_busca(texto)
        if alvo:
            prefixo, contem = [], []
            for i, v in enumerate(self.valores):
                n = _norm_busca(v)
                if n.startswith(alvo):
                    prefixo.append(i)
                elif alvo in n:
                    contem.append(i)
            self._visiveis = prefixo + contem
        else:
            self._visiveis = list(range(len(self.valores)))

        self.listbox.delete(0, "end")
        pos = None
        if self._visiveis:
            for i in self._visiveis:
                self.listbox.insert("end", self.valores[i])
            pos = 0
            if tudo and self._indice in self._visiveis:
                pos = self._visiveis.index(self._indice)   # destaca o item já escolhido
        else:
            self.listbox.insert("end", "Nenhum resultado encontrado")
            self.listbox.itemconfig(0, fg="#a0aec0")

        # altura da lista = nº de itens (até o máximo) e barra de rolagem só se precisar
        self.listbox.configure(height=max(1, min(self.linhas, len(self._visiveis) or 1)))
        self.scroll.pack_forget()
        self.listbox.pack_forget()
        if len(self._visiveis) > self.linhas:
            self.scroll.pack(side="right", fill="y", padx=(0, 1), pady=1)
        self.listbox.pack(side="left", fill="both", expand=True, padx=1, pady=1)
        self._mostrar_lista()

        # Só agora, com a altura já aplicada, destaca/rola até o item (senão a lista abre rolada)
        self.listbox.update_idletasks()
        if pos is not None:
            self._selecionar_posicao(pos)

    def _mostrar_lista(self):
        self.janela.update_idletasks()
        x = self.caixa.winfo_rootx() - self.janela.winfo_rootx()
        y = self.caixa.winfo_rooty() - self.janela.winfo_rooty() + self.caixa.winfo_height()
        self.lista_box.place(x=x, y=y, width=self.caixa.winfo_width())
        self.lista_box.lift()
        self._aberta = True

    def _selecionar_posicao(self, pos):
        self.listbox.selection_clear(0, "end")
        self.listbox.selection_set(pos)
        self.listbox.activate(pos)
        self.listbox.see(pos)

    # ---------- escolha / saída do campo ----------
    def _escolher(self, i):
        self._indice = i
        self._definir_texto(self.valores[i])
        self._esconder_lista()
        self.entry.icursor("end")
        if self.ao_selecionar:
            self.ao_selecionar()

    def _finalizar_edicao(self):
        """Fecha a lista e garante que o campo mostre um item válido da lista."""
        self._esconder_lista()
        if self._indice >= 0:
            self._definir_texto(self.valores[self._indice])
            return
        alvo = _norm_busca(self._texto())
        if alvo:
            exatos = [i for i, v in enumerate(self.valores) if _norm_busca(v) == alvo]
            if len(exatos) == 1:
                self._escolher(exatos[0])
                return
        self._definir_texto("")   # texto que não corresponde a nenhum cliente: descarta

    # ---------- eventos ----------
    def _ao_focar(self, _e=None):
        self._tirar_placeholder()
        self.caixa.config(highlightbackground=COR_PRIMARIA)
        if self._indice >= 0:
            self.entry.select_range(0, "end")

    def _agendar_perda_foco(self, _e=None):
        try:
            self.janela.after(150, self._ao_perder_foco)
        except tk.TclError:
            pass

    def _ao_perder_foco(self):
        try:
            if self._tem_foco():
                return
            self.caixa.config(highlightbackground=COR_BORDA)
            self._finalizar_edicao()
        except tk.TclError:
            pass   # janela já foi fechada

    def _ao_digitar(self, _e=None):
        if self._placeholder_ativo:
            return
        texto = self.entry.get()
        if texto == self._ultimo_texto:
            return   # tecla de navegação, sem mudança no texto
        self._ultimo_texto = texto
        self._indice = -1
        self._atualizar_lista(texto)

    def _alternar(self, _e=None):
        if self._lista_visivel():
            self._finalizar_edicao()
            return
        self.entry.focus_set()
        self._tirar_placeholder()
        self._atualizar_lista("", tudo=True)

    def _seta_baixo(self, _e=None):
        if not self._lista_visivel():
            self._tirar_placeholder()
            self._atualizar_lista("", tudo=True)
        elif self._visiveis:
            atual = self.listbox.curselection()
            pos = min(len(self._visiveis) - 1, (atual[0] + 1) if atual else 0)
            self._selecionar_posicao(pos)
        return "break"

    def _seta_cima(self, _e=None):
        if self._lista_visivel() and self._visiveis:
            atual = self.listbox.curselection()
            pos = max(0, (atual[0] - 1) if atual else 0)
            self._selecionar_posicao(pos)
        return "break"

    def _enter(self, _e=None):
        if self._lista_visivel() and self._visiveis:
            atual = self.listbox.curselection()
            if atual:
                self._escolher(self._visiveis[atual[0]])
        return "break"

    def _esc(self, _e=None):
        if self._lista_visivel():
            self._finalizar_edicao()
            return "break"

    def _clique_lista(self, e):
        if self._visiveis:
            pos = self.listbox.nearest(e.y)
            if 0 <= pos < len(self._visiveis):
                self._escolher(self._visiveis[pos])
        return "break"

    def _mouse_sobre_item(self, e):
        if self._visiveis:
            self._selecionar_posicao(self.listbox.nearest(e.y))

    def _clique_fora(self, e):
        if not self._lista_visivel():
            return
        w = e.widget
        while w is not None:
            if w is self.caixa or w is self.lista_box:
                return
            w = getattr(w, "master", None)
        self._finalizar_edicao()


# ============================================================
# TELA DE LANÇAMENTO (substitui o e-mail do Outlook)
# ============================================================
def abrir_tela_lancamento(titulo, hora_inicial, hora_final, minutos, pausa_min=0):
    win = tk.Toplevel(janela_principal)
    # Enquanto houver lançamento aberto, a atualização automática espera (não perde o tempo)
    telas_lancamento_abertas.add(win)
    win.bind("<Destroy>", lambda e: telas_lancamento_abertas.discard(win) if e.widget is win else None, add="+")
    win.title("Registrar Atividade - TimeSheet")
    win.configure(bg=COR_FUNDO)
    win.attributes("-topmost", True)
    win.resizable(False, False)
    # Altura ajustada à tela (notebooks menores): nunca passa da área visível
    altura_janela = min(740, win.winfo_screenheight() - 90)
    centralizar(win, 480, altura_janela)
    estilizar_combo(win)

    # Cabeçalho com gradiente roxo
    header = tk.Canvas(win, width=480, height=72, highlightthickness=0, bd=0)
    header.pack(fill="x")
    _desenhar_gradiente(header, 480, 72, "#667eea", "#764ba2")
    header.create_text(20, 36, anchor="w", text="⏱  Registrar Atividade",
                      fill=COR_TEXTO_HEADER, font=("Segoe UI", 14, "bold"))
    header.create_text(460, 26 if pausa_min else 36, anchor="e",
                      text=f"{hora_inicial} - {hora_final}  ({minutos} min)",
                      fill="#eae6f7", font=("Segoe UI", 10))
    if pausa_min:
        header.create_text(460, 48, anchor="e",
                          text=f"⏸ pausa de {pausa_min} min descontada",
                          fill="#fefcbf", font=("Segoe UI", 8))

    # ---- RODAPÉ FIXO com os botões (empacotado ANTES do corpo, para nunca ser cortado) ----
    rodape = tk.Frame(win, bg="#ffffff", highlightbackground=COR_BORDA, highlightthickness=1)
    rodape.pack(side="bottom", fill="x")
    frame_btns = tk.Frame(rodape, bg="#ffffff")
    frame_btns.pack(fill="x", padx=24, pady=14)

    # Mensagens (erros/sucesso) ficam no rodapé fixo, sempre visíveis, acima dos botões
    lbl_status = tk.Label(rodape, text="", bg="#ffffff", fg=COR_ERRO, font=("Segoe UI", 9),
                         wraplength=430, justify="left")
    lbl_status.pack(fill="x", padx=24, pady=(10, 0), before=frame_btns)

    # ---- Corpo ----
    body = tk.Frame(win, bg=COR_FUNDO)
    body.pack(fill="both", expand=True, padx=24, pady=16)

    # Janela detectada (contexto)
    tk.Label(body, text="Janela detectada", bg=COR_FUNDO, fg=COR_TEXTO_SUAVE,
            font=("Segoe UI", 8)).pack(anchor="w")
    tk.Label(body, text=(titulo[:60] + "...") if len(titulo) > 60 else titulo,
            bg=COR_FUNDO, fg=COR_TEXTO, font=("Segoe UI", 9, "italic"),
            wraplength=430, justify="left").pack(anchor="w", pady=(0, 10))

    def label_campo(txt):
        tk.Label(body, text=txt, bg=COR_FUNDO, fg=COR_TEXTO_SUAVE,
                font=("Segoe UI", 9)).pack(anchor="w", pady=(9, 3))

    # Combos
    label_campo("Cliente")
    # Lista de clientes pode ser grande: campo com busca (digite para filtrar)
    combo_cliente = ComboBusca(body, win, placeholder="Digite para buscar o cliente...")
    combo_cliente.pack(fill="x")

    label_campo("Área")
    combo_area = ttk.Combobox(body, state="readonly", style="TS.TCombobox", font=("Segoe UI", 10))
    combo_area.pack(fill="x", ipady=3)

    label_campo("Tipo de registro")
    combo_tipo = ttk.Combobox(body, state="readonly", style="TS.TCombobox", font=("Segoe UI", 10))
    combo_tipo.pack(fill="x", ipady=3)

    label_campo("Solicitante")
    combo_solicitante = ttk.Combobox(body, style="TS.TCombobox", font=("Segoe UI", 10))
    combo_solicitante.pack(fill="x", ipady=3)

    # Horário (editável)
    frame_horas = tk.Frame(body, bg=COR_FUNDO)
    frame_horas.pack(fill="x", pady=(10, 0))
    col1 = tk.Frame(frame_horas, bg=COR_FUNDO); col1.pack(side="left", expand=True, fill="x", padx=(0, 6))
    col2 = tk.Frame(frame_horas, bg=COR_FUNDO); col2.pack(side="left", expand=True, fill="x", padx=(6, 0))
    tk.Label(col1, text="Início", bg=COR_FUNDO, fg=COR_TEXTO_SUAVE, font=("Segoe UI", 9)).pack(anchor="w")
    entry_inicio = tk.Entry(col1, bg="#ffffff", fg=COR_TEXTO, insertbackground=COR_TEXTO,
                           relief="flat", font=("Segoe UI", 11), justify="center",
                           highlightbackground=COR_BORDA, highlightthickness=1, highlightcolor=COR_PRIMARIA)
    entry_inicio.insert(0, hora_inicial); entry_inicio.pack(fill="x", ipady=5)
    tk.Label(col2, text="Fim (ajustado pela pausa)" if pausa_min else "Fim", bg=COR_FUNDO,
            fg=COR_TEXTO_SUAVE, font=("Segoe UI", 9)).pack(anchor="w")
    entry_fim = tk.Entry(col2, bg="#ffffff", fg=COR_TEXTO, insertbackground=COR_TEXTO,
                        relief="flat", font=("Segoe UI", 11), justify="center",
                        highlightbackground=COR_BORDA, highlightthickness=1, highlightcolor=COR_PRIMARIA)
    entry_fim.insert(0, hora_final); entry_fim.pack(fill="x", ipady=5)

    # Descrição
    label_campo("Descrição da atividade")
    txt_descricao = tk.Text(body, bg="#ffffff", fg=COR_TEXTO, insertbackground=COR_TEXTO,
                           relief="flat", font=("Segoe UI", 10), height=3, wrap="word",
                           highlightbackground=COR_BORDA, highlightthickness=1, highlightcolor=COR_PRIMARIA)
    txt_descricao.pack(fill="x")

    # Pré-preenche com uma sugestão automática baseada na janela detectada (sem IA).
    # O usuário pode editar livremente antes de registrar.
    sugestao = gerar_descricao_automatica(titulo)
    if sugestao:
        txt_descricao.insert("1.0", sugestao[:LIMITE_DESCRICAO])
    rodape_desc = tk.Frame(body, bg=COR_FUNDO)
    rodape_desc.pack(fill="x", pady=(3, 0))
    tk.Label(rodape_desc, text="Sugestão automática — edite à vontade antes de registrar.",
            bg=COR_FUNDO, fg=COR_TEXTO_SUAVE, font=("Segoe UI", 8)).pack(side="left")
    lbl_contador = tk.Label(rodape_desc, text="", bg=COR_FUNDO, fg=COR_TEXTO_SUAVE, font=("Segoe UI", 8))
    lbl_contador.pack(side="right")

    def _limitar_descricao(_evt=None):
        # Corta o excedente (digitado ou colado) e mostra o contador
        texto = txt_descricao.get("1.0", "end-1c")
        if len(texto) > LIMITE_DESCRICAO:
            posicao = txt_descricao.index("insert")
            txt_descricao.delete(f"1.0+{LIMITE_DESCRICAO}c", "end")
            try:
                txt_descricao.mark_set("insert", posicao)
            except tk.TclError:
                pass
            texto = txt_descricao.get("1.0", "end-1c")
        lbl_contador.config(text=f"{len(texto)}/{LIMITE_DESCRICAO}",
                            fg="#c0392b" if len(texto) >= LIMITE_DESCRICAO else COR_TEXTO_SUAVE)

    txt_descricao.bind("<KeyRelease>", _limitar_descricao)
    txt_descricao.bind("<<Paste>>", lambda e: body.after(10, _limitar_descricao))
    _limitar_descricao()

    # Guarda os dados carregados
    dados = {"clientes": [], "areas": [], "tipos": []}

    # Carrega os combos a partir da API (em thread para não travar a UI)
    def carregar_dados():
        try:
            dados["clientes"] = api.get_clientes()
            dados["areas"] = api.get_areas()
            dados["tipos"] = api.get_tipos()
            win.after(0, preencher_combos)
        except requests.exceptions.HTTPError as e:
            status = e.response.status_code if e.response is not None else None
            if status == 401:
                msg = ("Sua sessão expirou e não foi possível renovar. Use 'Trocar de usuário' "
                       "no ícone da bandeja para entrar de novo e depois clique em Tentar novamente.")
            elif status == 404:
                msg = ("O TimeSheet ainda não está atualizado para o Monitor. "
                       "Avise o suporte (JRW) e tente novamente depois.")
            else:
                msg = "O TimeSheet não conseguiu enviar os dados. Tente novamente em instantes."
            win.after(0, lambda m=msg: falha_carregamento(m))
        except requests.exceptions.RequestException:
            win.after(0, lambda: falha_carregamento(
                "Sem conexão com o TimeSheet. Verifique sua internet e tente novamente."))
        except Exception:
            win.after(0, lambda: falha_carregamento(
                "Não foi possível carregar os dados do TimeSheet. Tente novamente."))

    def falha_carregamento(msg):
        # Mantém a tela aberta (o tempo cronometrado não se perde) e oferece nova tentativa
        lbl_status.config(fg=COR_ERRO, text=msg)
        btn_salvar.config(text="Tentar novamente", state="normal", command=recarregar)

    def recarregar():
        lbl_status.config(text="")
        btn_salvar.config(text="Carregando...", state="disabled")
        threading.Thread(target=carregar_dados, daemon=True).start()

    def preencher_combos():
        # Carregou com sucesso: devolve o botão à função de registrar
        lbl_status.config(text="")
        btn_salvar.config(text="Registrar no TimeSheet", state="normal", command=salvar)
        combo_cliente.set_valores([c["nome"] for c in dados["clientes"]])
        combo_area["values"] = [a["nome"] for a in dados["areas"]]
        combo_tipo["values"] = [t["nome"] for t in dados["tipos"]]
        if dados["tipos"]:
            combo_tipo.current(0)
        # pré-seleciona a área do usuário logado, se houver
        dep_id = sessao.get("departmentId")
        if dep_id:
            for i, a in enumerate(dados["areas"]):
                if a["id"] == dep_id:
                    combo_area.current(i)
                    break

    def ao_trocar_cliente(_evt=None):
        idx = combo_cliente.current()
        if idx < 0:
            return
        cliente_id = dados["clientes"][idx]["id"]
        def buscar():
            try:
                resp = api.get_solicitantes(cliente_id)
                nomes = resp.get("solicitantes", [])
                padrao = resp.get("solicitantePadrao", "")
                def aplicar():
                    combo_solicitante["values"] = nomes
                    combo_solicitante.set(padrao or "")
                win.after(0, aplicar)
            except Exception:
                pass
        threading.Thread(target=buscar, daemon=True).start()

    combo_cliente.ao_selecionar = ao_trocar_cliente

    def _hora_para_iso(txt):
        """Converte 'HH:MM' em 'HH:MM:00'. Retorna None se inválido."""
        txt = txt.strip()
        partes = txt.replace("h", ":").split(":")
        try:
            h = int(partes[0]); m = int(partes[1]) if len(partes) > 1 and partes[1] != "" else 0
            if 0 <= h <= 23 and 0 <= m <= 59:
                return f"{h:02d}:{m:02d}:00"
        except (ValueError, IndexError):
            pass
        return None

    def salvar():
        ic = combo_cliente.current()
        ia = combo_area.current()
        it = combo_tipo.current()
        if ic < 0:
            lbl_status.config(text="Selecione o cliente na lista."); return
        if ia < 0:
            lbl_status.config(text="Selecione a área."); return
        descricao = txt_descricao.get("1.0", "end").strip()
        if len(descricao) < 10:
            lbl_status.config(text="A descrição deve ter pelo menos 10 caracteres."); return
        if len(descricao) > LIMITE_DESCRICAO:
            lbl_status.config(text=f"A descrição deve ter no máximo {LIMITE_DESCRICAO} caracteres."); return
        ini = _hora_para_iso(entry_inicio.get())
        fim = _hora_para_iso(entry_fim.get())
        if not ini or not fim:
            lbl_status.config(text="Horário inválido. Use o formato HH:MM."); return

        payload = {
            "date": datetime.now().strftime("%Y-%m-%dT00:00:00"),
            "clientId": dados["clientes"][ic]["id"],
            "departmentId": dados["areas"][ia]["id"],
            "horaInicial": ini,
            "horaFinal": fim,
            "description": descricao,
            "solicitante": combo_solicitante.get().strip(),
            "recordType": dados["tipos"][it]["id"] if it >= 0 else 0
        }

        btn_salvar.config(text="Salvando...", state="disabled")
        win.update()

        def enviar():
            ok, msg = api.registrar_lancamento(payload)
            def resultado():
                if ok:
                    lbl_status.config(fg=COR_SUCESSO, text="✓ " + msg)
                    win.after(900, win.destroy)
                else:
                    lbl_status.config(fg=COR_ERRO, text=msg)
                    btn_salvar.config(text="Registrar no TimeSheet", state="normal")
            win.after(0, resultado)
        threading.Thread(target=enviar, daemon=True).start()

    # Botões no rodapé fixo (criado lá em cima, sempre visível)
    btn_salvar = criar_botao(frame_btns, "Registrar no TimeSheet", salvar, primario=True)
    btn_salvar.pack(side="right")
    criar_botao(frame_btns, "Descartar", win.destroy, primario=False).pack(side="right", padx=(0, 10))

    threading.Thread(target=carregar_dados, daemon=True).start()


# ============================================================
# TELA DE CONFIGURAÇÕES (programas monitorados)
# ============================================================
def abrir_configuracoes():
    """Tela para o usuário escolher quais programas o monitor deve observar."""
    win = tk.Toplevel(janela_principal)
    win.title("Configurações - TimeSheet")
    win.configure(bg=COR_FUNDO)
    win.attributes("-topmost", True)
    win.resizable(False, False)
    aplicar_icone_janela(win)
    # Altura ajustada à tela (notebooks menores): nunca passa da área visível
    centralizar(win, 460, min(800, win.winfo_screenheight() - 60))

    # Cabeçalho com gradiente roxo
    header = tk.Canvas(win, width=440, height=66, highlightthickness=0, bd=0)
    header.pack(fill="x")
    _desenhar_gradiente(header, 440, 66, "#667eea", "#764ba2")
    header.create_text(20, 33, anchor="w", text="⚙  Configurações",
                      fill=COR_TEXTO_HEADER, font=("Segoe UI", 14, "bold"))

    # Rodapé com os botões, fixo na base (empacotado antes do corpo para nunca ser cortado)
    frame_btns = tk.Frame(win, bg=COR_FUNDO)
    frame_btns.pack(side="bottom", fill="x", padx=24, pady=(0, 18))

    body = tk.Frame(win, bg=COR_FUNDO)
    body.pack(fill="both", expand=True, padx=24, pady=18)

    # ---------- Exibição ----------
    tk.Label(body, text="Exibição", bg=COR_FUNDO, fg=COR_TEXTO,
            font=("Segoe UI", 11, "bold")).pack(anchor="w")
    card_exibicao = tk.Frame(body, bg="#ffffff", highlightbackground=COR_BORDA, highlightthickness=1)
    card_exibicao.pack(fill="x", pady=(6, 16))
    var_cronometro = tk.BooleanVar(value=exibir_cronometro)
    tk.Checkbutton(
        card_exibicao, text="  Mostrar o cronômetro flutuante no canto da tela",
        variable=var_cronometro, bg="#ffffff", fg=COR_TEXTO, selectcolor="#ffffff",
        activebackground="#ffffff", activeforeground=COR_PRIMARIA,
        font=("Segoe UI", 10), anchor="w", cursor="hand2", highlightthickness=0, bd=0
    ).pack(fill="x", padx=16, pady=(10, 0))
    var_retomada = tk.BooleanVar(value=aviso_retomada_ativo)
    tk.Checkbutton(
        card_exibicao, text="  Avisar quando eu voltar a trabalhar durante uma pausa",
        variable=var_retomada, bg="#ffffff", fg=COR_TEXTO, selectcolor="#ffffff",
        activebackground="#ffffff", activeforeground=COR_PRIMARIA,
        font=("Segoe UI", 10), anchor="w", cursor="hand2", highlightthickness=0, bd=0
    ).pack(fill="x", padx=16, pady=(4, 0))
    var_inicio = tk.BooleanVar(value=iniciar_com_windows_ativo())
    tk.Checkbutton(
        card_exibicao, text="  Abrir o Monitor junto com o Windows",
        variable=var_inicio, bg="#ffffff", fg=COR_TEXTO, selectcolor="#ffffff",
        activebackground="#ffffff", activeforeground=COR_PRIMARIA,
        font=("Segoe UI", 10), anchor="w", cursor="hand2", highlightthickness=0, bd=0
    ).pack(fill="x", padx=16, pady=(4, 6))

    # Tecla de atalho da pausa (cada usuário escolhe a sua; fica no config do próprio usuário)
    linha_atalho = tk.Frame(card_exibicao, bg="#ffffff")
    linha_atalho.pack(fill="x", padx=16, pady=(0, 10))
    tk.Label(linha_atalho, text="Atalho para pausar (clique e tecle):", bg="#ffffff", fg=COR_TEXTO,
            font=("Segoe UI", 10)).pack(side="left")
    var_atalho = tk.StringVar(value=atalho_pausa)
    entry_atalho = tk.Entry(
        linha_atalho, textvariable=var_atalho, state="readonly", readonlybackground="#f7fafc",
        fg=COR_TEXTO, relief="flat", font=("Segoe UI", 10, "bold"), justify="center", width=13,
        highlightbackground=COR_BORDA, highlightthickness=1, highlightcolor=COR_PRIMARIA, cursor="hand2")
    entry_atalho.pack(side="left", padx=(8, 0), ipady=3)
    btn_padrao = criar_botao(linha_atalho, "Padrão", lambda: var_atalho.set(ATALHO_PADRAO), primario=False)
    btn_padrao.config(padx=8, pady=2, font=("Segoe UI", 9))
    btn_padrao.pack(side="left", padx=(6, 0))

    _TECLAS_MODIFICADORAS = {"Control_L", "Control_R", "Alt_L", "Alt_R", "Shift_L", "Shift_R",
                             "Win_L", "Win_R", "Caps_Lock", "Num_Lock", "Meta_L", "Meta_R",
                             "Super_L", "Super_R"}

    def capturar_atalho(e):
        if e.keysym == "Tab" and not (e.state & 0x4):
            return None                       # deixa o Tab mudar de campo
        if e.keysym in _TECLAS_MODIFICADORAS:
            return "break"
        mods = []
        if e.state & 0x4:
            mods.append("Ctrl")
        if e.state & 0x20000:
            mods.append("Alt")
        if e.state & 0x1:
            mods.append("Shift")
        k = e.keysym
        tecla = k.upper() if (len(k) == 1 and k.isalnum()) else (k if re.fullmatch(r"F\d{1,2}", k) else None)
        candidato = "+".join(mods + [tecla]) if tecla else ""
        if candidato and interpretar_atalho(candidato):
            var_atalho.set(candidato)
            lbl_status.config(text="")
        else:
            lbl_status.config(fg=COR_ERRO, text="Combinação não aceita. Use Ctrl ou Alt com uma "
                                                "letra ou número, ou uma tecla de F1 a F12.")
        return "break"

    entry_atalho.bind("<KeyPress>", capturar_atalho)
    entry_atalho.bind("<Button-1>", lambda e: entry_atalho.focus_set())
    # Enquanto o campo está em foco, o atalho atual é liberado (senão ele dispararia a pausa
    # em vez de ser capturado); ao sair do campo, volta a valer.
    entry_atalho.bind("<FocusIn>", lambda e: parar_atalho_pausa())
    entry_atalho.bind("<FocusOut>", lambda e: registrar_atalho_pausa(atalho_pausa))
    win.bind("<Destroy>", lambda e: (e.widget is win) and registrar_atalho_pausa(atalho_pausa))

    # ---------- Programas monitorados ----------
    tk.Label(body, text="Programas monitorados", bg=COR_FUNDO, fg=COR_TEXTO,
            font=("Segoe UI", 11, "bold")).pack(anchor="w")
    tk.Label(body, text="Selecione quais aplicativos o monitor deve observar para sugerir o registro de horas.",
            bg=COR_FUNDO, fg=COR_TEXTO_SUAVE, font=("Segoe UI", 9),
            wraplength=390, justify="left").pack(anchor="w", pady=(2, 8))

    # Cartão com os checkboxes, em duas colunas
    card = tk.Frame(body, bg="#ffffff", highlightbackground=COR_BORDA, highlightthickness=1)
    card.pack(fill="x")
    card.grid_columnconfigure(0, weight=1, uniform="prog")
    card.grid_columnconfigure(1, weight=1, uniform="prog")

    vars_check = {}
    for n, (rotulo, termo) in enumerate(PROGRAMAS_DISPONIVEIS):
        var = tk.BooleanVar(value=(termo in palavras_alvo))
        vars_check[termo] = var
        cb = tk.Checkbutton(
            card, text="  " + rotulo, variable=var,
            bg="#ffffff", fg=COR_TEXTO, selectcolor="#ffffff",
            activebackground="#ffffff", activeforeground=COR_PRIMARIA,
            font=("Segoe UI", 10), anchor="w", cursor="hand2",
            highlightthickness=0, bd=0
        )
        cb.grid(row=n // 2, column=n % 2, sticky="w", padx=(16, 8), pady=4)

    # ---------- Outros programas (personalizados) ----------
    termos_catalogo = {termo for _, termo in PROGRAMAS_DISPONIVEIS}
    # Personalizados = o que está monitorado e não pertence ao catálogo
    personalizados = [t for t in palavras_alvo if t not in termos_catalogo]

    tk.Label(body, text="Outros programas", bg=COR_FUNDO, fg=COR_TEXTO,
            font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(16, 0))
    tk.Label(body, text="Digite uma palavra que apareça no título da janela do programa (ex.: autocad, power bi).",
            bg=COR_FUNDO, fg=COR_TEXTO_SUAVE, font=("Segoe UI", 9),
            wraplength=390, justify="left").pack(anchor="w", pady=(2, 8))

    # Campo + botão adicionar
    linha_add = tk.Frame(body, bg=COR_FUNDO)
    linha_add.pack(fill="x")
    entry_novo = tk.Entry(linha_add, bg="#ffffff", fg=COR_TEXTO, insertbackground=COR_TEXTO,
                         relief="flat", font=("Segoe UI", 10),
                         highlightbackground=COR_BORDA, highlightthickness=1, highlightcolor=COR_PRIMARIA)
    entry_novo.pack(side="left", fill="x", expand=True, ipady=4)

    # Lista dos personalizados (com remover)
    lista_frame = tk.Frame(body, bg=COR_FUNDO)
    lista_frame.pack(fill="x", pady=(8, 0))

    def render_personalizados():
        for w in lista_frame.winfo_children():
            w.destroy()
        for termo in personalizados:
            chip = tk.Frame(lista_frame, bg="#eef1fb")
            chip.pack(side="left", padx=(0, 6), pady=3)
            tk.Label(chip, text=termo, bg="#eef1fb", fg=COR_PRIMARIA_ESCURA,
                    font=("Segoe UI", 9)).pack(side="left", padx=(8, 4), pady=3)
            btn_x = tk.Label(chip, text="✕", bg="#eef1fb", fg="#a0aec0",
                           font=("Segoe UI", 9, "bold"), cursor="hand2")
            btn_x.pack(side="left", padx=(0, 6))
            btn_x.bind("<Button-1>", lambda e, t=termo: remover_personalizado(t))

    def adicionar_personalizado():
        novo = entry_novo.get().strip().lower()
        if not novo:
            return
        if novo in termos_catalogo:
            lbl_status.config(fg=COR_ERRO, text=f'"{novo}" já existe na lista acima. Basta marcá-lo.')
            return
        if novo in personalizados:
            lbl_status.config(fg=COR_ERRO, text=f'"{novo}" já foi adicionado.')
            return
        personalizados.append(novo)
        entry_novo.delete(0, "end")
        lbl_status.config(text="")
        render_personalizados()

    def remover_personalizado(termo):
        if termo in personalizados:
            personalizados.remove(termo)
            render_personalizados()

    criar_botao(linha_add, "Adicionar", adicionar_personalizado, primario=False).pack(side="left", padx=(8, 0))
    entry_novo.bind("<Return>", lambda e: adicionar_personalizado())
    render_personalizados()

    lbl_status = tk.Label(body, text="", bg=COR_FUNDO, fg=COR_SUCESSO, font=("Segoe UI", 9),
                         wraplength=390, justify="left")
    lbl_status.pack(anchor="w", pady=(12, 0))

    def salvar_config():
        marcados = [termo for termo, var in vars_check.items() if var.get()]
        selecionados = marcados + personalizados
        if not selecionados:
            lbl_status.config(fg=COR_ERRO, text="Selecione ou adicione ao menos um programa.")
            return
        if var_atalho.get() != atalho_pausa:
            ok_atalho, msg_atalho = registrar_atalho_pausa(var_atalho.get())
            if not ok_atalho:
                registrar_atalho_pausa(atalho_pausa)      # mantém o atalho anterior
                lbl_status.config(fg=COR_ERRO, text=msg_atalho)
                return
            definir_atalho_pausa(var_atalho.get())
        if var_inicio.get() != iniciar_com_windows_ativo():
            if not definir_iniciar_com_windows(var_inicio.get()):
                lbl_status.config(fg=COR_ERRO, text="Não foi possível alterar a abertura com o Windows.")
                return
        if (salvar_palavras_alvo(selecionados)
                and definir_exibir_cronometro(var_cronometro.get())
                and definir_aviso_retomada(var_retomada.get())):
            lbl_status.config(fg=COR_SUCESSO, text="✓ Configurações salvas. Já valem para as próximas detecções.")
            win.after(1200, win.destroy)
        else:
            lbl_status.config(fg=COR_ERRO, text="Não foi possível salvar as configurações.")

    # Botões (o rodapé fixo foi criado lá em cima)
    criar_botao(frame_btns, "Salvar", salvar_config, primario=True).pack(side="right")
    criar_botao(frame_btns, "Fechar", win.destroy, primario=False).pack(side="right", padx=(0, 10))


# ============================================================
# TELA DE AJUDA (instruções de uso, com prints embutidos)
# ============================================================
_ajuda_imgs = []  # mantém referência das imagens para não serem coletadas


def abrir_tela_ajuda():
    win = tk.Toplevel(janela_principal)
    win.title("Ajuda - Como usar o Monitor")
    win.configure(bg=COR_FUNDO)
    win.resizable(False, True)
    aplicar_icone_janela(win)
    centralizar(win, 560, 720)

    # Cabeçalho com gradiente roxo
    header = tk.Canvas(win, width=560, height=70, highlightthickness=0, bd=0)
    header.pack(fill="x")
    _desenhar_gradiente(header, 560, 70, "#667eea", "#764ba2")
    header.create_text(24, 35, anchor="w", text="❔  Como usar o Monitor de Produtividade",
                      fill=COR_TEXTO_HEADER, font=("Segoe UI", 13, "bold"))

    # Área rolável (Canvas + Scrollbar)
    container = tk.Frame(win, bg=COR_FUNDO)
    container.pack(fill="both", expand=True)
    canvas = tk.Canvas(container, bg=COR_FUNDO, highlightthickness=0)
    scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
    conteudo = tk.Frame(canvas, bg=COR_FUNDO)
    conteudo.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=conteudo, anchor="nw", width=540)
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True, padx=(16, 0))
    scrollbar.pack(side="right", fill="y")

    # rolagem com a roda do mouse
    def _scroll(e):
        canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
    canvas.bind_all("<MouseWheel>", _scroll)
    win.protocol("WM_DELETE_WINDOW", lambda: (canvas.unbind_all("<MouseWheel>"), win.destroy()))

    def titulo(txt):
        tk.Label(conteudo, text=txt, bg=COR_FUNDO, fg=COR_PRIMARIA_ESCURA,
                font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(18, 4), padx=8)

    def paragrafo(txt):
        tk.Label(conteudo, text=txt, bg=COR_FUNDO, fg=COR_TEXTO, font=("Segoe UI", 10),
                wraplength=500, justify="left").pack(anchor="w", padx=8, pady=(0, 6))

    def imagem(nome, larg=480):
        try:
            img = Image.open(caminho_asset(nome))
            # redimensiona mantendo proporção
            w, h = img.size
            if w > larg:
                img = img.resize((larg, int(h * larg / w)), Image.LANCZOS)
            foto = ImageTk.PhotoImage(img)
            _ajuda_imgs.append(foto)
            lbl = tk.Label(conteudo, image=foto, bg=COR_FUNDO, bd=1, relief="solid")
            lbl.pack(anchor="w", padx=8, pady=(2, 10))
        except Exception:
            pass

    # ---- Conteúdo ----
    paragrafo("Este guia mostra, passo a passo, como usar o Monitor de Produtividade "
             "para registrar suas horas automaticamente no TimeSheet.")

    titulo("1. Fazer login")
    paragrafo("Ao abrir o aplicativo, entre com o mesmo usuário e senha que você usa no TimeSheet.")
    imagem("TelaLogin.png", 300)
    paragrafo("Mantenha marcada a opção 'Manter conectado neste computador': nas próximas vezes o "
             "Monitor abre junto com o Windows e já entra sozinho, sem pedir usuário e senha. "
             "Se a sua senha for alterada no TimeSheet, ele pede o login de novo. "
             "O usuário e a senha ficam guardados só neste computador, protegidos pela criptografia do Windows.")

    titulo("2. Trabalhe normalmente")
    paragrafo("Ao abrir um programa monitorado (Word, Excel, PDF, navegador...), o app pergunta "
             "se deseja iniciar o cronômetro.")
    imagem("Telaqueperguntasedesejamonitorar.png")
    paragrafo("Confirmando, um cronômetro discreto aparece no canto inferior direito, "
             "acompanhando o tempo. Passe o mouse sobre ele para ver o que está sendo monitorado. "
             "Se preferir não vê-lo, desmarque a opção em Configurações: o tempo continua sendo contado.")
    imagem("cronometro.png", 220)

    titulo("3. Registrar a atividade")
    paragrafo("Ao fechar a janela (ou usar 'Parar e registrar' na bandeja), abre-se a tela de "
             "lançamento. Selecione cliente, área, tipo e solicitante — a descrição já vem sugerida. "
             "Confira e clique em 'Registrar no TimeSheet'.")
    imagem("Salnadoatividade.png", 320)

    titulo("4. Ícone na bandeja")
    paragrafo("O aplicativo fica na bandeja do Windows, perto do relógio.")
    imagem("iconebandeija.png", 300)
    paragrafo("Clicando com o botão direito, você acessa as opções: parar e registrar, pausar o "
             "monitoramento, configurações, esta ajuda, verificar atualizações, trocar de usuário e sair. "
             "Quando houver uma nova versão, o Monitor avisa ao abrir e atualiza sozinho, se você aceitar. "
             "Em computadores compartilhados, use 'Trocar de usuário' para entrar com outra conta.")
    imagem("opcoesbandeija.png", 300)

    titulo("5. Escolher o que monitorar")
    paragrafo("Em 'Configurações', marque os programas que deseja monitorar. Você também pode "
             "adicionar outros programas digitando uma palavra do título da janela.")
    imagem("configuraesAppmonitorados.png", 360)

    titulo("6. Pausar o cronômetro")
    paragrafo(f"Atendeu uma ligação ou foi tratar de outro assunto? Pressione {atalho_pausa} (em qualquer "
             "programa) para pausar o cronômetro e a mesma tecla para retomar. Enquanto estiver pausado, "
             "o cronômetro e a borda ficam amarelos e o tempo não conta. Também dá para pausar e retomar "
             "pelo ícone da bandeja.")
    paragrafo("Ao registrar, o horário de início continua o real e o de fim é ajustado, descontando o "
             "tempo pausado. A tela avisa quantos minutos foram descontados, e você pode editar o horário. "
             "Se a janela for fechada durante uma pausa, o tempo pausado também não é contado. "
             "A tecla de atalho pode ser trocada em Configurações.")
    paragrafo("Esqueceu de retomar? Se você voltar a trabalhar no programa pausado, o Monitor pergunta, "
             "num aviso no canto da tela (sem atrapalhar a digitação), se deseja retomar o cronômetro. "
             "Se abrir outro programa monitorado durante a pausa, ele não inicia uma segunda contagem: "
             "pergunta se você quer continuar a atividade pausada, encerrá-la e iniciar uma nova, ou "
             "manter a pausa. Se escolher continuar pausado, ele não pergunta de novo por 10 minutos. "
             "Esse aviso pode ser desligado em Configurações.")
    imagem("TelaPerguntapausa.png", 340)

    # Botão fechar
    rodape = tk.Frame(win, bg=COR_FUNDO)
    rodape.pack(fill="x", pady=10)
    criar_botao(rodape, "Fechar", lambda: (canvas.unbind_all("<MouseWheel>"), win.destroy()),
               primario=True).pack()


# ============================================================
# BANDEJA
# ============================================================
# ============================================================
# ATUALIZAÇÃO AUTOMÁTICA DO MONITOR
# ============================================================
# Ao abrir (e a cada 6 h, pois o Monitor fica aberto o dia todo), consulta a versão no
# TimeSheet. Havendo uma nova, pergunta no máximo uma vez por dia. "Atualizar agora" baixa o
# instalador pela API (só com login), confere o tamanho e o SHA-256, fecha o Monitor, instala
# em silêncio e o instalador abre o Monitor de novo (o login guardado continua valendo).
ATUALIZACAO_PRIMEIRA_MS = 8000
ATUALIZACAO_RECHECAR_MS = 6 * 60 * 60 * 1000
ATUALIZACAO_ADIAR_OCUPADO_MS = 10 * 60 * 1000   # atividade em andamento: tenta de novo em 10 min

telas_lancamento_abertas = set()
_atualizacao = {"janela": None, "baixando": False}


def _versao_tupla(v):
    try:
        return tuple(int(p) for p in str(v).strip().split("."))
    except Exception:
        return (0,)


def _ocupado_com_atividade():
    return cronometro_rodando or bool(telas_lancamento_abertas)


def verificar_atualizacao(manual=False):
    """Consulta a versão no servidor (em segundo plano). 'manual' = pedido pela bandeja."""
    if not manual:
        janela_principal.after(ATUALIZACAO_RECHECAR_MS, verificar_atualizacao)
    if _atualizacao["janela"] is not None or _atualizacao["baixando"]:
        return
    if not (sessao.get("token") or api._login):
        return

    def trabalho():
        try:
            info = api.get_versao()
        except Exception:
            if manual:
                janela_principal.after(0, lambda: messagebox.showwarning(
                    "Atualização", "Não foi possível verificar agora. Confira sua conexão e tente novamente."))
            return
        nova = info.get("disponivel") and _versao_tupla(info.get("versao")) > _versao_tupla(APP_VERSAO)
        if nova:
            janela_principal.after(0, lambda: oferecer_atualizacao(info, manual))
        elif manual:
            janela_principal.after(0, lambda: messagebox.showinfo(
                "Atualização", f"Você já está usando a versão mais recente ({APP_VERSAO})."))

    threading.Thread(target=trabalho, daemon=True).start()


def oferecer_atualizacao(info, manual=False):
    versao_nova = str(info.get("versao"))
    if not manual:
        adiada = carregar_configuracoes().get("atualizacao_adiada") or {}
        if adiada.get("versao") == versao_nova and adiada.get("data") == datetime.now().strftime("%Y-%m-%d"):
            return                                   # já perguntou hoje sobre esta versão
        if _ocupado_com_atividade():
            # Nunca interrompe uma atividade: pergunta quando ela terminar
            janela_principal.after(ATUALIZACAO_ADIAR_OCUPADO_MS, lambda: oferecer_atualizacao(info))
            return
    if _atualizacao["janela"] is not None:
        return

    win = tk.Toplevel(janela_principal)
    _atualizacao["janela"] = win
    win.title("Atualização - Monitor de Produtividade")
    win.configure(bg=COR_FUNDO)
    win.attributes("-topmost", True)
    win.resizable(False, False)
    aplicar_icone_janela(win)
    centralizar(win, 440, 300)

    header = tk.Canvas(win, width=440, height=62, highlightthickness=0, bd=0)
    header.pack(fill="x")
    _desenhar_gradiente(header, 440, 62, "#667eea", "#764ba2")
    header.create_text(20, 31, anchor="w", text="⬆  Nova versão disponível",
                       fill=COR_TEXTO_HEADER, font=("Segoe UI", 13, "bold"))

    rodape = tk.Frame(win, bg=COR_FUNDO)
    rodape.pack(side="bottom", fill="x", padx=24, pady=(0, 18))
    corpo = tk.Frame(win, bg=COR_FUNDO)
    corpo.pack(fill="both", expand=True, padx=24, pady=16)

    tk.Label(corpo, text=f"Versão instalada: {APP_VERSAO}     →     Nova versão: {versao_nova}",
             bg=COR_FUNDO, fg=COR_TEXTO, font=("Segoe UI", 10, "bold")).pack(anchor="w")
    tk.Label(corpo, text=("A atualização leva cerca de um minuto. O Monitor fecha, instala a nova versão "
                          "e abre de novo sozinho."),
             bg=COR_FUNDO, fg=COR_TEXTO_SUAVE, font=("Segoe UI", 9), wraplength=390,
             justify="left").pack(anchor="w", pady=(8, 10))
    lbl_status = tk.Label(corpo, text="", bg=COR_FUNDO, fg=COR_TEXTO_SUAVE, font=("Segoe UI", 9),
                          wraplength=390, justify="left")
    lbl_status.pack(anchor="w")
    barra = ttk.Progressbar(corpo, mode="determinate", length=390, maximum=100)

    def fechar():
        _atualizacao["janela"] = None
        try:
            win.destroy()
        except Exception:
            pass

    def depois():
        if _atualizacao["baixando"]:
            return
        gravar_config(atualizacao_adiada={"versao": versao_nova, "data": datetime.now().strftime("%Y-%m-%d")})
        fechar()

    def atualizar():
        if _atualizacao["baixando"]:
            return
        if _ocupado_com_atividade():
            lbl_status.config(fg=COR_ERRO, text="Há uma atividade em andamento. Registre-a antes de atualizar.")
            return
        _atualizacao["baixando"] = True
        btn_atualizar.config(state="disabled", text="Baixando...")
        btn_depois.config(state="disabled")
        barra.pack(anchor="w", pady=(4, 0))
        lbl_status.config(fg=COR_TEXTO_SUAVE, text="Baixando a atualização...")
        threading.Thread(target=_baixar_e_instalar, args=(info, win, lbl_status, barra, falhou),
                         daemon=True).start()

    def falhou(msg):
        _atualizacao["baixando"] = False
        barra.pack_forget()
        lbl_status.config(fg=COR_ERRO, text=msg)
        btn_atualizar.config(state="normal", text="Tentar novamente")
        btn_depois.config(state="normal")

    btn_atualizar = criar_botao(rodape, "Atualizar agora", atualizar, primario=True)
    btn_atualizar.pack(side="right")
    btn_depois = criar_botao(rodape, "Depois", depois, primario=False)
    btn_depois.pack(side="right", padx=(0, 10))
    win.protocol("WM_DELETE_WINDOW", depois)


def _baixar_e_instalar(info, win, lbl_status, barra, falhou):
    """Roda em thread: baixa, confere e dispara a instalação silenciosa."""
    import hashlib as _hl
    import subprocess
    import tempfile

    def na_tela(func):
        try:
            janela_principal.after(0, func)
        except Exception:
            pass

    destino = os.path.join(tempfile.gettempdir(), info.get("arquivo") or f"MonitorTimeSheet-Setup-{info.get('versao')}.exe")
    ultimo = {"pct": -1}

    def progresso(baixado, total):
        total = total or info.get("tamanho") or 0
        pct = int(baixado * 100 / total) if total else 0
        if pct != ultimo["pct"]:
            ultimo["pct"] = pct
            txt = f"Baixando... {baixado / 1048576:.1f} de {total / 1048576:.1f} MB"
            na_tela(lambda: (barra.config(value=pct), lbl_status.config(text=txt)))

    try:
        api.baixar_instalador(destino, progresso)
    except Exception:
        na_tela(lambda: falhou("Não foi possível baixar a atualização. Confira sua conexão e tente novamente."))
        return

    # Confere se o arquivo chegou inteiro e é exatamente o publicado no TimeSheet
    try:
        h = _hl.sha256()
        with open(destino, "rb") as f:
            for bloco in iter(lambda: f.read(1024 * 1024), b""):
                h.update(bloco)
        tamanho_ok = not info.get("tamanho") or os.path.getsize(destino) == int(info["tamanho"])
        hash_ok = not info.get("sha256") or h.hexdigest().lower() == str(info["sha256"]).lower()
    except Exception:
        tamanho_ok = hash_ok = False
    if not (tamanho_ok and hash_ok):
        try:
            os.remove(destino)
        except Exception:
            pass
        na_tela(lambda: falhou("O arquivo baixado veio incompleto ou diferente do publicado. Tente novamente."))
        return

    if not getattr(sys, "frozen", False):
        # Rodando o .py (desenvolvimento): não instala por cima do app instalado
        def aviso_dev():
            _atualizacao["baixando"] = False
            barra.config(value=100)
            lbl_status.config(fg=COR_SUCESSO, text=f"Modo de desenvolvimento: instalador baixado e conferido em\n{destino}\n(não foi executado).")
        na_tela(aviso_dev)
        return

    # Espera o Monitor fechar (libera o arquivo e a instância única) e instala em silêncio.
    # O instalador abre o Monitor de novo ao terminar (entrada [Run] com WizardSilent).
    comando = (f'cmd /c ping -n 4 127.0.0.1 >nul & "{destino}" '
               f'/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-')
    # O instalador (e o Monitor que ele reabre) herdam o ambiente deste processo. Sem limpar as
    # variáveis internas do PyInstaller, o Monitor reaberto se acha "subprocesso" deste e falha com
    # "Security validation failure". PYINSTALLER_RESET_ENVIRONMENT é a forma oficial (PyInstaller 6.10+).
    ambiente = {k: v for k, v in os.environ.items()
                if not k.upper().startswith(("_PYI_", "_MEIPASS"))}
    ambiente["PYINSTALLER_RESET_ENVIRONMENT"] = "1"
    try:
        subprocess.Popen(comando, creationflags=0x08000000, env=ambiente)   # CREATE_NO_WINDOW
    except Exception:
        na_tela(lambda: falhou("Não foi possível iniciar a instalação. Tente novamente."))
        return

    def encerrar_para_atualizar():
        lbl_status.config(fg=COR_SUCESSO, text="Instalando... o Monitor vai abrir de novo em instantes.")
        win.update()
        try:
            parar_atalho_pausa()
            if icone_bandeja:
                icone_bandeja.stop()
        except Exception:
            pass
        janela_principal.after(800, lambda: os._exit(0))
    na_tela(encerrar_para_atualizar)


def verificar_atualizacao_bandeja(icon, item):
    janela_principal.after(0, lambda: verificar_atualizacao(manual=True))


def abrir_config_thread(icon, item):
    """Abre a tela de config na thread do tkinter (a bandeja roda em outra thread)."""
    janela_principal.after(0, abrir_configuracoes)


def abrir_ajuda(icon, item):
    """Abre a tela de ajuda (instruções de uso) na thread do tkinter."""
    janela_principal.after(0, abrir_tela_ajuda)


def alternar_pausa(icon, item):
    global monitoramento_pausado
    monitoramento_pausado = not monitoramento_pausado
    atualizar_menu_bandeja()


def parar_e_registrar(icon, item):
    """Para manualmente o cronômetro em andamento e abre a tela de lançamento.
    Útil quando a janela não 'fecha' (ex.: troca de aba no navegador)."""
    if cronometro_rodando:
        # Executa na thread do tkinter (a bandeja roda em outra thread)
        janela_principal.after(0, parar_cronometro_automatico)


_login_aberto = {"v": False}


def _abrir_login(aviso=None, login_inicial=""):
    """Abre a tela de login (uma só por vez). Retorna True se o usuário entrou."""
    if _login_aberto["v"]:
        return False
    _login_aberto["v"] = True
    try:
        ok = tela_login(janela_principal, aviso=aviso, login_inicial=login_inicial)
    finally:
        _login_aberto["v"] = False
    atualizar_menu_bandeja()
    return ok


def trocar_usuario(icon, item):
    """Bandeja > Trocar de usuário: abre o login; ao entrar, o login guardado é substituído."""
    janela_principal.after(0, _trocar_usuario)


def _trocar_usuario():
    if cronometro_rodando:
        messagebox.showinfo(
            "Trocar de usuário",
            "Há uma atividade em andamento.\n\nUse 'Parar e registrar' na bandeja antes de trocar de usuário.")
        return
    _abrir_login()


def credenciais_recusadas(login_antigo):
    """O servidor recusou o login guardado (ex.: senha alterada). Pede o login de novo, com aviso."""
    janela_principal.after(0, lambda: _abrir_login(
        aviso="Sua senha foi alterada no TimeSheet. Entre novamente para continuar.",
        login_inicial=login_antigo or ""))


def tentar_reconectar():
    """Sem rede na abertura (ex.: Windows ainda conectando): tenta entrar de novo a cada 30 s."""
    if sessao.get("token") or not api._login:
        return

    def trabalho():
        if api.renovar_sessao():
            janela_principal.after(0, atualizar_menu_bandeja)
        elif api.ultimo_erro != "credenciais":
            janela_principal.after(30000, tentar_reconectar)

    threading.Thread(target=trabalho, daemon=True).start()


def texto_usuario():
    if sessao.get("userName"):
        return f"Usuário: {sessao['userName']}"
    if api._login:
        return f"Usuário: {api._login} (reconectando...)"
    return "Usuário: não autenticado"


def obter_texto_status():
    if monitoramento_pausado:
        return "Status: PAUSADO"
    elif cronometro_rodando and cronometro_pausado:
        return f"⏸ PAUSADO: {titulo_monitorado[:30]}..."
    elif cronometro_rodando and titulo_monitorado:
        return f"Monitorando: {titulo_monitorado[:35]}..."
    else:
        return "Status: Aguardando janela..."


def pausar_cronometro_bandeja(icon, item):
    """Pausa/retoma o cronômetro pela bandeja (a bandeja roda em outra thread)."""
    janela_principal.after(0, alternar_pausa_cronometro)


def _montar_menu_bandeja():
    texto_pausa = "Retomar Monitoramento" if monitoramento_pausado else "Pausar Monitoramento"
    texto_cronometro = (f"Retomar cronômetro ({atalho_pausa})" if cronometro_pausado
                        else f"Pausar cronômetro ({atalho_pausa})")
    return pystray.Menu(
        pystray.MenuItem(f"Monitor TimeSheet  v{APP_VERSAO}", lambda: None, enabled=False),
        pystray.MenuItem(texto_usuario(), lambda: None, enabled=False),
        pystray.MenuItem(obter_texto_status(), lambda: None, enabled=False),
        pystray.MenuItem('Parar e registrar', parar_e_registrar,
                         enabled=lambda item: cronometro_rodando),
        pystray.MenuItem(texto_cronometro, pausar_cronometro_bandeja,
                         enabled=lambda item: cronometro_rodando),
        pystray.MenuItem(texto_pausa, alternar_pausa),
        pystray.MenuItem('Configurações', abrir_config_thread),
        pystray.MenuItem('Ajuda', abrir_ajuda),
        pystray.MenuItem('Verificar atualizações', verificar_atualizacao_bandeja),
        pystray.MenuItem('Trocar de usuário', trocar_usuario),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem('Sair do Timesheet', sair_completamente)
    )


def atualizar_menu_bandeja():
    if icone_bandeja:
        icone_bandeja.menu = _montar_menu_bandeja()


def sair_completamente(icon, item):
    if icon:
        icon.stop()
    janela_principal.quit()
    os._exit(0)


def configurar_bandeja():
    global icone_bandeja
    imagem = criar_imagem_bandeja()
    icone_bandeja = pystray.Icon("AppTimesheet", imagem, "Monitor de Timesheet", _montar_menu_bandeja())
    icone_bandeja.run()


# ============================================================
# EXECUÇÃO PRINCIPAL
# ============================================================
if __name__ == "__main__":
    # Instância única: se o app já estiver rodando, avisa e encerra esta segunda cópia.
    if ja_esta_rodando():
        try:
            _aviso = tk.Tk()
            _aviso.withdraw()
            messagebox.showinfo(
                "Monitor de Produtividade",
                "O Monitor de Produtividade já está em execução.\n\n"
                "Procure o ícone do relógio na bandeja do sistema (perto do relógio do Windows)."
            )
            _aviso.destroy()
        except Exception:
            pass
        os._exit(0)

    # Define um AppUserModelID próprio para o Windows usar NOSSO ícone na barra
    # de tarefas (em vez do ícone do python.exe). Só tem efeito no Windows.
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("JRW.TimeSheet.Monitor.1")
    except Exception:
        pass

    # Raiz ÚNICA do tkinter (evita o erro "Tcl_AsyncDelete: wrong thread"
    # que acontece ao usar dois tk.Tk() com threads). Fica oculta.
    janela_principal = tk.Tk()
    janela_principal.title("TimeSheet - Monitor de Produtividade")
    janela_principal.withdraw()
    aplicar_icone_janela(janela_principal)  # ícone do relógio em vez da pena

    # 1) Login: usa o login guardado ("Manter conectado") ou mostra a tela de login
    api.ao_credenciais_recusadas = credenciais_recusadas
    precisa_tela, aviso_login, usuario_login = True, None, ""
    guardado = carregar_login_guardado()
    if guardado:
        ok_auto, _ = api.login(*guardado)
        if ok_auto:
            precisa_tela = False
        elif api.ultimo_erro == "credenciais":
            # Senha alterada no TimeSheet: descarta o login guardado e pede de novo
            apagar_login_guardado()
            aviso_login = "Sua senha foi alterada no TimeSheet. Entre novamente para continuar."
            usuario_login = guardado[0]
        else:
            # Sem internet (ou servidor fora) agora: abre o Monitor e tenta de novo sozinho
            api.definir_credenciais(*guardado)
            precisa_tela = False
            janela_principal.after(30000, tentar_reconectar)
    if precisa_tela and not tela_login(janela_principal, aviso=aviso_login, login_inicial=usuario_login):
        # Usuário fechou a janela sem logar
        janela_principal.destroy()
        os._exit(0)

    # 2) Carrega os programas monitorados salvos (ou usa o padrão na 1ª vez)
    carregar_palavras_alvo()
    carregar_preferencias()

    # Atalho global de pausa do cronômetro (configurável em Configurações)
    atalho_pausa = carregar_configuracoes().get("atalho_pausa") or ATALHO_PADRAO
    if not interpretar_atalho(atalho_pausa):
        atalho_pausa = ATALHO_PADRAO
    _ok_atalho, _msg_atalho = registrar_atalho_pausa(atalho_pausa)
    if not _ok_atalho:
        janela_principal.after(1500, lambda: messagebox.showwarning("Atalho de pausa", _msg_atalho))

    # 3) Inicia o monitoramento e a bandeja, usando a MESMA raiz
    verificar_janelas()
    verificar_retomada()
    janela_principal.after(ATUALIZACAO_PRIMEIRA_MS, verificar_atualizacao)
    threading.Thread(target=configurar_bandeja, daemon=True).start()

    janela_principal.mainloop()
