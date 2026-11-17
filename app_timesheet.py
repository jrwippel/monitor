import time
import tkinter as tk
from tkinter import simpledialog, messagebox
import win32gui
import win32process
import psutil
import json
import os
import math
from datetime import datetime
import threading
import pystray
from PIL import Image, ImageDraw

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

palavras_alvo = ["word", "excel", "powerpoint", "outlook", "pdf", "chrome", "edge", "firefox", "eproc", "projudi"]
termos_ignorar = ["lista de atalhos", "menu", "iniciar", "pesquisa", "configurações"]

# --- ÍCONE DE RELÓGIO PARA A BANDEJA ---
def criar_imagem_bandeja():
    """Gera um ícone de relógio estilizado para a bandeja do sistema"""
    img = Image.new('RGB', (64, 64), color="#152232")
    d = ImageDraw.Draw(img)
    d.ellipse([8, 8, 56, 56], outline="#4EA4DF", width=4)
    d.line([32, 32, 32, 18], fill="#4EA4DF", width=4)
    d.line([32, 32, 44, 32], fill="#4EA4DF", width=3)
    return img

# --- GERENCIAMENTO DE CONFIGURAÇÕES ---
def carregar_configuracoes():
    if os.path.exists(arquivo_config):
        try:
            with open(arquivo_config, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            pass
    return {}

def salvar_configuracoes(email):
    try:
        data = {"email_usuario": email}
        with open(arquivo_config, "w", encoding="utf-8") as f:
            json.dump(data, f)
    except:
        pass

def obter_email_usuario():
    config = carregar_configuracoes()
    email = config.get("email_usuario", "")
    if not email:
        root = tk.Tk()
        root.withdraw()
        email = simpledialog.askstring(
            "Configuração Inicial - Timesheet", 
            "Bem-vindo ao Monitor de Timesheet!\n\nPor favor, insira o seu e-mail corporativo:"
        )
        root.destroy()
        if email and "@" in email:
            salvar_configuracoes(email.strip())
        else:
            return ""
    return email

# --- TENTATIVA DE CAPTURAR O CAMINHO DO ARQUIVO (REDE U:\) ---
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
    except:
        pass
    return ""

# --- LÓGICA DE MONITORAMENTO INTELIGENTE ---
def verificar_janelas():
    global cronometro_rodando, monitoramento_pausado, hwnd_monitorado, titulo_monitorado, caminho_arquivo_monitorado, janelas_ja_perguntadas
    
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
                            matches = [palavra for palavra in palavras_alvo if palavra in titulo_minusculo]
                            if matches and hwnd_atual not in janelas_ja_perguntadas:
                                janelas_ja_perguntadas.add(hwnd_atual)
                                perguntar_iniciar_cronometro(hwnd_atual, titulo_atual)
        except:
            pass

    janela_principal.after(4000, verificar_janelas)

def perguntar_iniciar_cronometro(hwnd, titulo_janela):
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    resposta = messagebox.askyesno(
        "Monitoramento Ativo", 
        f"Detectada janela de trabalho:\n'{titulo_janela[:45]}...'\n\nDeseja iniciar o cronômetro para esta janela?",
        parent=root
    )
    root.destroy()
    if resposta:
        iniciar_cronometro(hwnd, titulo_janela)

def iniciar_cronometro(hwnd, titulo_janela):
    global cronometro_rodando, tempo_inicio, hwnd_monitorado, titulo_monitorado, caminho_arquivo_monitorado, hora_inicio_texto
    cronometro_rodando = True
    hwnd_monitorado = hwnd
    titulo_monitorado = titulo_janela
    caminho_arquivo_monitorado = obter_caminho_do_processo(hwnd)
    tempo_inicio = time.time() 
    hora_inicio_texto = datetime.now().strftime("%H:%M")
    atualizar_menu_bandeja()

def parar_cronometro_automatico():
    global cronometro_rodando, hwnd_monitorado, titulo_monitorado, caminho_arquivo_monitorado, hora_inicio_texto
    
    cronometro_rodando = False
    
    tempo_fim = time.time()
    
    # ARREDONDAMENTO INTELIGENTE PARA CIMA (Ceil)
    # Garante que qualquer fração iniciada conte o minuto cheio correspondente
    minutos_gastos = math.ceil((tempo_fim - tempo_inicio) / 60)
    if minutos_gastos == 0: 
        minutos_gastos = 1 
        
    hora_fim_texto = datetime.now().strftime("%H:%M")
    
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    
    assunto = simpledialog.askstring(
        "Lançamento de Horas", 
        f"A janela foi fechada!\nTempo registrado: {minutos_gastos} min.\n\nQual o Assunto ou Cliente?", 
        parent=root
    )
    root.destroy()
    
    if assunto is None or assunto.strip() == "":
        assunto_final = f"CANCELADO - {titulo_monitorado}"
    else:
        assunto_final = assunto.strip()
    
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        email = outlook.CreateItem(0)
        email.To = obter_email_usuario()
        email.Subject = f"TIMESHEET: {assunto_final}"
        
        corpo_email = (
            f"Registro automático de horas:\n\n"
            f"Atividade/Cliente: {assunto_final}\n"
            f"Janela: {titulo_monitorado}\n"
        )
        if caminho_arquivo_monitorado:
            corpo_email += f"Caminho do arquivo: {caminho_arquivo_monitorado}\n"
            
        corpo_email += (
            f"Horário: {hora_inicio_texto} às {hora_fim_texto}\n"
            f"Tempo total: {minutos_gastos} minuto(s)."
        )
        
        email.Body = corpo_email
        email.Send()
    except Exception:
        pass
            
    hwnd_monitorado = None
    titulo_monitorado = ""
    caminho_arquivo_monitorado = ""
    atualizar_menu_bandeja()

# --- FUNÇÕES DE CONTROLE DA BANDEJA ---
def alternar_pausa(icon, item):
    global monitoramento_pausado
    monitoramento_pausado = not monitoramento_pausado
    atualizar_menu_bandeja()

def obter_texto_status():
    if monitoramento_pausado:
        return "Status: PAUSADO"
    elif cronometro_rodando and titulo_monitorado:
        return f"Monitorando: {titulo_monitorado[:35]}..."
    else:
        return "Status: Aguardando janela..."

def atualizar_menu_bandeja():
    if icone_bandeja:
        texto_pausa = "Retomar Monitoramento" if monitoramento_pausado else "Pausar Monitoramento"
        icone_bandeja.menu = pystray.Menu(
            pystray.MenuItem(obter_texto_status(), lambda: None, enabled=False),
            pystray.MenuItem(texto_pausa, alternar_pausa),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem('Sair do Timesheet', sair_completamente)
        )

def sair_completamente(icon, item):
    if icon:
        icon.stop()
    janela_principal.quit()
    os._exit(0)

def configurar_bandeja():
    global icone_bandeja
    imagem = criar_imagem_bandeja()
    menu = pystray.Menu(
        pystray.MenuItem(obter_texto_status(), lambda: None, enabled=False),
        pystray.MenuItem('Pausar Monitoramento', alternar_pausa),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem('Sair do Timesheet', sair_completamente)
    )
    icone_bandeja = pystray.Icon("AppTimesheet", imagem, "Monitor de Timesheet", menu)
    icone_bandeja.run()

# --- EXECUÇÃO PRINCIPAL ---
if __name__ == "__main__":
    email_inicial = obter_email_usuario()
    
    janela_principal = tk.Tk()
    janela_principal.title("App Timesheet")
    janela_principal.withdraw()
    
    verificar_janelas()
    
    threading.Thread(target=configurar_bandeja, daemon=True).start()
    
    janela_principal.mainloop()