import tkinter as tk
from tkinter import ttk, simpledialog, messagebox 
import pyperclip
import json
import os
import psutil
import time
import win32com.client
# NOVO: Importamos a biblioteca que lê o relógio e o calendário do Windows
from datetime import datetime 

# --- 1. CONFIGURAÇÃO DE PASTAS ---
pasta_usuario = os.path.expanduser("~")
pasta_documentos = os.path.join(pasta_usuario, "Documents")

arquivo_config = os.path.join(pasta_documentos, "configuracoes_assistente.json")
arquivo_tarefas = os.path.join(pasta_documentos, "minhas_tarefas.txt")

# --- VARIÁVEIS GLOBAIS ---
dados_app = {
    "palavras_chave": ["prazo", "petição", "revisar"],
    "email_usuario": ""
}

# Controle da Área de Transferência (Agenda)
monitorando_copia = False
texto_anterior = ""

# Controle do Timesheet
cronometro_rodando = False
tempo_inicio = 0
hora_inicio_texto = "" # NOVO: Vai guardar o texto bonitinho da hora de início
janela_timer = None 
pids_notificados = set()
programas_alvo = ["WINWORD.EXE", "EXCEL.EXE"] 

# --- LÓGICA DE DADOS E CONFIGURAÇÃO ---

def carregar_dados():
    global dados_app
    if os.path.exists(arquivo_config):
        with open(arquivo_config, "r", encoding="utf-8") as f:
            dados_app = json.load(f)
    atualizar_lista_visual()
    if dados_app["email_usuario"]:
        entrada_email.insert(0, dados_app["email_usuario"])

def salvar_dados():
    with open(arquivo_config, "w", encoding="utf-8") as f:
        json.dump(dados_app, f)

# --- LÓGICA DA ÁREA DE TRANSFERÊNCIA (AGENDA) ---

def adicionar_palavra():
    nova_palavra = entrada_palavra.get().lower().strip()
    if nova_palavra and nova_palavra not in dados_app["palavras_chave"]:
        dados_app["palavras_chave"].append(nova_palavra)
        salvar_dados()
        atualizar_lista_visual()
        entrada_palavra.delete(0, tk.END)

def remover_palavra():
    selecao = lista_box.curselection()
    if selecao:
        palavra = lista_box.get(selecao[0])
        dados_app["palavras_chave"].remove(palavra)
        salvar_dados()
        atualizar_lista_visual()

def atualizar_lista_visual():
    lista_box.delete(0, tk.END)
    for p in dados_app["palavras_chave"]:
        lista_box.insert(tk.END, p)

def alternar_monitoramento_copia():
    global monitorando_copia
    monitorando_copia = not monitorando_copia 
    
    if monitorando_copia:
        btn_ligar.config(text="Desligar Leitura de Texto", bg="#d9534f", fg="white")
        lbl_status.config(text="Leitura: Vigiando cópias...", fg="green")
    else:
        btn_ligar.config(text="Ligar Leitura de Texto", bg="#5cb85c", fg="white")
        lbl_status.config(text="Leitura: Pausada", fg="red")

def verificar_area_de_transferencia():
    global texto_anterior
    if monitorando_copia:
        try:
            texto_atual = str(pyperclip.paste()).lower()
            if texto_atual != texto_anterior and texto_atual.strip() != "":
                texto_anterior = texto_atual
                if any(p in texto_atual for p in dados_app["palavras_chave"]):
                    mostrar_alerta_copia(texto_atual)
        except Exception:
            pass 
    janela.after(2000, verificar_area_de_transferencia)

def mostrar_alerta_copia(texto):
    resposta = messagebox.askyesno(
        "Assistente TDAH",
        f"Vi que você copiou algo importante.\nQuer anotar isso na sua lista de tarefas?\n\nTrecho: '{texto[:40]}...'"
    )
    if resposta:
        with open(arquivo_tarefas, "a", encoding="utf-8") as arquivo:
            arquivo.write(f"- Tarefa identificada: {texto}\n")

# --- LÓGICA DO TIMESHEET (MONITOR DE PROGRAMAS) ---

def salvar_email():
    email_digitado = entrada_email.get().strip()
    if "@" in email_digitado:
        dados_app["email_usuario"] = email_digitado
        salvar_dados()
        messagebox.showinfo("Sucesso", "E-mail salvo com sucesso!")
    else:
        messagebox.showerror("Erro", "Por favor, digite um e-mail válido.")

def verificar_programas():
    global cronometro_rodando, pids_notificados
    
    pids_ativos = [p.pid for p in psutil.process_iter()]
    pids_notificados = set([pid for pid in pids_notificados if pid in pids_ativos])

    for processo in psutil.process_iter(['name', 'pid']):
        try:
            nome = processo.info['name']
            pid = processo.info['pid']
            
            if nome and nome.upper() in programas_alvo:
                if pid not in pids_notificados:
                    pids_notificados.add(pid)
                    if not cronometro_rodando and janela_timer is None:
                        perguntar_iniciar_cronometro(nome)
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass

    janela.after(5000, verificar_programas)

def perguntar_iniciar_cronometro(nome_programa):
    resposta = messagebox.askyesno(
        "Timesheet Automático", 
        f"Notei que você abriu o {nome_programa}.\nDeseja iniciar o cronômetro para o Timesheet?"
    )
    if resposta:
        iniciar_cronometro()

def iniciar_cronometro():
    global cronometro_rodando, tempo_inicio, janela_timer, hora_inicio_texto
    cronometro_rodando = True
    
    # Salva o tempo matemático para fazer a conta depois
    tempo_inicio = time.time() 
    
    # NOVO: Salva o texto da hora atual (Ex: "21/10/2026 às 14:30")
    hora_inicio_texto = datetime.now().strftime("%d/%m/%Y às %H:%M")
    
    janela_timer = tk.Toplevel(janela)
    janela_timer.title("Timesheet")
    janela_timer.geometry("200x100")
    janela_timer.attributes("-topmost", True) 
    
    tk.Label(janela_timer, text="Gravando horas...", font=("Arial", 10, "bold"), fg="red").pack(pady=10)
    tk.Button(janela_timer, text="Parar e Registrar", command=parar_cronometro, bg="#f0ad4e").pack(pady=5)

# NOVO: A função de e-mail agora recebe também a hora de início e a hora do fim
def enviar_email_timesheet(assunto, minutos, hora_inicio, hora_fim):
    email_destino = dados_app["email_usuario"]
    if not email_destino:
        messagebox.showerror("Erro", "Você não configurou seu e-mail na aba Timesheet! E-mail não enviado.")
        return False
        
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        email = outlook.CreateItem(0)
        email.To = email_destino 
        email.Subject = f"TIMESHEET: {assunto}"
        
        # NOVO: O corpo do e-mail agora tem as horas detalhadas
        corpo_texto = (
            f"Registro automático de horas:\n\n"
            f"Assunto/Atividade: {assunto}\n\n"
            f"Início: {hora_inicio}\n"
            f"Fim: {hora_fim}\n"
            f"Tempo total: {minutos} minuto(s).\n\n"
            f"Enviado via Assistente."
        )
        email.Body = corpo_texto
        
        email.Send() 
        return True 
    except Exception as e:
        print(f"Erro: {e}")
        return False

def parar_cronometro():
    global cronometro_rodando, janela_timer, hora_inicio_texto
    cronometro_rodando = False
    
    # Salva o tempo matemático e calcula a diferença
    tempo_fim = time.time()
    minutos_gastos = int((tempo_fim - tempo_inicio) // 60)
    if minutos_gastos == 0: minutos_gastos = 1 
        
    # NOVO: Pega a hora exata em que você apertou "Parar"
    hora_fim_texto = datetime.now().strftime("%d/%m/%Y às %H:%M")
        
    janela_timer.destroy()
    janela_timer = None
    
    assunto = simpledialog.askstring("Registrar Timesheet", f"Tempo: {minutos_gastos} min.\nQual o assunto/cliente deste trabalho?")
    if assunto:
        # NOVO: Enviamos as horas junto para o e-mail
        sucesso = enviar_email_timesheet(assunto, minutos_gastos, hora_inicio_texto, hora_fim_texto)
        if sucesso:
            messagebox.showinfo("Sucesso", f"Timesheet enviado para {dados_app['email_usuario']}!")
    else:
        messagebox.showwarning("Aviso", "Registro cancelado.")

# --- CONSTRUÇÃO DA INTERFACE VISUAL (SEPARAÇÃO EM ABAS) ---

janela = tk.Tk()
janela.title("Meu Assistente")
janela.geometry("380x450")
janela.attributes("-topmost", True) 

sistema_abas = ttk.Notebook(janela)

aba_agenda = ttk.Frame(sistema_abas)
aba_timesheet = ttk.Frame(sistema_abas)

sistema_abas.add(aba_agenda, text="Minha Agenda (Foco)")
sistema_abas.add(aba_timesheet, text="Timesheet (Escritório)")
sistema_abas.pack(expand=True, fill="both", padx=10, pady=10)

# === ABA 1 (AGENDA) ===
tk.Label(aba_agenda, text="Monitor de Cópias de Textos", font=("Arial", 11, "bold")).pack(pady=10)

lbl_status = tk.Label(aba_agenda, text="Leitura: Pausada", fg="red", font=("Arial", 10, "bold"))
lbl_status.pack(pady=5)

btn_ligar = tk.Button(aba_agenda, text="Ligar Leitura de Texto", bg="#5cb85c", fg="white", font=("Arial", 10, "bold"), command=alternar_monitoramento_copia)
btn_ligar.pack(pady=5, ipadx=10, ipady=5)

tk.Label(aba_agenda, text="Dicionário de Palavras-chave:").pack(pady=(10, 0))
entrada_palavra = tk.Entry(aba_agenda, width=25, font=("Arial", 11))
entrada_palavra.pack(pady=5)
tk.Button(aba_agenda, text="Adicionar Palavra", command=adicionar_palavra).pack(pady=2)

lista_box = tk.Listbox(aba_agenda, width=30, height=6, font=("Arial", 10))
lista_box.pack(pady=5)
tk.Button(aba_agenda, text="Remover Selecionada", command=remover_palavra).pack(pady=2)

# === ABA 2 (TIMESHEET) ===
tk.Label(aba_timesheet, text="Controle de Horas Automático", font=("Arial", 11, "bold")).pack(pady=10)

tk.Label(aba_timesheet, text="Monitor de Word e Excel: ATIVO", fg="blue", font=("Arial", 9, "bold")).pack(pady=5)
tk.Label(aba_timesheet, text="(Basta abrir um dos programas para iniciar)").pack(pady=(0, 20))

tk.Label(aba_timesheet, text="Seu E-mail Corporativo de Destino:").pack()
entrada_email = tk.Entry(aba_timesheet, width=35)
entrada_email.pack(pady=5)
tk.Button(aba_timesheet, text="Salvar E-mail", command=salvar_email).pack(pady=2)

# --- Inicialização ---
carregar_dados()
verificar_area_de_transferencia()
verificar_programas()

janela.mainloop()