import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import win32com.client as win32
import pythoncom
import pandas as pd
import re
import threading
import time

def extrair_link_exportacao(url):
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
    if match:
        id_planilha = match.group(1)
        return f"https://docs.google.com/spreadsheets/d/{id_planilha}/export?format=xlsx"
    else:
        raise ValueError("URL do Google Sheets inválida ou incorreta.")

def limpar_emails(texto_email):
    """
    Higieniza a string de e-mails:
    - Remove quebras de linha e espaços extras.
    - Substitui vírgulas e múltiplos espaços por ponto e vírgula.
    - Remove pontos e vírgulas duplicados e ao final da string.
    """
    if not texto_email or str(texto_email).lower() in ["nan", "none"]:
        return ""
    
    texto_limpo = re.sub(r'[\r\n,]+', ';', str(texto_email).strip())
    
    partes = [p.strip() for p in texto_limpo.split(';') if p.strip()]
    
    return ";".join(partes)

corpo_mensagem_1 = """
<p>Prezados,</p>
<p>Solicito o envio das guias para compensação, conforme os prazos abaixo:</p>
<p>• <b>INSS:</b> vencimento 18/09 (envio até 15/09, às 16h – horário de Brasília)</p>
<p>• <b>PIS/COFINS:</b> vencimento 25/09 (envio até 21/09, às 16h – horário de Brasília)</p>
<p>Enviar as guias para: <a href="mailto:vanessa.odorizzi@grupostudio.com.br">vanessa.odorizzi@grupostudio.com.br</a>; 
<a href="mailto:priscila.figueiredo@grupostudio.com.br">priscila.figueiredo@grupostudio.com.br</a>; 
<a href="mailto:gabriela.sasso@grupostudio.com.br">gabriela.sasso@grupostudio.com.br</a>; 
<a href="mailto:carolina.fernandes@grupostudio.com.br">carolina.fernandes@grupostudio.com.br</a> 
ou via WhatsApp: 51 3086-6491.</p>
<p>Reforçamos que, para que as compensações sejam realizadas, é obrigatório que a procuração eletrônica da empresa esteja liberada para a Studio Fiscal (CNPJ Studio Operacional: 23.448.109/0001-91). É fundamental que as guias sejam enviadas dentro do prazo para que as compensações ocorram da melhor forma.</p>
<p style="font-size: 11px; color: #555555; text-align: justify; line-height: 1.4;">
<i><b>Aviso de Responsabilidade:</b> As guias DARF e os cálculos de retificação e compensação (PER/DCOMP) encaminhados foram elaborados com base nas informações, documentos e dados fiscais fornecidos pelo cliente, observados a legislação tributária vigente e os mais rigorosos padrões técnicos e éticos. A conferência dos valores, das competências e a quitação tempestiva das guias são de responsabilidade exclusiva do contribuinte. Eventuais alterações na legislação, divergências em bases de cálculo, glosas ou questionamentos por parte da Receita Federal e demais órgãos fiscalizadores não estão sob nosso controle e precisam ser informados, pois podem ensejar riscos de cobrança, multa ou indeferimento de compensação. Destacamos ainda que todos os riscos das operações são apontados no momento da apresentação, não tendo o Grupo Studio e suas empresas qualquer responsabilidade em relação às eventuais consequências decorrentes dos riscos dessas operações.</i>
</p>
<p>Ficamos à disposição.</p>
"""

corpo_mensagem_2 = """
<p>Prezados,</p>
<p>Solicito o envio das guias para compensação, conforme os prazos abaixo:</p>
<p>• <b>INSS:</b> vencimento 18/09 (envio até 15/09, às 16h – horário de Brasília)</p>
<p>• <b>PIS/COFINS:</b> vencimento 25/09 (envio até 21/09, às 16h – horário de Brasília)</p>
<p><b>Observação:</b> a empresa possui, entre os créditos disponíveis, créditos escriturais exclusivos para PIS/COFINS (códigos 5856 e 6912). Para compensar usando esse saldo, é necessário enviar o EFD Contribuições junto com as guias.</p>
<p>Enviar as guias para: <a href="mailto:vanessa.odorizzi@grupostudio.com.br">vanessa.odorizzi@grupostudio.com.br</a>; 
<a href="mailto:priscila.figueiredo@grupostudio.com.br">priscila.figueiredo@grupostudio.com.br</a>; 
<a href="mailto:gabriela.sasso@grupostudio.com.br">gabriela.sasso@grupostudio.com.br</a>; 
<a href="mailto:carolina.fernandes@grupostudio.com.br">carolina.fernandes@grupostudio.com.br</a> 
ou via WhatsApp: 51 3086-6491.</p>
<p>Reforçamos que, para que as compensações sejam realizadas, é obrigatório que a procuração eletrônica da empresa esteja liberada para a Studio Fiscal (CNPJ Studio Operacional: 23.448.109/0001-91). É fundamental que as guias sejam enviadas dentro do prazo para que as compensações ocorram da melhor forma.</p>
<p style="font-size: 11px; color: #555555; text-align: justify; line-height: 1.4;">
<i><b>Aviso de Responsabilidade:</b> As guias DARF e os cálculos de retificação e compensação (PER/DCOMP) encaminhados foram elaborados com base nas informações, documentos e dados fiscais fornecidos pelo cliente, observados a legislação tributária vigente e os mais rigorosos padrões técnicos e éticos. A conferência dos valores, das competências e a quitação tempestiva das guias são de responsabilidade exclusiva do contribuinte. Eventuais alterações na legislação, divergências em bases de cálculo, glosas ou questionamentos por parte da Receita Federal e demais órgãos fiscalizadores não estão sob nosso controle e precisam ser informados, pois podem ensejar riscos de cobrança, multa ou indeferimento de compensação. Destacamos ainda que todos os riscos das operações são apontados no momento da apresentação, não tendo o Grupo Studio e suas empresas qualquer responsabilidade em relação às eventuais consequências decorrentes dos riscos dessas operações.</i>
</p>
<p>Ficamos à disposição.</p>
"""

corpo_mensagem_3 = """
<p>Prezados,</p>
<p>Solicito o envio das guias para compensação, conforme o prazo abaixo:</p>
<p>• <b>PIS/COFINS:</b> vencimento 25/09 (envio até 21/09, às 16h – horário de Brasília)</p>
<p><b>Observação:</b> a empresa possui apenas créditos escriturais exclusivos para compensação de PIS/COFINS (códigos 5856 e 6912). Para compensar usando esse saldo, é necessário enviar o EFD Contribuições junto com as guias.</p>
<p>Enviar as guias para: <a href="mailto:vanessa.odorizzi@grupostudio.com.br">vanessa.odorizzi@grupostudio.com.br</a>; 
<a href="mailto:priscila.figueiredo@grupostudio.com.br">priscila.figueiredo@grupostudio.com.br</a>; 
<a href="mailto:gabriela.sasso@grupostudio.com.br">gabriela.sasso@grupostudio.com.br</a>; 
<a href="mailto:carolina.fernandes@grupostudio.com.br">carolina.fernandes@grupostudio.com.br</a> 
ou via WhatsApp: 51 3086-6491.</p>
<p>Reforçamos que, para que as compensações sejam realizadas, é obrigatório que a procuração eletrônica da empresa esteja liberada para a Studio Fiscal (CNPJ Studio Operacional: 23.448.109/0001-91). É fundamental que as guias sejam enviadas dentro do prazo para que as compensações ocorram da melhor forma.</p>
<p style="font-size: 11px; color: #555555; text-align: justify; line-height: 1.4;">
<i><b>Aviso de Responsabilidade:</b> As guias DARF e os cálculos de retificação e compensação (PER/DCOMP) encaminhados foram elaborados com base nas informações, documentos e dados fiscais fornecidos pelo cliente, observados a legislação tributária vigente e os mais rigorosos padrões técnicos e éticos. A conferência dos valores, das competências e a quitação tempestiva das guias são de responsabilidade exclusiva do contribuinte. Eventuais alterações na legislação, divergências em bases de cálculo, glosas ou questionamentos por parte da Receita Federal e demais órgãos fiscalizadores não estão sob nosso controle e precisam ser informados, pois podem ensejar riscos de cobrança, multa ou indeferimento de compensação. Destacamos ainda que todos os riscos das operações são apontados no momento da apresentação, não tendo o Grupo Studio e suas empresas qualquer responsabilidade em relação às eventuais consequências decorrentes dos riscos dessas operações.</i>
</p>
<p>Ficamos à disposição.</p>
"""

class EmailSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Disparador de E-mails - Grupo Studio")
        self.root.geometry("680x520")
        self.root.resizable(False, False)

        style = ttk.Style()
        style.theme_use("clam")

        lbl_instrucao = tk.Label(root, text="Cole abaixo a URL da planilha do Google Sheets:", font=("Segoe UI", 10, "bold"))
        lbl_instrucao.pack(anchor="w", padx=20, pady=(20, 5))

        self.txt_url = ttk.Entry(root, width=80, font=("Segoe UI", 10))
        self.txt_url.pack(padx=20, pady=5, fill="x")

        self.btn_iniciar = ttk.Button(root, text="🚀 Iniciar Disparo de E-mails", command=self.iniciar_processo_thread)
        self.btn_iniciar.pack(padx=20, pady=15, fill="x")

        lbl_log = tk.Label(root, text="Status e Registros de Envio:", font=("Segoe UI", 10, "bold"))
        lbl_log.pack(anchor="w", padx=20, pady=(10, 5))

        self.log_area = scrolledtext.ScrolledText(root, width=80, height=15, font=("Consolas", 9))
        self.log_area.pack(padx=20, pady=(0, 20), fill="both", expand=True)

    def log(self, mensagem):
        self.log_area.insert(tk.END, mensagem + "\n")
        self.log_area.see(tk.END)

    def iniciar_processo_thread(self):
        url = self.txt_url.get().strip()
        if not url:
            messagebox.showwarning("Aviso", "Por favor, insira o link do Google Sheets.")
            return

        self.btn_iniciar.config(state="disabled")
        self.log_area.delete(1.0, tk.END)
        
        threading.Thread(target=self.processar_envio, args=(url,), daemon=True).start()

    def processar_envio(self, url):
        pythoncom.CoInitialize()
        try:
            self.log("Acessando a planilha do Google Sheets...")
            link_direto = extrair_link_exportacao(url)

            df_base = pd.read_excel(link_direto, sheet_name="base dados", header=None)

            COLUNA_EMPRESA_MENSAGEM = 0  
            COLUNA_EMPRESA_BASE = 0     
            COLUNA_JOB_BASE = 3          
            COLUNA_EMAIL_P = 1         
            COLUNA_CC_Q = 2        
            COLUNA_JOB = 'job'

            tarefas = [
                {"aba": "MENSAGEM 1", "corpo": corpo_mensagem_1},
                {"aba": "MENSAGEM 2", "corpo": corpo_mensagem_2},
                {"aba": "MENSAGEM 3", "corpo": corpo_mensagem_3}
            ]

            outlook = win32.Dispatch('outlook.application')

            for tarefa in tarefas:
                nome_aba = tarefa["aba"]
                corpo_html = tarefa["corpo"]

                self.log(f"\n==================================================")
                self.log(f" Processando: {nome_aba}")
                self.log(f"==================================================")

                try:
                    df_mensagem = pd.read_excel(link_direto, sheet_name=nome_aba, header=None)
                except Exception as e:
                    self.log(f"Erro ao ler a aba '{nome_aba}': {e}. Pulando...")
                    continue

                for index, linha_msg in df_mensagem.iloc[1:].iterrows():
                    empresa = str(linha_msg[COLUNA_EMPRESA_MENSAGEM]).strip()
                    
                    if not empresa or empresa == "nan" or empresa.lower() == "empresa":
                        continue
                        
                    dados_empresa = df_base[df_base[COLUNA_EMPRESA_BASE].astype(str).str.strip().str.upper() == empresa.upper()]
                    
                    if not dados_empresa.empty:
                        email_destinatario_raw = str(dados_empresa.iloc[0, COLUNA_EMAIL_P])
                        email_cc_raw = str(dados_empresa.iloc[0, COLUNA_CC_Q])
                        job_valor = str(dados_empresa.iloc[0, COLUNA_JOB_BASE]).strip()
                        
                        if job_valor == "nan":
                            job_valor = ""
                            
                        email_destinatario = limpar_emails(email_destinatario_raw)
                        email_cc = limpar_emails(email_cc_raw)
                        
                        if not email_destinatario:
                            self.log(f"[{nome_aba}] AVISO: Empresa '{empresa}' sem e-mail válido.")
                            continue
                            
                        linha_msg[COLUNA_JOB] = job_valor

                        try:
                            email = outlook.CreateItem(0)
                            email.To = email_destinatario
                            
                            if email_cc:
                                email.CC = email_cc
                                
                            email.Subject = f"Solicitação de Guias para Compensação - {empresa}   {linha_msg[COLUNA_JOB]}"
                            
                            email.GetInspector.Activate()
                            assinatura_outlook = email.HTMLBody
                            
                            email.HTMLBody = corpo_html + assinatura_outlook
                            
                            email.Send()
                            self.log(f"[{nome_aba}] SUCESSO -> {empresa} ({email_destinatario})")
                            
                            time.sleep(15)

                        except Exception as err_envio:
                            try:
                                email_fallback = outlook.CreateItem(0)
                                email_fallback.To = email_destinatario
                                if email_cc:
                                    email_fallback.CC = email_cc
                                email_fallback.Subject = f"Solicitação de Guias para Compensação - {empresa}   {linha_msg[COLUNA_JOB]}"
                                email_fallback.HTMLBody = corpo_html
                                email_fallback.Send()
                                self.log(f"[{nome_aba}] SUCESSO (fallback sem assinatura) -> {empresa}")
                                time.sleep(15)
                            except Exception as err_fallback:
                                self.log(f"[{nome_aba}] ERRO ao enviar para '{empresa}': {err_fallback}")
                                time.sleep(15)
                    else:
                        self.log(f"[{nome_aba}] AVISO: Empresa '{empresa}' não foi encontrada na aba 'base dados'.")

            self.log("\n✅ TODO O PROCESSO FOI CONCLUÍDO COM SUCESSO!")
            messagebox.showinfo("Sucesso", "Envio das Mensagens 1, 2 e 3 concluído!")

        except Exception as e:
            self.log(f"\n❌ ERRO CRÍTICO: {e}")
            messagebox.showerror("Erro", f"Ocorreu uma falha: {e}")
            
        finally:
            pythoncom.CoUninitialize()
            self.btn_iniciar.config(state="normal")

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSenderApp(root)
    root.mainloop()