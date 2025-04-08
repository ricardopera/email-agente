import tkinter as tk
from tkinter import ttk, simpledialog, messagebox, BooleanVar
import threading
import sys
import os
import logging
from src.email_processor import EmailProcessor, logger

class EmailApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Agente de Extração de Emails")
        self.root.geometry("600x580")  # Aumentando a altura para acomodar novos controles
        self.root.resizable(True, True)
        
        self.email_processor = EmailProcessor()
        
        # Configurar o estilo
        self.style = ttk.Style()
        self.style.configure("TFrame", background="#f0f0f0")
        self.style.configure("TLabel", background="#f0f0f0", font=("Arial", 10))
        self.style.configure("TButton", font=("Arial", 10))
        self.style.configure("Header.TLabel", font=("Arial", 12, "bold"))
        
        # Carregar configurações salvas
        self.load_saved_config()
        
        self.create_widgets()
        
    def load_saved_config(self):
        """Carrega as configurações salvas anteriormente"""
        try:
            config = self.email_processor.load_config()
            self.saved_email = config.get('email_user', '')
            self.saved_server = config.get('imap_host', 'mail.itajai.sc.gov.br')  # Alterando o valor padrão
            self.saved_subject = config.get('search_subject', '')
            logger.info("Configurações carregadas na interface")
        except Exception as e:
            logger.error(f"Erro ao carregar configurações na interface: {str(e)}")
            self.saved_email = ''
            self.saved_server = 'mail.itajai.sc.gov.br'  # Alterando o valor padrão
            self.saved_subject = ''
        
    def create_widgets(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding=(20, 10))
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        ttk.Label(main_frame, text="Extração de Informações de Emails", 
                  style="Header.TLabel").grid(row=0, column=0, columnspan=2, pady=10)
        
        # Campos de entrada
        ttk.Label(main_frame, text="Endereço de Email:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.email_var = tk.StringVar(value=self.saved_email)
        ttk.Entry(main_frame, textvariable=self.email_var, width=40).grid(row=1, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(main_frame, text="Senha:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.password_var = tk.StringVar()
        password_entry = ttk.Entry(main_frame, textvariable=self.password_var, width=40, show='*')
        password_entry.grid(row=2, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(main_frame, text="Servidor IMAP:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.server_var = tk.StringVar(value=self.saved_server)
        ttk.Entry(main_frame, textvariable=self.server_var, width=40).grid(row=3, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(main_frame, text="Porta:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.port_var = tk.StringVar(value="993")
        ttk.Entry(main_frame, textvariable=self.port_var, width=10).grid(row=4, column=1, sticky=tk.W, pady=5)
        
        # Opções de conexão
        connection_frame = ttk.LabelFrame(main_frame, text="Opções de Conexão", padding=(10, 5))
        connection_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # Opção SSL/TLS
        self.use_ssl_var = BooleanVar(value=True)
        ttk.Checkbutton(connection_frame, text="Usar SSL (desmarque para TLS)", 
                        variable=self.use_ssl_var).grid(row=0, column=0, sticky=tk.W, pady=2)
        
        # Timeout
        ttk.Label(connection_frame, text="Timeout (segundos):").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.timeout_var = tk.StringVar(value="60")  # Aumentando o timeout padrão
        ttk.Entry(connection_frame, textvariable=self.timeout_var, width=5).grid(row=1, column=1, sticky=tk.W, pady=2)
        
        # Campo para o assunto
        ttk.Label(main_frame, text="Texto do Assunto:").grid(row=6, column=0, sticky=tk.W, pady=5)
        self.subject_var = tk.StringVar(value=self.saved_subject)
        ttk.Entry(main_frame, textvariable=self.subject_var, width=40).grid(row=6, column=1, sticky=tk.W, pady=5)
        
        # Saída de log
        ttk.Label(main_frame, text="Log de Operações:").grid(row=7, column=0, columnspan=2, sticky=tk.W, pady=(10, 5))
        
        self.log_text = tk.Text(main_frame, height=10, width=60, wrap=tk.WORD)
        self.log_text.grid(row=8, column=0, columnspan=2, pady=5)
        
        # Scroll para o log
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=8, column=2, sticky=(tk.N, tk.S))
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        # Botões
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=9, column=0, columnspan=2, pady=15)
        
        ttk.Button(button_frame, text="Processar Emails", command=self.start_processing).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="Teste de Conexão", command=self.test_connection).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="Sair", command=self.root.quit).pack(side=tk.LEFT, padx=10)
    
    def log(self, message):
        """Adiciona uma mensagem ao log"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)  # Rola para o final
        self.root.update_idletasks()  # Atualiza a interface
        logger.info(message)  # Também registra no arquivo de log
    
    def test_connection(self):
        """Testa a conexão com o servidor IMAP"""
        email = self.email_var.get().strip()
        password = self.password_var.get().strip()
        server = self.server_var.get().strip()
        
        try:
            port = int(self.port_var.get().strip())
        except ValueError:
            messagebox.showerror("Erro", "A porta deve ser um número!")
            return
            
        try:
            timeout = int(self.timeout_var.get().strip())
        except ValueError:
            messagebox.showerror("Erro", "O timeout deve ser um número!")
            return
        
        use_ssl = self.use_ssl_var.get()
        
        if not email or not password or not server:
            messagebox.showerror("Erro", "Email, senha e servidor são obrigatórios!")
            return
            
        self.log(f"Testando conexão com {server}:{port} (SSL: {use_ssl}, Timeout: {timeout}s)...")
        
        # Iniciar thread para não bloquear a interface
        threading.Thread(target=self._test_connection_thread, 
                         args=(email, password, server, port, use_ssl, timeout), 
                         daemon=True).start()
    
    def _test_connection_thread(self, email, password, server, port, use_ssl, timeout):
        """Thread para testar a conexão"""
        if self.email_processor.connect_to_server(email, password, server, port, use_ssl, timeout):
            self.log("Teste de conexão bem-sucedido!")
            messagebox.showinfo("Sucesso", "Conexão estabelecida com sucesso!")
            self.email_processor.close_connection()
        else:
            self.log("Falha no teste de conexão. Verifique os logs para mais detalhes.")
    
    def start_processing(self):
        """Inicia o processamento dos emails em uma thread separada"""
        # Validar campos
        email = self.email_var.get().strip()
        password = self.password_var.get().strip()
        server = self.server_var.get().strip()
        subject = self.subject_var.get().strip()
        
        try:
            port = int(self.port_var.get().strip())
        except ValueError:
            messagebox.showerror("Erro", "A porta deve ser um número!")
            return
            
        try:
            timeout = int(self.timeout_var.get().strip())
        except ValueError:
            messagebox.showerror("Erro", "O timeout deve ser um número!")
            return
        
        use_ssl = self.use_ssl_var.get()
        
        if not email or not password or not server or not subject:
            messagebox.showerror("Erro", "Todos os campos são obrigatórios!")
            return
        
        # Salvar as configurações para uso futuro
        self.email_processor.search_subject = subject
        self.email_processor.save_config(email, server, subject)
        
        # Iniciar thread para não bloquear a interface
        threading.Thread(target=self.process_emails, 
                         args=(email, password, server, port, use_ssl, timeout, subject), 
                         daemon=True).start()
    
    def process_emails(self, email, password, server, port, use_ssl, timeout, subject):
        """Processa os emails com os parâmetros fornecidos"""
        try:
            self.log(f"Conectando ao servidor {server}:{port} (SSL: {use_ssl}, Timeout: {timeout}s)...")
            if not self.email_processor.connect_to_server(email, password, server, port, use_ssl, timeout):
                self.log("Falha ao conectar ao servidor de email.")
                return
            
            self.log("Conexão estabelecida com sucesso!")
            self.log(f"Buscando emails não lidos com assunto contendo: '{subject}'...")
            
            email_ids = self.email_processor.search_emails(subject)
            
            if not email_ids:
                self.log("Nenhum email encontrado com os critérios especificados.")
                self.email_processor.close_connection()
                return
            
            self.log(f"Encontrados {len(email_ids)} emails. Processando...")
            processed = self.email_processor.process_emails(email_ids)
            
            self.log(f"Processados {processed} emails.")
            self.log("Salvando dados em arquivo Excel...")
            
            if self.email_processor.save_to_excel():
                self.log("Dados salvos com sucesso!")
            else:
                self.log("Não foi possível salvar os dados.")
                
            self.email_processor.close_connection()
            self.log("Conexão fechada. Processamento concluído.")
            
        except Exception as e:
            error_msg = f"Erro: {str(e)}"
            self.log(error_msg)
            logger.error(error_msg, exc_info=True)
            self.email_processor.close_connection()

def resource_path(relative_path):
    """Obtém o caminho absoluto para recursos, funciona para dev e para o PyInstaller"""
    try:
        # PyInstaller cria uma pasta temporária e armazena o caminho em _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailApp(root)
    root.mainloop()