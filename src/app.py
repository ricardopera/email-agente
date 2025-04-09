import tkinter as tk
from tkinter import ttk, simpledialog, messagebox, BooleanVar, Checkbutton, StringVar, Frame, Listbox, MULTIPLE, Toplevel
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
        
        # Padrões de campos disponíveis para extração
        self.available_fields = {
            'Número do Processo': 'Número processo CNJ',
            'Valor Líquido': 'Valor liquido transferido para parte',
            'Juiz': 'Juiz(a) autorizador(a)',
            'Chefe de Cartório': 'Chefe de cartório responsável',
            'Subconta': 'Subconta',
            'Valor Solicitado': 'Valor do pedido solicitado',
            'Valor Total': 'Valor total do pedido efetuado',
            'Tipo de Saque': 'Tipo de saque',
            'Beneficiado': 'Beneficiado',
            'CPF/CNPJ': 'CPF/CNPJ',
            'Data do Pedido': 'Data do pedido',
            'Data da Liberação': 'Data da liberação',
            'Banco': 'Banco',
            'Agência': 'Agência',
            'Conta': 'Conta',
            'Comprovante': 'Comprovante de liberação'
        }
        
        # Campos selecionados para extração (padrão)
        self.selected_fields = ['Número do Processo', 'Valor Líquido']
        
        # Carregar configurações salvas
        self.load_saved_config()
        
        self.create_widgets()
        
    def load_saved_config(self):
        """Carrega as configurações salvas anteriormente"""
        try:
            config = self.email_processor.load_config()
            self.saved_email = config.get('email_user', '')
            self.saved_server = config.get('imap_host', 'mail.itajai.sc.gov.br')  # Alterando o valor padrão
            self.saved_subject = config.get('search_subject', 'Confirmacao de transferencia bancaria')
            
            # Carregar campos de extração salvos
            if 'fields_to_extract' in config:
                self.selected_fields = config.get('fields_to_extract', ['Número do Processo', 'Valor Líquido'])
                self.email_processor.fields_to_extract = self.selected_fields
            
            logger.info("Configurações carregadas na interface")
        except Exception as e:
            logger.error(f"Erro ao carregar configurações na interface: {str(e)}")
            self.saved_email = ''
            self.saved_server = 'mail.itajai.sc.gov.br'  # Alterando o valor padrão
            self.saved_subject = 'Confirmacao de transferencia bancaria'  # Valor padrão do assunto
        
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
        
        # Botão para configurar campos de extração
        extraction_frame = ttk.Frame(main_frame)
        extraction_frame.grid(row=7, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(extraction_frame, text="Campos para Extração:").pack(side=tk.LEFT, padx=(0, 10))
        self.selected_fields_display = ttk.Label(extraction_frame, text=self.format_selected_fields())
        self.selected_fields_display.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(extraction_frame, text="Configurar", command=self.open_field_selector).pack(side=tk.LEFT)
        
        # Saída de log
        ttk.Label(main_frame, text="Log de Operações:").grid(row=8, column=0, columnspan=2, sticky=tk.W, pady=(10, 5))
        
        self.log_text = tk.Text(main_frame, height=10, width=60, wrap=tk.WORD)
        self.log_text.grid(row=9, column=0, columnspan=2, pady=5)
        
        # Scroll para o log
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=9, column=2, sticky=(tk.N, tk.S))
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        # Botões
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=10, column=0, columnspan=2, pady=15)
        
        ttk.Button(button_frame, text="Processar Emails", command=self.start_processing).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="Teste de Conexão", command=self.test_connection).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="Sair", command=self.root.quit).pack(side=tk.LEFT, padx=10)
    
    def format_selected_fields(self):
        """Formata os campos selecionados para exibição"""
        if not self.selected_fields:
            return "Nenhum campo selecionado"
        
        if len(self.selected_fields) <= 2:
            return ", ".join(self.selected_fields)
        else:
            return ", ".join(self.selected_fields[:2]) + f" e mais {len(self.selected_fields) - 2}..."
    
    def open_field_selector(self):
        """Abre a janela de seleção de campos para extração"""
        selector = Toplevel(self.root)
        selector.title("Selecionar Campos para Extração")
        selector.geometry("400x400")
        selector.transient(self.root)  # Faz a janela aparecer acima da janela principal
        
        ttk.Label(selector, text="Selecione os campos que deseja extrair dos emails:",
                 style="Header.TLabel").pack(pady=10, padx=20)
        
        # Criar listbox com scrollbar
        frame = ttk.Frame(selector)
        frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        scrollbar = ttk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        listbox = Listbox(frame, selectmode=MULTIPLE, height=15, width=40,
                         yscrollcommand=scrollbar.set)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        
        # Preencher a listbox com os campos disponíveis
        for field_name in self.available_fields:
            listbox.insert(tk.END, f"{field_name}: {self.available_fields[field_name]}")
            # Se o campo estiver na lista de selecionados, marcar ele
            if field_name in self.selected_fields:
                idx = list(self.available_fields.keys()).index(field_name)
                listbox.selection_set(idx)
        
        # Botões
        button_frame = ttk.Frame(selector)
        button_frame.pack(pady=15)
        
        ttk.Button(button_frame, text="Confirmar", command=lambda: self.save_selected_fields(listbox, selector)).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="Cancelar", command=selector.destroy).pack(side=tk.LEFT, padx=10)
        
        # Tornar a janela modal
        selector.grab_set()
        self.root.wait_window(selector)
    
    def save_selected_fields(self, listbox, selector):
        """Salva os campos selecionados"""
        selected_indices = listbox.curselection()
        field_names = list(self.available_fields.keys())
        
        self.selected_fields = [field_names[i] for i in selected_indices]
        
        if not self.selected_fields:
            messagebox.showwarning("Aviso", "É necessário selecionar ao menos um campo.")
            return
        
        # Atualizar os padrões de extração no EmailProcessor
        self.email_processor.fields_to_extract = self.selected_fields
        
        # Atualizar a exibição na interface
        self.selected_fields_display.config(text=self.format_selected_fields())
        
        # Atualizar os padrões de extração
        extraction_patterns = {}
        for field_name in self.selected_fields:
            description = self.available_fields[field_name]
            # Criar um padrão regex básico para cada campo
            pattern = r'{}[\\s:]*(.+?)(?=\\n|$)'.format(re.escape(description))
            extraction_patterns[field_name] = pattern
            
        # Definir padrões mais específicos para certos campos
        if 'Número do Processo' in self.selected_fields:
            extraction_patterns['Número do Processo'] = r'(?:Número processo CNJ|Processo)[\s:]*([\d.-]+)'
        
        if 'Valor Líquido' in self.selected_fields:
            extraction_patterns['Valor Líquido'] = r'Valor liquido transferido para parte:[\s]*R\$([\d.,]+)'
            
        self.email_processor.extraction_patterns = extraction_patterns
        
        self.log(f"Configurados {len(self.selected_fields)} campos para extração: {', '.join(self.selected_fields)}")
        
        # Fechar a janela de seleção
        selector.destroy()
    
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
        self.email_processor.save_config(email, server, subject, self.selected_fields)
        
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