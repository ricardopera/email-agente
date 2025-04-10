import tkinter as tk
from tkinter import ttk, simpledialog, messagebox, BooleanVar, Checkbutton, StringVar, Frame, Listbox, MULTIPLE, Toplevel, filedialog
import threading
import sys
import os
import logging
import re
import pandas as pd
from src.email_processor import EmailProcessor, logger

class EmailApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Agente de Extração de Emails")
        # Definindo tamanho mínimo inicial, mas permitindo ajuste automático
        self.root.geometry("600x500")  # Tamanho inicial mínimo
        self.root.minsize(600, 500)    # Tamanho mínimo permitido
        self.root.resizable(True, True)
        
        self.email_processor = EmailProcessor()
        
        # Configurar o estilo
        self.style = ttk.Style()
        self.style.configure("TFrame", background="#f0f0f0")
        self.style.configure("TLabel", background="#f0f0f0", font=("Arial", 10))
        self.style.configure("TButton", font=("Arial", 10))
        self.style.configure("Header.TLabel", font=("Arial", 12, "bold"))
        
        # Campos personalizados para extração com tipo de formato
        self.custom_fields = [
            {"name": "Número processo CNJ", "pattern": r'(?:Número processo CNJ|Processo)[\s:]*([\d.-]+)', "format": "texto"},
            {"name": "Valor liquido transferido para parte", "pattern": r'Valor liquido transferido para parte:[\s]*R\$([\d.,]+)', "format": "número"},
        ]
        
        # Campos adicionais para extração de arquivo Excel
        self.additional_excel_file = ""  # Caminho para o arquivo Excel adicional
        self.additional_fields = []  # Campos adicionais para extração do Excel
        self.key_field = ""  # Campo chave para relacionar os dados
        
        # Carregar configurações salvas
        self.load_saved_config()
        
        self.create_widgets()
        
    def load_saved_config(self):
        """Carrega as configurações salvas anteriormente"""
        try:
            config = self.email_processor.load_config()
            self.saved_email = config.get('email_user', '')
            self.saved_server = config.get('imap_host', 'mail.itajai.sc.gov.br')
            self.saved_subject = config.get('search_subject', 'Confirmacao de transferencia bancaria')
            
            # Carregar campos de extração salvos
            if 'custom_fields' in config:
                self.custom_fields = config.get('custom_fields', self.custom_fields)
                self.email_processor.custom_fields = self.custom_fields
            
            logger.info("Configurações carregadas na interface")
        except Exception as e:
            logger.error(f"Erro ao carregar configurações na interface: {str(e)}")
            self.saved_email = ''
            self.saved_server = 'mail.itajai.sc.gov.br'
            self.saved_subject = 'Confirmacao de transferencia bancaria'
        
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
        self.timeout_var = tk.StringVar(value="60")
        ttk.Entry(connection_frame, textvariable=self.timeout_var, width=5).grid(row=1, column=1, sticky=tk.W, pady=2)
        
        # Campo para o assunto
        ttk.Label(main_frame, text="Texto do Assunto:").grid(row=6, column=0, sticky=tk.W, pady=5)
        self.subject_var = tk.StringVar(value=self.saved_subject)
        ttk.Entry(main_frame, textvariable=self.subject_var, width=40).grid(row=6, column=1, sticky=tk.W, pady=5)
        
        # Frame para os campos de extração personalizados
        extraction_frame = ttk.LabelFrame(main_frame, text="Campos para Extração", padding=(10, 5))
        extraction_frame.grid(row=7, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # Header para a lista de campos
        header_frame = ttk.Frame(extraction_frame)
        header_frame.pack(fill=tk.X, padx=5, pady=(5, 0))
        
        ttk.Label(header_frame, text="Nome do Campo", width=30).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(header_frame, text="Formato", width=10).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(header_frame, text="").pack(side=tk.LEFT)  # Espaço para o botão de remover
        
        # Container para os campos dinâmicos
        self.fields_container = ttk.Frame(extraction_frame)
        self.fields_container.pack(fill=tk.BOTH, padx=5, pady=5)
        
        # Botões para gerenciar campos
        btn_frame = ttk.Frame(extraction_frame)
        btn_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(btn_frame, text="+", width=3, command=self.add_field).pack(side=tk.LEFT)
        
        # Carrega os campos personalizados salvos
        self.field_widgets = []
        for field in self.custom_fields:
            self.add_field_widget(field["name"], field.get("format", "texto"))
        
        # Frame para os campos adicionais a serem extraídos de um arquivo Excel
        additional_frame = ttk.LabelFrame(main_frame, text="Campos Adicionais (Arquivo Excel)", padding=(10, 5))
        additional_frame.grid(row=8, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # Seleção do arquivo Excel
        excel_frame = ttk.Frame(additional_frame)
        excel_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(excel_frame, text="Arquivo Excel:").pack(side=tk.LEFT, padx=(0, 5))
        self.excel_file_var = tk.StringVar()
        excel_entry = ttk.Entry(excel_frame, textvariable=self.excel_file_var, width=40)
        excel_entry.pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)
        
        # Botão para selecionar o arquivo
        excel_btn = ttk.Button(excel_frame, text="Selecionar", command=self.select_excel_file)
        excel_btn.pack(side=tk.LEFT)
        
        # Campo chave
        key_frame = ttk.Frame(additional_frame)
        key_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(key_frame, text="Campo Chave:").pack(side=tk.LEFT, padx=(0, 5))
        self.key_field_var = tk.StringVar()
        ttk.Entry(key_frame, textvariable=self.key_field_var, width=30).pack(side=tk.LEFT, padx=(0, 5))
        
        # Header para campos adicionais
        add_header_frame = ttk.Frame(additional_frame)
        add_header_frame.pack(fill=tk.X, padx=5, pady=(10, 0))
        
        ttk.Label(add_header_frame, text="Campos Adicionais", width=30).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(add_header_frame, text="").pack(side=tk.LEFT)  # Espaço para o botão de remover
        
        # Container para os campos adicionais dinâmicos
        self.additional_fields_container = ttk.Frame(additional_frame)
        self.additional_fields_container.pack(fill=tk.BOTH, padx=5, pady=5)
        
        # Botões para gerenciar campos adicionais
        add_btn_frame = ttk.Frame(additional_frame)
        add_btn_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(add_btn_frame, text="+", width=3, command=self.add_additional_field).pack(side=tk.LEFT)
        
        # Inicializa os widgets de campos adicionais
        self.additional_field_widgets = []
        # Se já tiver campos adicionais salvos, carregá-los
        if self.additional_fields:
            for field in self.additional_fields:
                self.add_additional_field_widget(field)
        else:
            # Adicionar pelo menos um campo adicional vazio
            self.add_additional_field_widget("")
        
        # Saída de log
        ttk.Label(main_frame, text="Log de Operações:").grid(row=9, column=0, columnspan=2, sticky=tk.W, pady=(10, 5))
        
        self.log_text = tk.Text(main_frame, height=10, width=60, wrap=tk.WORD)
        self.log_text.grid(row=10, column=0, columnspan=2, pady=5)
        
        # Scroll para o log
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=10, column=2, sticky=(tk.N, tk.S))
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        # Botões
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=11, column=0, columnspan=2, pady=15)
        
        ttk.Button(button_frame, text="Processar Emails", command=self.start_processing).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="Teste de Conexão", command=self.test_connection).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="Sair", command=self.root.quit).pack(side=tk.LEFT, padx=10)
        
        # Ajustar tamanho da janela após criar todos os widgets
        self.adjust_window_size()
    
    def add_field(self):
        """Adiciona um novo campo de extração personalizado"""
        self.add_field_widget("")
    
    def add_field_widget(self, default_text="", default_format="texto"):
        """Adiciona um widget de campo personalizado ao container com seleção de formato"""
        frame = ttk.Frame(self.fields_container)
        frame.pack(fill=tk.X, pady=2)
        
        field_var = StringVar(value=default_text)
        entry = ttk.Entry(frame, textvariable=field_var, width=30)
        entry.pack(side=tk.LEFT, padx=(0, 5))
        
        # Adicionar combo box para seleção de formato
        format_var = StringVar(value=default_format)
        format_combo = ttk.Combobox(frame, textvariable=format_var, width=10, state="readonly")
        format_combo['values'] = ('texto', 'número', 'data')
        format_combo.pack(side=tk.LEFT, padx=(0, 5))
        
        remove_btn = ttk.Button(frame, text="-", width=2, 
                              command=lambda f=frame, v=field_var, fv=format_var: self.remove_field(f, v, fv))
        remove_btn.pack(side=tk.LEFT)
        
        self.field_widgets.append((frame, field_var, entry, format_var))
        
        # Após adicionar um novo campo, reajustar o tamanho da janela
        self.adjust_window_size()
        
        return field_var
    
    def remove_field(self, frame, field_var, format_var):
        """Remove um campo personalizado"""
        # Não permitir remover se tiver apenas um campo
        if len(self.field_widgets) <= 1:
            messagebox.showwarning("Aviso", "Não é possível remover todos os campos. Pelo menos um deve existir.")
            return
            
        # Remover o widget
        for i, (f, v, e, fv) in enumerate(self.field_widgets):
            if f == frame and v == field_var:
                frame.destroy()
                self.field_widgets.pop(i)
                break
        
        # Após remover um campo, reajustar o tamanho da janela
        self.adjust_window_size()
    
    def add_additional_field(self):
        """Adiciona um novo campo adicional para extração do Excel"""
        self.add_additional_field_widget("")
    
    def add_additional_field_widget(self, default_text=""):
        """Adiciona um widget de campo adicional ao container"""
        frame = ttk.Frame(self.additional_fields_container)
        frame.pack(fill=tk.X, pady=2)
        
        field_var = StringVar(value=default_text)
        entry = ttk.Entry(frame, textvariable=field_var, width=30)
        entry.pack(side=tk.LEFT, padx=(0, 5))
        
        remove_btn = ttk.Button(frame, text="-", width=2, 
                              command=lambda f=frame, v=field_var: self.remove_additional_field(f, v))
        remove_btn.pack(side=tk.LEFT)
        
        self.additional_field_widgets.append((frame, field_var, entry))
        
        # Após adicionar um novo campo, reajustar o tamanho da janela
        self.adjust_window_size()
        
        return field_var
    
    def remove_additional_field(self, frame, field_var):
        """Remove um campo adicional"""
        # Não permitir remover se tiver apenas um campo
        if len(self.additional_field_widgets) <= 1:
            messagebox.showwarning("Aviso", "Não é possível remover todos os campos adicionais. Pelo menos um deve existir.")
            return
            
        # Remover o widget
        for i, (f, v, e) in enumerate(self.additional_field_widgets):
            if f == frame and v == field_var:
                frame.destroy()
                self.additional_field_widgets.pop(i)
                break
        
        # Após remover um campo, reajustar o tamanho da janela
        self.adjust_window_size()
    
    def select_excel_file(self):
        """Abre um diálogo para selecionar o arquivo Excel"""
        file_path = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx *.xls")])
        if file_path:
            self.excel_file_var.set(file_path)
            
            # Se o arquivo for selecionado, podemos mostrar as colunas disponíveis
            try:
                if file_path.lower().endswith(('.xlsx', '.xls')):
                    # Carregar o arquivo Excel para visualizar colunas disponíveis
                    df = pd.read_excel(file_path)
                    if not df.empty:
                        columns = df.columns.tolist()
                        self.show_available_columns(columns)
            except Exception as e:
                logger.error(f"Erro ao ler arquivo Excel: {str(e)}")
                messagebox.showerror("Erro", f"Não foi possível ler o arquivo Excel: {str(e)}")
    
    def show_available_columns(self, columns):
        """Mostra uma janela com as colunas disponíveis no Excel para referência"""
        if not columns:
            return
            
        # Criar uma janela popup para mostrar as colunas disponíveis
        columns_window = Toplevel(self.root)
        columns_window.title("Colunas Disponíveis")
        columns_window.geometry("300x300")
        
        ttk.Label(columns_window, text="Colunas disponíveis no arquivo Excel:").pack(pady=10)
        
        # Lista de colunas
        columns_listbox = Listbox(columns_window, width=40, height=15)
        columns_listbox.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)
        
        # Preencher a lista
        for col in columns:
            columns_listbox.insert(tk.END, col)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(columns_listbox, orient="vertical", command=columns_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        columns_listbox.config(yscrollcommand=scrollbar.set)
        
        # Botão de fechar
        ttk.Button(columns_window, text="Fechar", command=columns_window.destroy).pack(pady=10)
    
    def adjust_window_size(self):
        """Ajusta o tamanho da janela com base no conteúdo atual"""
        # Garantir que todas as atualizações pendentes sejam processadas
        self.root.update_idletasks()
        
        # Altura base e incremento por campo
        base_height = 600  # Altura base com um campo
        field_height = 35  # Altura estimada por campo
        
        # Calcular a altura necessária baseada no número total de campos
        total_fields = len(self.field_widgets) + len(self.additional_field_widgets)
        required_height = base_height + ((total_fields - 1) * field_height)
        
        # Adicionar espaço extra para garantir que tudo fique visível
        padding = 120
        required_height += padding
        
        # Obter a largura atual (manter a mesma)
        current_width = self.root.winfo_width()
        width = max(current_width, 600)
        
        # Limitar a altura máxima para não exceder o tamanho da tela
        max_height = self.root.winfo_screenheight() - 100
        height = min(required_height, max_height)
        
        # Garantir altura mínima
        height = max(height, 600)
        
        # Centralizar a janela na tela
        x = (self.root.winfo_screenwidth() - width) // 2
        y = (self.root.winfo_screenheight() - height) // 2
        
        # Definir geometria final
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        
        # Forçar atualização da interface
        self.root.update()
    
    def get_custom_fields(self):
        """Obtém todos os campos personalizados inseridos pelo usuário com seu formato"""
        fields = []
        for _, var, _, format_var in self.field_widgets:
            field_name = var.get().strip()
            field_format = format_var.get().strip()
            
            if field_name:
                # Criar um padrão que capture especificamente o texto entre ": " e a quebra de linha
                pattern = r'{}:\s*([^\r\n]+)'.format(re.escape(field_name))
                fields.append({
                    "name": field_name, 
                    "pattern": pattern,
                    "format": field_format
                })
        return fields
    
    def get_additional_fields(self):
        """Obtém todos os campos adicionais inseridos pelo usuário"""
        fields = []
        for _, var, _ in self.additional_field_widgets:
            field_name = var.get().strip()
            if field_name:
                fields.append(field_name)
        return fields
    
    def save_additional_fields_config(self):
        """Salva as configurações de campos adicionais"""
        self.additional_excel_file = self.excel_file_var.get().strip()
        self.key_field = self.key_field_var.get().strip()
        self.additional_fields = self.get_additional_fields()
        
        # Atualizar as configurações no processador de email
        self.email_processor.additional_excel_file = self.additional_excel_file
        self.email_processor.key_field = self.key_field
        self.email_processor.additional_fields = self.additional_fields
    
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
            
        # Obter campos personalizados
        self.custom_fields = self.get_custom_fields()
        if not self.custom_fields:
            messagebox.showerror("Erro", "É necessário definir pelo menos um campo para extração!")
            return
            
        # Atualizar os campos de extração no processador
        self.email_processor.custom_fields = self.custom_fields
        
        # Salvar as configurações para uso futuro
        self.email_processor.search_subject = subject
        self.email_processor.save_config(email, server, subject, self.custom_fields)
        
        # Salvar configurações de campos adicionais
        self.save_additional_fields_config()
        
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