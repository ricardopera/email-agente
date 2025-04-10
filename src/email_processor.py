import imaplib
import email
import re
import os
import json
import logging
import pandas as pd
from email.header import decode_header
from datetime import datetime
import chardet

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    filename='email_extrator.log'
)

logger = logging.getLogger('email_extrator')

class EmailProcessor:
    def __init__(self):
        self.imap_server = None
        self.email_user = None
        self.email_pass = None
        self.imap_host = None
        self.imap_port = 993  # Definindo a porta padrão IMAPS
        self.search_subject = None
        
        # Campos personalizados para extração
        self.custom_fields = [
            {"name": "Número processo CNJ", "pattern": r'(?:Número processo CNJ|Processo)[\s:]*([\d.-]+)', "format": "texto"},
            {"name": "Valor liquido transferido para parte", "pattern": r'Valor liquido transferido para parte:[\s]*R\$([\d.,]+)', "format": "número"}
        ]
        
        # Campos adicionais para extração de arquivo Excel
        self.additional_excel_file = ""  # Caminho para o arquivo Excel adicional
        self.additional_fields = []  # Campos adicionais para extração do Excel
        self.key_field = ""  # Campo chave para relacionar os dados
        self.reference_data = None  # DataFrame com dados de referência
        
        self.extracted_data = []
        self.config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'config.json')
        
    def save_config(self, email_user, imap_host, search_subject, custom_fields=None):
        """Salva as configurações para uso futuro"""
        try:
            config = {
                'email_user': email_user,
                'imap_host': imap_host,
                'search_subject': search_subject
            }
            
            # Salvar os campos personalizados, se fornecidos
            if custom_fields:
                config['custom_fields'] = custom_fields
            
            # Salvar as configurações de campos adicionais
            config['additional_excel_file'] = self.additional_excel_file
            config['key_field'] = self.key_field
            config['additional_fields'] = self.additional_fields
            
            with open(self.config_file, 'w') as f:
                json.dump(config, f)
                
            logger.info("Configurações salvas com sucesso")
            return True
        except Exception as e:
            logger.error(f"Erro ao salvar configurações: {str(e)}")
            return False
            
    def load_config(self):
        """Carrega as configurações salvas"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                
                # Carregar configurações de campos adicionais se existirem
                if 'additional_excel_file' in config:
                    self.additional_excel_file = config['additional_excel_file']
                if 'key_field' in config:
                    self.key_field = config['key_field']
                if 'additional_fields' in config:
                    self.additional_fields = config['additional_fields']
                
                logger.info("Configurações carregadas com sucesso")
                return config
            else:
                logger.info("Arquivo de configuração não encontrado")
                return {}
        except Exception as e:
            logger.error(f"Erro ao carregar configurações: {str(e)}")
            return {}

    def connect_to_server(self, email_user, email_pass, imap_host, imap_port=993, use_ssl=True, timeout=30):
        """Conecta ao servidor IMAP"""
        try:
            logger.info(f"Tentando conectar ao servidor {imap_host}:{imap_port} (SSL: {use_ssl}, Timeout: {timeout}s)")
            
            # Verificação de rede antes de tentar conexão
            self._check_network_connectivity(imap_host, imap_port)
            
            # Salva as configurações antes da conexão
            self.email_user = email_user
            self.email_pass = email_pass
            self.imap_host = imap_host
            self.imap_port = imap_port
            
            # Definindo timeout para a conexão
            imaplib.IMAP4.TIMEOUT = timeout
            
            # Tentativa de conectar via SSL ou TLS conforme configurado
            if use_ssl:
                logger.debug(f"Criando conexão SSL para {imap_host}:{imap_port}")
                # Usando socket com timeout explícito para mais controle
                import socket
                socket.setdefaulttimeout(timeout)
                self.imap_server = imaplib.IMAP4_SSL(imap_host, imap_port)
            else:
                # Conectar sem SSL e então iniciar TLS
                logger.debug(f"Criando conexão sem SSL para {imap_host}:{imap_port}")
                import socket
                socket.setdefaulttimeout(timeout)
                self.imap_server = imaplib.IMAP4(imap_host, imap_port)
                logger.debug("Iniciando conexão TLS")
                self.imap_server.starttls()
            
            logger.debug(f"Tentando login com usuário: {email_user}")
            self.imap_server.login(email_user, email_pass)
            
            logger.info("Conexão e login realizados com sucesso")
            return True
        except ConnectionRefusedError:
            error_msg = f"Conexão recusada pelo servidor {imap_host}:{imap_port}. Verifique se o endereço e porta estão corretos."
            logger.error(error_msg)
            messagebox.showerror("Erro de Conexão", error_msg)
            return False
        except TimeoutError:
            error_msg = f"Tempo limite excedido ao conectar a {imap_host}:{imap_port}. Verifique sua conexão e se o servidor está acessível.\n\nDica: Aumente o valor de timeout ou verifique se há firewalls bloqueando a conexão."
            logger.error(error_msg)
            messagebox.showerror("Erro de Timeout", error_msg)
            return False
        except socket.timeout:
            error_msg = f"Socket timeout ao conectar a {imap_host}:{imap_port}. Aumentar o timeout pode resolver este problema (WinError 10060)."
            logger.error(error_msg)
            messagebox.showerror("Erro de Timeout", error_msg)
            return False
        except imaplib.IMAP4.error as e:
            if 'LOGIN failed' in str(e):
                error_msg = "Falha no login. Verifique seu nome de usuário e senha."
            else:
                error_msg = f"Erro de IMAP: {str(e)}"
            logger.error(error_msg)
            messagebox.showerror("Erro IMAP", error_msg)
            return False
        except Exception as e:
            error_msg = f"Não foi possível conectar ao servidor: {str(e)}"
            logger.error(error_msg)
            logger.error(traceback.format_exc())
            messagebox.showerror("Erro de Conexão", error_msg)
            return False

    def _check_network_connectivity(self, host, port):
        """Verifica a conectividade de rede antes de tentar a conexão IMAP"""
        import socket
        
        logger.info(f"Verificando conectividade de rede com {host}:{port}")
        
        try:
            # Tenta resolver o nome do host
            ip_address = socket.gethostbyname(host)
            logger.info(f"Resolvido {host} para {ip_address}")
            
            # Testa se consegue abrir uma conexão TCP
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.settimeout(5)  # Timeout curto apenas para teste
            
            result = s.connect_ex((ip_address, port))
            s.close()
            
            if result == 0:
                logger.info(f"Porta {port} está aberta no host {host}")
            else:
                logger.warning(f"Porta {port} parece estar fechada ou bloqueada no host {host} (código {result})")
                logger.warning("Isso pode indicar um firewall bloqueando a conexão - considere verificar as configurações de rede")
        except socket.gaierror:
            logger.error(f"Não foi possível resolver o nome de host: {host}")
            logger.warning("Verifique se o nome do servidor está correto ou se há problemas com o DNS")
        except Exception as e:
            logger.error(f"Erro ao verificar conectividade: {str(e)}")

    def clean_text(self, text):
        """Limpa o texto removendo caracteres especiais e normalizando espaços"""
        if text is None:
            return ""
        # Remove quebras de linha e espaços extras
        text = re.sub(r'\s+', ' ', text.strip())
        return text

    def decode_email_subject(self, subject):
        """Decodifica o assunto do email, tratando adequadamente caracteres do português"""
        if subject is None:
            return ""
        decoded_parts = []
        for part, encoding in decode_header(subject):
            if isinstance(part, bytes):
                if encoding:
                    try:
                        # Tenta usar a codificação especificada primeiro
                        decoded_parts.append(part.decode(encoding))
                    except (UnicodeDecodeError, LookupError):
                        try:
                            # Tenta UTF-8 que é comum para caracteres especiais do português
                            decoded_parts.append(part.decode('utf-8'))
                        except UnicodeDecodeError:
                            # Fallback para latin-1 (ISO-8859-1) que é comum em emails em português
                            try:
                                decoded_parts.append(part.decode('latin-1'))
                            except:
                                # Último recurso: substituição de caracteres não reconhecidos
                                decoded_parts.append(part.decode('utf-8', errors='replace'))
                else:
                    try:
                        # Tenta UTF-8 primeiro
                        decoded_parts.append(part.decode('utf-8'))
                    except UnicodeDecodeError:
                        try:
                            # Tenta latin-1 que é comum para português
                            decoded_parts.append(part.decode('latin-1'))
                        except:
                            # Último recurso com substituição de caracteres
                            decoded_parts.append(part.decode('utf-8', errors='replace'))
            else:
                decoded_parts.append(part)
        return ''.join(decoded_parts)

    def get_email_content(self, msg):
        """Extrai o conteúdo do email (texto), tratando adequadamente caracteres do português"""
        content = ""
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                
                # Pular anexos
                if "attachment" in content_disposition:
                    continue
                
                # Obter conteúdo de texto
                if content_type == "text/plain" or content_type == "text/html":
                    try:
                        # Tentar obter a codificação especificada no email
                        charset = part.get_content_charset()
                        if charset:
                            body = part.get_payload(decode=True).decode(charset)
                        else:
                            # Se não especificado, tentar UTF-8
                            body = part.get_payload(decode=True).decode('utf-8')
                        content += body
                    except UnicodeDecodeError:
                        try:
                            # UTF-8 falhou, tentar latin-1 (comum para português)
                            body = part.get_payload(decode=True).decode('latin-1')
                            content += body
                        except:
                            # Último recurso: substituição de caracteres
                            body = part.get_payload(decode=True).decode('utf-8', errors='replace')
                            content += body
        else:
            # Mensagens não multipart
            content_type = msg.get_content_type()
            if content_type == "text/plain" or content_type == "text/html":
                try:
                    # Tentar obter a codificação especificada no email
                    charset = msg.get_content_charset()
                    if charset:
                        content = msg.get_payload(decode=True).decode(charset)
                    else:
                        # Se não especificado, tentar UTF-8
                        content = msg.get_payload(decode=True).decode('utf-8')
                except UnicodeDecodeError:
                    try:
                        # UTF-8 falhou, tentar latin-1 (comum para português)
                        content = msg.get_payload(decode=True).decode('latin-1')
                    except:
                        # Último recurso: substituição de caracteres
                        content = msg.get_payload(decode=True).decode('utf-8', errors='replace')
        
        return content

    def extract_fields(self, text):
        """Extrai os campos personalizados do texto do email baseado no formato especificado"""
        extracted_fields = {}
        
        # Aplicar cada padrão de extração personalizado
        for field in self.custom_fields:
            field_name = field["name"]
            field_format = field.get("format", "texto")  # Formato padrão é texto
            
            # Criar um padrão que capture especificamente o texto entre ": " e a quebra de linha
            # Primeiro tentamos encontrar o campo exato com dois-pontos
            pattern = r'{}:\s*([^\r\n]+)'.format(re.escape(field_name))
            match = re.search(pattern, text)
            
            # Se não encontrar com o formato exato, tenta um formato mais flexível
            if not match:
                pattern = r'{}\s*:?\s*([^\r\n]+)'.format(re.escape(field_name))
                match = re.search(pattern, text)
            
            if match:
                raw_value = match.group(1).strip()
                # Processar o valor conforme o formato especificado
                processed_value = self.process_value(raw_value, field_format)
                extracted_fields[field_name] = processed_value
            else:
                # Se não encontrar, definir como vazio
                extracted_fields[field_name] = ""
                logger.warning(f"Campo '{field_name}' não encontrado no texto do email")
                
        return extracted_fields
        
    def process_value(self, value, format_type):
        """
        Processa o valor extraído de acordo com o formato especificado
        - texto: mantém como texto
        - número: converte para número (formato brasileiro para decimal)
        - data: converte para formato de data
        """
        if not value:
            return value
            
        try:
            if format_type == "número":
                # Remover símbolos de moeda (R$, $, etc) e espaços
                value = re.sub(r'[R$\s]', '', value)
                # Substituir pontos de milhar e vírgulas decimais para formato numérico
                # Ex: "1.234,56" -> 1234.56
                if ',' in value:
                    # Remove pontos (separadores de milhar)
                    value = value.replace('.', '')
                    # Substitui vírgula por ponto (separador decimal)
                    value = value.replace(',', '.')
                return float(value)
                
            elif format_type == "data":
                # Converte data do formato brasileiro (dd/mm/yyyy) para formato ISO
                if '/' in value:
                    day, month, year = value.split('/')
                    return f"{year}-{month.zfill(2)}-{day.zfill(2)}"
                return value
                
            else:  # formato texto ou não especificado
                return value
                
        except Exception as e:
            logger.warning(f"Erro ao processar valor '{value}' para formato '{format_type}': {str(e)}")
            return value

    def search_emails(self, search_subject):
        """Busca emails não lidos com o assunto específico, suportando caracteres do português"""
        try:
            self.imap_server.select('INBOX')
            
            # Codificação para lidar com caracteres especiais no assunto da busca
            # Para IMAP, usamos a codificação UTF-8 em strings literais
            try:
                # Usar caracteres literais para garantir que caracteres especiais sejam tratados corretamente
                # Isso é necessário para suportar acentos e outros caracteres especiais do português na pesquisa
                self.imap_server.literal = search_subject.encode('utf-8')
                status, messages = self.imap_server.search(None, 'UNSEEN SUBJECT')
            except (AttributeError, UnicodeEncodeError):
                # Fallback para o método tradicional se o anterior falhar
                # Isso pode não funcionar perfeitamente com caracteres especiais
                logger.warning("Usando método de busca alternativo para sujeito com caracteres especiais")
                status, messages = self.imap_server.search(None, '(UNSEEN SUBJECT "{}")'.format(search_subject))
            
            if status != 'OK':
                messagebox.showwarning("Aviso", "Não foi possível buscar emails.")
                return []
            
            email_ids = messages[0].split()
            
            if not email_ids:
                messagebox.showinfo("Informação", "Nenhum email não lido encontrado com o assunto especificado.")
                return []
                
            return email_ids
        except Exception as e:
            logger.error(f"Erro ao buscar emails: {str(e)}", exc_info=True)
            messagebox.showerror("Erro", f"Erro ao buscar emails: {str(e)}")
            return []

    def process_emails(self, email_ids):
        """Processa cada email encontrado"""
        total_emails = len(email_ids)
        processed = 0
        
        # Limpar dados extraídos anteriormente
        self.extracted_data = []
        
        # Carregar dados de referência do Excel se campos adicionais estiverem configurados
        reference_data_loaded = False
        if self.additional_excel_file and self.key_field and self.additional_fields:
            reference_data_loaded = self.load_reference_data()
            if reference_data_loaded:
                logger.info(f"Dados de referência carregados do arquivo: {self.additional_excel_file}")
            else:
                logger.warning(f"Não foi possível carregar dados de referência: {self.additional_excel_file}")
        
        logger.info(f"Iniciando processamento de {total_emails} emails")
        
        for email_id in email_ids:
            try:
                logger.debug(f"Buscando email com ID: {email_id}")
                status, data = self.imap_server.fetch(email_id, '(RFC822)')
                
                if status != 'OK':
                    logger.warning(f"Falha ao buscar email com ID {email_id}, status: {status}")
                    continue
                    
                logger.debug(f"Email com ID {email_id} recuperado com sucesso")
                raw_email = data[0][1]
                msg = email.message_from_bytes(raw_email)
                
                # Extrair informações do email
                subject = self.decode_email_subject(msg['Subject'])
                sender = msg['From']
                date = msg['Date']
                
                logger.info(f"Processando email: '{subject}' de {sender}")
                
                # Obter conteúdo do email
                content = self.get_email_content(msg)
                
                # Extrair campos personalizados
                extracted_fields = self.extract_fields(content)
                
                if extracted_fields:
                    # Se temos dados de referência carregados, buscar campos adicionais
                    if reference_data_loaded and self.key_field in self.custom_fields:
                        # Buscar o valor do campo chave entre os campos extraídos
                        key_field_name = self.key_field
                        key_value = None
                        
                        # Encontrar o campo chave entre os extraídos
                        for field in self.custom_fields:
                            if field["name"] == key_field_name and field["name"] in extracted_fields:
                                key_value = extracted_fields[field["name"]]
                                break
                        
                        if key_value:
                            # Buscar dados adicionais usando o valor do campo chave
                            additional_data = self.get_additional_fields_data(key_value)
                            if additional_data:
                                logger.info(f"Dados adicionais encontrados para a chave '{key_value}'")
                                # Adicionar os campos adicionais aos dados extraídos
                                extracted_fields.update(additional_data)
                            else:
                                logger.warning(f"Nenhum dado adicional encontrado para a chave '{key_value}'")
                        else:
                            logger.warning(f"Campo chave '{key_field_name}' não encontrado ou vazio nos dados extraídos")
                    
                    # Adicionar metadados do email apenas se solicitado
                    # (mantemos apenas os campos especificados pelo usuário)
                    logger.info(f"Campos extraídos: {extracted_fields}")
                    self.extracted_data.append(extracted_fields)
                else:
                    logger.warning(f"Nenhum campo personalizado encontrado no email com assunto: {subject}")
                    
                processed += 1
                logger.debug(f"Email processado: {processed}/{total_emails}")
                
            except Exception as e:
                logger.error(f"Erro ao processar email ID {email_id}: {str(e)}", exc_info=True)
        
        logger.info(f"Processamento finalizado. {processed} emails processados, {len(self.extracted_data)} registros extraídos.")
        return processed

    def save_to_excel(self, filename=None):
        """Salva os dados extraídos em um arquivo Excel com os formatos adequados"""
        if not self.extracted_data:
            messagebox.showinfo("Informação", "Nenhum dado para salvar.")
            return False
            
        try:
            # Criar DataFrame com os dados extraídos
            df = pd.DataFrame(self.extracted_data)
            
            # Definir nome do arquivo Excel
            if filename is None:
                # Usar nome padrão com timestamp para evitar sobrescrever arquivos existentes
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                default_filename = f"processos_extraidos_{timestamp}.xlsx"
                # Verificar se temos arquivo Excel padrão no diretório atual
                excel_files = [f for f in os.listdir(os.getcwd()) if f.endswith('.xlsx') and f.startswith('processos_extraidos')]
                
                # Se existir algum arquivo Excel de dados extraídos, usar o mais recente
                if excel_files:
                    filename = sorted(excel_files)[-1]  # Pega o mais recente alfabeticamente (geralmente pelo timestamp)
                    logger.info(f"Encontrado arquivo existente: {filename}")
                else:
                    filename = default_filename
                    logger.info(f"Criando novo arquivo: {filename}")
            
            # Caminho completo para o arquivo
            full_path = os.path.join(os.getcwd(), filename)
            
            # Verificar se o arquivo já existe
            if os.path.exists(full_path):
                try:
                    # Criar um escritor Excel com formato
                    with pd.ExcelWriter(full_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                        # Ler o arquivo existente para obter os dados
                        existing_df = pd.read_excel(full_path)
                        
                        # Garantir que temos todas as colunas necessárias
                        for col in df.columns:
                            if col not in existing_df.columns:
                                existing_df[col] = ""
                        
                        # Garantir que novos dados têm as mesmas colunas do arquivo existente
                        for col in existing_df.columns:
                            if col not in df.columns:
                                df[col] = ""
                        
                        # Concatenar os dados existentes com os novos
                        combined_df = pd.concat([existing_df, df], ignore_index=True)
                        
                        # Configurar os formatos para cada coluna baseado nos formatos definidos
                        for field in self.custom_fields:
                            field_name = field["name"]
                            field_format = field.get("format", "texto")
                            
                            # Aplicar formatos específicos às colunas do DataFrame
                            if field_name in combined_df.columns:
                                if field_format == "número":
                                    # Converter para float caso ainda não seja
                                    combined_df[field_name] = pd.to_numeric(combined_df[field_name], errors='coerce')
                                elif field_format == "data":
                                    # Converter para datetime
                                    combined_df[field_name] = pd.to_datetime(combined_df[field_name], errors='coerce')
                        
                        # Salvar o arquivo com todos os dados
                        combined_df.to_excel(writer, index=False, sheet_name='Sheet1')
                        
                    logger.info(f"Acrescentando dados ao arquivo existente: {full_path}")
                except Exception as e:
                    # Se houver erro ao ler o arquivo existente, criar um novo
                    logger.warning(f"Não foi possível atualizar o arquivo existente: {str(e)}")
                    logger.info(f"Criando novo arquivo: {full_path}")
                    
                    # Criar um novo arquivo com os formatos adequados
                    with pd.ExcelWriter(full_path, engine='openpyxl') as writer:
                        # Configurar os formatos para cada coluna
                        for field in self.custom_fields:
                            field_name = field["name"]
                            field_format = field.get("format", "texto")
                            
                            # Aplicar formatos específicos às colunas do DataFrame
                            if field_name in df.columns:
                                if field_format == "número":
                                    # Converter para float caso ainda não seja
                                    df[field_name] = pd.to_numeric(df[field_name], errors='coerce')
                                elif field_format == "data":
                                    # Converter para datetime
                                    df[field_name] = pd.to_datetime(df[field_name], errors='coerce')
                        
                        df.to_excel(writer, index=False)
            else:
                # O arquivo não existe, criar um novo
                logger.info(f"Criando novo arquivo: {full_path}")
                
                with pd.ExcelWriter(full_path, engine='openpyxl') as writer:
                    # Configurar os formatos para cada coluna
                    for field in self.custom_fields:
                        field_name = field["name"]
                        field_format = field.get("format", "texto")
                        
                        # Aplicar formatos específicos às colunas do DataFrame
                        if field_name in df.columns:
                            if field_format == "número":
                                # Converter para float caso ainda não seja
                                df[field_name] = pd.to_numeric(df[field_name], errors='coerce')
                            elif field_format == "data":
                                # Converter para datetime
                                df[field_name] = pd.to_datetime(df[field_name], errors='coerce')
                    
                    df.to_excel(writer, index=False)
            
            messagebox.showinfo("Sucesso", f"Dados salvos com sucesso em {full_path}")
            return True
        except Exception as e:
            error_msg = f"Erro ao salvar arquivo Excel: {str(e)}"
            logger.error(error_msg, exc_info=True)
            messagebox.showerror("Erro", error_msg)
            return False

    def close_connection(self):
        """Fecha a conexão com o servidor IMAP"""
        if self.imap_server:
            try:
                self.imap_server.close()
                self.imap_server.logout()
            except:
                pass

    def load_reference_data(self):
        """Carrega os dados do arquivo Excel de referência"""
        if not self.additional_excel_file or not os.path.exists(self.additional_excel_file):
            logger.warning(f"Arquivo de referência não encontrado: {self.additional_excel_file}")
            return False
            
        try:
            # Carregar o arquivo Excel em um DataFrame
            self.reference_data = pd.read_excel(self.additional_excel_file)
            
            # Verificar se o campo chave existe no DataFrame
            if self.key_field and self.key_field not in self.reference_data.columns:
                logger.warning(f"Campo chave '{self.key_field}' não encontrado no arquivo Excel")
                return False
                
            # Verificar se os campos adicionais existem no DataFrame
            missing_fields = []
            for field in self.additional_fields:
                if field not in self.reference_data.columns:
                    missing_fields.append(field)
                    
            if missing_fields:
                logger.warning(f"Campos adicionais não encontrados no arquivo Excel: {', '.join(missing_fields)}")
                return False
                
            logger.info(f"Dados de referência carregados com sucesso: {len(self.reference_data)} registros")
            return True
        except Exception as e:
            logger.error(f"Erro ao carregar arquivo de referência: {str(e)}")
            return False

    def get_additional_fields_data(self, key_value):
        """
        Busca dados adicionais no arquivo Excel de referência com base no valor do campo chave
        
        Args:
            key_value: O valor do campo chave a ser buscado no Excel
            
        Returns:
            Um dicionário com os campos adicionais encontrados ou vazio se não encontrar
        """
        if self.reference_data is None or not self.key_field or not self.additional_fields:
            return {}
        
        try:
            # Verifica se o valor da chave está no DataFrame
            matches = self.reference_data[self.reference_data[self.key_field] == key_value]
            
            if matches.empty:
                logger.warning(f"Nenhum registro encontrado para o valor chave: {key_value}")
                return {}
            
            # Obtém o primeiro registro correspondente (assume-se chave única)
            first_match = matches.iloc[0]
            
            # Extrai os campos adicionais
            result = {}
            for field in self.additional_fields:
                if field in first_match:
                    result[field] = first_match[field]
            
            logger.info(f"Dados adicionais encontrados para chave '{key_value}': {result}")
            return result
        except Exception as e:
            logger.error(f"Erro ao buscar dados adicionais para '{key_value}': {str(e)}")
            return {}