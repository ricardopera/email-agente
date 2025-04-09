import imaplib
import email
from email.header import decode_header
import re
import pandas as pd
import os
import tkinter as tk
from tkinter import simpledialog, messagebox
import getpass
import sys
import traceback
import json
import logging
from datetime import datetime

# Configuração de logs
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    filename='email_extrator.log',
    filemode='a'
)
logger = logging.getLogger('EmailExtrator')

class EmailProcessor:
    def __init__(self):
        self.imap_server = None
        self.email_user = None
        self.email_pass = None
        self.imap_host = None
        self.imap_port = 993  # Definindo a porta padrão IMAPS
        self.search_subject = None
        # Padrões de extração padrão
        self.extraction_patterns = {
            'Número do Processo': r'(?:Número do Processo CNJ|Processo)[\s:]*([\d.-]+)',
            'Valor Líquido': r'Valor liquido transferido para parte:[\s]*R\$([\d.,]+)'
        }
        # Campos a serem extraídos (padrão)
        self.fields_to_extract = ['Número do Processo', 'Valor Líquido']
        self.extracted_data = []
        self.config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'config.json')
        
    def save_config(self, email_user, imap_host, search_subject, fields_to_extract=None):
        """Salva as configurações para uso futuro"""
        try:
            config = {
                'email_user': email_user,
                'imap_host': imap_host,
                'search_subject': search_subject
            }
            
            # Salvar os campos para extração, se fornecidos
            if fields_to_extract:
                config['fields_to_extract'] = fields_to_extract
            
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
            self.save_config(email_user, imap_host, self.search_subject, self.fields_to_extract)
            
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
        """Decodifica o assunto do email"""
        if subject is None:
            return ""
        decoded_parts = []
        for part, encoding in decode_header(subject):
            if isinstance(part, bytes):
                if encoding:
                    try:
                        decoded_parts.append(part.decode(encoding))
                    except:
                        decoded_parts.append(part.decode('utf-8', errors='replace'))
                else:
                    decoded_parts.append(part.decode('utf-8', errors='replace'))
            else:
                decoded_parts.append(part)
        return ''.join(decoded_parts)

    def get_email_content(self, msg):
        """Extrai o conteúdo do email (texto)"""
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
                        body = part.get_payload(decode=True).decode('utf-8', errors='replace')
                        content += body
                    except:
                        pass
        else:
            # Mensagens não multipart
            content_type = msg.get_content_type()
            if content_type == "text/plain" or content_type == "text/html":
                try:
                    content = msg.get_payload(decode=True).decode('utf-8', errors='replace')
                except:
                    pass
        
        return content

    def extract_fields(self, text):
        """Extrai os campos especificados do texto do email"""
        extracted_fields = {}
        for field, pattern in self.extraction_patterns.items():
            match = re.search(pattern, text)
            if match:
                extracted_fields[field] = match.group(1)
        return extracted_fields

    def search_emails(self, search_subject):
        """Busca emails não lidos com o assunto específico"""
        try:
            self.imap_server.select('INBOX')
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
            messagebox.showerror("Erro", f"Erro ao buscar emails: {str(e)}")
            return []

    def process_emails(self, email_ids):
        """Processa cada email encontrado"""
        total_emails = len(email_ids)
        processed = 0
        
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
                
                # Extrair campos especificados
                extracted_fields = self.extract_fields(content)
                
                if extracted_fields:
                    logger.info(f"Campos extraídos: {extracted_fields}")
                    extracted_fields.update({
                        'Assunto': subject,
                        'Remetente': sender,
                        'Data': date
                    })
                    self.extracted_data.append(extracted_fields)
                else:
                    logger.warning(f"Nenhum campo especificado encontrado no email com assunto: {subject}")
                    
                processed += 1
                logger.debug(f"Email processado: {processed}/{total_emails}")
                
            except Exception as e:
                logger.error(f"Erro ao processar email ID {email_id}: {str(e)}", exc_info=True)
        
        logger.info(f"Processamento finalizado. {processed} emails processados, {len(self.extracted_data)} campos extraídos.")
        return processed

    def save_to_excel(self, filename=None):
        """Salva os dados extraídos em um arquivo Excel"""
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
                    # Tentar ler o arquivo existente para acrescentar os novos dados
                    existing_df = pd.read_excel(full_path)
                    # Concatenar os dados existentes com os novos
                    combined_df = pd.concat([existing_df, df], ignore_index=True)
                    logger.info(f"Acrescentando dados ao arquivo existente: {full_path}")
                    # Salvar o arquivo com todos os dados
                    combined_df.to_excel(full_path, index=False)
                except Exception as e:
                    # Se houver erro ao ler o arquivo existente, criar um novo
                    logger.warning(f"Não foi possível ler o arquivo existente: {str(e)}")
                    logger.info(f"Criando novo arquivo: {full_path}")
                    df.to_excel(full_path, index=False)
            else:
                # O arquivo não existe, criar um novo
                logger.info(f"Criando novo arquivo: {full_path}")
                df.to_excel(full_path, index=False)
            
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