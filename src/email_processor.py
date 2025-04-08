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
        self.process_number_pattern = r'(?:Número do Processo CNJ|Processo)[\s:]*([\d.-]+)'
        self.extracted_data = []
        self.config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'config.json')
        
    def save_config(self, email_user, imap_host, search_subject):
        """Salva as configurações para uso futuro"""
        try:
            config = {
                'email_user': email_user,
                'imap_host': imap_host,
                'search_subject': search_subject
            }
            
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
            self.save_config(email_user, imap_host, self.search_subject)
            
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

    def extract_process_number(self, text):
        """Extrai o número do processo do texto do email"""
        match = re.search(self.process_number_pattern, text)
        if match:
            return match.group(1)
        return None

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
                
                # Extrair número do processo
                process_number = self.extract_process_number(content)
                
                if process_number:
                    logger.info(f"Número de processo encontrado: {process_number}")
                    self.extracted_data.append({
                        'Assunto': subject,
                        'Remetente': sender,
                        'Data': date,
                        'Número do Processo': process_number
                    })
                else:
                    logger.warning(f"Nenhum número de processo encontrado no email com assunto: {subject}")
                    
                processed += 1
                logger.debug(f"Email processado: {processed}/{total_emails}")
                
            except Exception as e:
                logger.error(f"Erro ao processar email ID {email_id}: {str(e)}", exc_info=True)
        
        logger.info(f"Processamento finalizado. {processed} emails processados, {len(self.extracted_data)} números de processo extraídos.")
        return processed

    def save_to_excel(self, filename=None):
        """Salva os dados extraídos em um arquivo Excel"""
        if not self.extracted_data:
            messagebox.showinfo("Informação", "Nenhum dado para salvar.")
            return False
            
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"processos_extraidos_{timestamp}.xlsx"
            
        try:
            df = pd.DataFrame(self.extracted_data)
            # Definir ordem das colunas
            df = df[['Número do Processo', 'Assunto', 'Remetente', 'Data']]
            
            # Salvar para Excel
            full_path = os.path.join(os.getcwd(), filename)
            df.to_excel(full_path, index=False)
            
            messagebox.showinfo("Sucesso", f"Dados salvos com sucesso em {full_path}")
            return True
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar arquivo Excel: {str(e)}")
            return False

    def close_connection(self):
        """Fecha a conexão com o servidor IMAP"""
        if self.imap_server:
            try:
                self.imap_server.close()
                self.imap_server.logout()
            except:
                pass