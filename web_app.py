#!/usr/bin/env python3
"""
Web version of the Email Agent application.
Provides a web interface for email processing with IMAP server connection.
All processing happens locally through this web server.
"""

from flask import Flask, render_template, request, jsonify, send_file, session
import os
import json
import logging
import threading
import tempfile
from datetime import datetime
from src.email_processor import EmailProcessor
import pandas as pd

# Configurar logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.secret_key = os.urandom(24)  # For session management

# Global email processor instance
email_processor = EmailProcessor()

# Global variables to track processing state
processing_status = {
    'is_processing': False,
    'progress': 0,
    'message': '',
    'total_emails': 0,
    'processed_emails': 0,
    'extracted_data': []
}

@app.route('/')
def index():
    """Main page with the email processing interface"""
    return render_template('index.html')

@app.route('/api/load_config', methods=['GET'])
def load_config():
    """Load saved configuration"""
    try:
        config = email_processor.load_config()
        return jsonify({
            'success': True,
            'config': config
        })
    except Exception as e:
        logger.error(f"Error loading config: {str(e)}")
        return jsonify({
            'success': False,
            'message': f'Erro ao carregar configurações: {str(e)}'
        })

@app.route('/api/save_config', methods=['POST'])
def save_config():
    """Save configuration"""
    try:
        data = request.get_json()
        success = email_processor.save_config(
            data.get('email_user', ''),
            data.get('imap_host', ''),
            data.get('search_subject', ''),
            data.get('custom_fields', [])
        )
        
        return jsonify({
            'success': success,
            'message': 'Configurações salvas com sucesso' if success else 'Erro ao salvar configurações'
        })
    except Exception as e:
        logger.error(f"Error saving config: {str(e)}")
        return jsonify({
            'success': False,
            'message': f'Erro ao salvar configurações: {str(e)}'
        })

@app.route('/api/test_connection', methods=['POST'])
def test_connection():
    """Test IMAP connection"""
    try:
        data = request.get_json()
        email_user = data.get('email_user', '')
        email_pass = data.get('email_pass', '')
        imap_host = data.get('imap_host', '')
        imap_port = int(data.get('imap_port', 993))
        use_ssl = data.get('use_ssl', True)
        timeout = int(data.get('timeout', 60))
        
        # Test connection in a thread to avoid blocking
        success = email_processor.connect_to_server(
            email_user, email_pass, imap_host, imap_port, use_ssl, timeout
        )
        
        if success:
            email_processor.close_connection()
            return jsonify({
                'success': True,
                'message': 'Conexão estabelecida com sucesso!'
            })
        else:
            return jsonify({
                'success': False,
                'message': 'Falha na conexão. Verifique os logs para mais detalhes.'
            })
    except Exception as e:
        logger.error(f"Error testing connection: {str(e)}")
        return jsonify({
            'success': False,
            'message': f'Erro ao testar conexão: {str(e)}'
        })

@app.route('/api/process_emails', methods=['POST'])
def process_emails():
    """Start email processing"""
    global processing_status
    
    if processing_status['is_processing']:
        return jsonify({
            'success': False,
            'message': 'Processamento já em andamento'
        })
    
    try:
        data = request.get_json()
        
        # Reset processing status
        processing_status = {
            'is_processing': True,
            'progress': 0,
            'message': 'Iniciando processamento...',
            'total_emails': 0,
            'processed_emails': 0,
            'extracted_data': []
        }
        
        # Start processing in a separate thread
        thread = threading.Thread(
            target=_process_emails_thread,
            args=(data,),
            daemon=True
        )
        thread.start()
        
        return jsonify({
            'success': True,
            'message': 'Processamento iniciado'
        })
    except Exception as e:
        processing_status['is_processing'] = False
        logger.error(f"Error starting email processing: {str(e)}")
        return jsonify({
            'success': False,
            'message': f'Erro ao iniciar processamento: {str(e)}'
        })

def _process_emails_thread(data):
    """Process emails in a separate thread"""
    global processing_status
    
    try:
        # Extract parameters
        email_user = data.get('email_user', '')
        email_pass = data.get('email_pass', '')
        imap_host = data.get('imap_host', '')
        imap_port = int(data.get('imap_port', 993))
        use_ssl = data.get('use_ssl', True)
        timeout = int(data.get('timeout', 60))
        search_subject = data.get('search_subject', '')
        custom_fields = data.get('custom_fields', [])
        
        # Update email processor with custom fields
        email_processor.custom_fields = custom_fields
        
        # Connect to server
        processing_status['message'] = f'Conectando ao servidor {imap_host}:{imap_port}...'
        if not email_processor.connect_to_server(email_user, email_pass, imap_host, imap_port, use_ssl, timeout):
            processing_status['is_processing'] = False
            processing_status['message'] = 'Falha ao conectar ao servidor'
            return
        
        # Search emails
        processing_status['message'] = f'Buscando emails com assunto: {search_subject}...'
        email_ids = email_processor.search_emails(search_subject)
        
        if not email_ids:
            processing_status['is_processing'] = False
            processing_status['message'] = 'Nenhum email encontrado'
            email_processor.close_connection()
            return
        
        processing_status['total_emails'] = len(email_ids)
        processing_status['message'] = f'Encontrados {len(email_ids)} emails. Processando...'
        
        # Process emails
        processed = email_processor.process_emails(email_ids)
        processing_status['processed_emails'] = processed
        
        # Save to Excel
        processing_status['message'] = 'Salvando dados em arquivo Excel...'
        if email_processor.save_to_excel():
            processing_status['message'] = f'Processamento concluído! {processed} emails processados.'
            processing_status['extracted_data'] = email_processor.extracted_data
        else:
            processing_status['message'] = 'Processamento concluído mas houve erro ao salvar o arquivo Excel.'
        
        email_processor.close_connection()
        processing_status['progress'] = 100
        
    except Exception as e:
        logger.error(f"Error in processing thread: {str(e)}")
        processing_status['message'] = f'Erro durante o processamento: {str(e)}'
        email_processor.close_connection()
    finally:
        processing_status['is_processing'] = False

@app.route('/api/processing_status', methods=['GET'])
def get_processing_status():
    """Get current processing status"""
    return jsonify(processing_status)

@app.route('/api/download_excel', methods=['GET'])
def download_excel():
    """Download the generated Excel file"""
    try:
        # Find the most recent Excel file
        excel_files = [f for f in os.listdir('.') if f.startswith('processos_extraidos_') and f.endswith('.xlsx')]
        if excel_files:
            latest_file = sorted(excel_files)[-1]
            return send_file(
                latest_file,
                as_attachment=True,
                download_name=latest_file,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            return jsonify({
                'success': False,
                'message': 'Nenhum arquivo Excel encontrado'
            }), 404
    except Exception as e:
        logger.error(f"Error downloading Excel: {str(e)}")
        return jsonify({
            'success': False,
            'message': f'Erro ao baixar arquivo: {str(e)}'
        }), 500

@app.route('/api/upload_excel', methods=['POST'])
def upload_excel():
    """Upload additional Excel file for reference data"""
    try:
        if 'file' not in request.files:
            return jsonify({
                'success': False,
                'message': 'Nenhum arquivo selecionado'
            })
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({
                'success': False,
                'message': 'Nenhum arquivo selecionado'
            })
        
        if file and file.filename.lower().endswith(('.xlsx', '.xls')):
            # Save uploaded file
            filename = f"uploaded_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{file.filename}"
            filepath = os.path.join(tempfile.gettempdir(), filename)
            file.save(filepath)
            
            # Read columns from Excel file
            df = pd.read_excel(filepath)
            columns = df.columns.tolist()
            
            # Update email processor with the uploaded file
            email_processor.additional_excel_file = filepath
            
            return jsonify({
                'success': True,
                'message': 'Arquivo carregado com sucesso',
                'columns': columns,
                'filepath': filepath
            })
        else:
            return jsonify({
                'success': False,
                'message': 'Apenas arquivos Excel (.xlsx, .xls) são permitidos'
            })
    except Exception as e:
        logger.error(f"Error uploading Excel: {str(e)}")
        return jsonify({
            'success': False,
            'message': f'Erro ao carregar arquivo: {str(e)}'
        })

if __name__ == '__main__':
    # Create templates directory
    template_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'templates')
    static_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'static')
    
    if not os.path.exists(template_dir):
        os.makedirs(template_dir)
    if not os.path.exists(static_dir):
        os.makedirs(static_dir)
    
    print("=== Email Agent Web Application ===")
    print("Iniciando servidor web local...")
    print("Acesse: http://localhost:5001")
    print("Pressione Ctrl+C para parar o servidor")
    print("=" * 40)
    
    app.run(debug=True, host='0.0.0.0', port=5001)