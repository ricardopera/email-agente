#!/usr/bin/env python3
"""
Demo script for Email Agent Web Application
Tests all major functionality without requiring real email credentials
"""

import json
import os
import sys

def test_web_version():
    """Test the web version components"""
    print("🧪 Testando componentes da versão web...")
    
    # Test 1: Import dependencies
    try:
        import flask
        import pandas
        import openpyxl
        from src.email_processor import EmailProcessor
        print("✅ Todas as dependências importadas com sucesso")
    except ImportError as e:
        print(f"❌ Erro na importação: {e}")
        return False
    
    # Test 2: Create EmailProcessor instance
    try:
        processor = EmailProcessor()
        print("✅ EmailProcessor criado com sucesso")
    except Exception as e:
        print(f"❌ Erro ao criar EmailProcessor: {e}")
        return False
    
    # Test 3: Test configuration save/load
    try:
        test_config = {
            'email_user': 'test@example.com',
            'imap_host': 'mail.example.com',
            'search_subject': 'Test Subject',
            'custom_fields': [
                {'name': 'Campo Teste', 'pattern': 'Campo Teste:\s*([^\r\n]+)', 'format': 'texto'}
            ]
        }
        
        # Save config
        processor.save_config(
            test_config['email_user'],
            test_config['imap_host'], 
            test_config['search_subject'],
            test_config['custom_fields']
        )
        
        # Load config
        loaded_config = processor.load_config()
        
        if loaded_config.get('email_user') == test_config['email_user']:
            print("✅ Configurações salvas e carregadas com sucesso")
        else:
            print("⚠️  Problema ao salvar/carregar configurações")
            
    except Exception as e:
        print(f"❌ Erro ao testar configurações: {e}")
        
    # Test 4: Test data processing functions
    try:
        # Test text processing
        test_text = "Campo Teste: Valor de teste\nOutro campo: 1.234,56\nData: 25/12/2023"
        
        processor.custom_fields = [
            {'name': 'Campo Teste', 'pattern': r'Campo Teste:\s*([^\r\n]+)', 'format': 'texto'},
            {'name': 'Outro campo', 'pattern': r'Outro campo:\s*([^\r\n]+)', 'format': 'número'},
            {'name': 'Data', 'pattern': r'Data:\s*([^\r\n]+)', 'format': 'data'}
        ]
        
        extracted = processor.extract_fields(test_text)
        
        if extracted and len(extracted) > 0:
            print("✅ Extração de campos funcionando")
            print(f"   Extraído: {extracted}")
        else:
            print("⚠️  Problema na extração de campos")
            
    except Exception as e:
        print(f"❌ Erro ao testar extração: {e}")
    
    # Test 5: Test value processing
    try:
        # Test number processing
        num_result = processor.process_value("1.234,56", "número")
        if abs(num_result - 1234.56) < 0.01:
            print("✅ Processamento de números funcionando")
        else:
            print(f"⚠️  Problema no processamento de números: {num_result}")
            
        # Test date processing
        date_result = processor.process_value("25/12/2023", "data")
        if date_result == "2023-12-25":
            print("✅ Processamento de datas funcionando")
        else:
            print(f"⚠️  Problema no processamento de datas: {date_result}")
            
    except Exception as e:
        print(f"❌ Erro ao testar processamento de valores: {e}")
    
    # Test 6: Web app components
    try:
        sys.path.append('.')
        from web_app import app
        
        # Test if Flask app was created
        if app and hasattr(app, 'route'):
            print("✅ Aplicação Flask criada com sucesso")
        else:
            print("⚠️  Problema na criação da aplicação Flask")
            
    except Exception as e:
        print(f"❌ Erro ao testar aplicação web: {e}")
    
    print("\n🎉 Testes concluídos!")
    return True

if __name__ == "__main__":
    print("=== Demo da Versão Web do Email Agente ===\n")
    
    # Change to script directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)
    
    success = test_web_version()
    
    if success:
        print("\n🚀 Para executar a versão web:")
        print("   python web_app.py")
        print("   Acesse: http://localhost:5001")
        print("\n📚 Para mais informações:")
        print("   Veja: README_WEB.md")
    else:
        print("\n❌ Alguns testes falharam. Verifique as dependências:")
        print("   pip install flask pandas openpyxl")