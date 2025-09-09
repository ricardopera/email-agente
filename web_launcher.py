#!/usr/bin/env python3
"""
Launcher script for Email Agent Web Application
"""
import os
import sys
import webbrowser
import time
import threading

def launch_web_app():
    """Launch the web application"""
    print("=== Email Agent Web Application ===")
    print("Versão Web do Agente de Extração de Emails")
    print("=" * 40)
    
    try:
        # Change to the application directory
        app_dir = os.path.dirname(os.path.abspath(__file__))
        os.chdir(app_dir)
        
        # Import and run the web app
        from web_app import app
        
        # Open browser after a short delay
        def open_browser():
            time.sleep(1.5)
            webbrowser.open('http://localhost:5000')
            
        # Start browser opening in background
        threading.Thread(target=open_browser, daemon=True).start()
        
        print("Servidor web iniciado em http://localhost:5000")
        print("O navegador será aberto automaticamente...")
        print("Pressione Ctrl+C para parar o servidor")
        print("=" * 40)
        
        # Run the Flask app
        app.run(debug=False, host='localhost', port=5000)
        
    except KeyboardInterrupt:
        print("\nServidor encerrado pelo usuário.")
    except ImportError as e:
        print(f"Erro de importação: {e}")
        print("Certifique-se de que todas as dependências estão instaladas:")
        print("pip install flask pandas openpyxl")
    except Exception as e:
        print(f"Erro ao iniciar a aplicação: {e}")
        
if __name__ == "__main__":
    launch_web_app()