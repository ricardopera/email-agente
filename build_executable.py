import os
import subprocess
import sys
import platform

def criar_executavel():
    print("Iniciando criação do executável do EmailExtrator...")
    
    # Verificar se PyInstaller está instalado
    try:
        import PyInstaller
        print("PyInstaller encontrado na versão:", PyInstaller.__version__)
    except ImportError:
        print("PyInstaller não encontrado. Instalando...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
            print("PyInstaller instalado com sucesso.")
        except Exception as e:
            print(f"Erro ao instalar PyInstaller: {str(e)}")
            return False
    
    # Verificar outras dependências
    dependencies = ["pandas", "pywin32", "openpyxl"]
    for dep in dependencies:
        try:
            __import__(dep)
            print(f"Dependência '{dep}' encontrada.")
        except ImportError:
            print(f"Dependência '{dep}' não encontrada. Instalando...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", dep])
                print(f"Dependência '{dep}' instalada com sucesso.")
            except Exception as e:
                print(f"Erro ao instalar {dep}: {str(e)}")
                return False
    
    # Construir o executável
    try:
        print("Construindo o executável...")
        # Usa o arquivo .spec existente
        result = subprocess.check_call(["pyinstaller", "EmailExtrator.spec", "--clean"])
        
        if result == 0:
            # Obter o caminho para o executável
            dist_dir = os.path.join(os.getcwd(), "dist")
            if platform.system() == "Windows":
                exe_path = os.path.join(dist_dir, "EmailExtrator.exe")
            else:
                exe_path = os.path.join(dist_dir, "EmailExtrator")
            
            if os.path.exists(exe_path):
                print("\n============================================")
                print(f"Executável criado com sucesso em: {exe_path}")
                print("============================================\n")
                return True
            else:
                print(f"Não foi possível encontrar o executável gerado em {exe_path}")
                return False
        else:
            print("Falha ao executar o PyInstaller.")
            return False
    
    except Exception as e:
        print(f"Erro durante a criação do executável: {str(e)}")
        return False

if __name__ == "__main__":
    sucesso = criar_executavel()
    if sucesso:
        input("Pressione Enter para sair...")
    else:
        input("Ocorreram erros durante a criação do executável. Pressione Enter para sair...")