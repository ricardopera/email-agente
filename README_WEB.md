# Email Agente - Versão Web

Uma versão web do aplicativo de extração automática de informações de emails, permitindo processamento local através do navegador.

![Interface Web do Email Agente](https://github.com/user-attachments/assets/577e5792-85b8-42b0-b753-f19729548dfc)

## 🌟 Principais Características

- **Interface Web Moderna**: Interface responsiva e intuitiva usando Bootstrap 5
- **Processamento Local**: Todo o processamento acontece localmente na máquina do usuário
- **Compatibilidade Total**: Mantém todas as funcionalidades da versão desktop
- **Suporte a Português**: Completo suporte a caracteres acentuados e especiais
- **Configuração Persistente**: Salva e carrega configurações automaticamente
- **Log em Tempo Real**: Acompanhe o progresso do processamento em tempo real

## 🚀 Como Usar a Versão Web

### Pré-requisitos

- Python 3.6 ou superior
- Conexão com a internet (apenas para baixar dependências na primeira execução)

### Instalação e Execução

1. **Clone o repositório**:
```bash
git clone https://github.com/ricardopera/email-agente.git
cd email-agente
```

2. **Instale as dependências**:
```bash
pip install flask pandas openpyxl
```

3. **Inicie a aplicação web**:
```bash
python web_app.py
```
ou use o launcher:
```bash
python web_launcher.py
```

4. **Acesse no navegador**:
   - Abra seu navegador em: `http://localhost:5001`
   - A interface web será carregada automaticamente

## 📋 Funcionalidades da Versão Web

### 1. Configuração de Conexão
- Endereço de email e senha
- Servidor IMAP personalizado
- Configurações SSL/TLS
- Timeout configurável
- Teste de conexão integrado

### 2. Campos para Extração
- **Adição dinâmica**: Adicione quantos campos precisar
- **Tipos de formato**:
  - 📝 **Texto**: Para dados textuais
  - 🔢 **Número**: Converte formato brasileiro (1.234,56) para decimal
  - 📅 **Data**: Converte dd/mm/aaaa para formato internacional
- **Remoção fácil**: Remova campos desnecessários com um clique

### 3. Arquivo Excel Adicional (Opcional)
- Upload de arquivo Excel de referência
- Campo chave para relacionamento de dados
- Campos adicionais configuráveis
- Visualização das colunas disponíveis

### 4. Processamento em Tempo Real
- **Status visual**: Acompanhe o progresso do processamento
- **Log detalhado**: Veja todas as operações em tempo real
- **Barra de progresso**: Indicador visual do andamento
- **Controle de estado**: Botões habilitados/desabilitados conforme o contexto

### 5. Download de Resultados
- Download automático do arquivo Excel gerado
- Dados formatados corretamente (números, datas, texto)
- Compatibilidade com Excel e LibreOffice

## 🎯 Vantagens da Versão Web

### ✅ Facilidade de Uso
- **Não precisa instalar**: Funciona diretamente no navegador
- **Interface moderna**: Design limpo e intuitivo
- **Responsivo**: Funciona em desktop, tablet e mobile

### ✅ Funcionalidades Avançadas
- **Configuração visual**: Configure campos sem editar código
- **Feedback em tempo real**: Veja o que está acontecendo
- **Gestão de arquivos**: Upload e download integrados

### ✅ Manutenção
- **Logs integrados**: Sistema de logging completo
- **Tratamento de erros**: Mensagens de erro claras e úteis
- **Validação de dados**: Validação automática de campos obrigatórios

## 🔧 Estrutura do Projeto Web

```
email-agente/
├── web_app.py              # Aplicação Flask principal
├── web_launcher.py         # Script de inicialização
├── templates/
│   └── index.html          # Interface web principal
├── static/
│   ├── css/
│   │   └── style.css       # Estilos personalizados
│   └── js/
│       └── app.js          # JavaScript da aplicação
└── src/
    ├── app.py              # Versão desktop (original)
    └── email_processor.py  # Lógica de processamento (compartilhada)
```

## 🎨 Interface da Aplicação

### Seções Principais

1. **📧 Configurações de Conexão**
   - Campos para email, senha, servidor
   - Opções SSL/TLS e timeout
   - Botão de teste de conexão

2. **🔍 Campos para Extração**
   - Lista dinâmica de campos personalizados
   - Seleção de formato para cada campo
   - Botões para adicionar/remover campos

3. **📊 Arquivo Excel Adicional**
   - Upload de arquivo de referência
   - Configuração de campo chave
   - Campos adicionais para extração

4. **⚡ Ações Principais**
   - Botão "Processar Emails"
   - Botão "Baixar Excel" (habilitado após processamento)

5. **📊 Status e Log**
   - Status do processamento em tempo real
   - Log detalhado de todas as operações
   - Barra de progresso visual

## 🔒 Segurança e Privacidade

- **Processamento Local**: Todas as credenciais e dados ficam na sua máquina
- **Sem envio externo**: Nenhum dado é enviado para servidores externos
- **Sessões seguras**: Usando Flask com chaves de sessão aleatórias
- **Logs locais**: Todos os logs são armazenados localmente

## 🆚 Comparação: Desktop vs Web

| Funcionalidade | Desktop (Tkinter) | Web (Flask) |
|---|---|---|
| Interface | ✅ Nativa | ✅ Moderna/Responsiva |
| Facilidade de Uso | ⭐⭐⭐ | ⭐⭐⭐⭐⭐ |
| Configuração | ✅ Formulários | ✅ Interface intuitiva |
| Processamento | ✅ Local | ✅ Local |
| Logs | ✅ Arquivo | ✅ Tempo real + arquivo |
| Portabilidade | Windows | Multiplataforma |
| Dependências | Tkinter | Navegador web |
| Extensibilidade | ⭐⭐⭐ | ⭐⭐⭐⭐⭐ |

## 🛠️ APIs da Aplicação Web

A versão web expõe as seguintes APIs REST:

- `GET /` - Interface principal
- `GET /api/load_config` - Carregar configurações
- `POST /api/save_config` - Salvar configurações
- `POST /api/test_connection` - Testar conexão IMAP
- `POST /api/process_emails` - Iniciar processamento
- `GET /api/processing_status` - Status do processamento
- `GET /api/download_excel` - Download do arquivo Excel
- `POST /api/upload_excel` - Upload de arquivo Excel adicional

## 📝 Exemplo de Uso

1. **Configure a Conexão**:
   - Email: `seu.email@empresa.com`
   - Servidor: `mail.empresa.com`
   - Porta: `993` (SSL) ou `143` (TLS)

2. **Defina os Campos**:
   - Nome: `Número do Processo` | Formato: `Texto`
   - Nome: `Valor da Transferência` | Formato: `Número`
   - Nome: `Data do Pagamento` | Formato: `Data`

3. **Teste a Conexão**: Clique em "Testar Conexão"

4. **Processe os Emails**: Clique em "Processar Emails"

5. **Baixe os Resultados**: Clique em "Baixar Excel" quando concluído

## 🐛 Solução de Problemas

### Erro de Porta em Uso
```bash
# Se a porta 5001 estiver em uso, edite web_app.py:
app.run(debug=True, host='0.0.0.0', port=5002)  # Use outra porta
```

### Erro de Dependências
```bash
pip install --upgrade flask pandas openpyxl
```

### Problemas de Conexão IMAP
- Verifique as configurações do servidor
- Teste com diferentes portas (993, 143, 587)
- Verifique se o email permite apps de terceiros

## 🚀 Melhorias Futuras

- [ ] Suporte a múltiplos servidores de email
- [ ] Templates de campos pré-configurados
- [ ] Export para outros formatos (CSV, JSON)
- [ ] Agendamento automático de processamento
- [ ] Dashboard com estatísticas
- [ ] Suporte a filtros avançados de email
- [ ] Integration com APIs de email (Gmail, Outlook)

---

## 💡 Dica de Produtividade

Para usar diariamente, crie um script de inicialização:

**Windows (iniciar.bat)**:
```batch
@echo off
cd /d "%~dp0"
python web_launcher.py
```

**Linux/Mac (iniciar.sh)**:
```bash
#!/bin/bash
cd "$(dirname "$0")"
python3 web_launcher.py
```

Dessa forma, basta clicar duas vezes no arquivo para iniciar a aplicação! 🎉