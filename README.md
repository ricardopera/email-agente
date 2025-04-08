# Email Agente

Aplicativo para extração automática de informações de emails, como números de processos judiciais.

## Funcionalidades

- Conecta-se a servidores de email via IMAP
- Busca emails não lidos com assunto específico
- Extrai números de processos judiciais do conteúdo dos emails
- Salva os dados extraídos em um arquivo Excel

## Requisitos

- Windows
- Python 3.6 ou superior (para desenvolvimento)
- Ou simplesmente use o executável compilado

## Instalação

### Usar o executável compilado

1. Baixe o arquivo executável da seção Releases
2. Execute o arquivo `EmailExtrator.exe`

### Para desenvolvedores

1. Clone este repositório
```
git clone https://github.com/ricardopera/email-agente.git
cd email-agente
```

2. Instale as dependências necessárias
```
pip install openpyxl pandas pyinstaller
```

3. Execute o aplicativo
```
python main.py
```

## Compilação do Executável

Para gerar um arquivo executável (.exe), use:

```
pyinstaller EmailExtrator.spec
```

O executável será gerado na pasta `dist`.

## Uso

1. Preencha os campos de:
   - Email e senha
   - Servidor IMAP (padrão: mail.itajai.sc.gov.br)
   - Porta (padrão: 993)
   - Timeout (aumentado para servidores mais lentos)
   - Texto a ser pesquisado no assunto dos emails

2. Use o botão "Teste de Conexão" para verificar se as configurações estão corretas
3. Clique em "Processar Emails" para buscar emails e extrair informações
4. Um arquivo Excel com os dados será gerado ao final do processamento

## Logging

O aplicativo gera logs detalhados no arquivo `email_extrator.log` que podem ser úteis para diagnosticar problemas de conexão.

## Licença

Este projeto está licenciado sob a licença MIT.