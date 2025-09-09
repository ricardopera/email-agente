// Email Agent Web Application JavaScript

let customFieldCounter = 0;
let additionalFieldCounter = 0;
let processingInterval = null;

// Initialize the application
document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
});

function initializeApp() {
    // Add default custom fields
    addCustomField('Número processo CNJ', 'texto');
    addCustomField('Valor liquido transferido para parte', 'número');
    
    // Add default additional field
    addAdditionalField();
    
    // Load saved configuration
    loadConfig();
    
    // Set up file upload handler
    document.getElementById('excel_file').addEventListener('change', handleFileUpload);
    
    logMessage('Aplicação inicializada', 'info');
}

function addCustomField(defaultName = '', defaultFormat = 'texto') {
    const container = document.getElementById('custom_fields_container');
    const fieldId = ++customFieldCounter;
    
    const fieldHtml = `
        <div class="custom-field-row" id="custom_field_${fieldId}">
            <div class="row align-items-center">
                <div class="col-md-6">
                    <label class="form-label fw-bold">Nome do Campo:</label>
                    <input type="text" class="form-control custom-field-name" 
                           placeholder="Ex: Número do processo" value="${defaultName}">
                </div>
                <div class="col-md-4">
                    <label class="form-label fw-bold">Formato:</label>
                    <select class="form-select custom-field-format">
                        <option value="texto" ${defaultFormat === 'texto' ? 'selected' : ''}>Texto</option>
                        <option value="número" ${defaultFormat === 'número' ? 'selected' : ''}>Número</option>
                        <option value="data" ${defaultFormat === 'data' ? 'selected' : ''}>Data</option>
                    </select>
                </div>
                <div class="col-md-2">
                    <label class="form-label">&nbsp;</label>
                    <div>
                        <button type="button" class="btn btn-outline-danger field-remove-btn" 
                                onclick="removeCustomField(${fieldId})" title="Remover campo">
                            <i class="fas fa-trash"></i>
                        </button>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    container.insertAdjacentHTML('beforeend', fieldHtml);
    logMessage(`Campo personalizado adicionado: ${defaultName || 'Novo campo'}`, 'info');
}

function removeCustomField(fieldId) {
    const customFields = document.querySelectorAll('.custom-field-row');
    if (customFields.length <= 1) {
        showMessage('Aviso', 'É necessário manter pelo menos um campo para extração.', 'warning');
        return;
    }
    
    const field = document.getElementById(`custom_field_${fieldId}`);
    if (field) {
        field.remove();
        logMessage('Campo personalizado removido', 'info');
    }
}

function addAdditionalField(defaultValue = '') {
    const container = document.getElementById('additional_fields_container');
    const fieldId = ++additionalFieldCounter;
    
    const fieldHtml = `
        <div class="additional-field-row mb-2" id="additional_field_${fieldId}">
            <div class="input-group">
                <input type="text" class="form-control additional-field-name" 
                       placeholder="Nome do campo adicional" value="${defaultValue}">
                <button type="button" class="btn btn-outline-danger" 
                        onclick="removeAdditionalField(${fieldId})" title="Remover campo">
                    <i class="fas fa-trash"></i>
                </button>
            </div>
        </div>
    `;
    
    container.insertAdjacentHTML('beforeend', fieldHtml);
    logMessage(`Campo adicional adicionado: ${defaultValue || 'Novo campo adicional'}`, 'info');
}

function removeAdditionalField(fieldId) {
    const additionalFields = document.querySelectorAll('.additional-field-row');
    if (additionalFields.length <= 1) {
        showMessage('Aviso', 'É necessário manter pelo menos um campo adicional.', 'warning');
        return;
    }
    
    const field = document.getElementById(`additional_field_${fieldId}`);
    if (field) {
        field.remove();
        logMessage('Campo adicional removido', 'info');
    }
}

function getCustomFields() {
    const fields = [];
    document.querySelectorAll('.custom-field-row').forEach(row => {
        const name = row.querySelector('.custom-field-name').value.trim();
        const format = row.querySelector('.custom-field-format').value;
        
        if (name) {
            fields.push({
                name: name,
                pattern: `${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}:\\s*([^\\r\\n]+)`,
                format: format
            });
        }
    });
    return fields;
}

function getAdditionalFields() {
    const fields = [];
    document.querySelectorAll('.additional-field-name').forEach(input => {
        const value = input.value.trim();
        if (value) {
            fields.push(value);
        }
    });
    return fields;
}

async function testConnection() {
    const button = event.target;
    const originalText = button.innerHTML;
    button.innerHTML = '<span class="loading-spinner"></span>Testando...';
    button.disabled = true;
    
    try {
        const data = {
            email_user: document.getElementById('email_user').value,
            email_pass: document.getElementById('email_pass').value,
            imap_host: document.getElementById('imap_host').value,
            imap_port: document.getElementById('imap_port').value,
            use_ssl: document.getElementById('use_ssl').checked,
            timeout: document.getElementById('timeout').value
        };
        
        if (!data.email_user || !data.email_pass || !data.imap_host) {
            showMessage('Erro', 'Email, senha e servidor são obrigatórios!', 'error');
            return;
        }
        
        logMessage('Testando conexão com servidor IMAP...', 'info');
        
        const response = await fetch('/api/test_connection', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(data)
        });
        
        const result = await response.json();
        
        if (result.success) {
            showMessage('Sucesso', result.message, 'success');
            logMessage('Teste de conexão bem-sucedido!', 'success');
        } else {
            showMessage('Erro', result.message, 'error');
            logMessage('Falha no teste de conexão', 'error');
        }
    } catch (error) {
        showMessage('Erro', `Erro ao testar conexão: ${error.message}`, 'error');
        logMessage(`Erro ao testar conexão: ${error.message}`, 'error');
    } finally {
        button.innerHTML = originalText;
        button.disabled = false;
    }
}

async function loadConfig() {
    try {
        logMessage('Carregando configurações salvas...', 'info');
        
        const response = await fetch('/api/load_config');
        const result = await response.json();
        
        if (result.success && result.config) {
            const config = result.config;
            
            // Load connection settings
            if (config.email_user) document.getElementById('email_user').value = config.email_user;
            if (config.imap_host) document.getElementById('imap_host').value = config.imap_host;
            if (config.search_subject) document.getElementById('search_subject').value = config.search_subject;
            
            // Load custom fields if available
            if (config.custom_fields && config.custom_fields.length > 0) {
                // Clear existing fields first
                document.getElementById('custom_fields_container').innerHTML = '';
                customFieldCounter = 0;
                
                // Add saved custom fields
                config.custom_fields.forEach(field => {
                    addCustomField(field.name, field.format || 'texto');
                });
            }
            
            logMessage('Configurações carregadas com sucesso!', 'success');
        } else {
            logMessage('Nenhuma configuração salva encontrada', 'warning');
        }
    } catch (error) {
        logMessage(`Erro ao carregar configurações: ${error.message}`, 'error');
    }
}

async function saveConfig() {
    try {
        const data = {
            email_user: document.getElementById('email_user').value,
            imap_host: document.getElementById('imap_host').value,
            search_subject: document.getElementById('search_subject').value,
            custom_fields: getCustomFields()
        };
        
        logMessage('Salvando configurações...', 'info');
        
        const response = await fetch('/api/save_config', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(data)
        });
        
        const result = await response.json();
        
        if (result.success) {
            showMessage('Sucesso', result.message, 'success');
            logMessage('Configurações salvas com sucesso!', 'success');
        } else {
            showMessage('Erro', result.message, 'error');
            logMessage('Erro ao salvar configurações', 'error');
        }
    } catch (error) {
        showMessage('Erro', `Erro ao salvar configurações: ${error.message}`, 'error');
        logMessage(`Erro ao salvar configurações: ${error.message}`, 'error');
    }
}

async function processEmails() {
    // Validation
    const email_user = document.getElementById('email_user').value.trim();
    const email_pass = document.getElementById('email_pass').value.trim();
    const imap_host = document.getElementById('imap_host').value.trim();
    const search_subject = document.getElementById('search_subject').value.trim();
    
    if (!email_user || !email_pass || !imap_host || !search_subject) {
        showMessage('Erro', 'Todos os campos de conexão são obrigatórios!', 'error');
        return;
    }
    
    const customFields = getCustomFields();
    if (customFields.length === 0) {
        showMessage('Erro', 'É necessário definir pelo menos um campo para extração!', 'error');
        return;
    }
    
    try {
        const data = {
            email_user: email_user,
            email_pass: email_pass,
            imap_host: imap_host,
            imap_port: document.getElementById('imap_port').value,
            use_ssl: document.getElementById('use_ssl').checked,
            timeout: document.getElementById('timeout').value,
            search_subject: search_subject,
            custom_fields: customFields
        };
        
        logMessage('Iniciando processamento de emails...', 'info');
        
        const response = await fetch('/api/process_emails', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(data)
        });
        
        const result = await response.json();
        
        if (result.success) {
            logMessage('Processamento iniciado com sucesso!', 'success');
            startProcessingStatusCheck();
        } else {
            showMessage('Erro', result.message, 'error');
            logMessage(`Erro ao iniciar processamento: ${result.message}`, 'error');
        }
    } catch (error) {
        showMessage('Erro', `Erro ao processar emails: ${error.message}`, 'error');
        logMessage(`Erro ao processar emails: ${error.message}`, 'error');
    }
}

function startProcessingStatusCheck() {
    document.getElementById('progress_bar').style.display = 'block';
    
    processingInterval = setInterval(async () => {
        try {
            const response = await fetch('/api/processing_status');
            const status = await response.json();
            
            updateProcessingStatus(status);
            
            if (!status.is_processing) {
                clearInterval(processingInterval);
                processingInterval = null;
                document.getElementById('progress_bar').style.display = 'none';
                
                if (status.extracted_data && status.extracted_data.length > 0) {
                    document.getElementById('download_btn').disabled = false;
                    logMessage(`Processamento concluído! ${status.extracted_data.length} registros extraídos.`, 'success');
                } else {
                    logMessage('Processamento concluído, mas nenhum dado foi extraído.', 'warning');
                }
            }
        } catch (error) {
            console.error('Error checking processing status:', error);
        }
    }, 1000);
}

function updateProcessingStatus(status) {
    const statusContainer = document.getElementById('processing_status');
    const progressBar = document.querySelector('.progress-bar');
    
    let statusClass = 'status-ready';
    let statusIcon = 'fas fa-circle';
    
    if (status.is_processing) {
        statusClass = 'status-processing processing-animation';
        statusIcon = 'fas fa-cog fa-spin';
    } else if (status.progress === 100) {
        statusClass = 'status-success';
        statusIcon = 'fas fa-check-circle';
    } else if (status.message && status.message.includes('Erro')) {
        statusClass = 'status-error';
        statusIcon = 'fas fa-exclamation-circle';
    }
    
    statusContainer.innerHTML = `
        <div class="${statusClass}">
            <i class="${statusIcon} me-2"></i>
            <strong>${status.message || 'Pronto para processar emails'}</strong>
        </div>
        ${status.total_emails > 0 ? `
            <div class="mt-2">
                <small>Emails encontrados: ${status.total_emails}</small><br>
                <small>Emails processados: ${status.processed_emails}</small>
            </div>
        ` : ''}
    `;
    
    if (status.progress > 0) {
        progressBar.style.width = `${status.progress}%`;
        progressBar.textContent = `${status.progress}%`;
    }
    
    // Log status changes
    if (status.message && status.message !== statusContainer.dataset.lastMessage) {
        logMessage(status.message, status.is_processing ? 'info' : 'success');
        statusContainer.dataset.lastMessage = status.message;
    }
}

async function downloadExcel() {
    try {
        logMessage('Baixando arquivo Excel...', 'info');
        
        const response = await fetch('/api/download_excel');
        
        if (response.ok) {
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = url;
            a.download = `processos_extraidos_${new Date().toISOString().slice(0,19).replace(/:/g, '-')}.xlsx`;
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
            
            logMessage('Arquivo Excel baixado com sucesso!', 'success');
        } else {
            const result = await response.json();
            showMessage('Erro', result.message, 'error');
        }
    } catch (error) {
        showMessage('Erro', `Erro ao baixar arquivo: ${error.message}`, 'error');
        logMessage(`Erro ao baixar arquivo: ${error.message}`, 'error');
    }
}

async function handleFileUpload() {
    const fileInput = document.getElementById('excel_file');
    const file = fileInput.files[0];
    
    if (!file) return;
    
    const formData = new FormData();
    formData.append('file', file);
    
    try {
        logMessage(`Carregando arquivo Excel: ${file.name}...`, 'info');
        
        const response = await fetch('/api/upload_excel', {
            method: 'POST',
            body: formData
        });
        
        const result = await response.json();
        
        if (result.success) {
            logMessage('Arquivo Excel carregado com sucesso!', 'success');
            if (result.columns) {
                showMessage('Sucesso', 
                    `Arquivo carregado! Colunas disponíveis: ${result.columns.join(', ')}`, 
                    'success');
            }
        } else {
            showMessage('Erro', result.message, 'error');
            logMessage(`Erro ao carregar arquivo: ${result.message}`, 'error');
        }
    } catch (error) {
        showMessage('Erro', `Erro ao carregar arquivo: ${error.message}`, 'error');
        logMessage(`Erro ao carregar arquivo: ${error.message}`, 'error');
    }
}

function logMessage(message, type = 'info') {
    const logContainer = document.getElementById('log_container');
    const timestamp = new Date().toLocaleTimeString();
    
    const logEntry = document.createElement('div');
    logEntry.className = `log-entry log-${type}`;
    logEntry.innerHTML = `<strong>[${timestamp}]</strong> ${message}`;
    
    logContainer.appendChild(logEntry);
    logContainer.scrollTop = logContainer.scrollHeight;
    
    // Keep only last 100 log entries
    const entries = logContainer.children;
    if (entries.length > 100) {
        logContainer.removeChild(entries[0]);
    }
}

function clearLog() {
    document.getElementById('log_container').innerHTML = 
        '<div class="text-muted">Log limpo - aguardando operações...</div>';
    logMessage('Log de operações limpo', 'info');
}

function showMessage(title, message, type = 'info') {
    // Use browser native alert since Bootstrap may not be available
    alert(`${title}: ${message}`);
}