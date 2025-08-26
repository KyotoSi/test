// Конфигурация API
const API_BASE_URL = '/api/letters';

// Элементы DOM
const reportingFileInput = document.getElementById('reporting-file');
const sedFileInput = document.getElementById('sed-file');
const uploadBtn = document.getElementById('upload-btn');
const processBtn = document.getElementById('process-btn');
const downloadAllBtn = document.getElementById('download-all-btn');

const reportingStatus = document.getElementById('reporting-status');
const sedStatus = document.getElementById('sed-status');
const statusIndicator = document.getElementById('status-indicator');
const statusText = document.getElementById('status-text');

const processSection = document.getElementById('process-section');
const resultsSection = document.getElementById('results-section');
const resultsSummary = document.getElementById('results-summary');
const lettersList = document.getElementById('letters-list');

const loadingModal = document.getElementById('loading-modal');
const errorModal = document.getElementById('error-modal');
const loadingText = document.getElementById('loading-text');
const errorText = document.getElementById('error-text');
const closeErrorModal = document.getElementById('close-error-modal');

// Состояние приложения
let appState = {
    reportingFileSelected: false,
    sedFileSelected: false,
    filesUploaded: false,
    dataProcessed: false,
    lettersData: []
};

// Инициализация
document.addEventListener('DOMContentLoaded', function() {
    initializeEventListeners();
    updateUI();
    checkServerStatus();
});

// Инициализация обработчиков событий
function initializeEventListeners() {
    // Обработчики файлов
    reportingFileInput.addEventListener('change', handleReportingFileSelect);
    sedFileInput.addEventListener('change', handleSedFileSelect);
    
    // Обработчики кнопок
    uploadBtn.addEventListener('click', handleUpload);
    processBtn.addEventListener('click', handleProcess);
    downloadAllBtn.addEventListener('click', handleDownloadAll);
    
    // Модальные окна
    closeErrorModal.addEventListener('click', hideErrorModal);
    errorModal.addEventListener('click', function(e) {
        if (e.target === errorModal) {
            hideErrorModal();
        }
    });
}

// Обработчик выбора файла отчетности
function handleReportingFileSelect(event) {
    const file = event.target.files[0];
    if (file) {
        appState.reportingFileSelected = true;
        updateFileStatus(reportingStatus, file.name, true);
    } else {
        appState.reportingFileSelected = false;
        updateFileStatus(reportingStatus, 'Файл не выбран', false);
    }
    updateUI();
}

// Обработчик выбора файла СЭД
function handleSedFileSelect(event) {
    const file = event.target.files[0];
    if (file) {
        appState.sedFileSelected = true;
        updateFileStatus(sedStatus, file.name, true);
    } else {
        appState.sedFileSelected = false;
        updateFileStatus(sedStatus, 'Файл не выбран', false);
    }
    updateUI();
}

// Обновление статуса файла
function updateFileStatus(statusElement, text, isSuccess) {
    const statusTextElement = statusElement.querySelector('.status-text');
    statusTextElement.textContent = text;
    
    statusElement.className = 'file-status ' + (isSuccess ? 'success' : 'default');
    
    // Обновляем родительский элемент
    const fileUpload = statusElement.closest('.file-upload');
    if (isSuccess) {
        fileUpload.classList.add('active');
    } else {
        fileUpload.classList.remove('active');
    }
}

// Обработчик загрузки файлов
async function handleUpload() {
    if (!appState.reportingFileSelected || !appState.sedFileSelected) {
        showError('Пожалуйста, выберите оба файла');
        return;
    }
    
    const formData = new FormData();
    formData.append('reporting_file', reportingFileInput.files[0]);
    formData.append('sed_file', sedFileInput.files[0]);
    
    showLoading('Загрузка файлов...');
    
    try {
        const response = await fetch(`${API_BASE_URL}/upload`, {
            method: 'POST',
            body: formData
        });
        
        const data = await response.json();
        
        if (response.ok) {
            appState.filesUploaded = true;
            updateStatus('Файлы успешно загружены', 'success');
            showSuccess('Файлы успешно загружены!');
        } else {
            throw new Error(data.error || 'Ошибка при загрузке файлов');
        }
    } catch (error) {
        showError('Ошибка при загрузке файлов: ' + error.message);
        updateStatus('Ошибка загрузки файлов', 'error');
    } finally {
        hideLoading();
        updateUI();
    }
}

// Обработчик обработки данных
async function handleProcess() {
    showLoading('Обработка данных и генерация писем...');
    
    try {
        const response = await fetch(`${API_BASE_URL}/process`, {
            method: 'POST'
        });
        
        const data = await response.json();
        
        if (response.ok) {
            appState.dataProcessed = true;
            appState.lettersData = data.letters_data || [];
            
            updateStatus(`Обработано ${data.letters_count} писем`, 'success');
            displayResults(data);
            showSuccess(`Успешно сгенерировано ${data.letters_count} писем!`);
        } else {
            throw new Error(data.error || 'Ошибка при обработке данных');
        }
    } catch (error) {
        showError('Ошибка при обработке данных: ' + error.message);
        updateStatus('Ошибка обработки данных', 'error');
    } finally {
        hideLoading();
        updateUI();
    }
}

// Обработчик скачивания всех писем
async function handleDownloadAll() {
    try {
        updateStatus('Подготовка архива...', 'processing');
        
        const response = await fetch(`${API_BASE_URL}/download_all`);
        
        if (response.ok) {
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = url;
            a.download = 'all_letters.zip';
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
            
            updateStatus('Архив успешно скачан', 'success');
        } else {
            const data = await response.json();
            throw new Error(data.error || 'Ошибка при скачивании архива');
        }
    } catch (error) {
        showError('Ошибка при скачивании архива: ' + error.message);
        updateStatus('Ошибка скачивания', 'error');
    }
}

// Скачивание отдельного файла
async function downloadFile(filename) {
    try {
        const response = await fetch(`${API_BASE_URL}/download/${filename}`);
        
        if (response.ok) {
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
        } else {
            const data = await response.json();
            throw new Error(data.error || 'Ошибка при скачивании файла');
        }
    } catch (error) {
        showError('Ошибка при скачивании файла: ' + error.message);
    }
}

// Отображение результатов
function displayResults(data) {
    // Обновляем сводку
    resultsSummary.innerHTML = `
        <h3><i class="fas fa-check-circle"></i> Обработка завершена успешно</h3>
        <p><strong>Количество писем:</strong> ${data.letters_count}</p>
        <p><strong>Файлов сгенерировано:</strong> ${data.files_generated ? data.files_generated.length : 0}</p>
        <p><strong>Время обработки:</strong> ${new Date().toLocaleString('ru-RU')}</p>
    `;
    
    // Отображаем список писем
    if (appState.lettersData && appState.lettersData.length > 0) {
        lettersList.innerHTML = appState.lettersData.map((letter, index) => `
            <div class="letter-item fade-in">
                <div class="letter-header">
                    <div class="letter-info">
                        <h4>${letter.contractor_name}</h4>
                        <p><strong>Заказ:</strong> ${letter.order_number}</p>
                        <p><strong>Сумма:</strong> ${formatCurrency(letter.total_amount)}</p>
                        <p><strong>Пени:</strong> ${formatCurrency(letter.total_penalty)}</p>
                        <p><strong>Позиций:</strong> ${letter.total_positions}</p>
                    </div>
                    <div class="letter-actions">
                        <button class="btn btn-small btn-primary" onclick="downloadFile('letter_${index + 1}_${letter.contractor_short_name}_${letter.order_number}.docx')">
                            <i class="fas fa-download"></i> Письмо
                        </button>
                        <button class="btn btn-small btn-success" onclick="downloadFile('appendix_${index + 1}_${letter.contractor_short_name}_${letter.order_number}.docx')">
                            <i class="fas fa-download"></i> Приложение
                        </button>
                    </div>
                </div>
            </div>
        `).join('');
    }
}

// Проверка статуса сервера
async function checkServerStatus() {
    try {
        const response = await fetch(`${API_BASE_URL}/status`);
        const data = await response.json();
        
        if (response.ok) {
            if (data.reporting_file_uploaded && data.sed_file_uploaded) {
                appState.filesUploaded = true;
                updateStatus('Файлы загружены, готов к обработке', 'success');
            }
            
            if (data.generated_letters_count > 0) {
                appState.dataProcessed = true;
                updateStatus(`Найдено ${data.generated_letters_count} сгенерированных писем`, 'success');
            }
        }
    } catch (error) {
        console.log('Не удалось получить статус сервера:', error.message);
    }
    
    updateUI();
}

// Обновление интерфейса
function updateUI() {
    // Кнопка загрузки
    uploadBtn.disabled = !(appState.reportingFileSelected && appState.sedFileSelected);
    
    // Секция обработки
    processSection.style.display = appState.filesUploaded ? 'block' : 'none';
    
    // Секция результатов
    resultsSection.style.display = appState.dataProcessed ? 'block' : 'none';
}

// Обновление статуса
function updateStatus(message, type = 'default') {
    statusText.textContent = message;
    statusIndicator.className = `status-indicator ${type}`;
    
    // Обновляем иконку
    const icon = statusIndicator.querySelector('i');
    switch (type) {
        case 'success':
            icon.className = 'fas fa-check-circle';
            break;
        case 'error':
            icon.className = 'fas fa-exclamation-triangle';
            break;
        case 'processing':
            icon.className = 'fas fa-spinner fa-spin';
            break;
        default:
            icon.className = 'fas fa-info-circle';
    }
}

// Показать загрузку
function showLoading(message = 'Загрузка...') {
    loadingText.textContent = message;
    loadingModal.style.display = 'block';
}

// Скрыть загрузку
function hideLoading() {
    loadingModal.style.display = 'none';
}

// Показать ошибку
function showError(message) {
    errorText.textContent = message;
    errorModal.style.display = 'block';
}

// Скрыть ошибку
function hideErrorModal() {
    errorModal.style.display = 'none';
}

// Показать успех
function showSuccess(message) {
    // Создаем временное уведомление
    const notification = document.createElement('div');
    notification.className = 'success-message fade-in';
    notification.innerHTML = `<i class="fas fa-check-circle"></i> ${message}`;
    
    document.querySelector('.main-content').insertBefore(notification, document.querySelector('.main-content').firstChild);
    
    // Удаляем через 5 секунд
    setTimeout(() => {
        if (notification.parentNode) {
            notification.parentNode.removeChild(notification);
        }
    }, 5000);
}

// Форматирование валюты
function formatCurrency(amount) {
    return new Intl.NumberFormat('ru-RU', {
        style: 'currency',
        currency: 'RUB'
    }).format(amount);
}

// Обработка ошибок глобально
window.addEventListener('error', function(event) {
    console.error('Глобальная ошибка:', event.error);
    showError('Произошла неожиданная ошибка. Пожалуйста, обновите страницу.');
});

// Обработка необработанных промисов
window.addEventListener('unhandledrejection', function(event) {
    console.error('Необработанная ошибка промиса:', event.reason);
    showError('Произошла ошибка при выполнении запроса.');
});

