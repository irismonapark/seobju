const MAX_FILE_SIZE = 4 * 1024 * 1024;
const REQUEST_TIMEOUT_MS = 90000;

document.addEventListener('DOMContentLoaded', function() {
    initializeFileUpload();
    initializeForm();
    initializeRefresh();
});

function initializeFileUpload() {
    const zone = document.getElementById('pdfDropZone');
    const input = document.getElementById('resumePdf');

    zone.addEventListener('click', () => input.click());

    zone.addEventListener('dragover', (e) => {
        e.preventDefault();
        zone.classList.add('dragover');
    });

    zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));

    zone.addEventListener('drop', (e) => {
        e.preventDefault();
        zone.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) {
            input.files = e.dataTransfer.files;
            handleFileSelect(input);
        }
    });

    input.addEventListener('change', () => handleFileSelect(input));
}

function handleFileSelect(input) {
    const file = input.files[0];
    const info = document.getElementById('pdfFileInfo');
    const button = document.getElementById('convertButton');

    if (!file) {
        info.style.display = 'none';
        button.disabled = true;
        return;
    }

    if (file.size > MAX_FILE_SIZE) {
        showErrorMessage(`"${file.name}" 파일이 4MB를 초과합니다.`);
        input.value = '';
        info.style.display = 'none';
        button.disabled = true;
        return;
    }

    document.getElementById('pdfFileName').textContent =
        `${file.name} (${formatFileSize(file.size)})`;
    info.style.display = 'flex';
    button.disabled = false;
}

function formatFileSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function initializeRefresh() {
    document.getElementById('refreshButton').addEventListener('click', resetForm);
}

function resetForm() {
    const input = document.getElementById('resumePdf');
    const info = document.getElementById('pdfFileInfo');
    const button = document.getElementById('convertButton');
    const resultMessage = document.getElementById('resultMessage');
    const zone = document.getElementById('pdfDropZone');

    input.value = '';
    info.style.display = 'none';
    resultMessage.style.display = 'none';
    resultMessage.innerHTML = '';
    zone.classList.remove('dragover');
    button.disabled = true;
    button.classList.remove('loading');
    const buttonIcon = button.querySelector('.material-icons');
    if (buttonIcon) buttonIcon.textContent = 'table_chart';
}

function initializeForm() {
    document.getElementById('convertForm').addEventListener('submit', async function(e) {
        e.preventDefault();

        const pdfFile = document.getElementById('resumePdf').files[0];
        if (!pdfFile) {
            showErrorMessage('알바몬 이력서 PDF 파일을 업로드해주세요.');
            return;
        }

        const button = document.getElementById('convertButton');
        const buttonIcon = button.querySelector('.material-icons');
        const resultMessage = document.getElementById('resultMessage');

        button.disabled = true;
        button.classList.add('loading');
        buttonIcon.textContent = 'hourglass_empty';
        resultMessage.style.display = 'none';

        const formData = new FormData();
        formData.append('resume_pdf', pdfFile);

        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);

        try {
            const response = await fetch('/convert', {
                method: 'POST',
                body: formData,
                signal: controller.signal,
            });

            const contentType = response.headers.get('content-type') || '';

            if (response.ok && isExcelResponse(contentType)) {
                const blob = await response.blob();
                const filename = getFilenameFromResponse(response) || '알바몬_이력서.xlsx';
                triggerFileDownload(blob, filename);
                showSuccessMessage(filename);
            } else {
                showErrorMessage(await parseErrorResponse(response));
            }
        } catch (error) {
            if (error.name === 'AbortError') {
                showErrorMessage('처리 시간이 초과되었습니다. 파일 크기를 줄이거나 다시 시도해주세요.');
            } else {
                showErrorMessage('서버와 통신 중 오류가 발생했습니다. 다시 시도해주세요.');
            }
        } finally {
            clearTimeout(timeoutId);
            button.disabled = false;
            button.classList.remove('loading');
            buttonIcon.textContent = 'table_chart';
        }
    });
}

async function parseErrorResponse(response) {
    try {
        const data = await response.json();
        if (data && data.error) {
            if (data.detail) return `${data.error} (${data.detail})`;
            return data.error;
        }
    } catch (_) {
        /* ignore */
    }
    if (response.status === 413) return 'PDF 파일이 너무 큽니다 (최대 4MB).';
    if (response.status >= 500) return '서버 오류가 발생했습니다. 잠시 후 다시 시도해주세요.';
    return '변환에 실패했습니다.';
}

function isExcelResponse(contentType) {
    return contentType.includes('spreadsheetml') ||
        contentType.includes('octet-stream') ||
        contentType.includes('vnd.ms-excel');
}

function getFilenameFromResponse(response) {
    const disposition = response.headers.get('content-disposition') || '';
    const utf8Match = disposition.match(/filename\*=UTF-8''([^;]+)/i);
    if (utf8Match) return decodeURIComponent(utf8Match[1]);
    const match = disposition.match(/filename="?([^";]+)"?/i);
    return match ? match[1] : null;
}

function triggerFileDownload(blob, filename) {
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.URL.revokeObjectURL(url);
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function showSuccessMessage(filename) {
    const resultMessage = document.getElementById('resultMessage');
    resultMessage.className = 'result-message success';
    resultMessage.innerHTML = `
        <h3>
            <span class="material-icons">check_circle</span>
            변환이 완료되었습니다!
        </h3>
        <p><strong>${escapeHtml(filename)}</strong> 파일이 다운로드되었습니다.</p>
    `;
    resultMessage.style.display = 'block';
    resultMessage.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function showErrorMessage(error) {
    const resultMessage = document.getElementById('resultMessage');
    resultMessage.className = 'result-message error';
    resultMessage.innerHTML = `
        <h3>
            <span class="material-icons">error</span>
            변환 중 오류가 발생했습니다
        </h3>
        <p>${escapeHtml(error)}</p>
    `;
    resultMessage.style.display = 'block';
    resultMessage.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}
