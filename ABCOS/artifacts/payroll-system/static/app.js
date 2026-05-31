document.addEventListener('DOMContentLoaded', () => {
  setupCard('bc');
  setupCard('pdf');
  setupCard('payslip');
});

const MAX_UPLOAD_MB = Number(document.body.dataset.maxUploadMb || 50);
const MAX_UPLOAD_BYTES = MAX_UPLOAD_MB * 1024 * 1024;
const REQUEST_TIMEOUT_MS = {
  bc: 60_000,
  pdf: 120_000,
  payslip: 120_000,
};

function setupCard(type) {
  const dropZone = document.getElementById(`drop-${type}`);
  const fileInput = document.getElementById(`${type}-file`);
  const fileNameEl = document.getElementById(`${type}-filename`);
  const btn = document.getElementById(`btn-${type}`);
  const resetBtn = document.getElementById(`reset-${type}`);
  const spinner = document.getElementById(`${type}-spinner`);
  const result = document.getElementById(`${type}-result`);

  const endpointMap = { bc: '/convert/bc', pdf: '/convert/pdf', payslip: '/convert/payslip-excel' };
  const endpoint = endpointMap[type] || '/convert/bc';

  dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
  });

  dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('drag-over');
  });

  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      setFile(files[0]);
    }
  });

  dropZone.addEventListener('click', (e) => {
    if (e.target.tagName !== 'LABEL' && e.target.tagName !== 'INPUT') {
      fileInput.click();
    }
  });

  fileInput.addEventListener('change', () => {
    if (fileInput.files.length > 0) {
      setFile(fileInput.files[0]);
    }
  });

  resetBtn.addEventListener('click', () => resetCard(type));

  function setFile(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['xls', 'xlsx'].includes(ext)) {
      showError(result, 'xls 또는 xlsx 파일만 업로드 가능합니다.');
      result.classList.remove('d-none');
      return;
    }
    if (file.size === 0) {
      showError(result, '빈 파일입니다. 다른 파일을 선택해 주세요.');
      result.classList.remove('d-none');
      return;
    }
    if (file.size > MAX_UPLOAD_BYTES) {
      showError(result, `파일 크기가 ${MAX_UPLOAD_MB}MB를 초과합니다.`);
      result.classList.remove('d-none');
      return;
    }
    fileNameEl.textContent = `📄 ${file.name} (${formatSize(file.size)})`;
    dropZone.classList.add('has-file');
    btn.disabled = false;
    resetBtn.classList.remove('d-none');
    result.classList.add('d-none');
    result.innerHTML = '';

    const dt = new DataTransfer();
    dt.items.add(file);
    fileInput.files = dt.files;
  }

  btn.addEventListener('click', async () => {
    if (!fileInput.files.length) return;

    btn.disabled = true;
    resetBtn.classList.add('d-none');
    spinner.classList.remove('d-none');
    result.classList.add('d-none');
    result.innerHTML = '';

    const formData = new FormData();
    formData.append('file', fileInput.files[0]);

    const controller = new AbortController();
    const timeoutMs = REQUEST_TIMEOUT_MS[type] || 60_000;
    const timeoutId = setTimeout(() => controller.abort(), timeoutMs);

    try {
      const response = await fetch(endpoint, {
        method: 'POST',
        body: formData,
        signal: controller.signal,
      });

      clearTimeout(timeoutId);
      spinner.classList.add('d-none');

      const contentType = response.headers.get('content-type') || '';

      if (!response.ok || contentType.includes('application/json')) {
        let errorMsg = '알 수 없는 오류가 발생했습니다.';
        try {
          const data = await response.json();
          errorMsg = data.error || errorMsg;
        } catch {
          if (response.status === 413) {
            errorMsg = `파일 크기가 ${MAX_UPLOAD_MB}MB를 초과합니다.`;
          }
        }
        showError(result, errorMsg);
        result.classList.remove('d-none');
        btn.disabled = false;
        resetBtn.classList.remove('d-none');
        return;
      }

      const disposition = response.headers.get('content-disposition') || '';

      let filename = type === 'bc' ? '청구서.xlsx' : '급여명세서.zip';
      const utf8Match = disposition.match(/filename\*=UTF-8''([^;\s]+)/i);
      if (utf8Match) {
        filename = decodeURIComponent(utf8Match[1]);
      } else {
        const plainMatch = disposition.match(/filename=['"]?([^;'"]+)/i);
        if (plainMatch) {
          filename = plainMatch[1];
        }
      }

      const blob = await response.blob();
      if (blob.size === 0) {
        showError(result, '다운로드 파일이 비어 있습니다. 다시 시도해 주세요.');
        result.classList.remove('d-none');
        btn.disabled = false;
        resetBtn.classList.remove('d-none');
        return;
      }

      const url = URL.createObjectURL(blob);

      const isZip = filename.endsWith('.zip');
      const icon = isZip ? 'bi-file-zip' : 'bi-file-earmark-excel';
      const btnClassMap = { bc: 'download-btn', pdf: 'download-btn download-btn-red', payslip: 'download-btn download-btn-green' };
      const btnClass = btnClassMap[type] || 'download-btn';

      result.innerHTML = `
        <div class="result-success">
          <div class="d-flex align-items-center gap-2 mb-2">
            <i class="bi bi-check-circle-fill text-success fs-5"></i>
            <strong>변환 완료!</strong>
          </div>
          <p class="mb-2 small text-muted">파일이 준비되었습니다. 아래 버튼을 클릭해 다운로드하세요.</p>
          <a href="${url}" download="${escapeHtml(filename)}" class="${btnClass} d-block w-100 text-center">
            <i class="bi ${icon}"></i>${escapeHtml(filename)}
          </a>
        </div>
      `;
      result.classList.remove('d-none');
      btn.disabled = false;
      resetBtn.classList.remove('d-none');

    } catch (err) {
      clearTimeout(timeoutId);
      spinner.classList.add('d-none');
      const message = err.name === 'AbortError'
        ? '요청 시간이 초과되었습니다. 파일 크기를 줄이거나 잠시 후 다시 시도해 주세요.'
        : `네트워크 오류: ${err.message}`;
      showError(result, message);
      result.classList.remove('d-none');
      btn.disabled = false;
      resetBtn.classList.remove('d-none');
    }
  });
}

function showError(el, msg) {
  el.innerHTML = `
    <div class="result-error">
      <div class="d-flex align-items-center gap-2 mb-1">
        <i class="bi bi-exclamation-triangle-fill text-danger fs-5"></i>
        <strong>오류 발생</strong>
      </div>
      <p class="mb-0 small">${escapeHtml(msg)}</p>
    </div>
  `;
  el.classList.remove('d-none');
}

function formatSize(bytes) {
  if (bytes < 1024) return bytes + 'B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + 'KB';
  return (bytes / (1024 * 1024)).toFixed(1) + 'MB';
}

function escapeHtml(str) {
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function resetCard(type) {
  const dropZone = document.getElementById(`drop-${type}`);
  const fileInput = document.getElementById(`${type}-file`);
  const fileNameEl = document.getElementById(`${type}-filename`);
  const btn = document.getElementById(`btn-${type}`);
  const resetBtn = document.getElementById(`reset-${type}`);
  const result = document.getElementById(`${type}-result`);

  fileInput.value = '';
  fileNameEl.textContent = '';
  dropZone.classList.remove('has-file');
  btn.disabled = true;
  resetBtn.classList.add('d-none');
  result.classList.add('d-none');
  result.innerHTML = '';
}
