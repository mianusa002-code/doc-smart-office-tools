/**
 * DocSmart Office Tools - Application Controller
 * Handles user interactions, file dropzone, previews, and conversion flow.
 */

document.addEventListener('DOMContentLoaded', () => {
  // Core state
  const converter = new ExcelToPdfConverter();
  let currentFile = null;
  let currentMetadata = null;
  let activePreviewSheet = null;
  let convertedPdfResult = null;

  // DOM Elements
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('fileInput');
  const btnBrowse = document.getElementById('btnBrowse');
  const btnSample = document.getElementById('btnSample');

  const fileCard = document.getElementById('fileCard');
  const fileNameEl = document.getElementById('fileName');
  const fileMetaEl = document.getElementById('fileMeta');
  const btnRemoveFile = document.getElementById('btnRemoveFile');

  const sheetSelectionBox = document.getElementById('sheetSelectionBox');
  const sheetTabsContainer = document.getElementById('sheetTabs');

  const previewToggle = document.getElementById('previewToggle');
  const previewContent = document.getElementById('previewContent');
  const previewTableContainer = document.getElementById('previewTableContainer');
  const previewRowCountEl = document.getElementById('previewRowCount');
  const previewColCountEl = document.getElementById('previewColCount');

  const optionsPanel = document.getElementById('optionsPanel');
  const optOrientation = document.getElementById('optOrientation');
  const optPageSize = document.getElementById('optPageSize');
  const optTheme = document.getElementById('optTheme');
  const optFontSize = document.getElementById('optFontSize');
  const optIncludeTitle = document.getElementById('optIncludeTitle');
  const optIncludeHeader = document.getElementById('optIncludeHeader');
  const optIncludeFooter = document.getElementById('optIncludeFooter');

  const convertActionArea = document.getElementById('convertActionArea');
  const btnConvert = document.getElementById('btnConvert');

  const progressCard = document.getElementById('progressCard');
  const progressText = document.getElementById('progressText');
  const progressPercent = document.getElementById('progressPercent');
  const progressBarFill = document.getElementById('progressBarFill');

  const resultCard = document.getElementById('resultCard');
  const btnDownloadPdf = document.getElementById('btnDownloadPdf');
  const btnPrintPdf = document.getElementById('btnPrintPdf');
  const btnConvertAnother = document.getElementById('btnConvertAnother');
  const pdfFrame = document.getElementById('pdfFrame');

  const toastContainer = document.getElementById('toastContainer');

  // ==========================================
  // Toast Notifications
  // ==========================================
  function showToast(message, type = 'info', duration = 4000) {
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;

    let iconSvg = '';
    if (type === 'error') {
      iconSvg = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#dc2626" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="8" x2="12" y2="12"></line><line x1="12" y1="16" x2="12.01" y2="16"></line></svg>`;
    } else if (type === 'success') {
      iconSvg = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#16a34a" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>`;
    } else {
      iconSvg = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#0078d4" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="16" x2="12" y2="12"></line><line x1="12" y1="8" x2="12.01" y2="8"></line></svg>`;
    }

    toast.innerHTML = `
      <span class="toast-icon">${iconSvg}</span>
      <span class="toast-message">${escapeHtml(message)}</span>
      <button class="toast-close" aria-label="Close">&times;</button>
    `;

    toast.querySelector('.toast-close').addEventListener('click', () => {
      toast.remove();
    });

    toastContainer.appendChild(toast);

    if (duration > 0) {
      setTimeout(() => {
        if (toast.parentElement) {
          toast.style.opacity = '0';
          toast.style.transition = 'opacity 300ms ease';
          setTimeout(() => toast.remove(), 300);
        }
      }, duration);
    }
  }

  function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  function formatBytes(bytes, decimals = 1) {
    if (!bytes || bytes === 0) return '0 Bytes';
    const k = 1024;
    const dm = decimals < 0 ? 0 : decimals;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
  }

  // ==========================================
  // Drag and Drop & File Input Handling
  // ==========================================
  btnBrowse.addEventListener('click', e => {
    e.stopPropagation();
    fileInput.click();
  });

  dropzone.addEventListener('click', () => {
    fileInput.click();
  });

  ['dragenter', 'dragover'].forEach(eventName => {
    dropzone.addEventListener(eventName, e => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.add('dragover');
    });
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropzone.addEventListener(eventName, e => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.remove('dragover');
    });
  });

  dropzone.addEventListener('drop', e => {
    const dt = e.dataTransfer;
    const files = dt.files;
    if (files && files.length > 0) {
      handleFileSelected(files[0]);
    }
  });

  fileInput.addEventListener('change', e => {
    if (fileInput.files && fileInput.files.length > 0) {
      handleFileSelected(fileInput.files[0]);
    }
  });

  btnRemoveFile.addEventListener('click', () => {
    resetApp();
    showToast('File removed', 'info');
  });

  // ==========================================
  // Sample Data Generator
  // ==========================================
  btnSample.addEventListener('click', e => {
    e.stopPropagation();
    generateSampleSpreadsheet();
  });

  function generateSampleSpreadsheet() {
    try {
      showToast('Generating sample corporate spreadsheet...', 'info', 2000);

      // Create a brand new workbook using SheetJS
      const wb = XLSX.utils.book_new();

      // Sheet 1: Quarterly Sales Performance
      const salesData = [
        ['Region', 'Product Line', 'Q1 Revenue ($)', 'Q2 Revenue ($)', 'Q3 Revenue ($)', 'Q4 Revenue ($)', 'Growth Rate', 'Status'],
        ['North America', 'Enterprise Cloud Suite', '$142,500', '$168,200', '$185,400', '$210,000', '+18.5%', 'Exceeded'],
        ['North America', 'AI Developer Tools', '$98,400', '$115,000', '$134,800', '$159,200', '+22.4%', 'Exceeded'],
        ['Europe / EMEA', 'Security & Compliance', '$88,200', '$92,500', '$97,100', '$105,400', '+9.2%', 'On Target'],
        ['Europe / EMEA', 'Enterprise Cloud Suite', '$115,600', '$121,000', '$130,500', '$142,800', '+12.1%', 'On Target'],
        ['Asia Pacific', 'AI Developer Tools', '$76,800', '$89,400', '$104,200', '$126,500', '+26.8%', 'Exceeded'],
        ['Asia Pacific', 'Hardware Appliances', '$54,100', '$52,800', '$49,600', '$48,200', '-5.2%', 'Under Review'],
        ['Latin America', 'Enterprise Cloud Suite', '$38,200', '$42,500', '$47,800', '$53,100', '+14.6%', 'On Target'],
        ['Total / Average', 'All Products Combined', '$613,800', '$681,400', '$749,400', '$845,200', '+15.8%', 'Strong Growth'],
      ];
      const ws1 = XLSX.utils.aoa_to_sheet(salesData);
      XLSX.utils.book_append_sheet(wb, ws1, 'Quarterly Sales');

      // Sheet 2: Department Budget & Expenses
      const expenseData = [
        ['Department', 'Budget Allocated ($)', 'Q1-Q2 Spent ($)', 'Q3-Q4 Spent ($)', 'Total Spent ($)', 'Remaining ($)', 'Variance (%)'],
        ['Research & Development', '$350,000', '$165,000', '$172,400', '$337,400', '$12,600', '-3.6%'],
        ['Sales & Marketing', '$280,000', '$138,500', '$139,200', '$277,700', '$2,300', '-0.8%'],
        ['Information Technology', '$190,000', '$88,400', '$94,100', '$182,500', '$7,500', '-3.9%'],
        ['Customer Operations', '$140,000', '$67,200', '$71,300', '$138,500', '$1,500', '-1.1%'],
        ['Human Resources & Admin', '$95,000', '$44,800', '$47,200', '$92,000', '$3,000', '-3.2%'],
        ['Total Operating Expenses', '$1,055,000', '$503,900', '$524,200', '$1,028,100', '$26,900', '-2.5%'],
      ];
      const ws2 = XLSX.utils.aoa_to_sheet(expenseData);
      XLSX.utils.book_append_sheet(wb, ws2, 'Department Expenses');

      // Generate binary buffer
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const sampleFile = new File([wbout], 'Corporate_Annual_Report_Sample.xlsx', {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });

      handleFileSelected(sampleFile);
      showToast('Sample Excel file loaded successfully!', 'success');
    } catch (err) {
      console.error(err);
      showToast('Could not load sample file: ' + err.message, 'error');
    }
  }

  // ==========================================
  // Process Selected File
  // ==========================================
  async function handleFileSelected(file) {
    if (!file) return;

    // Validate extension
    const validExtensions = ['.xlsx', '.xls', '.csv', '.ods', '.tsv'];
    const fileExt = '.' + file.name.split('.').pop().toLowerCase();
    if (!validExtensions.includes(fileExt)) {
      showToast(`Unsupported format (${fileExt}). Please upload an Excel (.xlsx, .xls) or CSV file.`, 'error', 5000);
      return;
    }

    currentFile = file;

    // Show loading state
    fileNameEl.textContent = 'Loading ' + file.name + '...';
    fileMetaEl.textContent = 'Reading spreadsheet data...';
    fileCard.classList.remove('hidden');
    dropzone.classList.add('hidden');

    try {
      currentMetadata = await converter.loadWorkbook(file, file.name);

      // Update file card details
      fileNameEl.textContent = currentMetadata.fileName;
      const totalSheets = currentMetadata.sheetNames.length;
      fileMetaEl.innerHTML = `
        <span>Size: <strong>${formatBytes(file.size)}</strong></span>
        <span>•</span>
        <span>Sheets: <strong>${totalSheets}</strong></span>
      `;

      // Build Sheet Selector Tabs
      renderSheetTabs(currentMetadata.sheetNames);

      // Auto-preview the first sheet
      activePreviewSheet = currentMetadata.sheetNames[0];
      renderTablePreview(activePreviewSheet);

      // Show options panel and convert button
      sheetSelectionBox.classList.remove('hidden');
      optionsPanel.classList.remove('hidden');
      convertActionArea.classList.remove('hidden');
      resultCard.classList.add('hidden');

      showToast(`Successfully loaded "${file.name}" with ${totalSheets} sheet(s).`, 'success');
    } catch (err) {
      console.error(err);
      showToast(err.message || 'Failed to read file.', 'error', 6000);
      resetApp();
    }
  }

  // ==========================================
  // Render Sheet Tabs
  // ==========================================
  function renderSheetTabs(sheetNames) {
    sheetTabsContainer.innerHTML = '';

    // "All Sheets" option if more than 1 sheet
    if (sheetNames.length > 1) {
      const allBtn = document.createElement('button');
      allBtn.type = 'button';
      allBtn.className = 'sheet-tab active';
      allBtn.dataset.sheet = 'all';
      allBtn.innerHTML = `
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18"/><path d="M9 21V9"/></svg>
        All Sheets (${sheetNames.length})
      `;
      allBtn.addEventListener('click', () => selectSheetTab('all', allBtn));
      sheetTabsContainer.appendChild(allBtn);
    }

    // Individual sheet tabs
    sheetNames.forEach((sheetName, index) => {
      const sheetInfo = currentMetadata.sheets[sheetName];
      const tabBtn = document.createElement('button');
      tabBtn.type = 'button';
      tabBtn.className = `sheet-tab ${sheetNames.length === 1 && index === 0 ? 'active' : ''}`;
      tabBtn.dataset.sheet = sheetName;
      tabBtn.innerHTML = `
        <span>${escapeHtml(sheetName)}</span>
        <span class="badge-count">${sheetInfo.rowCount} rows</span>
      `;

      tabBtn.addEventListener('click', () => selectSheetTab(sheetName, tabBtn));
      sheetTabsContainer.appendChild(tabBtn);
    });
  }

  function selectSheetTab(sheetKey, buttonEl) {
    document.querySelectorAll('.sheet-tab').forEach(b => b.classList.remove('active'));
    buttonEl.classList.add('active');

    if (sheetKey === 'all') {
      activePreviewSheet = currentMetadata.sheetNames[0];
    } else {
      activePreviewSheet = sheetKey;
    }
    renderTablePreview(activePreviewSheet);
  }

  // ==========================================
  // Render Table Preview
  // ==========================================
  function renderTablePreview(sheetName) {
    const sheet = currentMetadata.sheets[sheetName];
    if (!sheet || !sheet.rows || sheet.rows.length === 0) {
      previewTableContainer.innerHTML = '<div style="padding: 20px; text-align: center; color: var(--gray-500);">This sheet contains no rows to preview.</div>';
      previewRowCountEl.textContent = '0 rows';
      previewColCountEl.textContent = '0 columns';
      return;
    }

    previewRowCountEl.textContent = `${sheet.rowCount} rows`;
    previewColCountEl.textContent = `${sheet.colCount} columns`;

    const maxPreviewRows = 15;
    const rowsToDisplay = sheet.rows.slice(0, maxPreviewRows);

    let html = '<table class="preview-table">';

    // Table Header (Row 0)
    if (rowsToDisplay.length > 0) {
      html += '<thead><tr>';
      rowsToDisplay[0].forEach((col, idx) => {
        const headerText = col !== '' ? col : `Column ${idx + 1}`;
        html += `<th>${escapeHtml(headerText)}</th>`;
      });
      html += '</tr></thead>';
    }

    // Table Body
    html += '<tbody>';
    for (let i = 1; i < rowsToDisplay.length; i++) {
      html += '<tr>';
      rowsToDisplay[i].forEach(cell => {
        html += `<td>${escapeHtml(cell)}</td>`;
      });
      html += '</tr>';
    }
    html += '</tbody></table>';

    if (sheet.rows.length > maxPreviewRows) {
      html += `<div style="padding: 8px 12px; background: var(--gray-50); font-size: 0.8rem; color: var(--gray-500); text-align: center; border-top: 1px solid var(--gray-200);">Showing first ${maxPreviewRows} rows of ${sheet.rowCount} total rows. The entire sheet will be converted in your PDF.</div>`;
    }

    previewTableContainer.innerHTML = html;
  }

  // Collapsible preview toggle
  previewToggle.addEventListener('click', () => {
    const isShowing = previewContent.classList.contains('show');
    if (isShowing) {
      previewContent.classList.remove('show');
      previewToggle.classList.remove('expanded');
    } else {
      previewContent.classList.add('show');
      previewToggle.classList.add('expanded');
    }
  });

  // ==========================================
  // Conversion Execution & Progress Bar
  // ==========================================
  btnConvert.addEventListener('click', async () => {
    if (!currentFile || !currentMetadata) {
      showToast('Please upload an Excel file first.', 'warning');
      return;
    }

    // Determine target sheet
    const activeTab = document.querySelector('.sheet-tab.active');
    const targetSheet = activeTab ? activeTab.dataset.sheet : 'all';

    // Build options object
    const options = {
      targetSheet: targetSheet,
      orientation: optOrientation.value,
      pageSize: optPageSize.value,
      theme: optTheme.value,
      fontSize: optFontSize.value,
      includeTitle: optIncludeTitle.checked,
      includeHeader: optIncludeHeader.checked,
      includeFooter: optIncludeFooter.checked,
    };

    // UI State: Transition to progress
    btnConvert.disabled = true;
    btnConvert.innerHTML = `<span class="spinner"></span> Converting to PDF...`;
    progressCard.classList.remove('hidden');
    resultCard.classList.add('hidden');
    updateProgress('Initializing PDF engine...', 10);

    try {
      // Delay slightly for smooth UI frame render
      await new Promise(resolve => setTimeout(resolve, 150));

      // Run conversion
      convertedPdfResult = await converter.convertToPdf(options, (msg, pct) => {
        updateProgress(msg, pct);
      });

      // Show result
      displayConversionResult(convertedPdfResult);
      showToast('PDF conversion completed successfully!', 'success');
    } catch (err) {
      console.error('Conversion failed:', err);
      showToast('Conversion failed: ' + (err.message || 'Unknown error'), 'error', 6000);
      progressCard.classList.add('hidden');
    } finally {
      btnConvert.disabled = false;
      btnConvert.innerHTML = `
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="12" y1="18" x2="12" y2="12"/><line x1="9" y1="15" x2="15" y2="15"/></svg>
        Convert to PDF Now
      `;
    }
  });

  function updateProgress(message, percent) {
    progressText.textContent = message;
    progressPercent.textContent = `${percent}%`;
    progressBarFill.style.width = `${percent}%`;

    // Step dots highlights
    const step1 = document.getElementById('pStep1');
    const step2 = document.getElementById('pStep2');
    const step3 = document.getElementById('pStep3');
    const step4 = document.getElementById('pStep4');

    [step1, step2, step3, step4].forEach(s => s.classList.remove('active', 'completed'));

    if (percent < 30) {
      step1.classList.add('active');
    } else if (percent < 70) {
      step1.classList.add('completed');
      step2.classList.add('active');
    } else if (percent < 95) {
      step1.classList.add('completed');
      step2.classList.add('completed');
      step3.classList.add('active');
    } else {
      step1.classList.add('completed');
      step2.classList.add('completed');
      step3.classList.add('completed');
      step4.classList.add('active', 'completed');
    }
  }

  // ==========================================
  // Display Result & PDF Download
  // ==========================================
  function displayConversionResult(result) {
    progressCard.classList.add('hidden');
    resultCard.classList.remove('hidden');

    // Display PDF in iframe
    if (pdfFrame && result.url) {
      pdfFrame.src = result.url;
    }

    // Scroll smoothly to result card
    resultCard.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  btnDownloadPdf.addEventListener('click', () => {
    if (!convertedPdfResult || !convertedPdfResult.url) {
      showToast('No PDF available to download.', 'error');
      return;
    }

    const a = document.createElement('a');
    a.href = convertedPdfResult.url;
    a.download = convertedPdfResult.fileName || 'converted_spreadsheet.pdf';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);

    showToast(`Downloading "${convertedPdfResult.fileName}"...`, 'success');
  });

  btnPrintPdf.addEventListener('click', () => {
    if (!pdfFrame || !pdfFrame.contentWindow) {
      showToast('Print preview is not ready.', 'warning');
      return;
    }
    try {
      pdfFrame.contentWindow.print();
    } catch (e) {
      // Fallback: open URL in new tab for browser native print
      window.open(convertedPdfResult.url, '_blank');
    }
  });

  btnConvertAnother.addEventListener('click', () => {
    resetApp();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });

  // ==========================================
  // Reset Application State
  // ==========================================
  function resetApp() {
    currentFile = null;
    currentMetadata = null;
    activePreviewSheet = null;

    if (convertedPdfResult && convertedPdfResult.url) {
      URL.revokeObjectURL(convertedPdfResult.url);
      convertedPdfResult = null;
    }

    fileInput.value = '';
    pdfFrame.src = 'about:blank';

    fileCard.classList.add('hidden');
    sheetSelectionBox.classList.add('hidden');
    optionsPanel.classList.add('hidden');
    convertActionArea.classList.add('hidden');
    progressCard.classList.add('hidden');
    resultCard.classList.add('hidden');

    dropzone.classList.remove('hidden');
    previewTableContainer.innerHTML = '';
  }
});
