/**
 * DocSmart Office Tools - Excel to PDF Converter Engine
 * Uses SheetJS (xlsx) for spreadsheet parsing and jsPDF + AutoTable for PDF rendering.
 */

class ExcelToPdfConverter {
  constructor() {
    this.workbook = null;
    this.fileName = '';
    this.sheetsData = {};
  }

  /**
   * Load and parse an Excel workbook from an ArrayBuffer or File
   * @param {File|ArrayBuffer} fileOrBuffer
   * @param {string} fileName
   * @returns {Promise<Object>} workbook metadata
   */
  async loadWorkbook(fileOrBuffer, fileName = 'spreadsheet.xlsx') {
    this.fileName = fileName;
    this.sheetsData = {};

    let data;
    if (fileOrBuffer instanceof File || fileOrBuffer instanceof Blob) {
      data = await fileOrBuffer.arrayBuffer();
    } else {
      data = fileOrBuffer;
    }

    try {
      // Read workbook using SheetJS
      this.workbook = XLSX.read(data, {
        type: 'array',
        cellDates: true,
        cellNF: true,
        cellText: true,
      });
    } catch (err) {
      console.error('Error parsing Excel file:', err);
      throw new Error('Unable to read this file. Please make sure it is a valid .xlsx, .xls, or .csv document.');
    }

    if (!this.workbook.SheetNames || this.workbook.SheetNames.length === 0) {
      throw new Error('This Excel file does not contain any sheets.');
    }

    // Parse each sheet into 2D array representation
    for (const sheetName of this.workbook.SheetNames) {
      const worksheet = this.workbook.Sheets[sheetName];
      // Convert sheet to json 2D array (header: 1 returns array of arrays)
      const rawData = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        raw: false,
        defval: '',
        blankrows: false,
      });

      // Filter out completely empty trailing rows
      const cleanedRows = rawData.filter(row => Array.isArray(row) && row.some(cell => String(cell).trim() !== ''));

      // Calculate max columns
      let maxCols = 0;
      cleanedRows.forEach(row => {
        if (row.length > maxCols) maxCols = row.length;
      });

      // Pad rows so all rows have the same number of columns
      const normalizedRows = cleanedRows.map(row => {
        const padded = [...row];
        while (padded.length < maxCols) {
          padded.push('');
        }
        return padded.map(cell => (cell === null || cell === undefined ? '' : String(cell).trim()));
      });

      this.sheetsData[sheetName] = {
        name: sheetName,
        rows: normalizedRows,
        rowCount: normalizedRows.length,
        colCount: maxCols,
      };
    }

    return {
      fileName: this.fileName,
      sheetNames: this.workbook.SheetNames,
      sheets: this.sheetsData,
    };
  }

  /**
   * Convert parsed Excel sheets into a PDF document
   * @param {Object} options Conversion preferences
   * @param {Function} onProgress Progress callback (step, percent)
   * @returns {Promise<{blob: Blob, url: string, doc: any}>}
   */
  async convertToPdf(options = {}, onProgress = () => {}) {
    const {
      targetSheet = 'all', // 'all' or specific sheet name
      orientation = 'auto', // 'auto', 'portrait', 'landscape'
      pageSize = 'a4', // 'a4', 'letter', 'legal'
      theme = 'emerald', // 'emerald', 'navy', 'minimal', 'monochrome'
      fontSize = 'normal', // 'compact', 'normal', 'large'
      includeHeader = true,
      includeFooter = true,
      includeTitle = true,
      fitToPage = true,
    } = options;

    if (!this.workbook || Object.keys(this.sheetsData).length === 0) {
      throw new Error('No Excel workbook loaded. Please upload a file first.');
    }

    // Determine which sheets to process
    const sheetsToConvert = [];
    if (targetSheet === 'all') {
      for (const name of this.workbook.SheetNames) {
        if (this.sheetsData[name] && this.sheetsData[name].rowCount > 0) {
          sheetsToConvert.push(this.sheetsData[name]);
        }
      }
    } else {
      if (this.sheetsData[targetSheet]) {
        sheetsToConvert.push(this.sheetsData[targetSheet]);
      }
    }

    if (sheetsToConvert.length === 0) {
      throw new Error('The selected sheet contains no printable data.');
    }

    onProgress('Preparing layout...', 25);

    // Calculate maximum columns across sheets to decide auto orientation
    let maxColsOverall = 0;
    sheetsToConvert.forEach(s => {
      if (s.colCount > maxColsOverall) maxColsOverall = s.colCount;
    });

    let chosenOrientation = orientation;
    if (orientation === 'auto') {
      chosenOrientation = maxColsOverall > 6 ? 'landscape' : 'portrait';
    }

    // Setup jsPDF
    const jsPDFClass = (window.jspdf && window.jspdf.jsPDF) ? window.jspdf.jsPDF : (window.jsPDF || (typeof jsPDF !== 'undefined' ? jsPDF : null));
    if (!jsPDFClass) {
      throw new Error('PDF generator library (jsPDF) is not loaded. Please refresh the page.');
    }
    const doc = new jsPDFClass({
      orientation: chosenOrientation,
      unit: 'mm',
      format: pageSize,
    });

    // Theme color palettes
    const themeStyles = {
      emerald: {
        headBg: [16, 124, 65], // #107c41
        headText: [255, 255, 255],
        alternateBg: [247, 252, 248],
        borderColor: [209, 231, 214],
        titleColor: [16, 124, 65],
      },
      navy: {
        headBg: [30, 58, 138], // #1e3a8a
        headText: [255, 255, 255],
        alternateBg: [241, 245, 249],
        borderColor: [203, 213, 225],
        titleColor: [30, 58, 138],
      },
      minimal: {
        headBg: [71, 85, 105], // #475569
        headText: [255, 255, 255],
        alternateBg: [248, 250, 252],
        borderColor: [226, 232, 240],
        titleColor: [30, 41, 59],
      },
      monochrome: {
        headBg: [24, 24, 27], // #18181b
        headText: [255, 255, 255],
        alternateBg: [244, 244, 245],
        borderColor: [212, 212, 216],
        titleColor: [0, 0, 0],
      },
    };

    const activeTheme = themeStyles[theme] || themeStyles.emerald;

    // Font size mapping
    const fontSizes = {
      compact: { body: 7, head: 7.5, title: 13 },
      normal: { body: 8.5, head: 9, title: 15 },
      large: { body: 10, head: 10.5, title: 16 },
    };
    const activeFonts = fontSizes[fontSize] || fontSizes.normal;

    const totalSheets = sheetsToConvert.length;
    let sheetIndex = 0;

    for (const sheet of sheetsToConvert) {
      sheetIndex++;
      const percent = Math.round(30 + (sheetIndex / totalSheets) * 50);
      onProgress(`Rendering "${sheet.name}" (${sheetIndex} of ${totalSheets})...`, percent);

      // If not the first sheet, add a new page
      if (sheetIndex > 1) {
        doc.addPage(pageSize, chosenOrientation);
      }

      const pageWidth = doc.internal.pageSize.getWidth();
      const pageHeight = doc.internal.pageSize.getHeight();
      let startY = 16;

      // Draw Document & Sheet Title Header
      if (includeTitle) {
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(activeFonts.title);
        doc.setTextColor(...activeTheme.titleColor);
        doc.text(sheet.name, 14, startY);

        doc.setFont('helvetica', 'normal');
        doc.setFontSize(8.5);
        doc.setTextColor(120, 120, 120);

        const cleanBaseName = this.fileName.replace(/\.[^/.]+$/, '');
        const metaText = `${cleanBaseName} • ${sheet.rowCount} rows × ${sheet.colCount} columns`;
        doc.text(metaText, 14, startY + 5.5);

        // Right-aligned Date Timestamp
        if (includeHeader) {
          const dateStr = new Date().toLocaleDateString(undefined, {
            year: 'numeric',
            month: 'short',
            day: 'numeric',
          });
          doc.text(`Generated: ${dateStr}`, pageWidth - 14, startY + 5.5, { align: 'right' });
        }

        // Horizontal accent line
        doc.setDrawColor(...activeTheme.headBg);
        doc.setLineWidth(0.6);
        doc.line(14, startY + 8, pageWidth - 14, startY + 8);

        startY += 12;
      }

      // Prepare table data
      const rows = sheet.rows;
      if (rows.length === 0) continue;

      // First row as headers, remaining as body
      const tableHeaders = [rows[0]];
      const tableBody = rows.slice(1);

      // AutoTable render configuration
      doc.autoTable({
        head: tableHeaders,
        body: tableBody,
        startY: startY,
        margin: { top: 18, right: 12, bottom: 18, left: 12 },
        theme: 'grid',
        tableWidth: 'auto',
        styles: {
          font: 'helvetica',
          fontSize: activeFonts.body,
          cellPadding: { top: 2.2, right: 3, bottom: 2.2, left: 3 },
          overflow: 'linebreak',
          cellWidth: 'wrap',
          textColor: [30, 41, 59],
          lineColor: activeTheme.borderColor,
          lineWidth: 0.15,
        },
        headStyles: {
          fillColor: activeTheme.headBg,
          textColor: activeTheme.headText,
          fontSize: activeFonts.head,
          fontStyle: 'bold',
          halign: 'left',
          valign: 'middle',
          minCellHeight: 8,
        },
        alternateRowStyles: {
          fillColor: activeTheme.alternateBg,
        },
        didDrawPage: data => {
          // Add Running Header (if multi-page and on subsequent pages)
          if (includeHeader && data.pageNumber > 1) {
            doc.setFont('helvetica', 'italic');
            doc.setFontSize(8);
            doc.setTextColor(140, 140, 140);
            doc.text(`${sheet.name} (continued)`, 14, 10);
            doc.setDrawColor(220, 220, 220);
            doc.setLineWidth(0.2);
            doc.line(14, 12, pageWidth - 14, 12);
          }

          // Add Running Footer with Page Numbers
          if (includeFooter) {
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(8);
            doc.setTextColor(130, 130, 130);

            // Left footer: branding
            doc.text('DocSmart Office Tools', 14, pageHeight - 8);

            // Right footer: Page number placeholder (updated later)
            const pageStr = `Page ${doc.internal.getNumberOfPages()}`;
            doc.text(pageStr, pageWidth - 14, pageHeight - 8, { align: 'right' });
          }
        },
      });
    }

    onProgress('Finalizing PDF document...', 90);

    // Finalize page count in footer for all pages
    if (includeFooter) {
      const totalPages = doc.internal.getNumberOfPages();
      for (let i = 1; i <= totalPages; i++) {
        doc.setPage(i);
        const pageWidth = doc.internal.pageSize.getWidth();
        const pageHeight = doc.internal.pageSize.getHeight();
        
        // Blank out previous page text area cleanly
        doc.setFillColor(255, 255, 255);
        doc.rect(pageWidth - 45, pageHeight - 12, 35, 6, 'F');

        // Draw accurate "Page X of Y"
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(8);
        doc.setTextColor(130, 130, 130);
        doc.text(`Page ${i} of ${totalPages}`, pageWidth - 14, pageHeight - 8, { align: 'right' });
      }
    }

    onProgress('Done! Ready to view and download.', 100);

    // Generate PDF Blob and URL for viewing and downloading
    const pdfBlob = doc.output('blob');
    const pdfUrl = URL.createObjectURL(pdfBlob);

    return {
      doc,
      blob: pdfBlob,
      url: pdfUrl,
      fileName: this.getOutputPdfName(),
    };
  }

  /**
   * Helper to derive the suggested PDF filename
   */
  getOutputPdfName() {
    const base = this.fileName ? this.fileName.replace(/\.[^/.]+$/, '') : 'spreadsheet';
    return `${base}_converted.pdf`;
  }
}

// Export singleton instance or class to window
window.ExcelToPdfConverter = ExcelToPdfConverter;
