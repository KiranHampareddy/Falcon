// IIFE to encapsulate code and avoid polluting the global scope
(function() {
    'use strict';

    // --- TOOL CONFIGURATION ---
    const TOOL_CONFIG = {
        'tool1': { name: 'AdX and Private Auction', processor: processTool1 },
        'bidding': { name: 'Open Bidding Partners', processor: processBiddingData },
        'tool2': { name: 'Rubicon/Magnite', processor: processTool2 },
        'citrusPivot': { name: 'Pubmatic Openwrap', processor: processCitrusAdPivotData },
        'tool3': { name: 'Verizon/Yahoo', processor: processTool3 },
        'tool4': { name: 'Xandr/AppNexus', processor: processTool4 },
        'tool5': { name: 'TripleLift', processor: processTool5 },
        'tool6': { name: 'Cadent Live', processor: processTool6 },
        'pivot': { name: 'EPSILON Live', processor: processPivotData },
        'tool7': { name: 'Media.net (HB)', processor: processTool7 },
        'tool8': { name: 'Sharethrough (HB)', processor: processTool8 },
        'impressions': { name: 'Programmatic Avails', isMultiFile: true, processor: processImpressionsData },
    };

    // --- INITIALIZATION ---
    document.addEventListener('DOMContentLoaded', () => {
        setupAccordion();
        setupDropZones();
        setupToolButtons();
        showPlaceholder();
    });

    // --- GLOBAL UI & UTILITY FUNCTIONS ---
    const outputPanel = document.getElementById('outputPanel');

    function setupAccordion() {
        const accordion = document.getElementById('toolAccordion');
        accordion.addEventListener('click', (e) => {
            const header = e.target.closest('.tool-header');
            if (!header) return;
            const card = header.parentElement;
            const isActive = card.classList.contains('active');
            accordion.querySelectorAll('.tool-card').forEach(c => c.classList.remove('active'));
            if (!isActive) {
                card.classList.add('active');
            }
        });
    }

    function setupDropZones() {
        document.querySelectorAll('.drop-zone').forEach(zone => {
            const inputId = zone.dataset.inputId;
            if (!inputId) return;
            const input = document.getElementById(inputId);
            if (!input) return;
            zone.addEventListener('click', () => input.click());
            ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
                zone.addEventListener(eventName, e => { e.preventDefault(); e.stopPropagation(); }, false);
            });
            ['dragenter', 'dragover'].forEach(eventName => {
                zone.addEventListener(eventName, () => zone.classList.add('dragover'), false);
            });
            ['dragleave', 'drop'].forEach(eventName => {
                zone.addEventListener(eventName, () => zone.classList.remove('dragover'), false);
            });
            zone.addEventListener('drop', e => {
                input.files = e.dataTransfer.files;
                updateFileDisplay(zone, input.files[0]);
            }, false);
            input.addEventListener('change', () => {
                updateFileDisplay(zone, input.files[0]);
            });
        });
    }

    function setupToolButtons() {
        document.querySelectorAll('.tool-card button[data-action="process"]').forEach(button => {
            const card = button.closest('.tool-card');
            const toolKey = card.dataset.tool;
            const config = TOOL_CONFIG[toolKey];
            if (config) {
                button.addEventListener('click', () => {
                    outputPanel.scrollTo({ top: 0, behavior: 'smooth' });
                    if (config.isMultiFile) {
                        config.processor();
                    } else {
                        const inputId = card.querySelector('input[type="file"]').id;
                        processFile(inputId, config.name, config.processor);
                    }
                });
            }
        });
    }

    function updateFileDisplay(zone, file) {
        const textSpan = zone.querySelector('.drop-zone-text');
        if (file) {
            textSpan.innerHTML = `✓ ${file.name}`;
            zone.classList.add('has-file');
        } else {
            textSpan.innerHTML = 'Drag & Drop or Click';
            zone.classList.remove('has-file');
        }
    }

    function showPlaceholder() {
        outputPanel.innerHTML = `<div class="output-wrapper"><div class="output-placeholder"><svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" d="M7.5 14.25v2.25m3-4.5v4.5m3-6.75v6.75m3-9v9M6 20.25h12A2.25 2.25 0 0020.25 18V6A2.25 2.25 0 0018 3.75H6A2.25 2.25 0 003.75 6v12A2.25 2.25 0 006 20.25z" /></svg><h2>Your Results Will Appear Here</h2><p>Select a tool, upload your file(s), and click process.</p></div></div>`;
    }

    function showLoading() {
        outputPanel.innerHTML = `<div class="output-wrapper"><div class="loading-spinner"></div></div>`;
    }

    function showToast(message) {
        const toast = document.createElement('div');
        toast.className = 'toast-notification';
        toast.textContent = message;
        document.body.appendChild(toast);
        setTimeout(() => toast.classList.add('show'), 10);
        setTimeout(() => {
            toast.classList.remove('show');
            toast.addEventListener('transitionend', () => toast.remove());
        }, 2500);
    }

    function renderOutput(title, htmlContent, tableId, options = {}) {
        const { error = null, fileName = '' } = options;
        let content = '';
        if (error) {
            content = `<div class="output-wrapper"><p style="color:var(--accent-red); background:#fef2f2; padding:1em; border-radius:8px; max-width: 80%;"><b>Error in ${title}:</b> ${error}</p></div>`;
        } else if (htmlContent && tableId) {
            if (!htmlContent.includes(`id="${tableId}"`)) {
                htmlContent = htmlContent.replace('<table', `<table id="${tableId}"`);
            }
            content = `
              <div class="results-container">
                  <div class="results-header">
                      <div class="results-title-group">
                          <h2 class="results-title">${title}</h2>
                          ${fileName ? `<p class="results-subtitle">Source: ${fileName}</p>` : ''}
                      </div>
                  </div>
                  <p style="text-align:center; color: var(--text-light); font-size: 0.9rem; margin-bottom: 16px;">
                      ✨ <strong>Pro Tip:</strong> Click and drag to select cells like in Excel, then press <strong>Ctrl+C</strong> to copy your selection.
                  </p>
                  <div class="table-container">${htmlContent}</div>
              </div>`;
        }
        outputPanel.innerHTML = content;
        if (tableId && !error) {
            makeTableSelectable(tableId);
        }
    }

    function readFileAsArrayBuffer(file) {
        return new Promise((resolve, reject) => {
            if (!file) { return reject(new Error('Please select a file first.')); }
            const reader = new FileReader();
            reader.onload = e => resolve(e.target.result);
            reader.onerror = e => reject(new Error('Error reading file: ' + e.target.error.name));
            reader.readAsArrayBuffer(file);
        });
    }

    function makeTableSelectable(tableId) {
        const table = document.getElementById(tableId);
        if (!table) return;
        const tableContainer = table.closest('.table-container');
        if (!tableContainer) return;
        let isSelecting = false, startCell = null, endCell = null, scrollInterval = null;
        const cells = table.querySelectorAll('td, th');
        tableContainer.addEventListener('mousedown', (e) => {
            const targetCell = e.target.closest('td, th');
            if (!targetCell) return;
            e.preventDefault();
            isSelecting = true;
            startCell = targetCell;
            endCell = targetCell;
            cells.forEach(c => c.classList.remove('cell-selected'));
            updateSelection();
        });
        tableContainer.addEventListener('mousemove', (e) => {
            if (!isSelecting) return;
            const targetCell = e.target.closest('td, th');
            if (targetCell) { endCell = targetCell; }
            updateSelection();
            handleAutoScroll(e);
        });
        document.addEventListener('mouseup', () => {
            isSelecting = false;
            clearInterval(scrollInterval);
            scrollInterval = null;
        });
        tableContainer.addEventListener('mouseleave', () => {
            clearInterval(scrollInterval);
            scrollInterval = null;
        });
        function handleAutoScroll(e) {
            clearInterval(scrollInterval);
            scrollInterval = null;
            const containerRect = tableContainer.getBoundingClientRect();
            const hotZoneWidth = 50, scrollSpeed = 15;
            if (e.clientX > containerRect.right - hotZoneWidth && e.clientX < containerRect.right) {
                scrollInterval = setInterval(() => { tableContainer.scrollLeft += scrollSpeed; }, 24);
            } else if (e.clientX < containerRect.left + hotZoneWidth && e.clientX > containerRect.left) {
                scrollInterval = setInterval(() => { tableContainer.scrollLeft -= scrollSpeed; }, 24);
            }
        }
        function updateSelection() {
            cells.forEach(cell => cell.classList.remove('cell-selected'));
            if (!startCell || !endCell) return;
            const startRow = startCell.parentElement.rowIndex, endRow = endCell.parentElement.rowIndex;
            const startCol = startCell.cellIndex, endCol = endCell.cellIndex;
            const minRow = Math.min(startRow, endRow), maxRow = Math.max(startRow, endRow);
            const minCol = Math.min(startCol, endCol), maxCol = Math.max(startCol, endCol);
            for (let i = minRow; i <= maxRow; i++) {
                for (let j = minCol; j <= maxCol; j++) {
                    const row = table.rows[i];
                    if (row && row.cells[j]) { row.cells[j].classList.add('cell-selected'); }
                }
            }
        }
    }

    document.addEventListener('copy', (e) => {
        const selectedCells = document.querySelectorAll('.cell-selected');
        if (selectedCells.length === 0) return;
        e.preventDefault();
        const rows = new Map();
        selectedCells.forEach(cell => {
            const row = cell.parentElement;
            if (!rows.has(row)) { rows.set(row, []); }
            rows.get(row).push(cell);
        });
        const sortedRows = Array.from(rows.keys()).sort((a, b) => a.rowIndex - b.rowIndex);
        let tsv = sortedRows.map(row => {
            const cellsInRow = rows.get(row).sort((a, b) => a.cellIndex - b.cellIndex);
            return cellsInRow.map(cell => cell.innerText.replace(/\t/g, " ").replace(/\n/g, " ")).join('\t');
        }).join('\n');
        if (e.clipboardData) {
            e.clipboardData.setData('text/plain', tsv);
            showToast('Selected cells copied!');
        } else {
            showToast('Failed to copy. Your browser may be too old.');
        }
    });

    // --- STANDARDIZED FORMATTING HELPERS ---
    function formatDate(dateValue) {
        if (!dateValue && dateValue !== 0) return '';
        let date;
        if (typeof dateValue === 'number' && dateValue > 10000) {
            date = new Date(Math.round((dateValue - 25569) * 86400 * 1000));
        } else {
            date = new Date(dateValue);
        }
        if (isNaN(date.getTime())) { return String(dateValue); }
        const timezoneOffset = date.getTimezoneOffset() * 60 * 1000;
        date = new Date(date.getTime() + timezoneOffset);
        return `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`;
    }

    function formatNumber(num) {
        const val = Number(String(num).replace(/,/g, ''));
        return isNaN(val) ? '0' : val.toLocaleString('en-US');
    }

    function formatCurrency(num, precision = 2) {
        const val = parseFloat(String(num).replace(/[$,]/g, ''));
        return isNaN(val) ? (0).toFixed(precision) : val.toFixed(precision);
    }

    // --- CORE PROCESSING LOGIC ---
    async function processFile(inputId, toolName, processor) {
        showLoading();
        const file = document.getElementById(inputId).files[0];
        try {
            if (!file) throw new Error('Please select a file first.');
            const buffer = await readFileAsArrayBuffer(file);
            const workbook = XLSX.read(buffer, { type: 'array' });
            await processor(workbook, file.name);
        } catch (error) {
            renderOutput(toolName, null, null, { error: error.message, fileName: file ? file.name : '' });
        }
    }
    
    /**
     * Reusable helper function to find the most likely data sheet in a workbook.
     * @param {Object} workbook - The XLSX workbook object.
     * @param {string[]} keywords - An array of keywords to look for in sheet names, in order of priority.
     * @returns {string|null} The name of the found sheet, or null if no sheets exist.
     */
    function findDataSheet(workbook, keywords = []) {
        const sheetNames = workbook.SheetNames;
        // 1. Look for a sheet matching keywords in the order they are provided.
        for (const keyword of keywords) {
            const foundSheet = sheetNames.find(name => name.toLowerCase().trim().includes(keyword));
            if (foundSheet) return foundSheet;
        }
        // 2. If no keyword match, find the first sheet that is NOT a "Properties" or metadata tab.
        const nonMetadataSheet = sheetNames.find(name => !name.toLowerCase().includes('properties'));
        if (nonMetadataSheet) return nonMetadataSheet;
        
        // 3. As a last resort, if there are any sheets at all, return the very first one.
        return sheetNames.length > 0 ? sheetNames[0] : null;
    }


    // ------------------------------------------------------------------
    // FILE 1: AdX and Private Auction (processTool1)
    // Handles AdX & PMP Daily Report.
    // ------------------------------------------------------------------
    async function processTool1(workbook, fileName) {
        const toolName = 'AdX and Private Auction';
        try {
            const sheetName = findDataSheet(workbook, ['report data']);
            if (!sheetName) throw new Error(`Could not find a sheet named 'report data'.`);
            
            const json = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
            if (!json || json.length === 0) throw new Error("The data sheet is empty.");

            // Find column names dynamically and case-insensitively
            const originalHeaders = Object.keys(json[0]);
            const dateCol = originalHeaders.find(h => h.toLowerCase().includes('date'));
            const channelCol = originalHeaders.find(h => h.toLowerCase().includes('programmatic channel'));
            const impCol = originalHeaders.find(h => h.toLowerCase().includes('ad exchange impressions'));
            const revCol = originalHeaders.find(h => h.toLowerCase().includes('ad exchange revenue'));
            
            if (!dateCol || !channelCol || !impCol || !revCol) {
                throw new Error("Missing required columns. The file must contain columns for Date, Programmatic Channel, Impressions, and Revenue.");
            }

            const grouped = {};
            json.forEach(row => {
                const dateRaw = formatDate(row[dateCol]);
                const channel = (row[channelCol] || '').trim();
                if (!dateRaw || channel.toLowerCase() === 'total' || !channel) return;
                
                const imps = parseInt(row[impCol]);
                const rev = parseFloat(String(row[revCol]).replace(/[$,]/g, ''));
                if (isNaN(imps) || isNaN(rev)) return;

                if (!grouped[dateRaw]) grouped[dateRaw] = {};
                if (channel === 'Open Auction') {
                    grouped[dateRaw].adxImps = (grouped[dateRaw].adxImps || 0) + imps;
                    grouped[dateRaw].adxRev = (grouped[dateRaw].adxRev || 0) + rev;
                }
                if (channel === 'Private Auction') {
                    grouped[dateRaw].pmpImps = (grouped[dateRaw].pmpImps || 0) + imps;
                    grouped[dateRaw].pmpRev = (grouped[dateRaw].pmpRev || 0) + rev;
                }
            });

            const result = Object.entries(grouped).sort((a, b) => new Date(a[0]) - new Date(b[0])).map(([date, data]) => {
                const adxImps = data.adxImps || 0, adxRev = data.adxRev || 0;
                const pmpImps = data.pmpImps || 0, pmpRev = data.pmpRev || 0;
                return { Date: date, 'AdX Paid Imps': adxImps, 'AdX NET Revenue': adxRev, 'AdX eCPM': adxImps ? (adxRev / adxImps) * 1000 : 0, 'PMP Paid Imps': pmpImps, 'PMP NET Revenue': pmpRev, 'PMP eCPM': pmpImps ? (pmpRev / pmpImps) * 1000 : 0 };
            });

            if (result.length === 0) throw new Error('No valid data rows could be processed.');
            
            let html = `<thead><tr><th rowspan="2">Date</th><th colspan="3" class="centered">AdX</th><th colspan="3" class="centered">Private Auction (PMP)</th></tr><tr><th class="numeric">Paid Imps</th><th class="numeric">NET Revenue</th><th class="numeric">eCPM</th><th class="numeric">PMP Paid Imps</th><th class="numeric">PMP NET Revenue</th><th class="numeric">PMP eCPM</th></tr></thead><tbody>`;
            result.forEach(r => {
                html += `<tr><td>${r.Date}</td><td class="numeric">${formatNumber(r['AdX Paid Imps'])}</td><td class="numeric">$${formatCurrency(r['AdX NET Revenue'])}</td><td class="numeric">$${formatCurrency(r['AdX eCPM'])}</td><td class="numeric">${formatNumber(r['PMP Paid Imps'])}</td><td class="numeric">$${formatCurrency(r['PMP NET Revenue'])}</td><td class="numeric">$${formatCurrency(r['PMP eCPM'])}</td></tr>`;
            });
            renderOutput(toolName, `<table>${html}</tbody></table>`, 'outputTable1', { fileName, tableData: result });
        } catch (error) {
            renderOutput(toolName, null, null, { error: error.message, fileName });
        }
    }


    // ------------------------------------------------------------------
    // FILE 2: Open Bidding Partners (processBiddingData)
    // Handles Open Bidding reports with dynamic sheet and column names.
    // ------------------------------------------------------------------
    async function processBiddingData(workbook, fileName) {
        const toolName = "Open Bidding Partners";
        try {
            const sheetName = findDataSheet(workbook, ['open bidding', 'report data']);
            if (!sheetName) throw new Error(`Could not find a valid data sheet (e.g., '...Open Bidding' or 'Report data').`);

            const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "", raw: false, dateNF: 'm/d/yyyy' });
            if (!data || data.length < 1) throw new Error("Found data sheet is empty.");

            const PARTNER_ORDER = ['Sovrn OB', 'Rubicon OB', 'Media.net OB', 'PubMatic OB', 'TripleLift', 'Sharethrough OB', 'OneTag OB'];
            const PARTNER_MAPPING = { 'sov_eb': 'Sovrn OB', 'sovrn': 'Sovrn OB', 'rp eb': 'Rubicon OB', 'rubicon': 'Rubicon OB', 'media.net': 'Media.net OB', 'pubmatic': 'PubMatic OB', 'pubmatic (bidder)': 'PubMatic OB', 'triplelift': 'TripleLift', 'triplelift (bidder)': 'TripleLift', 'sharethrough': 'Sharethrough OB', 'onetag': 'OneTag OB' };
            
            const headers = data[0].map(h => String(h).trim().toLowerCase());
            
            const dateIdx = headers.findIndex(h => h.includes('date'));
            const nameIdx = headers.findIndex(h => h.includes('yield partner'));
            const impIdx = headers.findIndex(h => h.includes('impression'));
            const revIdx = headers.findIndex(h => h.includes('revenue'));
            
            if ([dateIdx, nameIdx, impIdx, revIdx].includes(-1)) {
                throw new Error(`Could not find required columns (Date, Yield Partner, Impressions, or Revenue).`);
            }

            const pivotedData = new Map();
            for (let i = 1; i < data.length; i++) {
                const row = data[i];
                if (!row[nameIdx]) continue;
                const canonicalPartner = PARTNER_MAPPING[String(row[nameIdx]).toLowerCase().trim()];
                if (!canonicalPartner) continue;
                const dateStr = formatDate(row[dateIdx]);
                if (!pivotedData.has(dateStr)) pivotedData.set(dateStr, {});
                const dateEntry = pivotedData.get(dateStr);
                dateEntry[canonicalPartner] = {
                    imps: Number(row[impIdx]) || 0,
                    rev: parseFloat(String(row[revIdx]).replace(/[^0-9.-]+/g, '')) || 0,
                };
            }

            let html = `<thead><tr><th rowspan="2">Date</th>${PARTNER_ORDER.map(col => `<th colspan="3" class="centered">${col}</th>`).join('')}</tr><tr>${PARTNER_ORDER.map(() => `<th class="numeric">Paid Imps</th><th class="numeric">NET Revenue</th><th class="numeric">eCPM</th>`).join('')}</tr></thead><tbody>`;
            const sortedDates = [...pivotedData.keys()].sort((a, b) => new Date(a) - new Date(b));
            const tableData = [];
            for (const date of sortedDates) {
                const dateData = pivotedData.get(date);
                let rowHtml = `<tr><td>${date}</td>`;
                let rowData = { Date: date };
                PARTNER_ORDER.forEach(colName => {
                    const d = dateData[colName] || { imps: 0, rev: 0 };
                    const ecpm = d.imps > 0 ? (d.rev / d.imps) * 1000 : 0;
                    rowData[`${colName} Paid Imps`] = d.imps;
                    rowData[`${colName} NET Revenue`] = d.rev;
                    rowData[`${colName} eCPM`] = ecpm;
                    rowHtml += `<td class="numeric">${formatNumber(d.imps)}</td><td class="numeric">$${formatCurrency(d.rev)}</td><td class="numeric">$${formatCurrency(ecpm)}</td>`;
                });
                rowHtml += '</tr>';
                html += rowHtml;
                tableData.push(rowData);
            }
            renderOutput('Open Bidding Horizontal Report', `<table>${html}</tbody></table>`, 'biddingResultTable', { fileName, tableData });
        } catch (error) {
            renderOutput(toolName, null, null, { error: error.message, fileName });
        }
    }


    // ------------------------------------------------------------------
    // FILE 3: Rubicon/Magnite (processTool2)
    // Handles Rubicon/Magnite reports.
    // ------------------------------------------------------------------
    async function processTool2(workbook, fileName) {
        const toolName = 'Rubicon/Magnite';
        try {
            const sheetName = findDataSheet(workbook, ["report"]);
            if (!sheetName) throw new Error('Could not find a sheet containing "report".');

            let jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
            if (!jsonData.length) throw new Error('The report sheet is empty.');
            
            const originalHeaders = Object.keys(jsonData[0]);
            const dateCol = originalHeaders.find(h => h.toLowerCase().includes("date"));
            if (!dateCol) throw new Error("Could not find a 'Date' column.");

            jsonData.forEach(row => row.jsDate = new Date(formatDate(row[dateCol])));
            jsonData.sort((a, b) => a.jsDate - b.jsDate);

            let html = `<thead><tr>${originalHeaders.map(h => `<th${typeof jsonData[0][h] === 'number' ? ' class="numeric"' : ''}>${h}</th>`).join('')}</tr></thead><tbody>`;
            jsonData.forEach(row => {
                html += '<tr>';
                originalHeaders.forEach(header => {
                    const val = row[header];
                    const isNumeric = typeof val === 'number' || (header !== dateCol && !isNaN(parseFloat(String(val).replace(/,/g, ''))));
                    html += `<td class="${isNumeric ? 'numeric' : ''}">`;
                    if (header === dateCol) {
                        html += formatDate(val);
                    } else if (isNumeric) {
                        html += formatNumber(val);
                    } else {
                        html += val !== undefined ? val : '';
                    }
                    html += `</td>`;
                });
                html += '</tr>';
            });
            const tableData = jsonData.map(row => {
                const newRow = { ...row };
                delete newRow.jsDate;
                return newRow;
            });
            renderOutput(toolName, `<table>${html}</tbody></table>`, 'outputTable2', { fileName, tableData });
        } catch(error) {
            renderOutput(toolName, null, null, { error: error.message, fileName });
        }
    }


    // ------------------------------------------------------------------
    // GENERIC PIVOT HELPER
    // A reusable function for Epsilon, Pubmatic, and other similar pivot-style reports.
    // ------------------------------------------------------------------
    // ------------------------------------------------------------------
// FILE 3.5: Generic Pivot Helper (processGenericPivot)
// This is a special helper function used by Pubmatic, Epsilon, etc.
// The bug that caused the "'reven' not found" error has been fixed here.
// ------------------------------------------------------------------
function processGenericPivot(workbook, fileName, config) {
    const sheetName = findDataSheet(workbook);
    if (!sheetName) throw new Error("Could not find a valid data sheet.");

    const raw = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "" });
    if (!raw.length) throw new Error("Sheet is empty.");

    let headerRowIndex = -1;
    let headers = [];
    // Find the first row that contains a cell with our revenue keyword.
    for (let i = 0; i < raw.length; i++) {
        const potentialHeaders = raw[i].map(c => String(c).toLowerCase());
        
        // MODIFICATION: This is the critical bug fix. 
        // Changed from .includes() to .some(c => c.includes())
        if (potentialHeaders.some(c => c.includes(config.revenueKeyword))) {
            headerRowIndex = i;
            headers = potentialHeaders;
            break;
        }
    }

    if (headerRowIndex === -1) throw new Error(`Could not find a header row containing '${config.revenueKeyword}'.`);
    
    const colMap = {
        d: headers.findIndex(h => h.includes('date')),
        i: headers.findIndex(h => h.includes(config.impressionsKeyword)),
        r: headers.findIndex(h => h.includes(config.revenueKeyword))
    };
    if (Object.values(colMap).some(i => i === -1)) throw new Error(`Required columns are missing. Check for Date, ${config.impressionsKeyword}, and ${config.revenueKeyword}.`);

    const pivot = {};
    raw.slice(headerRowIndex + 1).forEach(row => {
        if (row.length === 0 || row.every(cell => cell === "")) return; // Skip empty rows
        const date = formatDate(row[colMap.d]);
        if (!date) return;
        const p = pivot[date] || { i: 0, r: 0 };
        p.i += parseInt(String(row[colMap.i] || 0).replace(/,/g, '')) || 0;
        p.r += parseFloat(String(row[colMap.r] || 0).replace(/[$,]/g, '')) || 0;
        pivot[date] = p;
    });

    const rows = Object.entries(pivot).map(([date, data]) => ({
        'Date': date,
        [config.impressionsHeader]: data.i,
        [config.revenueHeader]: data.r
    })).sort((a, b) => new Date(a.Date) - new Date(b.Date));
    
    if (rows.length === 0) throw new Error("No valid data rows could be processed.");

    const totals = rows.reduce((acc, r) => ({
        i: acc.i + r[config.impressionsHeader],
        r: acc.r + r[config.revenueHeader]
    }), { i: 0, r: 0 });

    let html = `<thead><tr><th>Date</th><th class="numeric">${config.impressionsHeader}</th><th class="numeric">${config.revenueHeader}</th></tr></thead><tbody>`;
    rows.forEach(r => html += `<tr><td>${r.Date}</td><td class="numeric">${formatNumber(r[config.impressionsHeader])}</td><td class="numeric">$${formatCurrency(r[config.revenueHeader], config.precision)}</td></tr>`);
    html += `<tr style="font-weight:bold; background:var(--bg-light);"><td>Grand Total</td><td class="numeric">${formatNumber(totals.i)}</td><td class="numeric">$${formatCurrency(totals.r, config.precision)}</td></tr></tbody>`;
    
    renderOutput(config.title, `<table>${html}</table>`, config.tableId, {
        fileName,
        tableData: [...rows, { 'Date': 'Grand Total', [config.impressionsHeader]: totals.i, [config.revenueHeader]: totals.r }]
    });
}

// ------------------------------------------------------------------
// FILE 4: Pubmatic Openwrap (processCitrusAdPivotData)
// This function uses the generic pivot helper. The keywords have been
// made more flexible to handle both old and new file formats.
// ------------------------------------------------------------------
async function processCitrusAdPivotData(workbook, fileName) {
    const toolName = 'Pubmatic Openwrap';
    try {
        // MODIFICATION: Keywords are now shorter to match all file variations.
        processGenericPivot(workbook, fileName, {
            title: "Pubmatic Openwrap Summary",
            tableId: "citrusPivotResultTable",
            impressionsKeyword: 'impres', // Catches "Paid Impressions" and "Paid Impres"
            revenueKeyword: 'reven',     // Catches "Net Revenue($)" and "Net Reven"
            impressionsHeader: 'Sum of Paid Impressions',
            revenueHeader: 'Sum of Net Revenue($)',
            precision: 4
        });
    } catch(error) {
        renderOutput(toolName, null, null, { error: error.message, fileName });
    }
}


    // ------------------------------------------------------------------
    // FILE 5: Verizon/Yahoo (processTool3)
    // Handles Verizon/Yahoo reports.
    // ------------------------------------------------------------------
    async function processTool3(workbook, fileName) {
        const toolName = 'Verizon/Yahoo';
        try {
            const sheetName = findDataSheet(workbook, ["ad requests", "ads sent"]);
            if (!sheetName) throw new Error(`Could not find a sheet containing "Ad Requests" or "Ads Sent".`);

            const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "", raw: true, cellDates: true });
            if (!jsonData || jsonData.length === 0) throw new Error("The data sheet is empty.");
            
            const originalHeaders = Object.keys(jsonData[0]);
            const dateCol = originalHeaders.find(h => h.toLowerCase().trim().includes("date"));
            const impCol = originalHeaders.find(h => h.toLowerCase().trim().includes("impression"));
            const revCol = originalHeaders.find(h => h.toLowerCase().trim().includes("net revenue"));

            if (!dateCol || !impCol || !revCol) {
                throw new Error("Missing required columns. The file must contain columns for Date, Impression, and Net Revenue.");
            }

            const sorted = jsonData.map(row => ({
                Date: formatDate(row[dateCol]),
                Impressions: row[impCol],
                'Net Revenue': row[revCol]
            })).filter(r => r.Date && r.Impressions !== "" && r['Net Revenue'] !== "").sort((a, b) => new Date(a.Date) - new Date(b.Date));
            
            if (!sorted.length) throw new Error("No valid data rows could be processed.");

            let html = "<thead><tr><th>Date</th><th class='numeric'>Impressions</th><th class='numeric'>Net Revenue</th></tr></thead><tbody>";
            sorted.forEach(row => html += `<tr><td>${row.Date}</td><td class='numeric'>${formatNumber(row.Impressions)}</td><td class='numeric'>$${formatCurrency(row['Net Revenue'])}</td></tr>`);
            renderOutput(toolName, `<table>${html}</tbody></table>`, 'outputTable3', { fileName, tableData: sorted });
        } catch(error) {
            renderOutput(toolName, null, null, { error: error.message, fileName });
        }
    }


    // ------------------------------------------------------------------
    // FILE 6: Xandr/AppNexus (processTool4)
    // Handles Xandr/AppNexus reports.
    // ------------------------------------------------------------------
    async function processTool4(workbook, fileName) {
        const toolName = 'Xandr/AppNexus';
        try {
            const sheetName = findDataSheet(workbook, ["report"]);
            if (!sheetName) throw new Error("Could not find a sheet containing 'report'.");
            
            const json = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
            if (!json.length) throw new Error("The data sheet is empty.");

            const originalHeaders = Object.keys(json[0]);
            const dayCol = originalHeaders.find(h => h.toLowerCase() === 'day' || h.toLowerCase() === 'date');
            const impsCol = originalHeaders.find(h => h.toLowerCase() === 'imps' || h.toLowerCase() === 'impressions');
            const revCol = originalHeaders.find(h => h.toLowerCase() === 'revenue' || h.toLowerCase() === 'earnings');

            if (!dayCol || !impsCol || !revCol) throw new Error("Missing required columns (day/date, imps/impressions, or revenue/earnings).");

            const filtered = json.map(row => ({
                Day: formatDate(row[dayCol]),
                Imps: row[impsCol] || 0,
                Revenue: row[revCol] || 0
            })).sort((a, b) => new Date(a.Day) - new Date(b.Day));

            let html = "<thead><tr><th>Day</th><th class='numeric'>Imps</th><th class='numeric'>Revenue</th></tr></thead><tbody>";
            filtered.forEach(row => html += `<tr><td>${row.Day}</td><td class='numeric'>${formatNumber(row.Imps)}</td><td class='numeric'>$${formatCurrency(row.Revenue)}</td></tr>`);
            renderOutput(toolName, `<table>${html}</tbody></table>`, "outputTable4", { fileName, tableData: filtered });
        } catch(error) {
            renderOutput(toolName, null, null, { error: error.message, fileName });
        }
    }


    // ------------------------------------------------------------------
    // FILE 7: TripleLift (processTool5)
    // Handles TripleLift reports.
    // ------------------------------------------------------------------
    async function processTool5(workbook, fileName) {
        const toolName = 'TripleLift';
        try {
            const sheetName = findDataSheet(workbook);
            if (!sheetName) throw new Error("No valid data sheet found.");

            const json = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "", cellDates: true });
            if (!json.length) throw new Error("The data sheet is empty.");

            const originalHeaders = Object.keys(json[0]);
            const ymdCol = originalHeaders.find(h => h.toLowerCase().includes('ymd') || h.toLowerCase().includes('date'));
            const revCol = originalHeaders.find(h => h.toLowerCase().includes('revenue'));
            const renderedCol = originalHeaders.find(h => h.toLowerCase().includes('rendered'));

            if (!ymdCol || !revCol || !renderedCol) {
                throw new Error("Sheet must contain columns for Date (YMD), Revenue, and Rendered impressions.");
            }
            
            const tableData = json.map(row => ({
                YMD: formatDate(row[ymdCol]),
                RENDERED: row[renderedCol],
                REVENUE: row[revCol]
            })).sort((a,b) => new Date(a.YMD) - new Date(b.YMD));

            let html = "<thead><tr><th>YMD</th><th class='numeric'>RENDERED</th><th class='numeric'>REVENUE</th></tr></thead><tbody>";
            tableData.forEach(row => html += `<tr><td>${row.YMD}</td><td class='numeric'>${formatNumber(row.RENDERED)}</td><td class='numeric'>$${formatCurrency(row.REVENUE)}</td></tr>`);
            renderOutput(toolName, `<table>${html}</tbody></table>`, "outputTable5", { fileName, tableData });
        } catch(error) {
            renderOutput(toolName, null, null, { error: error.message, fileName });
        }
    }


    // ------------------------------------------------------------------
    // FILE 8: Cadent Live (processTool6)
    // Handles Cadent Live reports.
    // ------------------------------------------------------------------
    async function processTool6(workbook, fileName) {
        const toolName = 'Cadent Live';
        try {
            const sheetName = findDataSheet(workbook);
            if (!sheetName) throw new Error("No valid data sheet found.");

            const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { raw: false });
            if (!data || data.length === 0) throw new Error("The data sheet is empty.");

            const originalHeaders = Object.keys(data[0]);
            const dateCol = originalHeaders.find(h => h.toLowerCase().includes('date'));
            const renderCol = originalHeaders.find(h => h.toLowerCase().includes('renders'));
            const revCol = originalHeaders.find(h => h.toLowerCase().includes('publisher revenue'));
            
            if (!dateCol || !renderCol || !revCol) {
                 throw new Error("Could not find required columns (date, Ad Renders, Publisher Revenue).");
            }

            const sorted = data.map(r => ({
                Date: formatDate(new Date(r[dateCol])),
                'Ad Renders': Number(r[renderCol]),
                'Publisher Revenue': Number(r[revCol])
            })).sort((a, b) => new Date(a.Date) - new Date(b.Date));
            
            let html = "<thead><tr><th>Date</th><th class='numeric'>Ad Renders</th><th class='numeric'>Publisher Revenue</th></tr></thead><tbody>";
            sorted.forEach(row => html += `<tr><td>${row.Date}</td><td class='numeric'>${formatNumber(row['Ad Renders'])}</td><td class='numeric'>$${formatCurrency(row['Publisher Revenue'], 6)}</td></tr>`);
            renderOutput(toolName, `<table>${html}</tbody></table>`, "outputTable6", { fileName, tableData: sorted });
        } catch(error) {
            renderOutput(toolName, null, null, { error: error.message, fileName });
        }
    }


    // ------------------------------------------------------------------
    // FILE 9: EPSILON Live (processPivotData)
    // Uses the generic pivot helper.
    // ------------------------------------------------------------------
    async function processPivotData(workbook, fileName) {
        const toolName = 'EPSILON Live';
        try {
            processGenericPivot(workbook, fileName, {
                title: "EPSILON Live Pivot Summary",
                tableId: "pivotResultTable",
                impressionsKeyword: 'impressions',
                revenueKeyword: 'earning',
                impressionsHeader: 'Sum of Impressions',
                revenueHeader: 'Sum of Earnings USD',
                precision: 2
            });
        } catch(error) {
            renderOutput(toolName, null, null, { error: error.message, fileName });
        }
    }


    // ------------------------------------------------------------------
    // FILE 10: Media.net (HB) (processTool7)
    // Handles Media.net reports.
    // ------------------------------------------------------------------
    async function processTool7(workbook, fileName) {
        const toolName = 'Media.net (HB)';
        try {
            const sheetName = findDataSheet(workbook);
            if (!sheetName) throw new Error("No valid data sheet found.");

            const json = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
            if (!json || json.length === 0) throw new Error("The data sheet is empty.");
            
            const originalHeaders = Object.keys(json[0]);
            const dateCol = originalHeaders.find(h => h.toLowerCase().includes('date'));
            const revCol = originalHeaders.find(h => h.toLowerCase().includes('revenue'));
            const impCol = originalHeaders.find(h => h.toLowerCase().includes('impr')); // 'impr' catches 'Impr.'

            if (!dateCol || !revCol || !impCol) {
                throw new Error("Could not find required columns (Date, Revenue, and Ad Impr.).");
            }
            const tableData = json.map(r => ({
                Date: formatDate(r[dateCol]),
                'Ad Impr.': r[impCol],
                Revenue: r[revCol]
            })).sort((a,b) => new Date(a.Date) - new Date(b.Date));
            
            let html = "<thead><tr><th>Date</th><th class='numeric'>Ad Impr.</th><th class='numeric'>Revenue</th></tr></thead><tbody>";
            tableData.forEach(row => html += `<tr><td>${row.Date}</td><td class='numeric'>${formatNumber(row['Ad Impr.'])}</td><td class='numeric'>$${formatCurrency(row['Revenue'])}</td></tr>`);
            renderOutput(toolName, `<table>${html}</tbody></table>`, "outputTable7", { fileName, tableData });
        } catch(error) {
            renderOutput(toolName, null, null, { error: error.message, fileName });
        }
    }


    // ------------------------------------------------------------------
    // FILE 11: Sharethrough (HB) (processTool8)
    // Handles Sharethrough reports.
    // ------------------------------------------------------------------
    async function processTool8(workbook, fileName) {
        const toolName = 'Sharethrough (HB)';
        try {
            const sheetName = findDataSheet(workbook, ["in"]);
            if (!sheetName) throw new Error("No valid data sheet found.");

            const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
            if (!jsonData.length) throw new Error("The data sheet is empty.");
            
            const validRows = jsonData.map(row => {
                const lowerCaseRow = {};
                for (const key in row) {
                    lowerCaseRow[key.toLowerCase()] = row[key];
                }
                return {
                    Date: formatDate(lowerCaseRow["date"]),
                    Rendered: lowerCaseRow["rendered impressions"] || lowerCaseRow["rendered"] || lowerCaseRow["impressions"],
                    Earnings: lowerCaseRow["earnings"] || lowerCaseRow["revenue"] || lowerCaseRow["net revenue"]
                };
            }).filter(r => r.Date && r.Rendered !== undefined && r.Earnings !== undefined)
              .sort((a,b) => new Date(a.Date) - new Date(b.Date));

            if (!validRows.length) throw new Error("Could not find the required columns (Date, Rendered/Impressions, Earnings/Revenue).");

            let html = "<thead><tr><th>Date</th><th class='numeric'>Rendered</th><th class='numeric'>Earnings</th></tr></thead><tbody>";
            validRows.forEach(row => html += `<tr><td>${row.Date}</td><td class='numeric'>${formatNumber(row.Rendered)}</td><td class='numeric'>$${formatCurrency(row.Earnings)}</td></tr>`);
            renderOutput(toolName, `<table>${html}</tbody></table>`, "outputTable8", { fileName, tableData: validRows });
        } catch(error) {
            renderOutput(toolName, null, null, { error: error.message, fileName });
        }
    }


    // ------------------------------------------------------------------
    // FILE 12: Programmatic Avails (processImpressionsData)
    // Handles the two-file Programmatic Avails analysis.
    // ------------------------------------------------------------------
    async function processImpressionsData() {
        const toolName = 'Programmatic Avails';
        let file1, file2;
        try {
            file1 = document.getElementById('impressionsFile1').files[0];
            file2 = document.getElementById('impressionsFile2').files[0];

            if (!file1 || !file2) throw new Error('Please upload both Excel files.');
            const [buffer1, buffer2] = await Promise.all([readFileAsArrayBuffer(file1), readFileAsArrayBuffer(file2)]);
            const wb1 = XLSX.read(buffer1, { type: 'array' });
            const wb2 = XLSX.read(buffer2, { type: 'array' });

            const sheetName1 = findDataSheet(wb1, ['total avails']);
            if (!sheetName1) throw new Error(`Could not find a valid data sheet in "${file1.name}".`);
            const data1 = XLSX.utils.sheet_to_json(wb1.Sheets[sheetName1], { header: 1 });
            
            const sheetName2 = findDataSheet(wb2, ['paid imps']);
            if (!sheetName2) throw new Error(`Could not find a valid data sheet in "${file2.name}".`);
            const data2 = XLSX.utils.sheet_to_json(wb2.Sheets[sheetName2], { header: 1 });

            const h1 = data1[0].map(h => String(h).trim().toLowerCase());
            const h2 = data2[0].map(h => String(h).trim().toLowerCase());

            const dIdx1 = h1.indexOf("date"),
                  uIdx = h1.indexOf("unfilled impressions"),
                  fIdx = h1.indexOf("total impressions");
            if ([dIdx1, uIdx, fIdx].includes(-1)) {
                throw new Error(`The 'Total Avails' file is missing required columns. It must contain 'Date', 'Unfilled Impressions', and 'Total Impressions'.`);
            }

            const dIdx2 = h2.indexOf("date"),
                  pIdx = h2.indexOf("total impressions");
            if ([dIdx2, pIdx].includes(-1)) {
                throw new Error(`The 'Paid Impressions' file is missing required columns. It must contain 'Date' and 'Total Impressions'.`);
            }

            const map1 = new Map(), map2 = new Map();
            data1.slice(1).forEach(r => {
                const d = formatDate(r[dIdx1]);
                if (d) {
                    const v = map1.get(d) || { u: 0, f: 0 };
                    v.u += +r[uIdx] || 0;
                    v.f += +r[fIdx] || 0;
                    map1.set(d, v);
                }
            });
            data2.slice(1).forEach(r => {
                const d = formatDate(r[dIdx2]);
                if (d) map2.set(d, (map2.get(d) || 0) + (+r[pIdx] || 0));
            });

            const allDates = [...new Set([...map1.keys(), ...map2.keys()])].sort((a, b) => new Date(a) - new Date(b));
            let html = `<thead><tr><th>Date</th><th class="numeric">Unfilled</th><th class="numeric">Filled</th><th class="numeric">Total Avails</th><th class="numeric">Paid</th><th class="numeric diff">Difference</th></tr></thead><tbody>`;
            const tableData = [];
            allDates.forEach(date => {
                const s1 = map1.get(date) || { u: 0, f: 0 };
                const s2 = map2.get(date) || 0;
                const total = s1.u + s1.f;
                const rowData = { Date: date, Unfilled: s1.u, Filled: s1.f, 'Total Avails': total, Paid: s2, Difference: total - s2 };
                tableData.push(rowData);
                html += `<tr><td>${date}</td><td class="numeric">${formatNumber(s1.u)}</td><td class="numeric">${formatNumber(s1.f)}</td><td class="numeric">${formatNumber(total)}</td><td class="numeric">${formatNumber(s2)}</td><td class="numeric diff">${formatNumber(total - s2)}</td></tr>`;
            });
            renderOutput(toolName, `<table>${html}</tbody></table>`, "impressionsResultTable", { fileName: `${file1.name} & ${file2.name}`, tableData });
        } catch (error) {
            renderOutput(toolName, null, null, { error: error.message });
        }
    }

})();
