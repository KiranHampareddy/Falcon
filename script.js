// IIFE to encapsulate code and avoid polluting the global scope
(function() {
    'use strict';

    // --- TOOL CONFIGURATION (NO CHANGE) ---
    const TOOL_CONFIG = {
        'tool1': {
            name: 'AdX and Private Auction',
            processor: processTool1
        },
        'bidding': {
            name: 'Open Bidding Partners',
            processor: processBiddingData
        },
        'tool2': {
            name: 'Rubicon/Magnite',
            processor: processTool2
        },
        'citrusPivot': {
            name: 'Pubmatic Openwrap',
            processor: processCitrusAdPivotData
        },
        'tool3': {
            name: 'Verizon/Yahoo',
            processor: processTool3
        },
        'tool4': {
            name: 'Xandr/AppNexus',
            processor: processTool4
        },
        'tool5': {
            name: 'TripleLift',
            processor: processTool5
        },
        'tool6': {
            name: 'Cadent Live',
            processor: processTool6
        },
        'pivot': {
            name: 'EPSILON Live',
            processor: processPivotData
        },
        'tool7': {
            name: 'Media.net (HB)',
            processor: processTool7
        },
        'tool8': {
            name: 'Sharethrough (HB)',
            processor: processTool8
        },
        'impressions': {
            name: 'Programmatic Avails',
            isMultiFile: true,
            processor: processImpressionsData
        },
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
                zone.addEventListener(eventName, e => {
                    e.preventDefault();
                    e.stopPropagation();
                }, false);
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
                    outputPanel.scrollTo({
                        top: 0,
                        behavior: 'smooth'
                    });

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

    // **REFACTORED**: renderOutput now builds the enhanced header and requires more data
    function renderOutput(title, htmlContent, tableId, options = {}) {
        const {
            error = null, fileName = ''
        } = options;

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
            if (!file) {
                return reject(new Error('Please select a file first.'));
            }
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

        let isSelecting = false;
        let startCell = null;
        let endCell = null;
        let scrollInterval = null;

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
            if (targetCell) {
                endCell = targetCell;
            }
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
            const hotZoneWidth = 50;
            const scrollSpeed = 15;

            if (e.clientX > containerRect.right - hotZoneWidth && e.clientX < containerRect.right) {
                scrollInterval = setInterval(() => {
                    tableContainer.scrollLeft += scrollSpeed;
                }, 24);
            } else if (e.clientX < containerRect.left + hotZoneWidth && e.clientX > containerRect.left) {
                scrollInterval = setInterval(() => {
                    tableContainer.scrollLeft -= scrollSpeed;
                }, 24);
            }
        }

        function updateSelection() {
            cells.forEach(cell => cell.classList.remove('cell-selected'));

            if (!startCell || !endCell) return;

            const startRow = startCell.parentElement.rowIndex;
            const endRow = endCell.parentElement.rowIndex;
            const startCol = startCell.cellIndex;
            const endCol = endCell.cellIndex;

            const minRow = Math.min(startRow, endRow);
            const maxRow = Math.max(startRow, endRow);
            const minCol = Math.min(startCol, endCol);
            const maxCol = Math.max(startCol, endCol);

            for (let i = minRow; i <= maxRow; i++) {
                for (let j = minCol; j <= maxCol; j++) {
                    const row = table.rows[i];
                    if (row && row.cells[j]) {
                        row.cells[j].classList.add('cell-selected');
                    }
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
            if (!rows.has(row)) {
                rows.set(row, []);
            }
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

    // --- STANDARDIZED FORMATTING HELPERS (NO CHANGE) ---
    function formatDate(dateValue) {
        if (!dateValue && dateValue !== 0) return '';
        let date;
        if (typeof dateValue === 'number' && dateValue > 10000) {
            date = new Date(Math.round((dateValue - 25569) * 86400 * 1000));
        } else {
            date = new Date(dateValue);
        }
        if (isNaN(date.getTime())) {
            return String(dateValue);
        }
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
            const workbook = XLSX.read(buffer, {
                type: 'array'
            });
            await processor(workbook, file.name);
        } catch (error) {
            renderOutput(toolName, null, null, {
                error: error.message
            });
        }
    }

    // --- TOOL-SPECIFIC FUNCTIONS (Logically unchanged, but updated to pass data to the new renderOutput) ---
    async function processImpressionsData() {
        showLoading();
        outputPanel.scrollTo({
            top: 0,
            behavior: 'smooth'
        });
        const file1 = document.getElementById('impressionsFile1').files[0];
        const file2 = document.getElementById('impressionsFile2').files[0];
        try {
            if (!file1 || !file2) throw new Error('Please upload both Excel files.');
            const [buffer1, buffer2] = await Promise.all([readFileAsArrayBuffer(file1), readFileAsArrayBuffer(file2)]);
            const wb1 = XLSX.read(buffer1, {
                type: 'array'
            });
            const wb2 = XLSX.read(buffer2, {
                type: 'array'
            });
            const readSheet = (wb, fileName) => {
                if (!wb.SheetNames.includes("Report data")) throw new Error(`Sheet "Report data" not found in ${fileName}.`);
                return XLSX.utils.sheet_to_json(wb.Sheets["Report data"], {
                    header: 1
                });
            };
            const [data1, data2] = [readSheet(wb1, file1.name), readSheet(wb2, file2.name)];
            const h1 = data1[0].map(h => h.toString().trim().toLowerCase());
            const h2 = data2[0].map(h => h.toString().trim().toLowerCase());
            const dIdx1 = h1.indexOf("date"),
                uIdx = h1.indexOf("unfilled impressions"),
                fIdx = h1.indexOf("total impressions");
            const dIdx2 = h2.indexOf("date"),
                pIdx = h2.indexOf("total impressions");
            if ([dIdx1, uIdx, fIdx, dIdx2, pIdx].includes(-1)) throw new Error('Required columns missing.');
            const map1 = new Map(),
                map2 = new Map();
            data1.slice(1).forEach(r => {
                const d = formatDate(r[dIdx1]);
                if (d) {
                    const v = map1.get(d) || {
                        u: 0,
                        f: 0
                    };
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
                const s1 = map1.get(date) || {
                    u: 0,
                    f: 0
                };
                const s2 = map2.get(date) || 0;
                const total = s1.u + s1.f;
                const rowData = {
                    Date: date,
                    Unfilled: s1.u,
                    Filled: s1.f,
                    'Total Avails': total,
                    Paid: s2,
                    Difference: total - s2
                };
                tableData.push(rowData);
                html += `<tr><td>${date}</td><td class="numeric">${formatNumber(s1.u)}</td><td class="numeric">${formatNumber(s1.f)}</td><td class="numeric">${formatNumber(total)}</td><td class="numeric">${formatNumber(s2)}</td><td class="numeric diff">${formatNumber(total - s2)}</td></tr>`;
            });
            renderOutput("Programmatic Avails Analysis", `<table>${html}</tbody></table>`, "impressionsResultTable", {
                fileName: `${file1.name} & ${file2.name}`,
                tableData
            });
        } catch (error) {
            renderOutput("Programmatic Avails", null, null, {
                error: error.message
            });
        }
    }
    async function processBiddingData(workbook, fileName) {
        const toolName = "Open Bidding Partners";
        try {
            const PARTNER_ORDER = ['Sovrn OB', 'Rubicon OB', 'Media.net OB', 'PubMatic OB', 'TripleLift', 'Sharethrough OB', 'OneTag OB'];
            const PARTNER_MAPPING = {'sov_eb':'Sovrn OB', 'sovrn':'Sovrn OB', 'rp eb':'Rubicon OB', 'rubicon':'Rubicon OB', 'media.net':'Media.net OB', 'pubmatic':'PubMatic OB', 'pubmatic (bidder)':'PubMatic OB', 'triplelift':'TripleLift', 'triplelift (bidder)':'TripleLift', 'sharethrough':'Sharethrough OB', 'onetag':'OneTag OB'};
            if (!workbook.SheetNames.includes("Report data")) throw new Error(`Sheet "Report data" not found.`);
            const data = XLSX.utils.sheet_to_json(workbook.Sheets["Report data"], {
                header: 1,
                defval: "",
                raw: false,
                dateNF: 'm/d/yyyy'
            });
            const headers = data[0].map(h => String(h).trim().toLowerCase());
            const dateIdx = headers.findIndex(h => h.includes("date"));
            const nameIdx = headers.findIndex(h => h.includes("yield partner"));
            const impIdx = headers.findIndex(h => h.includes("yield group impression"));
            const revIdx = headers.findIndex(h => h.includes("yield group estimated revenue ($)"));
            if ([dateIdx, nameIdx, impIdx, revIdx].includes(-1)) {
                throw new Error(`Could not find required columns (Date, Yield Partner, etc).`);
            }
            const pivotedData = new Map();
            for (let i = 1; i < data.length; i++) {
                const row = data[i];
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
                let rowData = {
                    Date: date
                };
                PARTNER_ORDER.forEach(colName => {
                    const d = dateData[colName] || {
                        imps: 0,
                        rev: 0
                    };
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
            renderOutput('Open Bidding Horizontal Report', `<table>${html}</tbody></table>`, 'biddingResultTable', {
                fileName,
                tableData
            });
        } catch (error) {
            renderOutput(toolName, null, null, {
                error: error.message,
                fileName
            });
        }
    }

    function processPivotData(workbook, fileName) {
        const raw = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], {
            header: 1,
            defval: ""
        });
        if (!raw.length) throw new Error("Sheet is empty.");
        const headerRowIndex = raw.findIndex(r => r.map(c => String(c).toLowerCase()).some(c => c.includes("earning")));
        if (headerRowIndex === -1) throw new Error("Could not find header row with 'Earnings'.");
        const headers = raw[headerRowIndex].map(h => String(h).toLowerCase());
        const colMap = {
            d: headers.indexOf("date"),
            i: headers.indexOf("impressions"),
            e: headers.findIndex(h => h.includes("earning"))
        };
        if (Object.values(colMap).some(i => i === -1)) throw new Error("Columns 'Date', 'Impressions', or 'Earnings' missing.");
        const pivot = {};
        raw.slice(headerRowIndex + 1).forEach(row => {
            const date = formatDate(row[colMap.d]);
            if (!date) return;
            const p = pivot[date] || {
                i: 0,
                e: 0
            };
            p.i += parseInt(String(row[colMap.i] || 0).replace(/,/g, '')) || 0;
            p.e += parseFloat(String(row[colMap.e] || 0).replace(/[$,]/g, '')) || 0;
            pivot[date] = p;
        });
        const rows = Object.entries(pivot).map(([date, data]) => ({
            'Row Labels': date,
            'Sum of Impressions': data.i,
            'Sum of Earnings USD': data.e
        })).sort((a, b) => new Date(a['Row Labels']) - new Date(b['Row Labels']));
        const totals = rows.reduce((acc, r) => ({
            i: acc.i + r['Sum of Impressions'],
            e: acc.e + r['Sum of Earnings USD']
        }), {
            i: 0,
            e: 0
        });
        let html = `<thead><tr><th>Row Labels</th><th class="numeric">Sum of Impressions</th><th class="numeric">Sum of Earnings USD</th></tr></thead><tbody>`;
        rows.forEach(r => html += `<tr><td>${r['Row Labels']}</td><td class="numeric">${formatNumber(r['Sum of Impressions'])}</td><td class="numeric">$${formatCurrency(r['Sum of Earnings USD'])}</td></tr>`);
        html += `<tr style="font-weight:bold; background:var(--bg-light);"><td>Grand Total</td><td class="numeric">${formatNumber(totals.i)}</td><td class="numeric">$${formatCurrency(totals.e)}</td></tr></tbody>`;
        renderOutput("EPSILON Live Pivot Summary", `<table>${html}</table>`, "pivotResultTable", {
            fileName,
            tableData: [...rows, {
                'Row Labels': 'Grand Total',
                'Sum of Impressions': totals.i,
                'Sum of Earnings USD': totals.e
            }]
        });
    }

    function processCitrusAdPivotData(workbook, fileName) {
        const raw = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], {
            header: 1,
            defval: ""
        });
        const headerRowIndex = raw.findIndex(r => r.map(c => String(c).toLowerCase()).some(c => c.includes("net revenue")));
        if (headerRowIndex === -1) throw new Error("Required column 'Net Revenue($)' not found.");
        const headers = raw[headerRowIndex].map(h => String(h).toLowerCase());
        const colMap = {
            d: headers.indexOf("date"),
            r: headers.findIndex(h => h.includes("net revenue")),
            i: headers.findIndex(h => h.includes("paid impressions"))
        };
        if (Object.values(colMap).some(idx => idx === -1)) throw new Error("Missing 'Date', 'Paid Impressions', or 'Net Revenue($)'.");
        const pivot = {};
        raw.slice(headerRowIndex + 1).forEach(row => {
            const date = formatDate(row[colMap.d]);
            if (!date) return;
            const p = pivot[date] || {
                r: 0,
                i: 0
            };
            p.r += parseFloat(String(row[colMap.r] || 0).replace(/[$,]/g, '')) || 0;
            p.i += parseInt(String(row[colMap.i] || 0).replace(/,/g, '')) || 0;
            pivot[date] = p;
        });
        const rows = Object.entries(pivot).map(([d, data]) => ({
            Date: d,
            'Sum of Paid Impressions': data.i,
            'Sum of Net Revenue($)': data.r
        })).sort((a, b) => new Date(a.Date) - new Date(b.Date));
        const totals = rows.reduce((acc, r) => ({
            r: acc.r + r['Sum of Net Revenue($)'],
            i: acc.i + r['Sum of Paid Impressions']
        }), {
            r: 0,
            i: 0
        });
        let html = `<thead><tr><th>Date</th><th class="numeric">Sum of Paid Impressions</th><th class="numeric">Sum of Net Revenue($)</th></tr></thead><tbody>`;
        rows.forEach(r => html += `<tr><td>${r.Date}</td><td class="numeric">${formatNumber(r['Sum of Paid Impressions'])}</td><td class="numeric">$${formatCurrency(r['Sum of Net Revenue($)'], 4)}</td></tr>`);
        html += `<tr style="font-weight:bold; background:var(--bg-light);"><td>Grand Total</td><td class="numeric">${formatNumber(totals.i)}</td><td class="numeric">$${formatCurrency(totals.r, 4)}</td></tr></tbody>`;
        renderOutput("Pubmatic Openwrap Summary", `<table>${html}</table>`, "citrusPivotResultTable", {
            fileName,
            tableData: [...rows, {
                Date: 'Grand Total',
                'Sum of Paid Impressions': totals.i,
                'Sum of Net Revenue($)': totals.r
            }]
        });
    }

    function processTool1(workbook, fileName) {
        const sheetName = workbook.SheetNames.find(name => name.trim().toLowerCase() === 'report data');
        if (!sheetName) throw new Error(`Sheet "report data" not found.`);
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
            defval: ""
        });
        const grouped = {};
        rows.forEach(row => {
            const dateRaw = formatDate(row['Date'] || row['date']);
            const channel = (row['Programmatic channel'] || '').trim();
            if (!dateRaw || channel.toLowerCase() === 'total' || !channel) return;
            const imps = parseInt(row['Ad Exchange impressions']);
            const rev = parseFloat(String(row['Ad Exchange revenue ($)']).replace(/[$,]/g, ''));
            if (isNaN(imps) || isNaN(rev)) return;
            if (!grouped[dateRaw]) grouped[dateRaw] = {};
            if (channel === 'Open Auction') {
                grouped[dateRaw].adxImps = imps;
                grouped[dateRaw].adxRev = rev;
            }
            if (channel === 'Private Auction') {
                grouped[dateRaw].pmpImps = imps;
                grouped[dateRaw].pmpRev = rev;
            }
        });
        const result = Object.entries(grouped).sort((a, b) => new Date(a[0]) - new Date(b[0])).map(([date, data]) => {
            const adxImps = data.adxImps || 0,
                adxRev = data.adxRev || 0;
            const pmpImps = data.pmpImps || 0,
                pmpRev = data.pmpRev || 0;
            return {
                Date: date,
                'AdX Paid Imps': adxImps,
                'AdX NET Revenue': adxRev,
                'AdX eCPM': adxImps ? (adxRev / adxImps) * 1000 : 0,
                'PMP Paid Imps': pmpImps,
                'PMP NET Revenue': pmpRev,
                'PMP eCPM': pmpImps ? (pmpRev / pmpImps) * 1000 : 0
            };
        });
        if (result.length === 0) throw new Error('No valid data rows found.');
        let html = `<thead><tr><th rowspan="2">Date</th><th colspan="3" class="centered">AdX</th><th colspan="3" class="centered">Private Auction (PMP)</th></tr><tr><th class="numeric">Paid Imps</th><th class="numeric">NET Revenue</th><th class="numeric">eCPM</th><th class="numeric">PMP Paid Imps</th><th class="numeric">PMP NET Revenue</th><th class="numeric">PMP eCPM</th></tr></thead><tbody>`;
        result.forEach(r => {
            html += `<tr><td>${r.Date}</td><td class="numeric">${formatNumber(r['AdX Paid Imps'])}</td><td class="numeric">$${formatCurrency(r['AdX NET Revenue'])}</td><td class="numeric">$${formatCurrency(r['AdX eCPM'])}</td><td class="numeric">${formatNumber(r['PMP Paid Imps'])}</td><td class="numeric">$${formatCurrency(r['PMP NET Revenue'])}</td><td class="numeric">$${formatCurrency(r['PMP eCPM'])}</td></tr>`;
        });
        renderOutput('AdX & PMP Daily Report', `<table>${html}</tbody></table>`, 'outputTable1', {
            fileName,
            tableData: result
        });
    }

    function processTool2(workbook, fileName) {
        const sheetName = workbook.SheetNames.find(name => name.toLowerCase() === "report");
        if (!sheetName) throw new Error('Sheet "report" not found.');
        let jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
        if (!jsonData.length) throw new Error('"report" sheet is empty.');
        const headers = Object.keys(jsonData[0]);
        const dateCol = headers.find(h => h.toLowerCase().includes("date"));
        if (!dateCol) throw new Error("No 'Date' column found.");
        jsonData.forEach(row => row.jsDate = new Date(formatDate(row[dateCol])));
        jsonData.sort((a, b) => a.jsDate - b.jsDate);
        let html = `<thead><tr>${headers.map(h => `<th${typeof jsonData[0][h] === 'number' ? ' class="numeric"' : ''}>${h}</th>`).join('')}</tr></thead><tbody>`;
        jsonData.forEach(row => {
            html += '<tr>';
            headers.forEach(header => {
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
            const newRow = {
                ...row
            };
            delete newRow.jsDate;
            return newRow;
        });
        renderOutput("Rubicon/Magnite Report", `<table>${html}</tbody></table>`, 'outputTable2', {
            fileName,
            tableData
        });
    }

    function processTool3(workbook, fileName) {
        const sheetName = "Ad Requests, Ads Sent by Date";
        if (!workbook.Sheets[sheetName]) throw new Error(`Sheet "${sheetName}" not found.`);
        const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
            defval: "",
            raw: true,
            cellDates: true
        });
        const sorted = jsonData.map(row => {
            const r = {};
            for (let key in row) {
                const k = key.trim().toLowerCase();
                if (k.includes("date")) r.Date = row[key];
                else if (k.includes("impression")) r.Impressions = row[key];
                else if (k.includes("net revenue")) r['Net Revenue'] = row[key];
            }
            return r;
        }).filter(r => r.Date && r.Impressions !== "" && r['Net Revenue'] !== "").map(r => ({
            ...r,
            Date: formatDate(r.Date)
        })).sort((a, b) => new Date(a.Date) - new Date(b.Date));
        if (!sorted.length) throw new Error("No valid data found.");
        let html = "<thead><tr><th>Date</th><th class='numeric'>Impressions</th><th class='numeric'>Net Revenue</th></tr></thead><tbody>";
        sorted.forEach(row => html += `<tr><td>${row.Date}</td><td class='numeric'>${formatNumber(row.Impressions)}</td><td class='numeric'>$${formatCurrency(row['Net Revenue'])}</td></tr>`);
        renderOutput("Verizon/Yahoo Cleaned Report", `<table>${html}</tbody></table>`, 'outputTable3', {
            fileName,
            tableData: sorted
        });
    }

    function processTool4(workbook, fileName) {
        const sheetName = workbook.SheetNames.find(name => name.toLowerCase().includes("report"));
        if (!sheetName) throw new Error("Sheet containing 'report' not found.");
        const json = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
        if (!json.length) throw new Error("No valid data found.");
        const filtered = json.map(row => ({
            Day: formatDate(row.day),
            Imps: row.imps || 0,
            Revenue: row.revenue || 0
        }));
        let html = "<thead><tr><th>Day</th><th class='numeric'>Imps</th><th class='numeric'>Revenue</th></tr></thead><tbody>";
        filtered.forEach(row => html += `<tr><td>${row.Day}</td><td class='numeric'>${formatNumber(row.Imps)}</td><td class='numeric'>$${formatCurrency(row.Revenue)}</td></tr>`);
        renderOutput("Xandr Report", `<table>${html}</tbody></table>`, "outputTable4", {
            fileName,
            tableData: filtered
        });
    }

    function processTool5(workbook, fileName) {
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        if (!sheet) throw new Error("No sheet found.");
        const json = XLSX.utils.sheet_to_json(sheet, {
            defval: "",
            cellDates: true
        });
        if (!json.length) throw new Error("No data found.");
        if (!json[0].hasOwnProperty("YMD") || !json[0].hasOwnProperty("REVENUE") || !json[0].hasOwnProperty("RENDERED")) {
            throw new Error("Sheet must contain YMD, REVENUE, and RENDERED columns.");
        }
        const tableData = json.map(row => ({
            YMD: formatDate(row.YMD),
            RENDERED: row.RENDERED,
            REVENUE: row.REVENUE
        }));
        let html = "<thead><tr><th>YMD</th><th class='numeric'>RENDERED</th><th class='numeric'>REVENUE</th></tr></thead><tbody>";
        tableData.forEach(row => html += `<tr><td>${row.YMD}</td><td class='numeric'>${formatNumber(row.RENDERED)}</td><td class='numeric'>$${formatCurrency(row.REVENUE)}</td></tr>`);
        renderOutput("TripleLift Report", `<table>${html}</tbody></table>`, "outputTable5", {
            fileName,
            tableData
        });
    }

    function processTool6(workbook, fileName) {
        let json;
        for (const sheetName of workbook.SheetNames) {
            const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
                raw: false
            });
            if (data.length && data[0]['datehour'] && data[0]['Ad Renders'] && data[0]['Publisher Revenue']) {
                json = data;
                break;
            }
        }
        if (!json) throw new Error("No sheet with expected columns (datehour, Ad Renders, Publisher Revenue) found.");
        const sorted = json.map(r => ({
            Date: formatDate(new Date(r['datehour'])),
            'Ad Renders': Number(r['Ad Renders']),
            'Publisher Revenue': Number(r['Publisher Revenue'])
        })).sort((a, b) => new Date(a.Date) - new Date(b.Date));
        let html = "<thead><tr><th>Date</th><th class='numeric'>Ad Renders</th><th class='numeric'>Publisher Revenue</th></tr></thead><tbody>";
        sorted.forEach(row => html += `<tr><td>${row.Date}</td><td class='numeric'>${formatNumber(row['Ad Renders'])}</td><td class='numeric'>$${formatCurrency(row['Publisher Revenue'], 6)}</td></tr>`);
        renderOutput("Cadent Report", `<table>${html}</tbody></table>`, "outputTable6", {
            fileName,
            tableData: sorted
        });
    }

    function processTool7(workbook, fileName) {
        let foundData;
        for (const name of workbook.SheetNames) {
            const json = XLSX.utils.sheet_to_json(workbook.Sheets[name]);
            if (json.length && 'Date' in json[0] && 'Revenue' in json[0] && 'Ad Impr.' in json[0]) {
                foundData = json;
                break;
            }
        }
        if (!foundData) throw new Error("No sheet with 'Date', 'Revenue', and 'Ad Impr.' found.");
        const tableData = foundData.map(r => ({
            Date: formatDate(r['Date']),
            'Ad Impr.': r['Ad Impr.'],
            Revenue: r['Revenue']
        }));
        let html = "<thead><tr><th>Date</th><th class='numeric'>Ad Impr.</th><th class='numeric'>Revenue</th></tr></thead><tbody>";
        tableData.forEach(row => html += `<tr><td>${row.Date}</td><td class='numeric'>${formatNumber(row['Ad Impr.'])}</td><td class='numeric'>$${formatCurrency(row['Revenue'])}</td></tr>`);
        renderOutput("Media.net Report", `<table>${html}</tbody></table>`, "outputTable7", {
            fileName,
            tableData
        });
    }

    function processTool8(workbook, fileName) {
        let sheetName = workbook.SheetNames.find(name => name.toLowerCase() === "in") || workbook.SheetNames[0];
        const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
            defval: ""
        });
        if (!jsonData.length) throw new Error("No data found.");
        const validRows = jsonData.map(row => ({
            Date: formatDate(row["Date"] || row["date"] || row["DATE"]),
            Rendered: row["Rendered Impressions"] || row["Rendered"] || row["Impressions"] || row["RENDERED"],
            Earnings: row["Earnings"] || row["Revenue"] || row["Net Revenue"] || row["EARNING"]
        })).filter(r => r.Date && r.Rendered !== undefined && r.Earnings !== undefined);
        if (!validRows.length) throw new Error("Could not find columns (Date, Rendered, Earnings).");
        let html = "<thead><tr><th>Date</th><th class='numeric'>Rendered</th><th class='numeric'>Earnings</th></tr></thead><tbody>";
        validRows.forEach(row => html += `<tr><td>${row.Date}</td><td class='numeric'>${formatNumber(row.Rendered)}</td><td class='numeric'>$${formatCurrency(row.Earnings)}</td></tr>`);
        renderOutput("Sharethrough (HB) Report", `<table>${html}</tbody></table>`, "outputTable8", {
            fileName,
            tableData: validRows
        });
    }
})();
