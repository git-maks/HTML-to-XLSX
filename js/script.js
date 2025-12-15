/* =====================================================
   HTML-to-XLSX Converter - Main Application Logic
   ===================================================== */

// ============ Global State ============
const state = {
    currentInputMethod: 'upload',
    currentHTML: null,
    currentFile: null,
    detectedTables: [],
    selectedTableIndices: []
};

// ============ DOM Elements ============
const elements = {
    // Input method selectors
    selectorBtns: document.querySelectorAll('.selector-btn'),
    uploadArea: document.getElementById('uploadArea'),
    pasteArea: document.getElementById('pasteArea'),
    
    // File upload elements
    fileDropZone: document.getElementById('fileDropZone'),
    htmlFile: document.getElementById('htmlFile'),
    fileInfo: document.getElementById('fileInfo'),
    fileName: document.getElementById('fileName'),
    fileSize: document.getElementById('fileSize'),
    tableCount: document.getElementById('tableCount'),
    
    // Paste elements
    htmlPaste: document.getElementById('htmlPaste'),
    charCount: document.getElementById('charCount'),
    clearPasteBtn: document.getElementById('clearPasteBtn'),
    
    // Table selection
    tableSelection: document.getElementById('tableSelection'),
    tableOptions: document.getElementById('tableOptions'),
    
    // Action buttons
    convertBtn: document.getElementById('convertBtn'),
    resetBtn: document.getElementById('resetBtn'),
    
    // Status and download
    statusOutput: document.getElementById('statusOutput'),
    downloadSection: document.getElementById('downloadSection'),
    downloadLink: document.getElementById('downloadLink')
};

// ============ Initialization ============
document.addEventListener('DOMContentLoaded', () => {
    initializeEventListeners();
    updateStatus('ready', 'SYSTEM READY');
    updateStatus('info', 'AWAITING INPUT...');
});

// ============ Event Listeners Setup ============
function initializeEventListeners() {
    // Input method selector
    elements.selectorBtns.forEach(btn => {
        btn.addEventListener('click', () => switchInputMethod(btn.dataset.method));
    });
    
    // File upload events
    elements.fileDropZone.addEventListener('click', () => elements.htmlFile.click());
    elements.htmlFile.addEventListener('change', handleFileSelect);
    
    // Drag and drop events
    elements.fileDropZone.addEventListener('dragover', handleDragOver);
    elements.fileDropZone.addEventListener('dragleave', handleDragLeave);
    elements.fileDropZone.addEventListener('drop', handleFileDrop);
    
    // Paste events
    elements.htmlPaste.addEventListener('input', handlePasteInput);
    elements.clearPasteBtn.addEventListener('click', clearPaste);
    
    // Action buttons
    elements.convertBtn.addEventListener('click', handleConvert);
    elements.resetBtn.addEventListener('click', resetAll);
}

// ============ Input Method Switching ============
function switchInputMethod(method) {
    state.currentInputMethod = method;
    
    // Update button states
    elements.selectorBtns.forEach(btn => {
        if (btn.dataset.method === method) {
            btn.classList.add('active');
        } else {
            btn.classList.remove('active');
        }
    });
    
    // Show/hide input areas
    if (method === 'upload') {
        elements.uploadArea.classList.add('active');
        elements.pasteArea.classList.remove('active');
    } else {
        elements.uploadArea.classList.remove('active');
        elements.pasteArea.classList.add('active');
    }
    
    // Reset state
    elements.tableSelection.classList.remove('show');
    elements.downloadSection.classList.remove('show');
    updateStatus('info', `SWITCHED TO ${method.toUpperCase()} MODE`);
}

// ============ File Upload Handlers ============
function handleDragOver(e) {
    e.preventDefault();
    e.stopPropagation();
    elements.fileDropZone.classList.add('drag-over');
}

function handleDragLeave(e) {
    e.preventDefault();
    e.stopPropagation();
    elements.fileDropZone.classList.remove('drag-over');
}

function handleFileDrop(e) {
    e.preventDefault();
    e.stopPropagation();
    elements.fileDropZone.classList.remove('drag-over');
    
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        processFile(files[0]);
    }
}

function handleFileSelect(e) {
    const file = e.target.files[0];
    if (file) {
        processFile(file);
    }
}

function processFile(file) {
    // Validate file size (10MB limit)
    const maxSize = 10 * 1024 * 1024;
    if (file.size > maxSize) {
        updateStatus('error', `ERROR: FILE TOO LARGE (${formatFileSize(file.size)} > 10MB)`);
        return;
    }
    
    state.currentFile = file;
    
    // Display file info
    elements.fileName.textContent = `FILE: ${file.name}`;
    elements.fileSize.textContent = `SIZE: ${formatFileSize(file.size)}`;
    elements.fileInfo.classList.add('show');
    
    // Read file content
    const reader = new FileReader();
    reader.onload = (e) => {
        state.currentHTML = e.target.result;
        analyzeHTML(state.currentHTML);
        updateStatus('ready', `FILE LOADED: ${file.name}`);
    };
    reader.onerror = () => {
        updateStatus('error', 'ERROR: FAILED TO READ FILE');
    };
    reader.readAsText(file);
}

// ============ Paste Handlers ============
function handlePasteInput(e) {
    const text = e.target.value;
    elements.charCount.textContent = text.length.toLocaleString();
    
    if (text.length > 0) {
        state.currentHTML = text;
        analyzeHTML(text);
    }
}

function clearPaste() {
    elements.htmlPaste.value = '';
    elements.charCount.textContent = '0';
    state.currentHTML = null;
    state.detectedTables = [];
    elements.tableSelection.classList.remove('show');
    updateStatus('info', 'PASTE AREA CLEARED');
}

// ============ HTML Analysis ============
function analyzeHTML(htmlString) {
    try {
        // Parse HTML using DOMParser
        const parser = new DOMParser();
        const doc = parser.parseFromString(htmlString, 'text/html');
        
        // Check for parsing errors
        const parserError = doc.querySelector('parsererror');
        if (parserError) {
            updateStatus('error', 'WARNING: MALFORMED HTML DETECTED - ATTEMPTING RECOVERY...');
        }
        
        // Extract all tables
        const tables = doc.querySelectorAll('table');
        state.detectedTables = Array.from(tables);
        
        if (tables.length === 0) {
            updateStatus('error', 'ERROR: NO TABLES FOUND IN HTML');
            elements.tableCount.textContent = 'TABLES: 0 DETECTED';
            return;
        }
        
        // Update UI based on table count
        const countMsg = `TABLES: ${tables.length} DETECTED`;
        elements.tableCount.textContent = countMsg;
        updateStatus('ready', countMsg);
        
        // Show table selection if multiple tables
        if (tables.length > 1) {
            showTableSelection(tables);
        } else {
            elements.tableSelection.classList.remove('show');
            state.selectedTableIndices = [0];
        }
        
    } catch (error) {
        updateStatus('error', `ERROR: ${error.message}`);
    }
}

// ============ Table Selection UI ============
function showTableSelection(tables) {
    elements.tableOptions.innerHTML = '';
    
    tables.forEach((table, index) => {
        const option = document.createElement('div');
        option.className = 'table-option';
        
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.id = `table-${index}`;
        checkbox.value = index;
        checkbox.checked = true;
        
        const label = document.createElement('label');
        label.htmlFor = `table-${index}`;
        
        // Get table info
        const rows = table.querySelectorAll('tr').length;
        const caption = table.querySelector('caption');
        const tableId = table.id;
        const tableInfo = caption ? caption.textContent : (tableId || `Table ${index + 1}`);
        
        label.textContent = `${tableInfo} (${rows} rows)`;
        
        option.appendChild(checkbox);
        option.appendChild(label);
        elements.tableOptions.appendChild(option);
        
        // Event listener
        checkbox.addEventListener('change', () => {
            updateSelectedTables();
        });
    });
    
    elements.tableSelection.classList.add('show');
    updateSelectedTables();
}

function updateSelectedTables() {
    const checkboxes = elements.tableOptions.querySelectorAll('input[type=\"checkbox\"]');
    state.selectedTableIndices = Array.from(checkboxes)
        .filter(cb => cb.checked)
        .map(cb => parseInt(cb.value));
    
    if (state.selectedTableIndices.length === 0) {
        updateStatus('error', 'WARNING: NO TABLES SELECTED');
    } else {
        updateStatus('info', `${state.selectedTableIndices.length} TABLE(S) SELECTED FOR CONVERSION`);
    }
}

// ============ HTML to Data Extraction ============
function extractTableData(table) {
    const data = [];
    const mergedCells = [];
    
    // Get all rows
    const rows = table.querySelectorAll('tr');
    
    rows.forEach((row, rowIndex) => {
        const cells = row.querySelectorAll('th, td');
        const rowData = [];
        let colIndex = 0;
        
        cells.forEach(cell => {
            // Skip columns that are occupied by merged cells from previous rows
            while (isCellMerged(rowIndex, colIndex, mergedCells)) {
                rowData.push(null); // Placeholder for merged cell
                colIndex++;
            }
            
            // Get cell content (strip HTML tags, decode entities)
            const content = getCleanCellContent(cell);
            
            // Handle colspan and rowspan
            const colspan = parseInt(cell.getAttribute('colspan')) || 1;
            const rowspan = parseInt(cell.getAttribute('rowspan')) || 1;
            
            // Add cell data
            rowData.push(content);
            
            // Track merged cells
            if (colspan > 1 || rowspan > 1) {
                for (let r = 0; r < rowspan; r++) {
                    for (let c = 0; c < colspan; c++) {
                        if (r === 0 && c === 0) continue; // Skip the origin cell
                        mergedCells.push({
                            row: rowIndex + r,
                            col: colIndex + c,
                            originRow: rowIndex,
                            originCol: colIndex
                        });
                    }
                }
            }
            
            // Fill colspan cells
            for (let c = 1; c < colspan; c++) {
                rowData.push(null);
            }
            
            colIndex += colspan;
        });
        
        data.push(rowData);
    });
    
    return { data, mergedCells };
}

function isCellMerged(row, col, mergedCells) {
    return mergedCells.some(m => m.row === row && m.col === col);
}

function getCleanCellContent(cell) {
    // Get text content
    let content = cell.textContent || '';
    
    // Decode HTML entities
    const textarea = document.createElement('textarea');
    textarea.innerHTML = content;
    content = textarea.value;
    
    // Normalize whitespace
    content = content.replace(/\s+/g, ' ').trim();
    
    // Try to detect numbers
    const num = parseFloat(content);
    if (!isNaN(num) && content === num.toString()) {
        return num;
    }
    
    return content;
}

// ============ XLSX Generation ============
function generateXLSX(tablesToConvert) {
    const workbook = XLSX.utils.book_new();
    
    tablesToConvert.forEach((tableIndex, sheetIndex) => {
        const table = state.detectedTables[tableIndex];
        const { data, mergedCells } = extractTableData(table);
        
        // Create worksheet from data
        const worksheet = XLSX.utils.aoa_to_sheet(data);
        
        // Apply merged cells
        if (mergedCells.length > 0) {
            worksheet['!merges'] = [];
            const processedMerges = new Set();
            
            mergedCells.forEach(merge => {
                const mergeKey = `${merge.originRow},${merge.originCol}`;
                if (!processedMerges.has(mergeKey)) {
                    // Find the extent of the merge
                    const relatedMerges = mergedCells.filter(m => 
                        m.originRow === merge.originRow && m.originCol === merge.originCol
                    );
                    
                    if (relatedMerges.length > 0) {
                        const maxRow = Math.max(...relatedMerges.map(m => m.row), merge.originRow);
                        const maxCol = Math.max(...relatedMerges.map(m => m.col), merge.originCol);
                        
                        worksheet['!merges'].push({
                            s: { r: merge.originRow, c: merge.originCol },
                            e: { r: maxRow, c: maxCol }
                        });
                        
                        processedMerges.add(mergeKey);
                    }
                }
            });
        }
        
        // Determine sheet name
        const caption = table.querySelector('caption');
        const tableId = table.id;
        const sheetName = caption ? caption.textContent.substring(0, 31) : 
                         (tableId || `Table${sheetIndex + 1}`);
        
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    });
    
    return workbook;
}

// ============ Conversion Handler ============
function handleConvert() {
    if (!state.currentHTML) {
        updateStatus('error', 'ERROR: NO INPUT PROVIDED');
        return;
    }
    
    if (state.detectedTables.length === 0) {
        updateStatus('error', 'ERROR: NO TABLES TO CONVERT');
        return;
    }
    
    if (state.selectedTableIndices.length === 0) {
        updateStatus('error', 'ERROR: NO TABLES SELECTED');
        return;
    }
    
    try {
        updateStatus('processing', 'PROCESSING...');
        updateStatus('processing', 'EXTRACTING TABLE DATA...');
        
        // Generate XLSX
        const workbook = generateXLSX(state.selectedTableIndices);
        
        updateStatus('processing', 'GENERATING XLSX FILE...');
        
        // Create file blob
        const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([wbout], { type: 'application/octet-stream' });
        
        // Generate filename
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
        const filename = state.currentFile ? 
            state.currentFile.name.replace(/\.[^/.]+$/, '') + '.xlsx' :
            `converted_table_${timestamp}.xlsx`;
        
        // Create download link
        const url = URL.createObjectURL(blob);
        elements.downloadLink.href = url;
        elements.downloadLink.download = filename;
        
        // Show download section
        elements.downloadSection.classList.add('show');
        updateStatus('ready', 'CONVERSION COMPLETE!');
        updateStatus('ready', `FILE: ${filename}`);
        updateStatus('ready', `SIZE: ${formatFileSize(blob.size)}`);
        
    } catch (error) {
        updateStatus('error', `ERROR: ${error.message}`);
        console.error('Conversion error:', error);
    }
}

// ============ Reset Function ============
function resetAll() {
    // Reset state
    state.currentHTML = null;
    state.currentFile = null;
    state.detectedTables = [];
    state.selectedTableIndices = [];
    
    // Reset file upload
    elements.htmlFile.value = '';
    elements.fileInfo.classList.remove('show');
    
    // Reset paste
    elements.htmlPaste.value = '';
    elements.charCount.textContent = '0';
    
    // Hide sections
    elements.tableSelection.classList.remove('show');
    elements.downloadSection.classList.remove('show');
    
    // Reset status
    updateStatus('ready', 'SYSTEM RESET');
    updateStatus('info', 'AWAITING INPUT...');
}

// ============ Status Update Function ============
function updateStatus(type, message) {
    const statusLine = document.createElement('div');
    statusLine.className = `status-line status-${type}`;
    statusLine.innerHTML = `$ ${message}`;
    
    // Clear old messages if too many
    if (elements.statusOutput.children.length > 10) {
        elements.statusOutput.innerHTML = '';
    }
    
    elements.statusOutput.appendChild(statusLine);
    elements.statusOutput.scrollTop = elements.statusOutput.scrollHeight;
}

// ============ Utility Functions ============
function formatFileSize(bytes) {
    if (bytes === 0) return '0 BYTES';
    const k = 1024;
    const sizes = ['BYTES', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
}
