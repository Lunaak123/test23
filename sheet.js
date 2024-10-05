let data = []; // This holds the initial Excel data
let filteredData = []; // This holds the filtered data after user operations

// Function to load and display the Excel sheet initially
async function loadExcelSheet(fileUrl) {
    try {
        const response = await fetch(fileUrl);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
        const sheetName = workbook.SheetNames[0]; // Load the first sheet
        const sheet = workbook.Sheets[sheetName];

        data = XLSX.utils.sheet_to_json(sheet, { defval: null });
        filteredData = [...data];

        // Initially display the full sheet
        displaySheet(filteredData);
    } catch (error) {
        console.error("Error loading Excel sheet:", error);
    }
}

// Function to display the Excel sheet as an HTML table
function displaySheet(sheetData) {
    const sheetContentDiv = document.getElementById('sheet-content');
    sheetContentDiv.innerHTML = ''; // Clear existing content

    if (sheetData.length === 0) {
        sheetContentDiv.innerHTML = '<p>No data available</p>';
        return;
    }

    const table = document.createElement('table');

    // Create table headers
    const headerRow = document.createElement('tr');
    Object.keys(sheetData[0]).forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    // Create table rows
    sheetData.forEach(row => {
        const tr = document.createElement('tr');
        Object.values(row).forEach(cell => {
            const td = document.createElement('td');
            td.textContent = cell === null ? 'NULL' : cell;
            tr.appendChild(td);
        });
        table.appendChild(tr);
    });

    sheetContentDiv.appendChild(table);
}

// Function to apply the selected operations and update the table
function applyOperation() {
    const primaryColumn = document.getElementById('primary-column').value.trim();
    const operationColumnsInput = document.getElementById('operation-columns').value.trim();
    const operationType = document.getElementById('operation-type').value;
    const operation = document.getElementById('operation').value;

    if (!primaryColumn || !operationColumnsInput) {
        alert('Please enter the primary column and columns to operate on.');
        return;
    }

    // Convert the entered column names (e.g., A, B, C) to column headers
    const operationColumns = operationColumnsInput.split(',').map(col => col.trim());

    filteredData = data.filter(row => {
        // Check if the primary column is null or not
        const isPrimaryNull = row[primaryColumn] === null;

        if (operation === 'null') {
            if (operationType === 'and') {
                return isPrimaryNull && operationColumns.every(col => row[col] === null);
            } else {
                return isPrimaryNull || operationColumns.some(col => row[col] === null);
            }
        } else {
            if (operationType === 'and') {
                return !isPrimaryNull && operationColumns.every(col => row[col] !== null);
            } else {
                return !isPrimaryNull || operationColumns.some(col => row[col] !== null);
            }
        }
    });

    displaySheet(filteredData);
}

// Function to download the Excel sheet
function downloadExcel() {
    const filename = document.getElementById('filename').value.trim();
    const format = document.getElementById('file-format').value;

    if (!filename) {
        alert('Please enter a filename.');
        return;
    }

    let exportData = XLSX.utils.json_to_sheet(filteredData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, exportData, 'Sheet1');

    const fileType = format === 'xlsx' ? 'application/octet-stream' : 'text/csv';
    const fileExtension = format === 'xlsx' ? '.xlsx' : '.csv';

    XLSX.writeFile(wb, `${filename}${fileExtension}`);
}

// Event Listeners
document.getElementById('apply-operation').addEventListener('click', applyOperation);
document.getElementById('download-button').addEventListener('click', () => {
    document.getElementById('download-modal').style.display = 'block';
});
document.getElementById('confirm-download').addEventListener('click', downloadExcel);
document.getElementById('close-modal').addEventListener('click', () => {
    document.getElementById('download-modal').style.display = 'none';
});

// Load the Excel sheet when the page loads
window.onload = () => {
    loadExcelSheet('path/to/your/excel/file.xlsx'); // Replace with your Excel file path
};
