let exportData = [];

// Set current date on page load
document.addEventListener('DOMContentLoaded', function () {
    const dateInput = document.getElementById('reportDate');
    const today = new Date();
    const year = today.getFullYear();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const day = String(today.getDate()).padStart(2, '0');
    dateInput.value = `${year}-${month}-${day}`;
});

// Enhanced addProduct function with Location validation
async function addProduct() {
    const productSelect = document.getElementById('productSelect');
    const initialQty = document.getElementById('initialQty');
    const usedQty = document.getElementById('usedQty');
    const tableBody = document.getElementById('tableBody');
    const reportDate = document.getElementById('reportDate').value;
    const location = document.getElementById('location').value;

    const productName = productSelect.options[productSelect.selectedIndex].text;
    const initialValue = parseFloat(initialQty.value) || 0;
    const usedValue = parseFloat(usedQty.value) || 0;

    if (!location || location === '') {
        alert('Please select a Location.');
        return;
    }

    if (!initialQty.value.trim() || !usedQty.value.trim()) {
        alert('Please fill in both Initial Quantity and Used Quantity.');
        return;
    }

    if (productName && initialValue >= 0 && usedValue >= 0 && initialValue >= usedValue) {
        const newRow = document.createElement('tr');
        newRow.innerHTML = `
            <td>${productName}</td>
            <td>${usedValue.toFixed(2)}</td>
        `;
        tableBody.appendChild(newRow);

        const docNo = `RC-${String(exportData.length + 1).padStart(4, '0')}`;
        const [year, month, day] = reportDate.split('-');
        const excelDate = `${day}${month}${year}`;
        const descriptionHdr = 'Stock Issue';
        const itemCode = productName.split(' (')[0];
        const qty = usedValue.toFixed(2);
        const uom = 'UNIT';

        const dataToSend = {
            DOCNO: docNo,
            DOCDATE: excelDate,
            DESCRIPTION_HDR: descriptionHdr,
            DOCAMT: '',
            REASON: '',
            REMARK: '',
            ITEMCODE: itemCode,
            LOCATION: location,
            BATCH: '',
            PROJECT: '',
            DESCRIPTION_DTL: productName,
            QTY: qty,
            UOM: uom,
            AMOUNT: '',
            REMARK1: '',
            REMARK2: ''
        };

        exportData.push(dataToSend);

        // Send data to ngrok URL with correct endpoint
        try {
            const response = await fetch('https://terina-unrefracted-elbert.ngrok-free.dev/add-inventory', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(dataToSend)
            });
            if (response.ok) {
                console.log('Data saved to server');
            } else {
                console.error('Error saving data to server:', response.statusText);
                alert('Error saving data to server. Data is still available locally for export.');
            }
        } catch (error) {
            console.error('Error connecting to server:', error.message);
            alert('Error connecting to server. Data is still available locally for export.');
        }

        initialQty.value = '';
        usedQty.value = '';
        document.getElementById('location').value = '';
        productSelect.selectedIndex = 0; // Reset product select
    } else {
        alert('Please enter valid quantities (Initial ≥ Used ≥ 0).');
    }
}

// Enhanced exportToExcel function with corrected layout and font size 11
function exportToExcel() {
    console.log('Export process started. Checking XLSX:', typeof XLSX !== 'undefined' ? 'Loaded' : 'Not Loaded');
    if (typeof XLSX === 'undefined') {
        alert('XLSX library failed to load. Check network or script URL.');
        console.error('XLSX is undefined');
        return;
    }

    console.log('exportData length:', exportData.length);
    if (exportData.length === 0) {
        alert('No data to export. Please add products first.');
        console.log('Export aborted: No data');
        return;
    }

    // Define headers based on "StockIssue" template (16 columns)
    const headers = [
        'DOCNO', 'DOCDATE', 'DESCRIPTION_HDR', 'DOCAMT', 'REASON',
        'REMARK', 'ITEMCODE', 'LOCATION', 'BATCH', 'PROJECT',
        'DESCRIPTION_DTL', 'QTY', 'UOM', 'AMOUNT', 'REMARK1', 'REMARK2'
    ];

    // Load the template file
    const templatePath = './Inventory_Template.xlsx';
    const request = new XMLHttpRequest();
    request.open('GET', templatePath, true);
    request.responseType = 'arraybuffer';

    request.onload = function () {
        if (request.status === 200) {
            console.log('Template loaded successfully');
            const data = new Uint8Array(request.response);
            const wb = XLSX.read(data, { type: 'array', cellStyles: true });

            // Get the "StockIssue" sheet
            let ws = wb.Sheets['StockIssue'];

            // Log the initial content of the sheet for debugging
            console.log('Initial sheet content:', ws);

            // Clear all rows below header (row 1) to remove any residual data or notes
            const range = XLSX.utils.decode_range(ws['!ref'] || { s: { r: 0, c: 0 }, e: { r: 0, c: 15 } });
            for (let r = 1; r <= range.e.r; r++) { // Start from row 2 to the end
                for (let c = 0; c <= range.e.c; c++) {
                    const cellAddress = XLSX.utils.encode_cell({ r: r, c: c });
                    if (ws[cellAddress]) {
                        delete ws[cellAddress]; // Remove all existing data and styles
                    }
                }
            }

            // Write headers on row 1 (A1)
            const headerRow = 0; // 0-based index for row 1
            headers.forEach((header, colIndex) => {
                const cellAddress = XLSX.utils.encode_cell({ r: headerRow, c: colIndex });
                ws[cellAddress] = {
                    v: header + (header.includes('(') ? '' : ''),
                    t: 's',
                    s: { font: { sz: 11, name: 'Calibri' }, alignment: { horizontal: 'left' } }
                };
            });

            // Write data starting from row 2 (A2)
            const startRow = 1; // 0-based index for row 2
            exportData.forEach((rowData, index) => {
                const rowNum = startRow + index; // Start from row 2
                headers.forEach((header, colIndex) => {
                    const cellAddress = XLSX.utils.encode_cell({ r: rowNum, c: colIndex });
                    ws[cellAddress] = {
                        v: rowData[header] || '',
                        t: 's',
                        s: { font: { sz: 11, name: 'Calibri' }, alignment: { horizontal: 'left' } }
                    };
                });
            });

            // Update the range to include new data only (remove extra rows)
            const newRange = { s: { r: 0, c: 0 }, e: { r: startRow + exportData.length - 1, c: 15 } };
            ws['!ref'] = XLSX.utils.encode_range(newRange);

            // Set column widths to match template exactly
            ws['!cols'] = [
                { wch: 10.56 }, // A (DOCNO)
                { wch: 10.56 }, // B (DOCDATE)
                { wch: 22.11 }, // C (DESCRIPTION_HDR)
                { wch: 7.22 },  // D (DOCAMT)
                { wch: 10.89 }, // E (REASON)
                { wch: 11.22 }, // F (REMARK)
                { wch: 13.89 }, // G (ITEMCODE)
                { wch: 13.11 }, // H (LOCATION)
                { wch: 11.67 }, // I (BATCH)
                { wch: 11.56 }, // J (PROJECT)
                { wch: 20.56 }, // K (DESCRIPTION_DTL)
                { wch: 7.22 },  // L (QTY)
                { wch: 7.22 },  // M (UOM)
                { wch: 7.22 },  // N (AMOUNT)
                { wch: 13.56 }, // O (REMARK1)
                { wch: 13.56 }  // P (REMARK2)
            ];

            // Generate automatic filename based on current date
            const dateInput = document.getElementById('reportDate').value;
            const [year, month, day] = dateInput.split('-');
            const months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'SEP', 'OCT', 'NOV', 'DEC'];
            const fileDate = `${day.padStart(2, '0')}-${months[parseInt(month) - 1]}-${year}`;
            const fileName = `Inventory Report ${fileDate}.xlsx`;
            console.log('Attempting to write file:', fileName);

            try {
                XLSX.writeFile(wb, fileName, { bookType: 'xlsx', type: 'binary', cellStyles: true });
                console.log('File write completed. Data written:', exportData);
            } catch (error) {
                console.error('Export failed:', error);
                alert('Export failed. Check console for details: ' + error.message);
            }
        } else {
            alert('Failed to load Inventory_Template.xlsx. Ensure it exists in the same directory as index.html.');
            console.error('Template load failed:', request.status, request.statusText);
        }
    };

    request.onerror = function () {
        alert('Error loading Inventory_Template.xlsx. Check the file path or server setup.');
        console.error('Template load error:', request.statusText);
    };

    console.log('Requesting template:', templatePath);
    request.send();
}