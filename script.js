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
            const response = await fetch('https://terina-unrefracted-elbert.ngrok-free.devv ', {
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

    const headers = [
        'DOCNO', 'DOCDATE', 'DESCRIPTION_HDR', 'DOCAMT', 'REASON',
        'REMARK', 'ITEMCODE', 'LOCATION', 'BATCH', 'PROJECT',
        'DESCRIPTION_DTL', 'QTY', 'UOM', 'AMOUNT', 'REMARK1', 'REMARK2'
    ];

    // Load the template file from ngrok
    const templatePath = 'https://terina-unrefracted-elbert.ngrok-free.dev/Inventory_Template.xlsx';
    const request = new XMLHttpRequest();
    request.open('GET', templatePath, true);
    request.responseType = 'arraybuffer';

    request.onload = function () {
        if (request.status === 200) {
            console.log('Template loaded successfully');
            const data = new Uint8Array(request.response);
            const wb = XLSX.read(data, { type: 'array', cellStyles: true });

            let ws = wb.Sheets['StockIssue'];

            console.log('Initial sheet content:', ws);

            const range = XLSX.utils.decode_range(ws['!ref'] || { s: { r: 0, c: 0 }, e: { r: 0, c: 15 } });
            for (let r = 1; r <= range.e.r; r++) {
                for (let c = 0; c <= range.e.c; c++) {
                    const cellAddress = XLSX.utils.encode_cell({ r, c });
                    if (ws[cellAddress]) delete ws[cellAddress];
                }
            }

            const headerRow = 0;
            headers.forEach((header, colIndex) => {
                const cellAddress = XLSX.utils.encode_cell({ r: headerRow, c: colIndex });
                ws[cellAddress] = {
                    v: header,
                    t: 's',
                    s: { font: { sz: 11, name: 'Calibri' }, alignment: { horizontal: 'left' } }
                };
            });

            const startRow = 1;
            exportData.forEach((rowData, index) => {
                const rowNum = startRow + index;
                headers.forEach((header, colIndex) => {
                    const cellAddress = XLSX.utils.encode_cell({ r: rowNum, c: colIndex });
                    ws[cellAddress] = {
                        v: rowData[header] || '',
                        t: 's',
                        s: { font: { sz: 11, name: 'Calibri' }, alignment: { horizontal: 'left' } }
                    };
                });
            });

            const newRange = { s: { r: 0, c: 0 }, e: { r: startRow + exportData.length - 1, c: 15 } };
            ws['!ref'] = XLSX.utils.encode_range(newRange);

            ws['!cols'] = [
                { wch: 10.56 }, { wch: 10.56 }, { wch: 22.11 }, { wch: 7.22 }, { wch: 10.89 },
                { wch: 11.22 }, { wch: 13.89 }, { wch: 13.11 }, { wch: 11.67 }, { wch: 11.56 },
                { wch: 20.56 }, { wch: 7.22 }, { wch: 7.22 }, { wch: 7.22 }, { wch: 13.56 }, { wch: 13.56 }
            ];

            const dateInput = document.getElementById('reportDate').value;
            const [year, month, day] = dateInput.split('-');
            const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
            const fileDate = `${day.padStart(2, '0')} ${monthNames[parseInt(month) - 1]} ${year}`;
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
            alert('Failed to load Inventory_Template.xlsx. Ensure it exists on the server.');
            console.error('Template load failed:', request.status, request.statusText);
        }
    };

    request.onerror = function () {
        alert('Error loading Inventory_Template.xlsx. Check the server setup.');
        console.error('Template load error:', request.statusText);
    };

    console.log('Requesting template:', templatePath);
    request.send();
}