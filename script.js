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

// Enhanced exportToExcel function to use server-side export with blob size logging
async function exportToExcel() {
    if (exportData.length === 0) {
        alert('No data to export.');
        return;
    }

    try {
        console.log('Starting export process...');
        const response = await fetch('https://terina-unrefracted-elbert.ngrok-free.dev/export', {
            method: 'GET', // Matches the server route
            headers: {
                'Content-Type': 'application/json'
            }
        });

        console.log('Export response status:', response.status, response.statusText);
        console.log('Export response headers:', Object.fromEntries(response.headers)); // Log all headers

        if (response.ok) {
            const blob = await response.blob();
            console.log('Blob received, size:', blob.size, 'bytes'); // Debug blob size
            console.log('Blob type:', blob.type); // Debug blob type

            if (blob.size === 0) {
                throw new Error('Empty blob received from server');
            }

            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            const dateInput = document.getElementById('reportDate').value;
            const [year, month, day] = dateInput.split('-');
            const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
            const fileDate = `${day.padStart(2, '0')} ${monthNames[parseInt(month) - 1]} ${year}`;
            a.download = `Inventory Report ${fileDate}.xlsx`;
            a.click();
            window.URL.revokeObjectURL(url);
            console.log('Export completed successfully');
        } else {
            console.error('Export failed:', response.statusText);
            alert('Export failed. Check console for details.');
        }
    } catch (error) {
        console.error('Error during export:', error);
        alert('Error during export. Check console for details: ' + error.message);
    }
}
