let exportData = [];

// Replace with your ngrok URL (example: https://abc123.ngrok.io)
// For local testing: http://localhost:3000
const serverUrl = 'http://localhost:3000'; // CHANGE THIS TO YOUR NGROK URL

// Load data from server
async function loadData() {
    try {
        const response = await fetch(`${serverUrl}/get-inventory`);
        if (response.ok) {
            exportData = await response.json();
            updateTable();
            console.log('Data loaded from server');
        }
    } catch (e) {
        console.log('No server data – starting empty');
    }
}

// Update table
function updateTable() {
    const tableBody = document.getElementById('tableBody');
    tableBody.innerHTML = '';
    exportData.forEach(item => {
        const row = document.createElement('tr');
        row.innerHTML = `<td>${item.DESCRIPTION_DTL}</td><td>${item.QTY}</td>`;
        tableBody.appendChild(row);
    });
}

// Set today's date
document.addEventListener('DOMContentLoaded', function () {
    const dateInput = document.getElementById('reportDate');
    const today = new Date();
    const year = today.getFullYear();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const day = String(today.getDate()).padStart(2, '0');
    dateInput.value = `${year}-${month}-${day}`;
    loadData();
});

// Add product to server
async function addProduct() {
    const productSelect = document.getElementById('productSelect');
    const initialQty = document.getElementById('initialQty');
    const usedQty = document.getElementById('usedQty');
    const reportDate = document.getElementById('reportDate').value;
    const location = document.getElementById('location').value;

    const productName = productSelect.options[productSelect.selectedIndex].text;
    const initialValue = parseFloat(initialQty.value) || 0;
    const usedValue = parseFloat(usedQty.value) || 0;

    if (!location) return alert('Pilih Lokasi.');
    if (!initialQty.value || !usedQty.value) return alert('Isi kuantiti.');
    if (initialValue < usedValue) return alert('Initial ≥ Used.');

    const docNo = `RC-${String(exportData.length + 1).padStart(4, '0')}`;

    // FIXED: Removed invalid `report |`
    const [year, month, day] = reportDate.split('-');
    const excelDate = `${day}${month}${year}`;

    const dataToSend = {
        DOCNO: docNo,
        DOCDATE: excelDate,
        DESCRIPTION_HDR: "Stock Issue",
        ITEMCODE: productName.split(' (')[0],
        LOCATION: location,
        DESCRIPTION_DTL: productName,
        QTY: usedValue.toFixed(2),
        UOM: "UNIT"
    };

    try {
        const response = await fetch(`${serverUrl}/add-inventory`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(dataToSend)
        });

        if (response.ok) {
            await loadData(); // Refresh from server
            alert('Data saved to server!');
        } else {
            alert('Failed to save to server.');
        }
    } catch (e) {
        alert('Server connection failed: ' + e.message);
    }

    // Clear form
    initialQty.value = '';
    usedQty.value = '';
    document.getElementById('location').value = '';
    productSelect.selectedIndex = 0;
}

// Export Excel from server
async function exportToExcel() {
    const reportDate = document.getElementById('reportDate').value;
    if (exportData.length === 0) return alert('No data to export.');

    try {
        const response = await fetch(`${serverUrl}/export?date=${reportDate}`);
        if (!response.ok) throw new Error('No data or server error');

        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;

        // Use filename from server
        const contentDisposition = response.headers.get('Content-Disposition');
        let fileName = 'Inventory Report.xlsx';
        if (contentDisposition) {
            const match = contentDisposition.match(/filename="(.+)"/);
            if (match) fileName = match[1];
        }

        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        alert('Excel downloaded successfully!');
    } catch (e) {
        alert('Export failed: ' + e.message);
    }
}