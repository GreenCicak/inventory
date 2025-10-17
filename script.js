async function exportToExcel() {
    if (exportData.length === 0) {
        alert('No data to export.');
        return;
    }

    try {
        console.log('Starting export process...');
        const response = await fetch('https://terina-unrefracted-elbert.ngrok-free.dev/export', {
            method: 'GET',
            headers: {
                'Content-Type': 'application/json'
            }
        });

        console.log('Export response status:', response.status, response.statusText);

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
