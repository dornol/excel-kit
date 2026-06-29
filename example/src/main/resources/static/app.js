document.addEventListener('DOMContentLoaded', () => {
    const memoryElement = document.getElementById('memory');
    let latestErrors = [];
    let latestUploadFile = null;

    setInterval(() => {
        fetch('/memory')
            .then(async (response) => {
                if (response.ok) {
                    memoryElement.innerText = await response.text();
                } else {
                    memoryElement.innerText = 'Failed to retrieve memory usage.';
                }
            })
            .catch((err) => {
                memoryElement.innerText = 'Error retrieving memory usage.';
                console.error(err);
            });
    }, 300);

    document.getElementById('gc').addEventListener('click', () => {
        fetch('/gc');
    });

    document.getElementById('feature-filter').addEventListener('input', (event) => {
        const query = event.target.value.trim().toLowerCase();
        document.querySelectorAll('.section').forEach((section) => {
            section.style.display = section.innerText.toLowerCase().includes(query) ? '' : 'none';
        });
    });

    document.getElementById('clear-filter').addEventListener('click', () => {
        document.getElementById('feature-filter').value = '';
        document.querySelectorAll('.section').forEach((section) => {
            section.style.display = '';
        });
    });

    document.querySelectorAll('form[data-enhanced-upload="true"]').forEach((form) => {
        form.addEventListener('submit', async (event) => {
            event.preventDefault();
            const resultPanel = document.getElementById('upload-result');
            const summary = document.getElementById('upload-summary');
            resultPanel.style.display = 'block';
            summary.innerText = 'Uploading...';
            latestUploadFile = form.querySelector('input[type="file"]').files[0] || null;

            try {
                const response = await fetch(form.action, {
                    method: 'POST',
                    headers: { 'Accept': 'application/json' },
                    body: new FormData(form)
                });
                if (!response.ok) {
                    throw new Error(`HTTP ${response.status}`);
                }
                renderUploadResult(await response.json());
            } catch (err) {
                document.getElementById('upload-title').innerText = 'Upload failed';
                summary.innerText = err.message;
                document.getElementById('upload-rows').innerHTML = '';
                document.getElementById('upload-errors').innerHTML = '';
                document.getElementById('download-errors').classList.add('hidden');
                document.getElementById('server-errors-csv').classList.add('hidden');
                document.getElementById('server-errors-excel').classList.add('hidden');
                latestErrors = [];
            }
        });
    });

    document.getElementById('download-errors').addEventListener('click', () => {
        if (latestErrors.length === 0) {
            return;
        }
        const header = ['fileRowNum', 'columnIndex', 'headerName', 'cellValue', 'message'];
        const rows = [header.join(',')];
        latestErrors.forEach((error) => {
            if (error.cellErrors && error.cellErrors.length > 0) {
                error.cellErrors.forEach((cell) => rows.push([
                    error.fileRowNum,
                    cell.columnIndex,
                    cell.headerName,
                    cell.cellValue,
                    cell.message
                ].map(csvEscape).join(',')));
            } else {
                rows.push([error.fileRowNum, '', '', '', (error.messages || []).join('; ')]
                    .map(csvEscape).join(','));
            }
        });
        const blob = new Blob([rows.join('\n')], { type: 'text/csv;charset=utf-8' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = 'excel-kit-read-errors.csv';
        link.click();
        URL.revokeObjectURL(link.href);
    });

    document.getElementById('server-errors-csv').addEventListener('click', () => {
        downloadServerErrorReport('/showcase/read-errors-csv', 'read-errors.csv');
    });

    document.getElementById('server-errors-excel').addEventListener('click', () => {
        downloadServerErrorReport('/showcase/read-errors-excel', 'read-errors.xlsx');
    });

    function renderUploadResult(report) {
        latestErrors = report.errors || [];
        document.getElementById('upload-title').innerText = `Name-Based ${report.type} Read Result`;
        document.getElementById('upload-summary').innerText =
            `Success: ${report.successCount} rows, Errors: ${report.errorCount} rows`;
        renderRows(report.rows || []);
        renderErrors(latestErrors);
        document.getElementById('download-errors')
            .classList.toggle('hidden', latestErrors.length === 0);
        document.getElementById('server-errors-csv')
            .classList.toggle('hidden', latestErrors.length === 0 || latestUploadFile === null);
        document.getElementById('server-errors-excel')
            .classList.toggle('hidden', latestErrors.length === 0 || latestUploadFile === null);
    }

    function renderRows(rows) {
        const container = document.getElementById('upload-rows');
        if (rows.length === 0) {
            container.innerHTML = '';
            return;
        }
        const cells = rows.map((row) => `<tr>
            <td>${escapeHtml(row.name)}</td>
            <td>${escapeHtml(row.category)}</td>
            <td>${escapeHtml(row.price)}</td>
            <td>${escapeHtml(row.quantity)}</td>
            <td>${escapeHtml(row.discount)}</td>
        </tr>`).join('');
        container.innerHTML = `<table>
            <thead><tr><th>Name</th><th>Category</th><th>Price</th><th>Quantity</th><th>Discount</th></tr></thead>
            <tbody>${cells}</tbody>
        </table>`;
    }

    function renderErrors(errors) {
        const container = document.getElementById('upload-errors');
        if (errors.length === 0) {
            container.innerHTML = '';
            return;
        }
        const items = errors.map((error) => {
            const cells = error.cellErrors || [];
            if (cells.length === 0) {
                return `<li>fileRow=${escapeHtml(error.fileRowNum)}: ${escapeHtml((error.messages || []).join(', '))}</li>`;
            }
            return cells.map((cell) => `<li>fileRow=${escapeHtml(error.fileRowNum)}:
                [column=${escapeHtml(cell.columnIndex)}, header=${escapeHtml(cell.headerName)},
                value=${escapeHtml(cell.cellValue)}, message=${escapeHtml(cell.message)}]</li>`).join('');
        }).join('');
        container.innerHTML = `<ul class="error-list">${items}</ul>`;
    }

    function escapeHtml(value) {
        if (value === null || value === undefined) {
            return '';
        }
        return String(value)
            .replaceAll('&', '&amp;')
            .replaceAll('<', '&lt;')
            .replaceAll('>', '&gt;')
            .replaceAll('"', '&quot;');
    }

    function csvEscape(value) {
        if (value === null || value === undefined) {
            return '';
        }
        const text = String(value);
        if (/[",\n\r]/.test(text)) {
            return `"${text.replaceAll('"', '""')}"`;
        }
        return text;
    }

    async function downloadServerErrorReport(endpoint, filename) {
        if (latestUploadFile === null) {
            return;
        }
        const formData = new FormData();
        formData.append('file', latestUploadFile);
        const response = await fetch(endpoint, { method: 'POST', body: formData });
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }
        const blob = await response.blob();
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = filename;
        link.click();
        URL.revokeObjectURL(link.href);
    }
});
