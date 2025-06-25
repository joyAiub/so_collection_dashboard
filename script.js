var gk_isXlsx = false;
var gk_xlsxFileLookup = {};
var gk_fileData = {};

function filledCell(cell) {
    return cell !== '' && cell != null;
}

function loadFileData(filename) {
    if (gk_isXlsx && gk_xlsxFileLookup[filename]) {
        try {
            var workbook = XLSX.read(gk_fileData[filename], { type: 'base64' });
            var firstSheetName = workbook.SheetNames[0];
            var worksheet = workbook.Sheets[firstSheetName];
            var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, blankrows: false, defval: '' });
            var filteredData = jsonData.filter(row => row.some(filledCell));
            var headerRowIndex = filteredData.findIndex((row, index) =>
                row.filter(filledCell).length >= filteredData[index + 1]?.filter(filledCell).length
            );
            if (headerRowIndex === -1 || headerRowIndex > 25) {
                headerRowIndex = 0;
            }
            var csv = XLSX.utils.aoa_to_sheet(filteredData.slice(headerRowIndex));
            csv = XLSX.utils.sheet_to_csv(csv, { header: 1 });
            return csv;
        } catch (e) {
            console.error(e);
            return "";
        }
    }
    return gk_fileData[filename] || "";
}

let soData = [];
let coData = [];
let currentData = [];
let filteredData = [];
let filters = {};
let dashboardFilter = null;
const rowsPerPage = 30;
let currentPage = 1;
let pageLoadTime = new Date();
const possibleLines = ['A', 'B', 'C', 'D', 'F', 'G', 'K'];
const possibleServers = Array.from({length: 20}, (_, i) => (i + 1).toString());
let currentTitle = 'Collection';
let sortField = null;
let sortDirection = 'asc';

function formatTimestamp(date) {
    return date.toLocaleString('en-US', { 
        timeZone: 'Asia/Dhaka', 
        hour12: true, 
        year: 'numeric', 
        month: '2-digit', 
        day: '2-digit', 
        hour: '2-digit', 
        minute: '2-digit', 
        second: '2-digit' 
    });
}

function formatDateForFilename(date) {
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}-${month}-${year}`;
}

function updateTimestamps() {
    document.getElementById('current-timestamp').textContent = `Current Time: ${formatTimestamp(new Date())}`;
    document.getElementById('page-loaded-timestamp').textContent = `Page Loaded: ${formatTimestamp(pageLoadTime)}`;
}

function setDefaultDates() {
    const today = new Date().toISOString().split('T')[0];
    document.getElementById('from-date').value = today;
    document.getElementById('to-date').value = today;
}

function updateTableTitle() {
    const titleElement = document.getElementById('table-title');
    titleElement.textContent = currentTitle;
}

function updateOrderCounts() {
    const totalOrderCount = new Set(soData.map(item => item.do_number).filter(num => num && num !== '-')).size;
    document.getElementById('total-order-count').textContent = `Total Order Count: ${totalOrderCount}`;

    const divisionCounts = {};
    possibleLines.forEach(line => {
        const count = new Set(
            soData
                .filter(item => item.Division === line && item.do_number && item.do_number !== '-')
                .map(item => item.do_number)
        ).size;
        if (count > 0) {
            divisionCounts[line] = count;
        }
    });
    const divisionText = Object.entries(divisionCounts).map(([line, count]) => `${line}=${count}`).join(', ');
    document.getElementById('division-order-count').textContent = `Division Order Count: ${divisionText || 'None'}`;
}

function updateServerAndLineAllocation() {
    const soServerCounts = {};
    possibleServers.forEach(server => {
        const uniqueTransactions = new Set(
            soData
                .filter(item => item.serverAllocation === server && item.TransactionId)
                .map(item => item.TransactionId)
        );
        soServerCounts[server] = uniqueTransactions.size;
    });
    const soServersDiv = document.getElementById('so-servers');
    soServersDiv.innerHTML = '';
    possibleServers.forEach(server => {
        const button = document.createElement('button');
        button.className = `allocation-button ${['red', 'orange', 'yellow', 'lime', 'emerald'][server % 5]}`;
        button.setAttribute('data-type', 'SO');
        button.setAttribute('data-field', 'serverAllocation');
        button.setAttribute('data-value', server);
        button.textContent = `S${server}: ${soServerCounts[server]}`;
        button.addEventListener('click', (e) => {
            e.stopPropagation();
            dashboardFilter = { type: 'SO', field: 'serverAllocation', value: server };
            currentData = soData;
            currentTitle = `S${server} SO Server Allocation`;
            applyFilters();
            updateTableTitle();
        });
        soServersDiv.appendChild(button);
    });

    const coServerCounts = {};
    possibleServers.forEach(server => {
        const uniqueTransactions = new Set(
            coData
                .filter(item => item.serverAllocation === server && item.TransactionId)
                .map(item => item.TransactionId)
        );
        coServerCounts[server] = uniqueTransactions.size;
    });
    const coServersDiv = document.getElementById('co-servers');
    coServersDiv.innerHTML = '';
    possibleServers.forEach(server => {
        const button = document.createElement('button');
        button.className = `allocation-button ${['red', 'orange', 'yellow', 'lime', 'emerald'][server % 5]}`;
        button.setAttribute('data-type', 'CO');
        button.setAttribute('data-field', 'serverAllocation');
        button.setAttribute('data-value', server);
        button.textContent = `S${server}: ${coServerCounts[server]}`;
        button.addEventListener('click', (e) => {
            e.stopPropagation();
            dashboardFilter = { type: 'CO', field: 'serverAllocation', value: server };
            currentData = coData;
            currentTitle = `S${server} CO Server Allocation`;
            applyFilters();
            updateTableTitle();
        });
        coServersDiv.appendChild(button);
    });

    const soLineCounts = {};
    let soOthersCount = 0;
    possibleLines.forEach(line => {
        const uniqueTransactions = new Set(
            soData
                .filter(item => item.Division === line && item.TransactionId)
                .map(item => item.TransactionId)
        );
        soLineCounts[line] = uniqueTransactions.size;
    });
    const soAllTransactions = new Set(soData.map(item => item.TransactionId));
    const soLineTransactions = new Set(
        soData
            .filter(item => possibleLines.includes(item.Division) && item.TransactionId)
            .map(item => item.TransactionId)
    );
    soOthersCount = soAllTransactions.size - soLineTransactions.size;
    const soLinesDiv = document.getElementById('so-lines');
    soLinesDiv.innerHTML = '';
    possibleLines.forEach(line => {
        const button = document.createElement('button');
        button.className = `allocation-button ${['pink', 'rose', 'fuchsia', 'violet'][possibleLines.indexOf(line) % 4]}`;
        button.setAttribute('data-type', 'SO');
        button.setAttribute('data-field', 'Division');
        button.setAttribute('data-value', line);
        button.textContent = `${line}: ${soLineCounts[line]}`;
        button.addEventListener('click', (e) => {
            e.stopPropagation();
            dashboardFilter = { type: 'SO', field: 'Division', value: line };
            currentData = soData;
            currentTitle = `SO Transaction Division Wise ${line}`;
            applyFilters();
            updateTableTitle();
        });
        soLinesDiv.appendChild(button);
    });

    const coLineCounts = {};
    possibleLines.forEach(line => {
        const uniqueTransactions = new Set(
            coData
                .filter(item => item.line === line && item.TransactionId)
                .map(item => item.TransactionId)
        );
        coLineCounts[line] = uniqueTransactions.size;
    });
    const coLinesDiv = document.getElementById('co-lines');
    coLinesDiv.innerHTML = '';
    possibleLines.forEach(line => {
        const button = document.createElement('button');
        button.className = `allocation-button ${['pink', 'rose', 'fuchsia', 'violet'][possibleLines.indexOf(line) % 4]}`;
        button.setAttribute('data-type', 'CO');
        button.setAttribute('data-field', 'line');
        button.setAttribute('data-value', line);
        button.textContent = `${line}: ${coLineCounts[line]}`;
        button.addEventListener('click', (e) => {
            e.stopPropagation();
            dashboardFilter = { type: 'CO', field: 'line', value: line };
            currentData = coData;
            currentTitle = `CO Transaction Line Wise ${line}`;
            applyFilters();
            updateTableTitle();
        });
        coLinesDiv.appendChild(button);
    });
}

async function fetchData() {
    const loadingDiv = document.getElementById('loading');
    const errorDiv = document.getElementById('error');
    const fromDate = document.getElementById('from-date').value;
    const toDate = document.getElementById('to-date').value;
    const lineNumber = document.getElementById('line-number').value || '0';
    const url = `https://psa.mgi.org/api/getALLDataView/${fromDate}/${toDate}/${lineNumber}`;
    try {
        loadingDiv.style.display = 'block';
        errorDiv.style.display = 'none';
        const response = await fetch(url, {
            method: 'GET',
            headers: {
                'Content-Type': 'application/json'
            }
        });
        if (!response.ok) {
            throw new Error(`HTTP error! Status: ${response.status}`);
        }
        const rawData = await response.json();
        soData = Array.isArray(rawData.SO) ? rawData.SO : [];
        coData = Array.isArray(rawData.CO) ? rawData.CO : [];
        updateDashboard();
        currentData = coData;
        filteredData = coData;
        filters = {};
        dashboardFilter = null;
        currentTitle = 'Collection';
        displayTable();
        updateTableTitle();
        loadingDiv.style.display = 'none';
        updateTimestamps();
        setInterval(updateTimestamps, 1000);
    } catch (error) {
        loadingDiv.style.display = 'none';
        errorDiv.textContent = `Error fetching data: ${error.message}`;
        errorDiv.style.display = 'block';
    }
}

function updateDashboard() {
    const totalSo = new Set(soData.map(item => item.TransactionId)).size;
    const totalCo = new Set(coData.map(item => item.TransactionId)).size;
    const completedSoData = soData
        .filter(item => item.soNumber && item.soNumber !== '-' && /^[a-zA-Z0-9]+$/.test(item.soNumber));
    const completedCoData = coData
        .filter(item => item.coNumber && item.coNumber !== '-' && /^[a-zA-Z0-9]+$/.test(item.coNumber));
    const completedSo = new Set(completedSoData.map(item => item.soNumber)).size;
    const completedCo = new Set(completedCoData.map(item => item.coNumber)).size;
    const pendingSoData = soData
        .filter(item => !item.soNumber || item.soNumber === '-');
    const pendingCoData = coData
        .filter(item => !item.coNumber || item.coNumber === '-');
    const pendingSo = new Set(pendingSoData.map(item => item.TransactionId)).size;
    const pendingCo = new Set(pendingCoData.map(item => item.TransactionId)).size;

    document.getElementById('total-so-rows').textContent = `SO API Row: ${soData.length}`;
    document.getElementById('total-co-rows').textContent = `Collection API Row: ${coData.length}`;
    document.getElementById('total-row').textContent = `Total Row: ${soData.length + coData.length}`;
    document.getElementById('total-so').textContent = totalSo;
    document.getElementById('total-co').textContent = totalCo;
    document.getElementById('completed-so').textContent = completedSo;
    document.getElementById('completed-co').textContent = completedCo;
    document.getElementById('pending-so').textContent = pendingSo;
    document.getElementById('pending-co').textContent = pendingCo;

    updateServerAndLineAllocation();
    updateOrderCounts();

    document.querySelectorAll('#dashboard .card').forEach(div => {
        div.addEventListener('click', () => {
            const filter = div.getAttribute('data-filter');
            if (filter) {
                if (filter === 'total-so') {
                    currentData = soData;
                    filteredData = soData;
                    dashboardFilter = null;
                    currentTitle = 'Total SO';
                } else if (filter === 'total-co') {
                    currentData = coData;
                    filteredData = coData;
                    dashboardFilter = null;
                    currentTitle = 'Total Collection';
                } else if (filter === 'completed-so') {
                    currentData = soData;
                    filteredData = completedSoData;
                    dashboardFilter = { type: 'SO', field: 'soNumber', value: 'completed' };
                    currentTitle = 'Completed SO';
                } else if (filter === 'completed-co') {
                    currentData = coData;
                    filteredData = completedCoData;
                    dashboardFilter = { type: 'CO', field: 'coNumber', value: 'completed' };
                    currentTitle = 'Completed Collection';
                } else if (filter === 'pending-so') {
                    currentData = soData;
                    filteredData = pendingSoData;
                    dashboardFilter = { type: 'SO', field: 'soNumber', value: 'pending' };
                    currentTitle = 'Pending SO';
                } else if (filter === 'pending-co') {
                    currentData = coData;
                    filteredData = pendingCoData;
                    dashboardFilter = { type: 'CO', field: 'coNumber', value: 'pending' };
                    currentTitle = 'Pending Collection';
                }
                filters = {};
                currentPage = 1;
                sortField = null;
                sortDirection = 'asc';
                displayTable();
                updateTableTitle();
            }
        });
    });
}

function applyFilters() {
    filteredData = currentData;
    if (dashboardFilter && dashboardFilter.field !== 'soNumber' && dashboardFilter.field !== 'coNumber') {
        filteredData = filteredData.filter(item => {
            const value = item[dashboardFilter.field] ?? 'N/A';
            return value.toString() === dashboardFilter.value;
        });
    }
    filteredData = filteredData.filter(item => {
        return Object.keys(filters).every(field => {
            if (!filters[field]) return true;
            const value = item[field] ?? 'N/A';
            return value.toString().toLowerCase().includes(filters[field].toLowerCase());
        });
    });
    if (sortField) {
        filteredData.sort((a, b) => {
            const valueA = a[sortField] ?? 'N/A';
            const valueB = b[sortField] ?? 'N/A';
            if (sortDirection === 'asc') {
                return valueA.toString().localeCompare(valueB.toString(), undefined, { numeric: true });
            } else {
                return valueB.toString().localeCompare(valueA.toString(), undefined, { numeric: true });
            }
        });
    }
    currentPage = 1;
    displayTable();
}

function exportToExcel(data, fields, title) {
    try {
        const worksheet = XLSX.utils.json_to_sheet(data.map(item => {
            const row = {};
            fields.forEach(field => {
                row[field.replace(/_/g, ' ').replace(/\b\w/g, c => c.toUpperCase())] = item[field] ?? 'N/A';
            });
            return row;
        }));
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Data');
        const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([excelBuffer], { type: 'application/octet-stream' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        const today = formatDateForFilename(new Date());
        a.download = `${title} ${today}.xlsx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
    } catch (error) {
        console.error('Export to Excel failed:', error);
        document.getElementById('error').textContent = `Error exporting data: ${error.message}`;
        document.getElementById('error').style.display = 'block';
    }
}

function displayTable() {
    document.getElementById('loading').style.display = 'none';
    const table = document.getElementById('data-table');
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');
    const debugDiv = document.getElementById('debug');

    thead.children[0].innerHTML = '';
    thead.children[1].innerHTML = '';
    tbody.innerHTML = '';

    const headers = currentData === soData ? soFields : coFields;

    const serialTh = document.createElement('th');
    serialTh.textContent = 'Serial No';
    thead.children[0].appendChild(serialTh);

    headers.forEach(header => {
        const th = document.createElement('th');
        th.textContent = header.replace(/_/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
        if (header === sortField) {
            th.classList.add(sortDirection === 'asc' ? 'sort-asc' : 'sort-desc');
        }
        th.addEventListener('click', () => {
            if (sortField === header) {
                sortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
            } else {
                sortField = header;
                sortDirection = 'asc';
            }
            applyFilters();
        });
        thead.children[0].appendChild(th);
    });

    const serialTd = document.createElement('td');
    thead.children[1].appendChild(serialTd);

    headers.forEach(header => {
        const td = document.createElement('td');
        if (header !== 'process') {
            const input = document.createElement('input');
            input.type = 'text';
            input.placeholder = 'Search (press Enter)';
            input.value = filters[header] || '';
            input.addEventListener('keypress', (e) => {
                if (e.key === 'Enter') {
                    filters[header] = e.target.value;
                    applyFilters();
                }
            });
            td.appendChild(input);
        }
        thead.children[1].appendChild(td);
    });

    table.style.display = 'table';
    debugDiv.style.display = 'none';

    if (!filteredData || filteredData.length === 0) {
        const tr = document.createElement('tr');
        const td = document.createElement('td');
        td.className = 'no-results';
        td.colSpan = headers.length + 1;
        td.textContent = 'No results found';
        tr.appendChild(td);
        tbody.appendChild(tr);
        updatePagination();
        return;
    }

    const start = (currentPage - 1) * rowsPerPage;
    const end = start + rowsPerPage;
    const paginatedData = filteredData.slice(start, end);

    paginatedData.forEach((item, index) => {
        const tr = document.createElement('tr');
        const serialTd = document.createElement('td');
        serialTd.textContent = start + index + 1;
        tr.appendChild(serialTd);
        headers.forEach(header => {
            const td = document.createElement('td');
            td.textContent = item[header] ?? 'N/A';
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });

    updatePagination();
}

function updatePagination() {
    const totalPages = Math.ceil(filteredData.length / rowsPerPage);
    document.getElementById('page-info').textContent = `Page ${currentPage} of ${totalPages || 1}`;
    
    const firstPageBtn = document.getElementById('first-page');
    const prevPageBtn = document.getElementById('prev-page');
    const nextPageBtn = document.getElementById('next-page');
    const lastPageBtn = document.getElementById('last-page');
    const pageSelect = document.getElementById('page-select');

    firstPageBtn.disabled = currentPage === 1;
    prevPageBtn.disabled = currentPage === 1;
    nextPageBtn.disabled = currentPage === totalPages || totalPages === 0;
    lastPageBtn.disabled = currentPage === totalPages || totalPages === 0;

    pageSelect.innerHTML = '';
    for (let i = 1; i <= totalPages; i++) {
        const option = document.createElement('option');
        option.value = i;
        option.textContent = i;
        if (i === currentPage) {
            option.selected = true;
        }
        pageSelect.appendChild(option);
    }

    firstPageBtn.onclick = () => {
        if (currentPage !== 1) {
            currentPage = 1;
            displayTable();
        }
    };

    prevPageBtn.onclick = () => {
        if (currentPage > 1) {
            currentPage--;
            displayTable();
        }
    };

    nextPageBtn.onclick = () => {
        if (currentPage < totalPages) {
            currentPage++;
            displayTable();
        }
    };

    lastPageBtn.onclick = () => {
        if (currentPage !== totalPages) {
            currentPage = totalPages;
            displayTable();
        }
    };

    pageSelect.onchange = (e) => {
        currentPage = parseInt(e.target.value);
        displayTable();
    };
}

const soFields = ['process', 'do_number', 'Line_Id', 'acc_approved_date_time', 'order_type', 'Sales_Org', 'Destination_Channel', 'Division', 'Sold_To_Party', 'Material_Code', 'Target_Qty', 'Target_Qty_Pcs', 'plant_code', 'Discount', 'CCA', 'TransactionId', 'serverAllocation', 'soNumber', 'Division2'];
const coFields = ['process', 'pay_id', 'order_payment_id', 'payment_number', 'receiving_company_code', 'customerName', 'customer_code', 'bankName', 'bankBranchName', 'routing_code', 'instrument_no', 'currentTime', 'instrument_issue_date', 'instrument_amount', 'instrument_type', 'instrument_recpt_date', 'compnay_code', 'credit_control_area', 'line', 'allocaton_amount', 'bank_account_number', 'house_bank', 'orderNumber', 'approvedTime', 'TransactionId', 'serverAllocation', 'coNumber', 'allocation_sequence', 'approvedBy'];

document.getElementById('so-button').onclick = () => {
    currentData = soData;
    filteredData = soData;
    filters = {};
    dashboardFilter = null;
    currentPage = 1;
    sortField = null;
    sortDirection = 'asc';
    currentTitle = 'Total SO';
    displayTable();
    updateTableTitle();
};
document.getElementById('co-button').onclick = () => {
    currentData = coData;
    filteredData = coData;
    filters = {};
    dashboardFilter = null;
    currentPage = 1;
    sortField = null;
    sortDirection = 'asc';
    currentTitle = 'Collection';
    displayTable();
    updateTableTitle();
};
document.getElementById('export-data').onclick = () => {
    const fields = currentData === soData ? soFields : coFields;
    exportToExcel(filteredData, fields, currentTitle);
};
document.getElementById('reset-filter').onclick = () => {
    currentData = coData;
    filteredData = coData;
    filters = {};
    dashboardFilter = null;
    currentPage = 1;
    sortField = null;
    sortDirection = 'asc';
    currentTitle = 'Collection';
    displayTable();
    updateTableTitle();
};
document.getElementById('refresh-button').onclick = () => {
    pageLoadTime = new Date();
    fetchData();
};
document.getElementById('apply-date-line').onclick = () => {
    fetchData();
};

window.onload = () => {
    pageLoadTime = new Date();
    setDefaultDates();
    fetchData();
};