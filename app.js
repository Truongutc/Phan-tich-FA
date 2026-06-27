let revNpatChartInstance = null;
let epsEquityChartInstance = null;

document.addEventListener("DOMContentLoaded", () => {
    loadStockIndex();
    
    const selector = document.getElementById("stock-select");
    selector.addEventListener("change", (e) => {
        loadStockData(e.target.value);
    });
});

async function loadStockIndex() {
    try {
        const response = await fetch("data/index.json");
        const indexData = await response.json();
        
        const selector = document.getElementById("stock-select");
        selector.innerHTML = '<option value="" disabled selected>Chọn một mã cổ phiếu</option>';
        
        if (indexData.length === 0) {
            selector.innerHTML = '<option value="" disabled>Chưa có dữ liệu phân tích</option>';
            return;
        }
        
        indexData.forEach(stock => {
            const opt = document.createElement("option");
            opt.value = stock.ticker;
            opt.textContent = `${stock.ticker} - ${stock.companyName}`;
            selector.appendChild(opt);
        });
    } catch (error) {
        console.error("Lỗi khi tải danh sách cổ phiếu:", error);
        const selector = document.getElementById("stock-select");
        selector.innerHTML = '<option value="" disabled>Lỗi tải danh sách</option>';
    }
}

async function loadStockData(ticker) {
    try {
        // Toggle view visibility
        document.getElementById("welcome-view").classList.add("hidden");
        document.getElementById("analysis-view").classList.remove("hidden");
        
        const response = await fetch(`data/${ticker}.json`);
        const stock = await response.json();
        
        // Update header details
        document.getElementById("ticker-badge").textContent = stock.ticker;
        document.getElementById("company-name").textContent = stock.companyName;
        document.getElementById("company-sector").textContent = stock.sector;
        
        // Set download buttons links
        const btnPdf = document.getElementById("download-pdf");
        const btnExcel = document.getElementById("download-excel");
        
        if (stock.gdrivePdfUrl) {
            btnPdf.href = stock.gdrivePdfUrl;
            btnPdf.classList.remove("hidden");
        } else {
            btnPdf.classList.add("hidden");
        }
        
        if (stock.gdriveExcelUrl) {
            btnExcel.href = stock.gdriveExcelUrl;
            btnExcel.classList.remove("hidden");
        } else {
            btnExcel.classList.add("hidden");
        }
        
        // Update basic metrics
        document.getElementById("metric-price").textContent = `${formatNumber(stock.currentPrice)} VND`;
        document.getElementById("metric-mcap").textContent = `${formatNumber(Math.round(stock.marketCap / 1e9))} tỷ VND`;
        document.getElementById("metric-shares").textContent = `${formatNumber(stock.shares)} CP`;
        
        // Render table
        renderFinancialTable(stock.data);
        
        // Render charts
        renderCharts(stock.data);
        
    } catch (error) {
        console.error("Lỗi khi tải dữ liệu chi tiết của mã:", error);
    }
}

function renderFinancialTable(financialData) {
    const table = document.getElementById("financial-table");
    const theadRow = table.querySelector("thead tr");
    const tbody = table.querySelector("tbody");
    
    // Clear dynamic parts
    theadRow.innerHTML = "<th>Chỉ tiêu</th>";
    tbody.innerHTML = "";
    
    // Build year headers
    const years = financialData.years;
    years.forEach(year => {
        const th = document.createElement("th");
        th.textContent = year;
        theadRow.appendChild(th);
    });
    
    // Define rows data structure
    const rowDefinitions = [
        { label: "Doanh thu (tỷ VND)", key: "revenue", precision: 1 },
        { label: "Lợi nhuận sau thuế (tỷ VND)", key: "npat", precision: 1 },
        { label: "EPS (VND)", key: "eps", precision: 0 },
        { label: "Vốn chủ sở hữu (tỷ VND)", key: "equity", precision: 1 }
    ];
    
    rowDefinitions.forEach(def => {
        const tr = document.createElement("tr");
        
        const tdLabel = document.createElement("td");
        tdLabel.innerHTML = `<strong>${def.label}</strong>`;
        tr.appendChild(tdLabel);
        
        const values = financialData[def.key];
        values.forEach(val => {
            const tdVal = document.createElement("td");
            tdVal.textContent = val !== null && val !== undefined ? formatNumber(val, def.precision) : "-";
            tr.appendChild(tdVal);
        });
        
        tbody.appendChild(tr);
    });
}

function renderCharts(financialData) {
    // 1. Revenue & NPAT Chart
    const ctx1 = document.getElementById("revNpatChart").getContext("2d");
    if (revNpatChartInstance) {
        revNpatChartInstance.destroy();
    }
    
    revNpatChartInstance = new Chart(ctx1, {
        type: 'bar',
        data: {
            labels: financialData.years,
            datasets: [
                {
                    label: 'Doanh thu (tỷ)',
                    data: financialData.revenue,
                    backgroundColor: 'rgba(59, 130, 246, 0.65)',
                    borderColor: 'rgba(59, 130, 246, 1)',
                    borderWidth: 1,
                    yAxisID: 'y'
                },
                {
                    label: 'Lợi nhuận sau thuế (tỷ)',
                    data: financialData.npat,
                    type: 'line',
                    borderColor: 'rgba(236, 72, 153, 1)',
                    backgroundColor: 'rgba(236, 72, 153, 0.2)',
                    borderWidth: 3,
                    tension: 0.3,
                    fill: false,
                    yAxisID: 'y1'
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    labels: { color: '#f3f4f6' }
                }
            },
            scales: {
                x: {
                    grid: { color: 'rgba(255, 255, 255, 0.05)' },
                    ticks: { color: '#9ca3af' }
                },
                y: {
                    type: 'linear',
                    display: true,
                    position: 'left',
                    grid: { color: 'rgba(255, 255, 255, 0.05)' },
                    ticks: { color: '#9ca3af' },
                    title: { display: true, text: 'Doanh thu (tỷ)', color: '#9ca3af' }
                },
                y1: {
                    type: 'linear',
                    display: true,
                    position: 'right',
                    grid: { drawOnChartArea: false },
                    ticks: { color: '#9ca3af' },
                    title: { display: true, text: 'LNST (tỷ)', color: '#9ca3af' }
                }
            }
        }
    });

    // 2. EPS & Equity Chart
    const ctx2 = document.getElementById("epsEquityChart").getContext("2d");
    if (epsEquityChartInstance) {
        epsEquityChartInstance.destroy();
    }
    
    epsEquityChartInstance = new Chart(ctx2, {
        type: 'bar',
        data: {
            labels: financialData.years,
            datasets: [
                {
                    label: 'Vốn chủ sở hữu (tỷ)',
                    data: financialData.equity,
                    backgroundColor: 'rgba(16, 185, 129, 0.65)',
                    borderColor: 'rgba(16, 185, 129, 1)',
                    borderWidth: 1,
                    yAxisID: 'y'
                },
                {
                    label: 'EPS (VND)',
                    data: financialData.eps,
                    type: 'line',
                    borderColor: 'rgba(245, 158, 11, 1)',
                    backgroundColor: 'rgba(245, 158, 11, 0.2)',
                    borderWidth: 3,
                    tension: 0.3,
                    fill: false,
                    yAxisID: 'y1'
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    labels: { color: '#f3f4f6' }
                }
            },
            scales: {
                x: {
                    grid: { color: 'rgba(255, 255, 255, 0.05)' },
                    ticks: { color: '#9ca3af' }
                },
                y: {
                    type: 'linear',
                    display: true,
                    position: 'left',
                    grid: { color: 'rgba(255, 255, 255, 0.05)' },
                    ticks: { color: '#9ca3af' },
                    title: { display: true, text: 'VCSH (tỷ)', color: '#9ca3af' }
                },
                y1: {
                    type: 'linear',
                    display: true,
                    position: 'right',
                    grid: { drawOnChartArea: false },
                    ticks: { color: '#9ca3af' },
                    title: { display: true, text: 'EPS (VND)', color: '#9ca3af' }
                }
            }
        }
    });
}

function formatNumber(num, precision = 0) {
    if (num === null || num === undefined || isNaN(num)) return "-";
    return num.toLocaleString("vi-VN", {
        minimumFractionDigits: precision,
        maximumFractionDigits: precision
    });
}
