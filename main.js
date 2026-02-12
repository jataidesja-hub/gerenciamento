// State Management
let salesData = [];
let charts = {};
let CONFIG = {
    apiUrl: localStorage.getItem('gerenciamento_api_url') || ''
};

// DOM Elements
const tabs = document.querySelectorAll('.nav-item');
const tabContents = document.querySelectorAll('.tab-content');
const tabTitle = document.getElementById('tab-title');
const salesList = document.getElementById('sales-list');
const saleForm = document.getElementById('sale-form');
const notification = document.getElementById('notification');
const apiUrlInput = document.getElementById('api-url');

// Initialization
document.addEventListener('DOMContentLoaded', () => {
    initTabs();
    if (CONFIG.apiUrl) {
        apiUrlInput.value = CONFIG.apiUrl;
        refreshData();
    } else {
        showTab('settings');
        showNotification('Configure a URL da API para começar.', 'info');
    }
});

// Tab Navigation
function initTabs() {
    tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            const tabId = tab.getAttribute('data-tab');
            showTab(tabId);
        });
    });
}

function showTab(tabId) {
    tabs.forEach(t => t.classList.remove('active'));
    tabContents.forEach(c => c.classList.remove('active'));

    const activeTab = document.querySelector(`[data-tab="${tabId}"]`);
    const activeContent = document.getElementById(tabId);

    if (activeTab && activeContent) {
        activeTab.classList.add('active');
        activeContent.classList.add('active');
        tabTitle.textContent = activeTab.textContent.trim();
    }
}

// API Interaction
async function refreshData() {
    if (!CONFIG.apiUrl) return;

    showNotification('Carregando dados...', 'info');
    try {
        const response = await fetch(`${CONFIG.apiUrl}?action=getSales`);
        const data = await response.json();

        if (data.error) throw new Error(data.error);

        salesData = Array.isArray(data) ? data : [];
        renderDashboard();
        renderCharts();
        showNotification('Dados atualizados!', 'success');
    } catch (error) {
        console.error(error);
        showNotification('Erro ao carregar dados. Verifique a URL da API.', 'error');
    }
}

async function saveSale(sale) {
    if (!CONFIG.apiUrl) {
        showNotification('Configure a URL da API primeiro!', 'error');
        showTab('settings');
        return;
    }

    showNotification('Salvando...', 'info');
    try {
        const response = await fetch(CONFIG.apiUrl, {
            method: 'POST',
            body: JSON.stringify({ action: 'saveSale', sale })
        });

        showNotification('Venda processada!', 'success');
        resetForm();
        showTab('dashboard');
        setTimeout(refreshData, 1500);
    } catch (error) {
        console.error(error);
        showNotification('Erro ao salvar. Verifique o console.', 'error');
    }
}

// Rendering
function renderDashboard() {
    salesList.innerHTML = '';

    const totalSales = salesData.length;
    const totalRevenue = salesData.reduce((acc, s) => acc + (parseFloat(s['Valor Total (R$)']) || 0), 0);
    const pendingSales = salesData.filter(s => s['Status do Pagamento'] === 'Em aberto').length;

    document.getElementById('total-sales').textContent = totalSales;
    document.getElementById('total-revenue').textContent = `R$ ${totalRevenue.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}`;
    document.getElementById('pending-sales').textContent = pendingSales;

    salesData.slice().reverse().forEach(sale => {
        const tr = document.createElement('tr');
        const status = sale['Status do Pagamento'] || 'Pendente';
        const statusClass = `status-${status.toLowerCase().replace(/\s+/g, '-')}`;

        tr.innerHTML = `
            <td>#${sale['ID da Venda']?.toString().split('-').pop() || '---'}</td>
            <td><strong>${sale['Nome do Cliente'] || 'N/A'}</strong></td>
            <td><span class="status ${statusClass}">${status}</span></td>
            <td>${formatDate(sale['Data da Compra'])}</td>
            <td>R$ ${parseFloat(sale['Valor Total (R$)'] || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</td>
            <td>${sale['Responsável'] || '---'}</td>
            <td>
                <button class="btn btn-sm btn-outline" onclick="editSale(${sale.rowIndex})">✏️</button>
            </td>
        `;
        salesList.appendChild(tr);
    });
}

function renderCharts() {
    Object.values(charts).forEach(chart => chart.destroy());

    const ctxStatus = document.getElementById('status-chart').getContext('2d');
    const ctxRevenue = document.getElementById('revenue-chart').getContext('2d');

    const statusCounts = {};
    salesData.forEach(s => {
        const st = s['Status do Pagamento'] || 'Outros';
        statusCounts[st] = (statusCounts[st] || 0) + 1;
    });

    charts.status = new Chart(ctxStatus, {
        type: 'doughnut',
        data: {
            labels: Object.keys(statusCounts),
            datasets: [{
                data: Object.values(statusCounts),
                backgroundColor: ['#10b981', '#f59e0b', '#3b82f6', '#ef4444', '#94a3b8'],
                borderWidth: 0
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom', labels: { color: '#94a3b8', font: { family: 'Outfit' } } }
            }
        }
    });

    const revenueByMonth = {};
    salesData.forEach(s => {
        const date = new Date(s['Data da Compra']);
        if (!isNaN(date)) {
            const monthYear = date.toLocaleDateString('pt-BR', { month: 'short', year: '2-digit' });
            revenueByMonth[monthYear] = (revenueByMonth[monthYear] || 0) + (parseFloat(s['Valor Total (R$)']) || 0);
        }
    });

    charts.revenue = new Chart(ctxRevenue, {
        type: 'line',
        data: {
            labels: Object.keys(revenueByMonth),
            datasets: [{
                label: 'Faturamento R$',
                data: Object.values(revenueByMonth),
                borderColor: '#c1a173',
                backgroundColor: 'rgba(193, 161, 115, 0.2)',
                fill: true,
                tension: 0.4
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: { grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#94a3b8' } },
                x: { grid: { display: false }, ticks: { color: '#94a3b8' } }
            },
            plugins: {
                legend: { display: false }
            }
        }
    });
}

// Form Handling
saleForm.addEventListener('submit', (e) => {
    e.preventDefault();
    const formData = new FormData(saleForm);
    const sale = {};
    formData.forEach((value, key) => {
        sale[key] = value;
    });

    saveSale(sale);
});

function resetForm() {
    saleForm.reset();
    const rowInput = saleForm.querySelector('[name="rowIndex"]');
    if (rowInput) rowInput.remove();
}

// Settings
function saveSettings() {
    const url = apiUrlInput.value.trim();
    if (url) {
        localStorage.setItem('gerenciamento_api_url', url);
        CONFIG.apiUrl = url;
        showNotification('Configuração salva!', 'success');
        refreshData();
    }
}

// Utilities
function showNotification(message, type = 'info') {
    notification.textContent = message;
    notification.className = `notification show ${type}`;
    setTimeout(() => {
        notification.classList.remove('show');
    }, 3000);
}

function formatDate(dateStr) {
    if (!dateStr) return '---';
    try {
        const date = new Date(dateStr);
        if (isNaN(date)) return dateStr;
        return date.toLocaleDateString('pt-BR');
    } catch {
        return dateStr;
    }
}

function editSale(rowIndex) {
    const sale = salesData.find(s => s.rowIndex === rowIndex);
    if (!sale) return;

    showTab('new-sale');
    resetForm();

    Object.keys(sale).forEach(key => {
        const input = saleForm.querySelector(`[name="${key}"]`);
        if (input) {
            if (input.type === 'date' && sale[key]) {
                const d = new Date(sale[key]);
                if (!isNaN(d)) {
                    input.value = d.toISOString().split('T')[0];
                }
            } else {
                input.value = sale[key];
            }
        }
    });

    let rowInput = document.createElement('input');
    rowInput.type = 'hidden';
    rowInput.name = 'rowIndex';
    rowInput.value = rowIndex;
    saleForm.appendChild(rowInput);
}
