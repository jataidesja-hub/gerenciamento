// State Management
let salesData = [];
let CONFIG = {
    // URL de implantação fornecida pelo usuário
    apiUrl: localStorage.getItem('agrovale_api_url') || 'https://script.google.com/macros/library/d/144eIrGSNScm3OfB6Ag7lrPmIMJyHXZBahNnj_d3-eQRhJdZ4hi7g0SzL/1'
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

    activeTab.classList.add('active');
    activeContent.classList.add('active');
    tabTitle.textContent = activeTab.textContent.trim();
}

// API Interaction
async function refreshData() {
    if (!CONFIG.apiUrl) return;

    showNotification('Carregando dados...', 'info');
    try {
        const response = await fetch(`${CONFIG.apiUrl}?action=getSales`);
        const data = await response.json();

        if (data.error) throw new Error(data.error);

        salesData = data;
        renderDashboard();
        showNotification('Dados atualizados!', 'success');
    } catch (error) {
        console.error(error);
        showNotification('Erro ao carregar dados: ' + error.message, 'error');
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
            mode: 'no-cors', // Apps Script issues with CORS in some environments
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ action: 'saveSale', sale })
        });

        // Note: with no-cors we can't read the response body, but usually it succeeds
        showNotification('Venda salva com sucesso!', 'success');
        resetForm();
        showTab('dashboard');
        setTimeout(refreshData, 1000); // Reload data after a short delay
    } catch (error) {
        showNotification('Erro ao salvar: ' + error.message, 'error');
    }
}

// Rendering
function renderDashboard() {
    salesList.innerHTML = '';

    // Stats
    const totalSales = salesData.length;
    const totalRevenue = salesData.reduce((acc, s) => acc + (parseFloat(s['Valor Total (R$)']) || 0), 0);
    const pendingSales = salesData.filter(s => s['Status do Pagamento'] === 'Em aberto').length;

    document.getElementById('total-sales').textContent = totalSales;
    document.getElementById('total-revenue').textContent = `R$ ${totalRevenue.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}`;
    document.getElementById('pending-sales').textContent = pendingSales;

    // Table
    salesData.forEach(sale => {
        const tr = document.createElement('tr');
        const statusClass = `status-${sale['Status do Pagamento']?.toLowerCase().replace(' ', '-') || 'default'}`;

        tr.innerHTML = `
            <td>#${sale['ID da Venda'] || '---'}</td>
            <td><strong>${sale['Nome do Cliente'] || 'N/A'}</strong></td>
            <td><span class="status ${statusClass}">${sale['Status do Pagamento'] || 'Pendente'}</span></td>
            <td>${formatDate(sale['Data da Compra'])}</td>
            <td>R$ ${parseFloat(sale['Valor Total (R$)']).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</td>
            <td>${sale['Responsável'] || '---'}</td>
            <td>
                <button class="btn btn-sm btn-outline" onclick="editSale(${sale.rowIndex})">✏️</button>
            </td>
        `;
        salesList.appendChild(tr);
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
}

// Settings
function saveSettings() {
    const url = apiUrlInput.value.trim();
    if (url) {
        localStorage.setItem('agrovale_api_url', url);
        CONFIG.apiUrl = url;
        showNotification('URL da API salva!', 'success');
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
        return date.toLocaleDateString('pt-BR');
    } catch {
        return dateStr;
    }
}

function editSale(rowIndex) {
    const sale = salesData.find(s => s.rowIndex === rowIndex);
    if (!sale) return;

    showTab('new-sale');
    // Map object properties to form inputs
    Object.keys(sale).forEach(key => {
        const input = saleForm.querySelector(`[name="${key}"]`);
        if (input) {
            if (input.type === 'date') {
                input.value = new Date(sale[key]).toISOString().split('T')[0];
            } else {
                input.value = sale[key];
            }
        }
    });

    // Store rowIndex in form for updates
    let rowInput = saleForm.querySelector('[name="rowIndex"]');
    if (!rowInput) {
        rowInput = document.createElement('input');
        rowInput.type = 'hidden';
        rowInput.name = 'rowIndex';
        saleForm.appendChild(rowInput);
    }
    rowInput.value = rowIndex;
}
