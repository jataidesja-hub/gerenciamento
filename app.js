// State Management
let salesData = [];
let installmentsData = [];
let charts = {};
let filteredClients = null;
let CONFIG = {
    apiUrl: 'https://script.google.com/macros/s/AKfycbxqTVpqoUI2NV5IUWZRaYNQjBX-LetKZ6Tg37SxbZvxlkBUMxNTuTu5_hgVZcO93RIQmA/exec'
};

// DOM Elements
const tabs = document.querySelectorAll('.nav-item');
const tabContents = document.querySelectorAll('.tab-content');
const tabTitle = document.getElementById('tab-title');
const salesList = document.getElementById('sales-list');
const saleForm = document.getElementById('sale-form');
const notification = document.getElementById('notification');

// Initialization
document.addEventListener('DOMContentLoaded', () => {
    initTabs();
    if (CONFIG.apiUrl) {
        refreshData();
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

    if (tabId === 'clients') {
        renderClients();
    }
}

// API Interaction
async function refreshData() {
    if (!CONFIG.apiUrl) return;

    showNotification('Carregando dados...', 'info');
    try {
        const [salesRes, installRes] = await Promise.all([
            fetch(`${CONFIG.apiUrl}?action=getSales`),
            fetch(`${CONFIG.apiUrl}?action=getInstallments`)
        ]);

        const salesJson = await salesRes.json();
        const installJson = await installRes.json();

        if (salesJson.error) throw new Error(salesJson.error);

        salesData = Array.isArray(salesJson) ? salesJson : [];
        installmentsData = Array.isArray(installJson) ? installJson : [];

        renderDashboard();
        renderCharts();
        renderClients();
        showNotification('Dados atualizados!', 'success');
    } catch (error) {
        console.error(error);
        showNotification('Erro ao carregar dados. Verifique a URL da API.', 'error');
    }
}

async function saveSale(sale) {
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

async function markInstallmentPaid(saleId, installmentNumber) {
    showNotification('Registrando pagamento...', 'info');
    try {
        await fetch(CONFIG.apiUrl, {
            method: 'POST',
            body: JSON.stringify({ action: 'payInstallment', saleId, installmentNumber })
        });
        showNotification('Parcela marcada como paga!', 'success');
        setTimeout(refreshData, 1000);
    } catch (error) {
        console.error(error);
        showNotification('Erro ao registrar pagamento.', 'error');
    }
}

// Dashboard Rendering
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

// ===== CLIENTS TAB =====

function getClientGroups() {
    const groups = {};
    salesData.forEach(sale => {
        const name = (sale['Nome do Cliente'] || 'Sem Nome').trim();
        if (!groups[name]) {
            groups[name] = {
                name: name,
                city: sale['Cidade/UF'] || '',
                phone: sale['Telefone / WhatsApp'] || '',
                sales: []
            };
        }
        // Update contact info if more recent
        if (sale['Cidade/UF']) groups[name].city = sale['Cidade/UF'];
        if (sale['Telefone / WhatsApp']) groups[name].phone = sale['Telefone / WhatsApp'];
        groups[name].sales.push(sale);
    });
    return groups;
}

function getInstallmentsForSale(saleId) {
    return installmentsData.filter(i => i['ID Venda'] === saleId);
}

function renderClients() {
    const clientsList = document.getElementById('clients-list');
    if (!clientsList) return;

    const groups = getClientGroups();
    const clientNames = Object.keys(groups);

    // Apply search filter
    let displayClients = clientNames;
    if (filteredClients !== null) {
        displayClients = filteredClients;
    }

    // Calculate totals
    let totalPaid = 0;
    let totalPending = 0;

    installmentsData.forEach(inst => {
        if (inst['Status'] === 'Pago') totalPaid++;
        else totalPending++;
    });

    const totalClientsEl = document.getElementById('total-clients');
    const totalPaidEl = document.getElementById('total-paid-installments');
    const totalPendingEl = document.getElementById('total-pending-installments');

    if (totalClientsEl) totalClientsEl.textContent = clientNames.length;
    if (totalPaidEl) totalPaidEl.textContent = totalPaid;
    if (totalPendingEl) totalPendingEl.textContent = totalPending;

    if (displayClients.length === 0) {
        clientsList.innerHTML = `
            <div class="no-clients">
                <div class="empty-icon">👥</div>
                <p>Nenhum cliente encontrado</p>
            </div>`;
        return;
    }

    let html = '';
    displayClients.forEach(name => {
        const client = groups[name];
        if (!client) return;

        const totalValue = client.sales.reduce((acc, s) => acc + (parseFloat(s['Valor Total (R$)']) || 0), 0);
        const initials = name.split(' ').map(w => w[0]).join('').substring(0, 2).toUpperCase();
        const safeId = name.replace(/[^a-zA-Z0-9]/g, '_');

        // Get all installments for this client
        let clientPaid = 0;
        let clientTotal = 0;
        client.sales.forEach(sale => {
            const saleInstallments = getInstallmentsForSale(sale['ID da Venda']);
            saleInstallments.forEach(inst => {
                clientTotal++;
                if (inst['Status'] === 'Pago') clientPaid++;
            });
        });

        html += `
            <div class="client-card">
                <div class="client-header" onclick="toggleClientDetails('${safeId}')">
                    <div class="client-info">
                        <div class="client-avatar">${initials}</div>
                        <div>
                            <div class="client-name">${name}</div>
                            <div class="client-meta">${client.city || 'Sem cidade'} · ${client.phone || 'Sem telefone'}</div>
                        </div>
                    </div>
                    <div class="client-stats">
                        <div class="client-stat">
                            <div class="value">${client.sales.length}</div>
                            <div class="label">Compras</div>
                        </div>
                        <div class="client-stat">
                            <div class="value">R$ ${totalValue.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</div>
                            <div class="label">Total</div>
                        </div>
                        <button class="client-toggle" id="toggle-${safeId}">▼</button>
                    </div>
                </div>
                <div class="client-details" id="details-${safeId}">
                    ${renderClientSales(client)}
                </div>
            </div>`;
    });

    clientsList.innerHTML = html;
}

function renderClientSales(client) {
    let html = '';
    client.sales.forEach(sale => {
        const saleId = sale['ID da Venda'];
        const shortId = saleId?.toString().split('-').pop() || '---';
        const status = sale['Status do Pagamento'] || 'Pendente';
        const statusClass = `status-${status.toLowerCase().replace(/\s+/g, '-')}`;
        const totalValue = parseFloat(sale['Valor Total (R$)'] || 0);
        const installments = getInstallmentsForSale(saleId);
        const numParcelas = parseInt(sale['Parcelas']) || 1;

        const paidCount = installments.filter(i => i['Status'] === 'Pago').length;
        const pendingCount = installments.length - paidCount;
        const progressPct = installments.length > 0 ? Math.round((paidCount / installments.length) * 100) : 0;

        html += `
            <div class="sale-block">
                <div class="sale-block-header">
                    <h5>Venda #${shortId} — ${formatDate(sale['Data da Compra'])}</h5>
                    <div>
                        <span class="status ${statusClass}">${status}</span>
                        <span style="margin-left:8px; color: var(--text-muted); font-size:0.85rem;">
                            R$ ${totalValue.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                        </span>
                    </div>
                </div>`;

        if (installments.length > 0) {
            html += `
                <div style="margin-bottom:8px; font-size:0.85rem; color: var(--text-muted);">
                    ✅ ${paidCount} paga(s) · ⏳ ${pendingCount} pendente(s) de ${installments.length}
                </div>
                <div class="progress-bar-container">
                    <div class="progress-bar" style="width: ${progressPct}%"></div>
                </div>
                <div class="installments-grid" style="margin-top: 0.75rem;">`;

            installments.forEach(inst => {
                const isPaid = inst['Status'] === 'Pago';
                const instClass = isPaid ? 'paid' : 'pending';
                const instValue = parseFloat(inst['Valor (R$)'] || 0);

                html += `
                    <div class="installment-item ${instClass}">
                        <div>
                            <span class="installment-label">${isPaid ? '✅' : '⏳'} ${inst['Nº Parcela']}ª</span>
                            <div class="installment-value">R$ ${instValue.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</div>
                        </div>
                        ${!isPaid ? `<button class="installment-action" onclick="markInstallmentPaid('${saleId}', ${inst['Nº Parcela']})" title="Marcar como paga">💰</button>` : `<span style="color: var(--success); font-size: 0.8rem;">${formatDate(inst['Data Pagamento'])}</span>`}
                    </div>`;
            });

            html += '</div>';
        } else {
            // No installments tracked — show basic info
            html += `
                <div style="color: var(--text-muted); font-size:0.85rem; padding: 0.5rem 0;">
                    ${numParcelas} parcela(s) — Sem controle detalhado de parcelas
                </div>`;
        }

        html += '</div>';
    });

    return html;
}

function toggleClientDetails(safeId) {
    const details = document.getElementById(`details-${safeId}`);
    const toggle = document.getElementById(`toggle-${safeId}`);
    if (details && toggle) {
        details.classList.toggle('open');
        toggle.classList.toggle('open');
    }
}

function searchClients(query) {
    const groups = getClientGroups();
    const all = Object.keys(groups);

    if (!query || query.trim() === '') {
        filteredClients = null;
    } else {
        const q = query.toLowerCase();
        filteredClients = all.filter(name => {
            const client = groups[name];
            return name.toLowerCase().includes(q) ||
                (client.city && client.city.toLowerCase().includes(q)) ||
                (client.phone && client.phone.toLowerCase().includes(q));
        });
    }
    renderClients();
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

// Service Worker Registration
if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
        navigator.serviceWorker.register('./sw.js')
            .then(reg => console.log('SW registrado', reg))
            .catch(err => console.log('SW falhou', err));
    });
}

// Install Prompt Logic
let deferredPrompt;
const installBanner = document.getElementById('install-banner');
const btnInstall = document.getElementById('btn-install');
const btnCloseBanner = document.getElementById('btn-close-banner');

window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferredPrompt = e;
    setTimeout(() => {
        if (installBanner) installBanner.classList.add('show');
    }, 2000);
});

if (btnInstall) {
    btnInstall.addEventListener('click', async () => {
        if (deferredPrompt) {
            deferredPrompt.prompt();
            const { outcome } = await deferredPrompt.userChoice;
            console.log(`User response to the install prompt: ${outcome}`);
            deferredPrompt = null;
            installBanner.classList.remove('show');
        }
    });
}

if (btnCloseBanner) {
    btnCloseBanner.addEventListener('click', () => {
        installBanner.classList.remove('show');
    });
}
