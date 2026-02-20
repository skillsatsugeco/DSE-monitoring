const BACKEND_URL = "https://script.google.com/macros/s/AKfycbyGbGdLL8_VyfL-y9v7wFlQSEuGMHETics49oehlZud7ajn-QP_9Lo_DTZWbBh0UY8h/exec";

let allMarketData = [];
let mainChart = null;

document.addEventListener('DOMContentLoaded', () => {
    fetchMarketData();

    // Setup Search
    document.getElementById('marketSearch').addEventListener('input', (e) => {
        filterMarketTable(e.target.value);
    });
});

async function fetchMarketData() {
    const syncEl = document.getElementById('lastSync');
    try {
        syncEl.textContent = "Updating Market Map...";
        syncEl.style.color = "var(--text-secondary)";

        console.log("Fetching from:", BACKEND_URL);
        const response = await fetch(BACKEND_URL);

        if (!response.ok) throw new Error(`Server returned ${response.status}`);

        const json = await response.json();
        console.log("Data received:", json);

        if (json.status === "success" && json.data && json.data.length > 0) {
            allMarketData = getLatestPerSecurity(json.data);
            console.log("Processed data:", allMarketData);
            renderDashboard();
            hideLoader();
            updateLastSync();
        } else {
            const msg = json.message || (json.data && json.data.length === 0 ? "No data found in your spreadsheet" : "Invalid data format");
            syncEl.textContent = "⚠️ " + msg;
            syncEl.style.color = "var(--accent-red)";
            console.error("Data Error:", msg);
            hideLoader();
        }
    } catch (err) {
        console.error("Fetch Error:", err);
        syncEl.textContent = "⚠️ Connection Failed";
        syncEl.style.color = "var(--accent-red)";
        hideLoader();
    }
}

function hideLoader() {
    document.getElementById('loader').style.display = 'none';
}

function updateLastSync() {
    document.getElementById('lastSync').textContent = "Last sync: " + new Date().toLocaleTimeString();
}

function renderDashboard() {
    renderMainTable(allMarketData);
    renderMovers(allMarketData);
}

function renderMainTable(data) {
    const tbody = document.getElementById('marketBody');
    tbody.innerHTML = '';

    // Alphabetical sort
    const sorted = [...data].sort((a, b) => a.SECURITY.localeCompare(b.SECURITY));

    sorted.forEach(row => {
        const tr = document.createElement('tr');
        tr.className = 'clickable';
        tr.onclick = () => selectSecurity(row.SECURITY);

        const rvol = row.rvol ? row.rvol.toFixed(2) : '-';

        // Priority: Hype/Breakout > Momentum
        let displaySignal = row.momentumSignal || 'NEUTRAL';
        if (row.hypeRisk && row.hypeRisk !== 'NORMAL') {
            displaySignal = row.hypeRisk;
        }

        const score = row.tradeScore || 0;
        const scoreClass = score > 0 ? 'score-positive' : (score < 0 ? 'score-negative' : 'score-neutral');
        const sigClass = getSignalClass(displaySignal);

        tr.innerHTML = `
            <td style="border-left: 4px solid ${score > 3 ? 'var(--accent-green)' : 'transparent'};">
                <div style="font-weight: 700;">${row.SECURITY}</div>
                <div style="font-size: 0.7rem; color: var(--text-secondary);">${row.TIMESTAMP ? row.TIMESTAMP.toString().split(' ')[0] : ''}</div>
            </td>
            <td><strong>${row.LAST || '-'}</strong></td>
            <td class="${getColorClass(row.DoD)}">${formatPercent(row.DoD)}</td>
            <td style="color: ${parseFloat(rvol) > 1.5 ? 'var(--accent-blue)' : 'var(--text-secondary)'}">
                ${rvol}x
                <div style="font-size: 0.65rem; opacity: 0.6;">RVOL</div>
            </td>
            <td><span class="badge ${sigClass}">${displaySignal.replace('_', ' ')}</span></td>
            <td><div class="score-box ${scoreClass}">${score}</div></td>
        `;
        tbody.appendChild(tr);
    });
}

function renderMovers(data) {
    const tbody = document.getElementById('moversBody');
    tbody.innerHTML = '';

    // Sort by absolute DoD change
    const movers = [...data].sort((a, b) => Math.abs(b.DoD || 0) - Math.abs(a.DoD || 0)).slice(0, 5);

    movers.forEach(row => {
        const tr = document.createElement('tr');
        const volValue = row.VOL ?? row.VOLUME ?? row["Vol."] ?? 0;
        const formattedVol = typeof volValue === 'number' ? volValue.toLocaleString() : volValue;

        tr.innerHTML = `
            <td><strong>${row.SECURITY}</strong></td>
            <td class="${getColorClass(row.DoD)}">${formatPercent(row.DoD)}</td>
            <td style="text-align: right; color: var(--text-secondary); font-size: 0.8rem;">Vol: ${formattedVol}</td>
        `;
        tbody.appendChild(tr);
    });
}

function filterMarketTable(query) {
    const filtered = allMarketData.filter(item =>
        item.SECURITY.toLowerCase().includes(query.toLowerCase())
    );
    renderMainTable(filtered);
}

async function selectSecurity(symbol) {
    // UI Feedback
    const placeholder = document.querySelector('.chart-placeholder');
    const canvas = document.getElementById('mainChart');
    const label = document.getElementById('chartLabel');
    const symbolBadge = document.getElementById('chartSymbol');
    const signalsPanel = document.getElementById('signalDetails');

    // Find latest row for this security in allMarketData for signals
    const securityData = allMarketData.find(d => d.SECURITY === symbol);

    placeholder.innerHTML = `<div class="spinner" style="margin: 0 auto;"></div><p>Fetching history for ${symbol}...</p>`;
    placeholder.style.display = 'flex';
    canvas.style.display = 'none';
    signalsPanel.style.display = 'none';
    symbolBadge.style.display = 'inline-block';
    symbolBadge.textContent = symbol;
    label.textContent = "Intraday Performance";

    if (securityData) {
        document.getElementById('sigLiquidity').textContent = securityData.liquidityScore || 'N/A';
        document.getElementById('sigHype').textContent = (securityData.hypeRisk || 'NORMAL').replace('_', ' ');
        document.getElementById('sigAvgVol').textContent = (securityData.avgVol30 || 0).toLocaleString();
        document.getElementById('sigStable').textContent = securityData.stableTrend ? 'YES (ACCUMULATING)' : 'NO';

        // Color coding
        document.getElementById('sigLiquidity').style.color = securityData.liquidityScore === 'HIGH' ? 'var(--accent-green)' : 'var(--text-secondary)';
        document.getElementById('sigHype').style.color = securityData.hypeRisk === 'NORMAL' ? 'var(--text-secondary)' : 'var(--accent-red)';
    }

    try {
        const url = `${BACKEND_URL}${BACKEND_URL.includes('?') ? '&' : '?'}security=${encodeURIComponent(symbol)}`;
        const response = await fetch(url);
        const json = await response.json();

        if (json.status === "success" && json.data) {
            if (json.data.length > 0) {
                renderChart(json.data, symbol);
                placeholder.style.display = 'none';
                canvas.style.display = 'block';
                signalsPanel.style.display = 'block';
            } else {
                placeholder.innerHTML = `<p>🚀 No recent intraday trades found for ${symbol}.<br><span style="font-size: 0.8rem;">Showing last known daily close.</span></p>`;
                signalsPanel.style.display = 'block';
            }
        }
    } catch (e) {
        placeholder.innerHTML = `<p>⚠️ Error loading chart data.</p>`;
    }
}

function renderChart(history, symbol) {
    const ctx = document.getElementById('mainChart').getContext('2d');

    if (mainChart) mainChart.destroy();

    const labels = history.map(h => {
        if (typeof h.time === 'string' && h.time.length > 5) return h.time.substring(0, 5);
        return h.time;
    });
    const prices = history.map(h => h.price);

    mainChart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [{
                label: symbol + ' Price (TZS)',
                data: prices,
                borderColor: '#38bdf8',
                backgroundColor: 'rgba(56, 189, 248, 0.1)',
                borderWidth: 2,
                fill: true,
                tension: 0.3,
                pointRadius: 3,
                pointBackgroundColor: '#38bdf8'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: {
                    mode: 'index',
                    intersect: false,
                    backgroundColor: '#1e293b',
                    titleColor: '#94a3b8',
                    bodyColor: '#f8fafc',
                    borderColor: '#334155',
                    borderWidth: 1,
                    callbacks: {
                        label: function (context) {
                            const tick = history[context.dataIndex];
                            const price = context.parsed.y.toLocaleString();
                            const lines = [`Price: ${price}`];

                            lines.push(`Volume: ${tick.volume.toLocaleString()}`);

                            if (tick.bid || tick.ask) {
                                lines.push(`B/A: ${tick.bid.toLocaleString()} / ${tick.ask.toLocaleString()}`);
                                lines.push(`Qty: ${tick.bidQty.toLocaleString()} / ${tick.askQty.toLocaleString()}`);
                                if (tick.spread > 0) lines.push(`Spread: ${tick.spread.toLocaleString()}`);
                                if (tick.imbalance !== 0) {
                                    const side = tick.imbalance > 0 ? "Bulls" : "Bears";
                                    lines.push(`Imbalance: ${tick.imbalance} (${side})`);
                                }
                            }
                            return lines;
                        }
                    }
                }
            },
            scales: {
                y: {
                    grid: { color: 'rgba(255, 255, 255, 0.05)', drawBorder: false },
                    ticks: { color: '#94a3b8', font: { size: 10 } }
                },
                x: {
                    grid: { display: false },
                    ticks: { color: '#94a3b8', font: { size: 10 } }
                }
            }
        }
    });
}

function getLatestPerSecurity(data) {
    const latest = {};
    data.forEach(row => {
        if (!row.SECURITY) return;
        let ts = 0;
        try {
            ts = new Date(row.TIMESTAMP).getTime();
            if (isNaN(ts)) ts = 0;
        } catch (e) { ts = 0; }

        if (!latest[row.SECURITY] || ts > latest[row.SECURITY].ts) {
            latest[row.SECURITY] = { ...row, ts: ts };
        }
    });
    return Object.values(latest);
}

function formatPercent(val) {
    if (val === null || val === undefined || isNaN(val)) return "-";
    const p = (val * 100).toFixed(2);
    return (val > 0 ? "+" : "") + p + "%";
}

function getColorClass(val) {
    if (val === null || val === undefined || isNaN(val)) return "neutral";
    const num = parseFloat(val);
    if (Math.abs(num) < 0.000001) return "neutral";
    return num > 0 ? "up" : "down";
}

function getSignalClass(sig) {
    switch (sig) {
        case 'CONFIRMED_UP': return 'badge-confirmed';
        case 'WEAK_UP': return 'badge-weak';
        case 'STRONG_SELL': return 'badge-sell';
        case 'HYPE_RISK': return 'badge-hype';
        case 'BREAKOUT': return 'badge-breakout';
        default: return 'badge-neutral';
    }
}
