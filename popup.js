// REPLACE THIS with your deployed Web App URL from analytics_backend.js
const BACKEND_URL = "https://script.google.com/macros/s/AKfycbzujsNLQd1zzuVsjNi0C5vuqI7nPzirW-ERf6pcnVF5T57ooLRIauMWXZkkMnVNFTOa/exec";

document.addEventListener('DOMContentLoaded', () => {
    if (BACKEND_URL.includes("REPLACE")) {
        showError("Please configure the Backend URL in popup.js");
        return;
    }
    fetchData();
});

async function fetchData() {
    const status = document.getElementById('status');
    try {
        status.innerHTML = '<div class="spinner"></div>Fetching live analytics...';

        const response = await fetch(BACKEND_URL);
        if (!response.ok) throw new Error(`HTTP Error: ${response.status}`);

        const json = await response.json();

        if (json.status === "success" && json.data) {
            renderDashboard(json.data);
            updateTimestamp();
        } else {
            showError("Failed to load data: " + (json.message || "Unknown error"));
        }
    } catch (error) {
        console.error("Fetch Error:", error);
        showError("Connection Failed: Ensure you are logged into Google and the script is deployed as 'Anyone'.");
    }
}

function renderDashboard(data) {
    const status = document.getElementById('status');
    const dashboard = document.getElementById('dashboard');
    const moversBody = document.getElementById('moversBody');
    const mainBody = document.getElementById('mainBody');

    // Clear previous data
    moversBody.innerHTML = '';
    mainBody.innerHTML = '';

    // Hide loading
    status.style.display = 'none';
    dashboard.style.display = 'block';

    // Get latest data per security
    const latestData = getLatestPerSecurity(data);

    // Sort by DoD descending for "Top Movers" - this makes it feel more "live"
    const topMovers = [...latestData].sort((a, b) => {
        const valA = Math.abs(a.DoD || 0);
        const valB = Math.abs(b.DoD || 0);
        return valB - valA;
    });

    // Render Movers (Top 5)
    topMovers.slice(0, 5).forEach(item => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td><strong>${item.SECURITY}</strong></td>
            <td class="${getColorClass(item.DoD)}">${formatPercent(item.DoD)}</td>
            <td class="${getColorClass(item.MoM)}">${formatPercent(item.MoM)}</td>
            <td>${item.LAST || '-'}</td>
        `;
        moversBody.appendChild(row);
    });

    // Render Full List (Alphabetical)
    const sortedList = [...latestData].sort((a, b) => a.SECURITY.localeCompare(b.SECURITY));
    sortedList.forEach(item => {
        const row = document.createElement('tr');
        row.style.cursor = 'pointer'; // Make it look clickable
        row.onclick = () => showChart(item.SECURITY); // Add click handler
        row.innerHTML = `
            <td><strong>${item.SECURITY}</strong></td>
            <td class="${getColorClass(item.DoD)}">${formatPercent(item.DoD)}</td>
            <td class="${getColorClass(item.MoM)}">${formatPercent(item.MoM)}</td>
            <td class="${getColorClass(item.YoY)}">${formatPercent(item.YoY)}</td>
        `;
        mainBody.appendChild(row);
    });

    // Setup close button
    document.getElementById('closeChart').onclick = () => {
        document.getElementById('chartSection').style.display = 'none';
        window.scrollTo({ top: 0, behavior: 'smooth' });
    };
}

let performanceChart = null;

async function showChart(security) {
    const section = document.getElementById('chartSection');
    const title = document.getElementById('chartTitle');

    // Show section and scroll to it
    section.style.display = 'block';
    title.textContent = `${security} Performance (Today)`;
    section.scrollIntoView({ behavior: 'smooth' });

    console.log(`Fetching chart data for ${security}...`);

    try {
        const url = `${BACKEND_URL}${BACKEND_URL.includes('?') ? '&' : '?'}security=${encodeURIComponent(security)}`;
        console.log(`Request URL: ${url}`);

        const response = await fetch(url);
        if (!response.ok) throw new Error(`HTTP Error: ${response.status}`);

        const json = await response.json();
        console.log("Chart Response:", json);

        if (json.status === "success" && json.data) {
            if (json.data.length === 0) {
                console.warn(`No intraday data found for ${security}. Check if the security name matches exactly in the sheet.`);
            }
            renderChart(json.data);
        } else {
            console.error("Backend Error:", json.message);
        }
    } catch (e) {
        console.error("Chart fetch failed", e);
    }
}

function renderChart(history) {
    const ctx = document.getElementById('performanceChart').getContext('2d');

    if (performanceChart) performanceChart.destroy();

    const labels = history.map(h => {
        if (typeof h.time === 'string' && h.time.length > 5) return h.time.substring(0, 5);
        return h.time;
    });
    const prices = history.map(h => h.price);

    performanceChart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [{
                label: 'Price (TZS)',
                data: prices,
                borderColor: '#3498db',
                backgroundColor: 'rgba(52, 152, 219, 0.1)',
                fill: true,
                tension: 0.3,
                pointRadius: 2
            }]
        },
        options: {
            responsive: true,
            plugins: {
                legend: { display: false }
            },
            scales: {
                y: { grid: { display: false } },
                x: { grid: { display: false } }
            }
        }
    });
}

function getLatestPerSecurity(data) {
    const latest = {};
    data.forEach(row => {
        if (!row.SECURITY) return;

        // Robust date parsing
        let timestamp = 0;
        try {
            timestamp = new Date(row.TIMESTAMP).getTime();
            if (isNaN(timestamp)) timestamp = 0;
        } catch (e) {
            timestamp = 0;
        }

        if (!latest[row.SECURITY] || timestamp > latest[row.SECURITY].ts) {
            latest[row.SECURITY] = { ...row, ts: timestamp };
        }
    });
    return Object.values(latest);
}

function formatPercent(value) {
    if (value === null || value === undefined || isNaN(value)) return "-";
    const percent = (value * 100).toFixed(2);
    return (value > 0 ? "+" : "") + percent + "%";
}

function getColorClass(value) {
    if (value === null || value === undefined || isNaN(value)) return "neutral";
    if (Math.abs(value) < 0.0001) return "neutral";
    return value > 0 ? "positive" : "negative";
}

function updateTimestamp() {
    const header = document.querySelector('.header');
    let tsEl = document.getElementById('lastUpdated');
    if (!tsEl) {
        tsEl = document.createElement('div');
        tsEl.id = 'lastUpdated';
        tsEl.style.fontSize = '10px';
        tsEl.style.opacity = '0.7';
        tsEl.style.marginTop = '5px';
        header.appendChild(tsEl);
    }
    tsEl.textContent = "Last sync: " + new Date().toLocaleTimeString();
}

function showError(msg) {
    const status = document.getElementById('status');
    status.style.display = 'block';
    status.className = "error";
    status.innerHTML = `
        <div style="font-weight:bold; margin-bottom:5px;">⚠️ Error</div>
        <div style="font-size:11px;">${msg}</div>
        <button onclick="location.reload()" style="margin-top:10px; padding:4px 8px; cursor:pointer;">Retry</button>
    `;
}
