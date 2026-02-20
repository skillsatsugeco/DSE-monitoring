/**
 * DSE Analytics Backend
 * 
 * 1. Reads raw market data from the existing spreadsheet.
 * 2. Calculates DoD (1 day), MoM (21 days), and YoY (252 days) returns.
 * 3. Serves the processed data as JSON for the Chrome Extension.
 * 4. Can optionally write these analytics back to the sheet or generate charts.
 */

const SPREADSHEET_ID = "1kFx5T2t07OlAdFhp9ic906N5Ie3p18kiS2RgZG4pCDs"; // Provided by user

// Configuration for trading days
const TRADING_DAYS = {
    DOD: 1,
    MOM: 21,
    YOY: 252
};

/**
 * Handle POST requests from the Chrome Extension Scraper
 */
function doPost(e) {
    try {
        const rawData = JSON.parse(e.postData.contents);
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheets()[0];

        // Ensure headers exist if sheet is empty
        if (sheet.getLastRow() === 0) {
            sheet.appendRow(["SECURITY", "LAST", "TIME", "TIMESTAMP", "VOL", "BID", "ASK", "BID QTY", "ASK QTY"]);
        }

        // Handle single object or array of objects
        const dataArray = Array.isArray(rawData) ? rawData : [rawData];
        const now = new Date();

        dataArray.forEach(item => {
            // Mapping incoming data to sheet columns
            const clean = (v) => v ? (v.toString().replace(/,/g, '').trim()) : 0;

            sheet.appendRow([
                (item.security || item.ticker || item.symbol || "").toUpperCase(),
                clean(item.last || item.price || 0),
                item.time || Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm:ss"),
                now,
                clean(item.vol || item.volume || 0),
                clean(item.bid || 0),
                clean(item.ask || item.offer || 0),
                clean(item.bidQty || 0),
                clean(item.askQty || 0)
            ]);
        });

        return ContentService.createTextOutput(JSON.stringify({
            status: "success",
            rows_added: dataArray.length
        })).setMimeType(ContentService.MimeType.JSON);

    } catch (err) {
        return ContentService.createTextOutput(JSON.stringify({
            status: "error",
            message: err.toString()
        })).setMimeType(ContentService.MimeType.JSON);
    }
}

function doGet(e) {
    try {
        const security = e.parameter.security;
        let data;

        if (security) {
            // New: Fetch full intraday history for a specific stock
            data = getIntradayHistory(security);
        } else {
            // Default: Fetch aggregated analytics
            data = getComputedAnalytics();
        }

        return ContentService.createTextOutput(JSON.stringify({
            status: "success",
            data: data
        })).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
        return ContentService.createTextOutput(JSON.stringify({
            status: "error",
            message: err.toString()
        })).setMimeType(ContentService.MimeType.JSON);
    }
}

/**
 * Helper to fetch raw intraday data for the latest day for a specific stock.
 */
function getIntradayHistory(securitySymbol) {
    console.log("Fetching intraday for: " + securitySymbol);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheets()[0];
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return [];

    const headers = values[0].map(h => h.toString().trim().toUpperCase());

    // Flexible header detection
    const findIdx = (aliases) => {
        return headers.findIndex(h => aliases.includes(h));
    };

    const secIdx = findIdx(["SECURITY", "SYMBOL", "TICKER", "SEC"]);
    const lastIdx = findIdx(["LAST", "PRICE", "CLOSE", "LTP"]);
    const tsIdx = findIdx(["TIMESTAMP", "DATE", "TS"]);
    const timeIdx = findIdx(["TIME", "CLOCK"]);
    const volIdx = findIdx(["VOL", "VOLUME", "VOL.", "QTY"]);
    const bidIdx = findIdx(["BID", "BUY"]);
    const askIdx = findIdx(["ASK", "OFFER", "SELL"]);
    const bidQtyIdx = findIdx(["BID QTY", "BUY QTY", "BID_QTY"]);
    const askQtyIdx = findIdx(["ASK QTY", "SELL QTY", "ASK_QTY"]);

    if (secIdx === -1 || lastIdx === -1) {
        console.error("Critical columns missing. Found headers: " + headers.join(", "));
        return [];
    }

    // 1. Find the LATEST date that this specific security has data for
    let targetDateStr = null;
    const searchSymbol = securitySymbol.trim().toUpperCase();

    for (let i = values.length - 1; i >= 1; i--) {
        const rowSec = (values[i][secIdx] || "").toString().trim().toUpperCase();
        if (rowSec === searchSymbol) {
            const ts = values[i][tsIdx];
            if (!ts) continue;
            targetDateStr = ts instanceof Date ?
                Utilities.formatDate(ts, Session.getScriptTimeZone(), "yyyy-MM-dd") :
                ts.toString().split(" ")[0];
            if (targetDateStr) break;
        }
    }

    if (!targetDateStr) {
        console.log("No data found for: " + securitySymbol);
        return [];
    }

    // 2. Filter data for that security on that specific date
    let rawTicks = [];
    const cleanVal = (v) => {
        if (v === null || v === undefined || v === "") return 0;
        if (typeof v === 'number') return v;
        const str = v.toString().replace(/,/g, '').trim();
        return parseFloat(str) || 0;
    };

    values.slice(1).forEach(row => {
        const rowSec = (row[secIdx] || "").toString().trim().toUpperCase();
        const rowTS = row[tsIdx];
        if (!rowTS) return;

        const rowDateStr = rowTS instanceof Date ?
            Utilities.formatDate(rowTS, Session.getScriptTimeZone(), "yyyy-MM-dd") :
            rowTS.toString().split(" ")[0];

        if (rowSec === searchSymbol && rowDateStr === targetDateStr) {
            let label = "";
            if (rowTS instanceof Date) label = Utilities.formatDate(rowTS, Session.getScriptTimeZone(), "HH:mm");
            else {
                const d = new Date(rowTS);
                label = !isNaN(d.getTime()) ? Utilities.formatDate(d, Session.getScriptTimeZone(), "HH:mm") : rowTS.toString().split(" ")[1] || "??";
            }

            const sortTime = rowTS instanceof Date ? rowTS.getTime() : new Date(rowTS).getTime();
            const cleanVol = cleanVal(row[volIdx]);

            rawTicks.push({
                time: label,
                price: cleanVal(row[lastIdx]),
                sortVal: sortTime,
                rawVol: isNaN(cleanVol) ? 0 : cleanVol,
                bid: bidIdx !== -1 ? cleanVal(row[bidIdx]) : 0,
                ask: askIdx !== -1 ? cleanVal(row[askIdx]) : 0,
                bidQty: bidQtyIdx !== -1 ? cleanVal(row[bidQtyIdx]) : 0,
                askQty: askQtyIdx !== -1 ? cleanVal(row[askQtyIdx]) : 0
            });
        }
    });

    if (rawTicks.length === 0) return [];

    // 3. Sort ticks by capture time
    rawTicks.sort((a, b) => a.sortVal - b.sortVal);

    // 4. Calculate Spread/Imbalance & Deduplicate
    const results = [];
    let cumulativeVol = 0;

    rawTicks.forEach(tick => {
        cumulativeVol = Math.max(cumulativeVol, tick.rawVol);

        const spread = (tick.ask > 0 && tick.bid > 0) ? (tick.ask - tick.bid) : 0;
        const totalDepth = tick.bidQty + tick.askQty;
        const imbalance = totalDepth > 0 ? ((tick.bidQty - tick.askQty) / totalDepth).toFixed(2) : 0;

        const processedTick = {
            ...tick,
            volume: cumulativeVol,
            spread: spread,
            imbalance: parseFloat(imbalance)
        };

        if (results.length === 0) {
            results.push(processedTick);
        } else {
            const last = results[results.length - 1];
            // Only add if minute changed OR price changed
            // This prevents the chart from being too "noisy" with 10-second updates
            if (tick.time !== last.time || tick.price !== last.price) {
                results.push(processedTick);
            } else {
                // Update volume/imbalance for the latest state of this minute
                results[results.length - 1] = processedTick;
            }
        }
    });

    // 5. EXTEND TO PRESENT: If the last trade was 1 hour ago, add a point for "Now"
    const lastTick = results[results.length - 1];
    const now = new Date();
    const nowTimeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm");

    if (lastTick.time !== nowTimeStr && targetDateStr === Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd")) {
        results.push({
            ...lastTick,
            time: nowTimeStr,
            sortVal: now.getTime()
        });
    }

    console.log(`Returning ${results.length} unique points for timeline.`);
    return results;
}

/**
 * Main function to read raw data, compute returns, and return the structured array.
 */
function getComputedAnalytics() {
    console.log("Opening spreadsheet: " + SPREADSHEET_ID);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheets()[0];
    if (!sheet) throw new Error("No sheets found. Check Spreadsheet ID and permissions.");

    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return [];

    const headers = values[0].map(h => h.toString().trim().toUpperCase());

    // Mapping keys to indices
    const findIdx = (aliases) => headers.findIndex(h => aliases.includes(h));
    const map = {
        security: findIdx(["SECURITY", "SYMBOL", "TICKER", "SEC"]),
        last: findIdx(["LAST", "PRICE", "CLOSE", "LTP"]),
        vol: findIdx(["VOL", "VOLUME", "VOL.", "QTY"]),
        ts: findIdx(["TIMESTAMP", "DATE", "TS"]),
        time: findIdx(["TIME", "CLOCK"])
    };

    if (map.security === -1 || map.last === -1 || map.ts === -1) {
        throw new Error("Missing required columns: SECURITY, LAST, and TIMESTAMP/DATE are required.");
    }

    const rawData = values.slice(1).map(row => {
        const obj = {};
        obj.SECURITY = row[map.security];
        obj.LAST = row[map.last];
        obj.VOL = row[map.vol] !== -1 ? row[map.vol] : 0;
        obj.TIMESTAMP = row[map.ts];
        obj.TIME = map.time !== -1 ? row[map.time] : null;
        return obj;
    });

    console.log("Daily rows aggregated: " + rawData.length);
    const dailyData = aggregateDailyCloses(rawData);
    const withReturns = computeReturns(dailyData);
    return computeMarketSignals(withReturns);
}

/**
 * Aggregates intraday tick data into daily closes per SECURITY.
 */
function aggregateDailyCloses(data) {
    const daily = {};

    data.forEach((row, idx) => {
        try {
            const security = row.SECURITY ? row.SECURITY.toString().trim().toUpperCase() : null;
            if (!security || !row.TIMESTAMP) return;

            // Handle TIMESTAMP
            let dateStr = "";
            const ts = row.TIMESTAMP;
            if (ts instanceof Date) {
                dateStr = Utilities.formatDate(ts, Session.getScriptTimeZone(), "yyyy-MM-dd");
            } else {
                const d = new Date(ts);
                dateStr = !isNaN(d.getTime()) ? Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd") : ts.toString().split(" ")[0];
            }

            const key = `${security}_${dateStr}`;

            const getTimeValue = (r) => {
                try {
                    if (r.TIME instanceof Date) return r.TIME.getTime();
                    if (!r.TIME) return (r.TIMESTAMP instanceof Date) ? r.TIMESTAMP.getTime() : new Date(r.TIMESTAMP).getTime();
                    return new Date(`${dateStr} ${r.TIME}`).getTime() || ((r.TIMESTAMP instanceof Date) ? r.TIMESTAMP.getTime() : new Date(r.TIMESTAMP).getTime());
                } catch (e) { return 0; }
            };

            const currentTickTime = getTimeValue(row);
            const clean = (v) => {
                if (v === null || v === undefined) return 0;
                if (typeof v === 'number') return v;
                return parseFloat(v.toString().replace(/,/g, '')) || 0;
            };

            const cleanVol = clean(row.VOL);
            const cleanPrice = clean(row.LAST);

            if (!daily[key]) {
                daily[key] = { ...row, SECURITY: security, LAST: cleanPrice, VOL: cleanVol, _time: currentTickTime };
            } else if (currentTickTime >= daily[key]._time) {
                const totalVol = Math.max(daily[key].VOL, cleanVol);
                daily[key] = { ...row, SECURITY: security, LAST: cleanPrice, VOL: totalVol, _time: currentTickTime };
            } else {
                daily[key].VOL = Math.max(daily[key].VOL, cleanVol);
            }
        } catch (err) {
            console.error(`Row ${idx} error: ${err.message}`);
        }
    });

    return Object.values(daily);
}

/**
 * Core Logic: Groups by stock, sorts by time, computes returns.
 */
function computeReturns(data) {
    // Group by SECURITY
    const grouped = {};
    data.forEach(row => {
        const security = row.SECURITY;
        if (!security) return; // Skip empty rows
        if (!grouped[security]) grouped[security] = [];
        grouped[security].push(row);
    });

    // Sort each stock by time
    Object.values(grouped).forEach(rows => {
        rows.sort((a, b) => {
            const getTime = (r) => {
                if (r.TIMESTAMP instanceof Date) return r.TIMESTAMP.getTime();
                const d = new Date(r.TIMESTAMP);
                return isNaN(d.getTime()) ? 0 : d.getTime();
            };
            return getTime(a) - getTime(b);
        });
    });

    // Compute returns
    Object.values(grouped).forEach(rows => {
        rows.forEach((row, i) => {
            row.DoD = calcReturn(rows, i, TRADING_DAYS.DOD);
            row.MoM = calcReturn(rows, i, TRADING_DAYS.MOM);
            row.YoY = calcReturn(rows, i, TRADING_DAYS.YOY);
        });
    });

    // Flatten back to array
    const result = [];
    Object.values(grouped).forEach(rows => {
        result.push(...rows);
    });

    return result;
}

function calcReturn(rows, index, lag) {
    if (index < lag) return null;

    const clean = (v) => {
        if (v === null || v === undefined) return 0;
        if (typeof v === 'number') return v;
        return parseFloat(v.toString().replace(/,/g, '')) || 0;
    };

    const today = clean(rows[index].LAST);
    const past = clean(rows[index - lag].LAST);

    if (today === 0 || past === 0) return null;

    return (today - past) / past;
}

/**
 * NEW: Adaptive trading signals based on Relative Volume and Momentum.
 */
function computeMarketSignals(data) {
    // Group by SECURITY
    const grouped = {};
    data.forEach(row => {
        const security = row.SECURITY;
        if (!security) return;
        if (!grouped[security]) grouped[security] = [];
        grouped[security].push(row);
    });

    // Process each security
    Object.keys(grouped).forEach(security => {
        const rows = grouped[security];

        // Sort by timestamp
        rows.sort((a, b) => {
            const getT = (r) => (r.TIMESTAMP instanceof Date) ? r.TIMESTAMP.getTime() : new Date(r.TIMESTAMP).getTime();
            return getT(a) - getT(b);
        });

        rows.forEach((row, i) => {
            // 1. Average Volume (avgVol30)
            // Window: last 30 rows including current
            const windowSize = 30;
            const startIdx = Math.max(0, i - windowSize + 1);
            const windowRows = rows.slice(startIdx, i + 1);

            const sumVol = windowRows.reduce((sum, r) => {
                const vol = parseFloat(r.VOL) || 0;
                return sum + vol;
            }, 0);

            row.avgVol30 = windowRows.length > 0 ? sumVol / windowRows.length : 0;

            // 2. Relative Volume (RVOL)
            row.rvol = (row.avgVol30 && row.avgVol30 > 0) ? (row.VOL / row.avgVol30) : null;

            // 3. Liquidity Score
            if (row.avgVol30 >= 100000) row.liquidityScore = "HIGH";
            else if (row.avgVol30 >= 20000) row.liquidityScore = "MEDIUM";
            else row.liquidityScore = "LOW";

            // 4. Momentum Signal
            const rvol = row.rvol || 0;
            const dod = row.DoD || 0;
            if (dod > 0 && rvol > 1.5) row.momentumSignal = "CONFIRMED_UP";
            else if (dod > 0 && rvol <= 1) row.momentumSignal = "WEAK_UP";
            else if (dod < 0 && rvol > 1.5) row.momentumSignal = "STRONG_SELL";
            else row.momentumSignal = "NEUTRAL";

            // 5. Hype Risk
            if (dod > 0.08 && rvol < 1) row.hypeRisk = "HYPE_RISK";
            else if (dod > 0.08 && rvol > 2) row.hypeRisk = "BREAKOUT";
            else row.hypeRisk = "NORMAL";

            // 6. Stable Accumulation
            row.stableTrend = (dod >= 0 && dod <= 0.03 && rvol > 1.2);

            // 7. Trade Score
            let score = 0;
            if (row.liquidityScore === "HIGH") score += 2;
            if (row.momentumSignal === "CONFIRMED_UP") score += 2;
            if (row.stableTrend === true) score += 1;
            if (row.hypeRisk === "HYPE_RISK") score -= 2;
            if (row.momentumSignal === "STRONG_SELL") score -= 1;
            row.tradeScore = score;
        });
    });

    // Flatten back to array
    const result = [];
    Object.values(grouped).forEach(rows => {
        result.push(...rows);
    });

    return result;
}

/**
 * FEATURE 4: Daily Intraday Chart
 * Creates a performance chart for a specific security showing price movement for the current day.
 */
function createDailyPerformanceChart(securitySymbol) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SPREADSHEET_ID);
        const sourceSheet = ss.getSheets()[0];
        const data = sourceSheet.getDataRange().getValues();
        const headers = data[0];

        const secIdx = headers.indexOf("SECURITY");
        const lastIdx = headers.indexOf("LAST");
        const timeIdx = headers.indexOf("TIME");
        const tsIdx = headers.indexOf("TIMESTAMP");

        if (secIdx === -1 || lastIdx === -1) {
            return "Required columns (SECURITY, LAST) missing.";
        }

        // 1. Get today's date (latest date in sheet)
        const latestRow = data[data.length - 1];
        const latestTS = latestRow[tsIdx];
        const targetDateStr = latestTS instanceof Date ?
            Utilities.formatDate(latestTS, Session.getScriptTimeZone(), "yyyy-MM-dd") :
            latestTS.toString().split(" ")[0];

        // 2. Filter for the specific security and target date
        const chartData = [["Time", "Price"]];
        data.slice(1).forEach(row => {
            const rowSec = row[secIdx];
            const rowTS = row[tsIdx];
            const rowDateStr = rowTS instanceof Date ?
                Utilities.formatDate(rowTS, Session.getScriptTimeZone(), "yyyy-MM-dd") :
                rowTS.toString().split(" ")[0];

            if (rowSec === securitySymbol && rowDateStr === targetDateStr) {
                let label = row[timeIdx] || rowDateStr;
                // If label is a Date object (HH:MM:SS), format it
                if (label instanceof Date) label = Utilities.formatDate(label, Session.getScriptTimeZone(), "HH:mm");
                chartData.push([label, row[lastIdx]]);
            }
        });

        if (chartData.length < 3) return "Not enough intraday data for " + securitySymbol;

        // 3. Create a temporary sheet to hold chart data
        let chartSheet = ss.getSheetByName("ChartTemp");
        if (chartSheet) ss.deleteSheet(chartSheet);
        chartSheet = ss.insertSheet("ChartTemp");
        chartSheet.getRange(1, 1, chartData.length, 2).setValues(chartData);

        // 4. Build the chart
        const chart = chartSheet.newChart()
            .setChartType(Charts.ChartType.AREA)
            .addRange(chartSheet.getRange(1, 1, chartData.length, 2))
            .setPosition(2, 4, 0, 0)
            .setOption('title', `Intraday Performance: ${securitySymbol} (${targetDateStr})`)
            .setOption('hAxis', { title: 'Time of Day' })
            .setOption('vAxis', { title: 'Price (TZS)' })
            .setOption('legend', { position: 'none' })
            .build();

        chartSheet.insertChart(chart);
        return "Chart created successfully in 'ChartTemp' sheet.";

    } catch (err) {
        console.error("Chart Error: " + err.message);
        return "Error creating chart: " + err.message;
    }
}

/**
 * TEST FUNCTION
 * Run this in the editor to view the output in the logs.
 */
function testAnalytics() {
    const results = getComputedAnalytics();
    console.log("Total rows processed:", results.length);

    if (results.length > 0) {
        // Look for a row where DoD is NOT null (meaning we have at least 2 days of data for that stock)
        const sample = results.find(r => r.DoD !== null) || results[0];

        console.log("Sample Stock:", sample.SECURITY);
        console.log("Date:", sample.TIMESTAMP);
        console.log("Last Price:", sample.LAST);
        if (sample.DoD !== null) console.log("DoD Change:", (sample.DoD * 100).toFixed(2) + "%");

        console.log("--- Market Signals ---");
        console.log("Avg Vol 30:", sample.avgVol30);
        console.log("Relative Vol:", sample.rvol);
        console.log("Liquidity:", sample.liquidityScore);
        console.log("Momentum:", sample.momentumSignal);
        console.log("Hype Risk:", sample.hypeRisk);
        console.log("Stable Trend:", sample.stableTrend);
        console.log("Trade Score:", sample.tradeScore);

        // --- NEW: Test the daily chart ---
        console.log("Attempting to create intraday chart for:", sample.SECURITY);
        const chartStatus = createDailyPerformanceChart(sample.SECURITY);
        console.log("Chart outcome:", chartStatus);
    }
}
