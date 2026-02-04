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

    const headers = values[0];
    const secIdx = headers.indexOf("SECURITY");
    const lastIdx = headers.indexOf("LAST");
    const timeIdx = headers.indexOf("TIME");
    const tsIdx = headers.indexOf("TIMESTAMP");

    // 1. Find the LATEST date that this specific security has data for
    let targetDateStr = null;
    const searchSymbol = securitySymbol.trim().toUpperCase();

    for (let i = values.length - 1; i >= 1; i--) {
        const rowSec = (values[i][secIdx] || "").toString().trim().toUpperCase();
        if (rowSec === searchSymbol) {
            const ts = values[i][tsIdx];
            targetDateStr = ts instanceof Date ?
                Utilities.formatDate(ts, Session.getScriptTimeZone(), "yyyy-MM-dd") :
                ts.toString().split(" ")[0];
            break;
        }
    }

    if (!targetDateStr) {
        console.log("No data found for: " + securitySymbol);
        return [];
    }

    console.log("Target Date found: " + targetDateStr);

    // 2. Filter data for that security on that specific date
    let rawTicks = [];
    const findHeader = (name) => {
        const lowerName = name.toLowerCase();
        return headers.findIndex(h => h.toString().toLowerCase().trim() === lowerName);
    };

    const bidIdx = findHeader("BID");
    const bidQtyIdx = findHeader("BID QTY");
    const askIdx = findHeader("OFFER") === -1 ? findHeader("ASK") : findHeader("OFFER");
    const askQtyIdx = findHeader("ASK QTY");

    values.slice(1).forEach(row => {
        const rowSec = (row[secIdx] || "").toString().trim().toUpperCase();
        const rowTS = row[tsIdx];
        const rowDateStr = rowTS instanceof Date ?
            Utilities.formatDate(rowTS, Session.getScriptTimeZone(), "yyyy-MM-dd") :
            rowTS.toString().split(" ")[0];

        if (rowSec === searchSymbol && rowDateStr === targetDateStr) {
            // Use TIMESTAMP (capture time) for the chart axis so we see the full continuum
            const rowTS = row[tsIdx];
            let label = "";
            if (rowTS instanceof Date) label = Utilities.formatDate(rowTS, Session.getScriptTimeZone(), "HH:mm");
            else {
                // Try to parse string timestamp
                const d = new Date(rowTS);
                label = !isNaN(d.getTime()) ? Utilities.formatDate(d, Session.getScriptTimeZone(), "HH:mm") : rowTS.toString().split(" ")[1] || "??";
            }

            // Sortable time value
            const sortTime = rowTS instanceof Date ? rowTS.getTime() : new Date(rowTS).getTime();

            // Volume tracking
            const volIdx = findHeader("VOL") === -1 ? (findHeader("VOLUME") === -1 ? findHeader("VOL.") : findHeader("VOLUME")) : findHeader("VOL");
            const rawVolVal = row[volIdx] || 0;
            const cleanVol = typeof rawVolVal === 'string' ? parseFloat(rawVolVal.replace(/,/g, '')) : parseFloat(rawVolVal);

            // Depth data (Handle commas)
            const cleanVal = (v) => {
                if (typeof v === 'string') return parseFloat(v.replace(/,/g, '')) || 0;
                return parseFloat(v) || 0;
            };

            rawTicks.push({
                time: label,
                price: parseFloat(row[lastIdx]),
                sortVal: sortTime,
                rawVol: isNaN(cleanVol) ? 0 : cleanVol,
                bid: cleanVal(row[bidIdx]),
                ask: cleanVal(row[askIdx]),
                bidQty: cleanVal(row[bidQtyIdx]),
                askQty: cleanVal(row[askQtyIdx])
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
    if (!sheet) throw new Error("No sheets found in the spreadsheet.");

    console.log("Reading data from sheet: " + sheet.getName());
    const values = sheet.getDataRange().getValues();
    console.log("Total rows in sheet: " + values.length);

    if (values.length < 2) return [];

    const headers = values[0].map(h => h.toString().trim().toUpperCase());
    const rawData = values.slice(1).map(row => {
        const obj = {};
        headers.forEach((h, i) => {
            obj[h] = row[i];
        });
        return obj;
    });

    console.log("Aggregating daily closes...");
    const dailyData = aggregateDailyCloses(rawData);
    console.log("Daily rows aggregated: " + dailyData.length);

    console.log("Computing returns...");
    return computeReturns(dailyData);
}

/**
 * Aggregates intraday tick data into daily closes per SECURITY.
 */
function aggregateDailyCloses(data) {
    const daily = {};

    data.forEach((row, idx) => {
        try {
            const security = row.SECURITY;
            if (!security || !row.TIMESTAMP) return;

            // Handle TIMESTAMP (could be Date object from Apps Script or string)
            let dateStr = "";
            if (row.TIMESTAMP instanceof Date) {
                dateStr = Utilities.formatDate(row.TIMESTAMP, Session.getScriptTimeZone(), "yyyy-MM-dd");
            } else {
                const d = new Date(row.TIMESTAMP);
                if (!isNaN(d.getTime())) {
                    dateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
                } else {
                    dateStr = row.TIMESTAMP.toString().split(" ")[0];
                }
            }

            const key = `${security}_${dateStr}`;

            const getTimeValue = (r) => {
                try {
                    if (r.TIME instanceof Date) return r.TIME.getTime();
                    if (!r.TIME) return (r.TIMESTAMP instanceof Date) ? r.TIMESTAMP.getTime() : new Date(r.TIMESTAMP).getTime();
                    const d = new Date(`${dateStr} ${r.TIME}`);
                    return !isNaN(d.getTime()) ? d.getTime() : ((r.TIMESTAMP instanceof Date) ? r.TIMESTAMP.getTime() : new Date(r.TIMESTAMP).getTime());
                } catch (e) { return 0; }
            };

            const currentTickTime = getTimeValue(row);

            // Volume normalization (handle headers like VOL, VOLUME, Vol.)
            const rawVol = row.VOL || row.VOLUME || row["Vol."] || 0;
            const volNum = typeof rawVol === 'string' ? parseFloat(rawVol.replace(/,/g, '')) : parseFloat(rawVol);
            const cleanVol = isNaN(volNum) ? 0 : volNum;

            if (!daily[key]) {
                daily[key] = { ...row, VOL: cleanVol }; // Start with this row's volume
            } else {
                const existingTickTime = getTimeValue(daily[key]);

                // If this is a LATER tick, update the price/status info
                if (currentTickTime > existingTickTime) {
                    const totalVolSoFar = daily[key].VOL;
                    daily[key] = { ...row };
                    // If DSE data is cumulative, we just take the new one. 
                    // If it's trade-by-trade, we sum it. 
                    // Most traders prefer "Last known Volume" if it grows, or Sum if it's tick.
                    // Let's assume Max (cumulative) to be safe for DSE.
                    daily[key].VOL = Math.max(totalVolSoFar, cleanVol);
                } else {
                    // Even if it's an older tick, it might have volume info we missed
                    daily[key].VOL = Math.max(daily[key].VOL, cleanVol);
                }
            }
        } catch (err) {
            console.error(`Error processing row ${idx}: ${err.message}`);
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

    const today = parseFloat(rows[index].LAST);
    const past = parseFloat(rows[index - lag].LAST);

    if (isNaN(today) || isNaN(past) || past === 0) return null;

    return (today - past) / past;
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

        // --- NEW: Test the daily chart ---
        console.log("Attempting to create intraday chart for:", sample.SECURITY);
        const chartStatus = createDailyPerformanceChart(sample.SECURITY);
        console.log("Chart outcome:", chartStatus);
    }
}
