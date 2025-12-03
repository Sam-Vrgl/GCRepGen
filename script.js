document.getElementById('fileInput').addEventListener('change', handleFileSelect, false);
document.getElementById('plateNumber').addEventListener('input', handlePlateNumberChange, false);
document.getElementById('exportBtn').addEventListener('click', () => window.print(), false);

let currentData = null;

function handleFileSelect(evt) {
    const file = evt.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        // Assume first sheet
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        currentData = XLSX.utils.sheet_to_json(worksheet);

        updateDisplay();
    };
    reader.readAsArrayBuffer(file);
}

function handlePlateNumberChange(evt) {
    const plateNum = evt.target.value;
    document.getElementById('mainTitle').textContent = `PLATE ${plateNum}`;
    const plateNumSpan = document.getElementById('plateNumSpan');
    if (plateNumSpan) {
        plateNumSpan.textContent = plateNum;
    }
}

function updateDisplay() {
    if (!currentData) return;
    const stats = calculateStats(currentData);
    displayResults(stats, currentData);
}

function calculateStats(data) {
    const wells = {};

    // Group data by Well
    data.forEach(row => {
        const well = row.Well;
        if (!well) return;
        if (!wells[well]) {
            wells[well] = [];
        }
        if (row.AreaShape_EquivalentDiameter) {
            wells[well].push(row.AreaShape_EquivalentDiameter);
        }
    });

    const stats = [];

    const averageRow = { Metric: "Average (µm)" };
    const medianRow = { Metric: "Median" };
    const stdDevRow = { Metric: "Standard Deviation" };
    const stdDevPercentRow = { Metric: "Standard Deviation %" };

    // Generate all wells A1-D6
    const allWells = [];
    const rows = ['A', 'B', 'C', 'D'];
    for (let r of rows) {
        for (let c = 1; c <= 6; c++) {
            allWells.push(r + c);
        }
    }

    allWells.forEach(well => {
        const values = wells[well];
        if (!values || values.length === 0) {
            averageRow[well] = null;
            medianRow[well] = null;
            stdDevRow[well] = null;
            stdDevPercentRow[well] = null;
            return;
        }

        // Average
        const sum = values.reduce((a, b) => a + b, 0);
        const avg = sum / values.length;
        averageRow[well] = avg;

        // Median
        values.sort((a, b) => a - b);
        const mid = Math.floor(values.length / 2);
        const median = values.length % 2 !== 0 ? values[mid] : (values[mid - 1] + values[mid]) / 2;
        medianRow[well] = median;

        // Standard Deviation (Population)
        const squareDiffs = values.map(value => {
            const diff = value - avg;
            return diff * diff;
        });
        const avgSquareDiff = squareDiffs.reduce((a, b) => a + b, 0) / values.length;
        const stdDev = Math.sqrt(avgSquareDiff);
        stdDevRow[well] = stdDev;

        // Standard Deviation Percentage
        const stdDevPercent = (stdDev / avg) * 100;
        stdDevPercentRow[well] = stdDevPercent;
    });

    stats.push(averageRow);
    stats.push(medianRow);
    stats.push(stdDevRow);
    stats.push(stdDevPercentRow);

    return { stats, allWells, wellsData: wells };
}

function displayResults(resultData, rawData) {
    const stats = resultData.stats;
    const allWells = resultData.allWells;
    const wellsData = resultData.wellsData;

    document.getElementById('results').classList.remove('hidden');

    // 1. Generate Violin Plot
    setTimeout(() => {
        generateViolinPlot(wellsData, allWells);
    }, 0);

    // 2. Generate Main Plate Diagram
    const plateContainer = document.getElementById('plateDiagram');
    plateContainer.innerHTML = '';

    const avgRow = stats.find(r => r.Metric === "Average (µm)");
    const stdRow = stats.find(r => r.Metric === "Standard Deviation");
    const pctRow = stats.find(r => r.Metric === "Standard Deviation %");

    allWells.forEach(well => {
        const wellDiv = document.createElement('div');
        wellDiv.className = 'well-container';

        const avg = avgRow[well];
        const std = stdRow[well];
        const pct = pctRow[well];

        let circleContent = '';
        if (avg !== null && avg !== undefined) {
            circleContent = `
                <div class="percentage">${Math.round(pct)}%</div>
                <div class="stats-text">${Math.round(avg)} ± ${Math.round(std)} µm</div>
            `;
        } else {
            circleContent = `<div class="stats-text">0%</div>`;
        }

        wellDiv.innerHTML = `
            <div class="well-label">${well}</div>
            <div class="well-circle">
                ${circleContent}
            </div>
        `;
        plateContainer.appendChild(wellDiv);
    });

    // 3. Generate Placeholder Plate Diagram
    const plateContainer2 = document.getElementById('plateDiagram2');
    plateContainer2.innerHTML = '';
    allWells.forEach(well => {
        const wellDiv = document.createElement('div');
        wellDiv.className = 'well-container';
        wellDiv.innerHTML = `
            <div class="well-label">${well}</div>
            <div class="well-circle">
                <div class="percentage" style="font-size: 10px; color: #ccc;">0%</div>
            </div>
        `;
        plateContainer2.appendChild(wellDiv);
    });

    // 4. Generate Table
    const tableContainer = document.getElementById('statsTableContainer');
    tableContainer.innerHTML = '';

    const table = document.createElement('table');
    const thead = document.createElement('thead');
    const tbody = document.createElement('tbody');

    // Header Row
    const headerRow = document.createElement('tr');
    const thMetric = document.createElement('th');
    thMetric.textContent = '';
    headerRow.appendChild(thMetric);

    allWells.forEach(well => {
        const th = document.createElement('th');
        th.textContent = well;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);

    // Data Rows
    const rowsToShow = ["Average (µm)", "Median"];

    rowsToShow.forEach(metricName => {
        const rowData = stats.find(r => r.Metric === metricName);
        if (!rowData) return;

        const tr = document.createElement('tr');
        const tdMetric = document.createElement('td');
        tdMetric.textContent = metricName;
        tdMetric.style.fontWeight = 'bold';
        tdMetric.style.textAlign = 'left';
        tr.appendChild(tdMetric);

        allWells.forEach(well => {
            const td = document.createElement('td');
            const val = rowData[well];
            td.textContent = val !== null && val !== undefined ? Math.round(val) : '';
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });

    table.appendChild(thead);
    table.appendChild(tbody);
    tableContainer.appendChild(table);
}

function generateViolinPlot(wellsData, allWells) {
    const plotData = [];

    // Create a trace for each well
    // To make it look like the image (single color, side by side), 
    // we can use a single trace with 'x' as well names and 'y' as values, 
    // or multiple traces. Plotly violin handles 'x' categories well.

    const xValues = [];
    const yValues = [];

    allWells.forEach(well => {
        const values = wellsData[well];
        if (values && values.length > 0) {
            values.forEach(v => {
                xValues.push(well);
                yValues.push(v);
            });
        }
    });

    const trace = {
        type: 'violin',
        x: xValues,
        y: yValues,
        points: false,
        box: {
            visible: true
        },
        line: {
            color: 'black'
        },
        fillcolor: '#8dd3c7',
        opacity: 0.6,
        meanline: {
            visible: true
        },
        x0: "A1"
    };

    const layout = {
        title: "",
        yaxis: {
            zeroline: false,
            title: "Spheroid Diameter (µm)"
        },
        xaxis: {
            title: "Wells ID"
        },
        margin: {
            l: 50,
            r: 10,
            b: 50,
            t: 10
        },
        showlegend: false
    };

    Plotly.newPlot('violinPlot', [trace], layout, { 
    displayModeBar: false, 
    responsive: true 
});
}

window.onbeforeprint = function() {
    const plotDiv = document.getElementById('violinPlot');
    Plotly.relayout(plotDiv, {
        width: 600, 
        height: 400
    });
};

window.onafterprint = function() {
    const plotDiv = document.getElementById('violinPlot');
    Plotly.relayout(plotDiv, {
        width: null,
        height: 500
    });
};