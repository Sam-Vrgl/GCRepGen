document
  .getElementById("fileInput")
  .addEventListener("change", handleFileSelect, false);
document
  .getElementById("plateNumber")
  .addEventListener("input", handlePlateNumberChange, false);
document
  .getElementById("exportBtn")
  .addEventListener("click", () => window.print(), false);

let allReportsData = [];
const DETECTION_THRESHOLD = { min: 50, max: 300 }; //µm
function handleFileSelect(evt) {
  const files = evt.target.files;
  if (!files || files.length === 0) return;

  allReportsData = [];
  document.getElementById("results").innerHTML = "";
  document.getElementById("results").classList.add("hidden");

  const fileReaders = [];

  Array.from(files).forEach((file) => {
    fileReaders.push(
      new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function (e) {
          try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array" });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet);
            resolve({ fileName: file.name, data: jsonData });
          } catch (err) {
            console.error("Error parsing file", file.name, err);
            resolve(null);
          }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
      }),
    );
  });

  Promise.all(fileReaders)
    .then((results) => {
      allReportsData = results.filter((r) => r !== null);
      renderAllReports();
    })
    .catch((err) => {
      console.error("Error reading files:", err);
      alert("Error reading files. See console for details.");
    });
}
function handlePlateNumberChange(evt) {
  updatePlateNumbers();
}

function updatePlateNumbers() {
  const startPlateNum =
    parseInt(document.getElementById("plateNumber").value, 10) || 1;

  allReportsData.forEach((_, index) => {
    const plateNum = startPlateNum + index;
    const mainTitle = document.getElementById(`mainTitle-${index}`);
    const plateNumSpan = document.getElementById(`plateNumSpan-${index}`);

    if (mainTitle) mainTitle.textContent = `PLATE ${plateNum}`;
    if (plateNumSpan) plateNumSpan.textContent = plateNum;
  });
}

function renderAllReports() {
  if (allReportsData.length === 0) return;

  const resultsContainer = document.getElementById("results");
  resultsContainer.innerHTML = "";
  resultsContainer.classList.remove("hidden");

  allReportsData.forEach((reportData, index) => {
    createReportDOM(reportData.data, index);
  });

  updatePlateNumbers();
}

function createReportDOM(data, index) {
  const stats = calculateStats(data);
  const resultsContainer = document.getElementById("results");

  let day = "--";
  if (data.length > 0) {
    const firstRowDay = data[0]["Day"];
    if (firstRowDay !== undefined) {
      day = firstRowDay.toString().replace(/D/i, "");
    }
  }

  const reportPage = document.createElement("div");
  reportPage.className = "report-page";
  reportPage.id = `report-${index}`;

  reportPage.innerHTML = `
        <div class="dashboard-grid">
            <div class="left-panel">
                <h3 class="chart-title">Analysis of the distribution of spheroid diameters per well on
                    plate <span id="plateNumSpan-${index}">1</span> at Day <span id="dayNumSpan-${index}">${day}</span> (J-2 before shipment)</h3>
                <div id="violinPlot-${index}" class="violin-plot"></div>
            </div>

            <div class="right-panel">
                <h1 id="mainTitle-${index}" class="plate-title">PLATE 1</h1>

                <div class="plates-row">
                    <div class="plate-item">
                        <div id="plateDiagram-${index}" class="plate-grid"></div>
                        <p class="caption">Average organoids diameter and standard deviation values and
                            standard deviation percentage</p>
                    </div>

                    <div class="plate-item">
                        <div id="plateDiagram2-${index}" class="plate-grid placeholder-plate"></div>
                        <p class="caption">Number of conform organoids per well (well formed organoids, not passed in 2D under the
                            pattern)</p>
                    </div>
                </div>
            </div>
        </div>

        <div id="statsTableContainer-${index}"></div>
    `;

  resultsContainer.appendChild(reportPage);

  setTimeout(() => {
    generateViolinPlot(`violinPlot-${index}`, stats.wellsData, stats.allWells);
  }, 0);

  renderPlateDiagram(stats, `plateDiagram-${index}`);

  renderManualPlate(`plateDiagram2-${index}`, stats.allWells, stats.stats);

  renderStatsTable(stats, `statsTableContainer-${index}`);
}

function calculateStats(data) {
  const wells = {};

  data.forEach((row) => {
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
  const numSpheroids = { Metric: "Number of Spheroids" };

  const allWells = [];
  const rows = ["A", "B", "C", "D"];
  for (let r of rows) {
    for (let c = 1; c <= 6; c++) {
      allWells.push(r + c);
    }
  }

  allWells.forEach((well) => {
    const values = wells[well];
    if (!values || values.length === 0) {
      averageRow[well] = null;
      medianRow[well] = null;
      stdDevRow[well] = null;
      stdDevPercentRow[well] = null;
      return;
    }
    let count = 0;
    values.forEach((value) => {
      if (
        value >= DETECTION_THRESHOLD.min &&
        value <= DETECTION_THRESHOLD.max
      ) {
        count += 1;
      }
    });
    numSpheroids[well] = count;

    const sum = values.reduce((a, b) => a + b, 0);
    const avg = sum / values.length;
    averageRow[well] = avg;

    values.sort((a, b) => a - b);
    const mid = Math.floor(values.length / 2);
    const median =
      values.length % 2 !== 0
        ? values[mid]
        : (values[mid - 1] + values[mid]) / 2;
    medianRow[well] = median;

    const squareDiffs = values.map((value) => {
      const diff = value - avg;
      return diff * diff;
    });
    const avgSquareDiff =
      squareDiffs.reduce((a, b) => a + b, 0) / values.length;
    const stdDev = Math.sqrt(avgSquareDiff);
    stdDevRow[well] = stdDev;

    const stdDevPercent = (stdDev / avg) * 100;
    stdDevPercentRow[well] = stdDevPercent;
  });

  stats.push(averageRow);
  stats.push(medianRow);
  stats.push(stdDevRow);
  stats.push(stdDevPercentRow);
  stats.push(numSpheroids);

  return { stats, allWells, wellsData: wells };
}

function renderPlateDiagram(statsData, containerId) {
  const container = document.getElementById(containerId);
  const { stats, allWells } = statsData;

  const avgRow = stats.find((r) => r.Metric === "Average (µm)");
  const stdRow = stats.find((r) => r.Metric === "Standard Deviation");
  const pctRow = stats.find((r) => r.Metric === "Standard Deviation %");
  //   const spheroidRow = stats.find((r) => r.Metric === "Number of spheroids%");

  allWells.forEach((well) => {
    const wellDiv = document.createElement("div");
    wellDiv.className = "well-container";

    const avg = avgRow[well];
    const std = stdRow[well];
    const pct = pctRow[well];
    // const sphero = spheroidRow[well];

    let circleContent = "";
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
    container.appendChild(wellDiv);
  });
}

function renderManualPlate(containerId, allWells, stats) {
  const container = document.getElementById(containerId);
  const sphRow = stats.find((r) => r.Metric === "Number of Spheroids");
  allWells.forEach((well) => {
    const wellDiv = document.createElement("div");
    const spheroidNumber = sphRow[well];
    wellDiv.className = "well-container";
    wellDiv.innerHTML = `
            <div class="well-label">${well}</div>
            <div class="well-circle" style="flex-direction: row; gap: 1px;">
                <input type="number"
                       class="manual-input" 
                       min="0" 
                       max="100" 
                       oninput="handleManualInput(this)"
                       value="${spheroidNumber}"
                       id="ValidSpheroidCount-${well}">
            </div>
        `;
    //<span class="percent-symbol">%</span>
    container.appendChild(wellDiv);

    const manualPlateInput = document.getElementById(
      `ValidSpheroidCount-${well}`,
    );
    handleManualInput(manualPlateInput);
  });
}

function renderStatsTable(statsData, containerId) {
  const container = document.getElementById(containerId);
  const { stats, allWells } = statsData;

  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const tbody = document.createElement("tbody");

  const headerRow = document.createElement("tr");
  const thMetric = document.createElement("th");
  thMetric.textContent = "";
  headerRow.appendChild(thMetric);

  allWells.forEach((well) => {
    const th = document.createElement("th");
    th.textContent = well;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  const rowsToShow = ["Average (µm)", "Median"];

  rowsToShow.forEach((metricName) => {
    const rowData = stats.find((r) => r.Metric === metricName);
    if (!rowData) return;

    const tr = document.createElement("tr");
    const tdMetric = document.createElement("td");
    tdMetric.textContent = metricName;
    tdMetric.style.fontWeight = "bold";
    tdMetric.style.textAlign = "left";
    tr.appendChild(tdMetric);

    allWells.forEach((well) => {
      const td = document.createElement("td");
      const val = rowData[well];
      td.textContent = val !== null && val !== undefined ? Math.round(val) : "";
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
  container.appendChild(table);
}

function generateViolinPlot(elementId, wellsData, allWells) {
  const plotData = [];

  const xValues = [];
  const yValues = [];

  allWells.forEach((well) => {
    const values = wellsData[well];
    if (values && values.length > 0) {
      values.forEach((v) => {
        xValues.push(well);
        yValues.push(v);
      });
    }
  });

  const trace = {
    type: "violin",
    x: xValues,
    y: yValues,
    points: false,
    width: 0.85,

    box: {
      visible: true,
    },
    line: {
      color: "black",
      width: 1,
    },
    fillcolor: "#fbcfda",
    opacity: 0.6,
    meanline: {
      visible: true,
    },
    x0: "A1",
  };

  const layout = {
    title: "",
    yaxis: {
      zeroline: false,
      title: "Spheroid Diameter (µm)",
      range: [50, 500],
      fixedrange: true,
    },
    xaxis: {
      title: "Wells ID",
    },
    margin: {
      l: 50,
      r: 10,
      b: 50,
      t: 10,
    },
    showlegend: false,
  };

  Plotly.newPlot(elementId, [trace], layout, {
    displayModeBar: false,
    responsive: true,
  });
}

function handleManualInput(inputElement) {
  //function of the number of spheroids
  const value = parseFloat(inputElement.value);
  const circle = inputElement.closest(".well-circle");

  circle.classList.remove("status-green", "status-black");

  if (isNaN(value) || inputElement.value === "") {
    return;
  }

  if (value < 50) {
    circle.classList.add("status-black");
  } else {
    circle.classList.add("status-green");
  }
}

window.onbeforeprint = function () {
  const plotDivs = document.querySelectorAll(".violin-plot");
  plotDivs.forEach((plotDiv) => {
    Plotly.relayout(plotDiv, {
      width: 600,
      height: 400,
    });
  });
};

window.onafterprint = function () {
  const plotDivs = document.querySelectorAll(".violin-plot");
  plotDivs.forEach((plotDiv) => {
    Plotly.relayout(plotDiv, {
      width: null,
      height: 500,
    });
  });
};
