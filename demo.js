const XLSX = require('xlsx');

// Read the file
const workbook = XLSX.readFile('segmentation_results_bare.xlsx');

// Get the first sheet name
const sheetName = workbook.SheetNames[0];

// Get the worksheet
const worksheet = workbook.Sheets[sheetName];

const fs = require('fs');

// Convert to JSON
const jsonData = XLSX.utils.sheet_to_json(worksheet);

console.log("JSON Data:");
console.log(jsonData);

// Save JSON to file
fs.writeFileSync('output.json', JSON.stringify(jsonData, null, 2));
console.log("JSON data saved to output.json");

// Convert to CSV
const csvData = XLSX.utils.sheet_to_csv(worksheet);
console.log("\nCSV Data:");
console.log(csvData);

// Save CSV to file
fs.writeFileSync('output.csv', csvData);
console.log("CSV data saved to output.csv");

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

    // Calculate stats for each well
    // We want output with columns A1...D6, so we'll create objects for Average and Median
    const averageRow = { Metric: "Average Diameter" };
    const medianRow = { Metric: "Median Diameter" };
    const stdDevRow = { Metric: "Standard Deviation" };

    // Sort well keys to ensure order (optional but nice)
    const sortedWells = Object.keys(wells).sort((a, b) => {
        const aRow = a.charAt(0);
        const bRow = b.charAt(0);
        const aCol = parseInt(a.slice(1));
        const bCol = parseInt(b.slice(1));
        if (aRow !== bRow) return aRow.localeCompare(bRow);
        return aCol - bCol;
    });

    sortedWells.forEach(well => {
        const values = wells[well];
        if (values.length === 0) {
            averageRow[well] = null;
            medianRow[well] = null;
            stdDevRow[well] = null;
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

        // Standard Deviation
        const squareDiffs = values.map(value => {
            const diff = value - avg;
            return diff * diff;
        });
        const avgSquareDiff = squareDiffs.reduce((a, b) => a + b, 0) / values.length;
        const stdDev = Math.sqrt(avgSquareDiff);
        stdDevRow[well] = stdDev;
    });

    stats.push(averageRow);
    stats.push(medianRow);
    stats.push(stdDevRow);

    return stats;
}

const stats = calculateStats(jsonData);
console.log("\nStats:");
console.log(stats);

fs.writeFileSync('stats.json', JSON.stringify(stats, null, 2));
console.log("Stats saved to stats.json");

