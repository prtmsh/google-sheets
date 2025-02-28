document.addEventListener("DOMContentLoaded", () => {
    const spreadsheet = document.querySelector(".spreadsheet");
    const table = document.createElement("table");

    // Create header row
    const headerRow = document.createElement("tr");
    headerRow.appendChild(document.createElement("th")); // Empty cell
    for (let i = 1; i <= 10; i++) {
        const th = document.createElement("th");
        th.textContent = String.fromCharCode(64 + i); // A to J
        headerRow.appendChild(th);
    }
    table.appendChild(headerRow);

    // Create data rows
    for (let i = 1; i <= 10; i++) {
        const tr = document.createElement("tr");
        const th = document.createElement("th");
        th.textContent = i;
        tr.appendChild(th);
        for (let j = 1; j <= 10; j++) {
            const td = document.createElement("td");
            td.setAttribute("contenteditable", "true");
            td.dataset.row = i;
            td.dataset.col = String.fromCharCode(64 + j);
            tr.appendChild(td);
        }
        table.appendChild(tr);
    }

    spreadsheet.appendChild(table);
});

let selectedCell = null;

document.addEventListener("click", (e) => {
    if (e.target.tagName === "TD") {
        selectedCell = e.target;
        const formulaInput = document.getElementById("formula-input");
        formulaInput.value = selectedCell.textContent;
    }
});

document.getElementById("formula-input").addEventListener("input", (e) => {
    if (selectedCell) {
        selectedCell.textContent = e.target.value;
    }
});

document.getElementById("bold-btn").addEventListener("click", () => {
    if (selectedCell) {
        const isBold = selectedCell.style.fontWeight === "bold";
        selectedCell.style.fontWeight = isBold ? "normal" : "bold";
    }
});

document.getElementById("italic-btn").addEventListener("click", () => {
    if (selectedCell) {
        const isItalic = selectedCell.style.fontStyle === "italic";
        selectedCell.style.fontStyle = isItalic ? "normal" : "italic";
    }
});

function evaluateFormula(formula, cell) {
    if (formula.startsWith("=")) {
        const func = formula.slice(1).toUpperCase();
        const rangeMatch = func.match(/\((.*?)\)/);
        if (!rangeMatch) return "Invalid formula";
        const [start, end] = rangeMatch[1].split(":");
        const values = getRangeValues(start, end);

        switch (func.split("(")[0]) {
            case "SUM":
                return values.reduce((a, b) => a + b, 0);
            case "AVERAGE":
                return values.reduce((a, b) => a + b, 0) / values.length;
            case "MAX":
                return Math.max(...values);
            case "MIN":
                return Math.min(...values);
            case "COUNT":
                return values.length;
            default:
                return "Invalid function";
        }
    }
    return formula;
}

function getRangeValues(start, end) {
    const startCol = start.charCodeAt(0) - 64;
    const startRow = parseInt(start.slice(1));
    const endCol = end.charCodeAt(0) - 64;
    const endRow = parseInt(end.slice(1));

    const values = [];
    for (let i = startRow; i <= endRow; i++) {
        for (let j = startCol; j <= endCol; j++) {
            const cell = document.querySelector(
                `td[data-row="${i}"][data-col="${String.fromCharCode(64 + j)}"]`
            );
            values.push(parseFloat(cell.textContent) || 0);
        }
    }
    return values;
}

document.getElementById("formula-input").addEventListener("change", (e) => {
    if (selectedCell) {
        const formula = e.target.value;
        const result = evaluateFormula(formula, selectedCell);
        selectedCell.textContent = result;
    }
});

function applyDataQualityFunction(func, cell) {
    let value = cell.textContent;
    switch (func) {
        case "TRIM":
            value = value.trim();
            break;
        case "UPPER":
            value = value.toUpperCase();
            break;
        case "LOWER":
            value = value.toLowerCase();
            break;
    }
    cell.textContent = value;
}

function getRangeCells(start, end) {
    const startCol = start.charCodeAt(0) - 64;
    const startRow = parseInt(start.slice(1));
    const endCol = end.charCodeAt(0) - 64;
    const endRow = parseInt(end.slice(1));

    const cells = [];
    for (let i = startRow; i <= endRow; i++) {
        const row = [];
        for (let j = startCol; j <= endCol; j++) {
            const cell = document.querySelector(
                `td[data-row="${i}"][data-col="${String.fromCharCode(64 + j)}"]`
            );
            row.push(cell);
        }
        cells.push(row);
    }
    return cells;
}

document
    .getElementById("remove-duplicates-btn")
    .addEventListener("click", () => {
        const start = document.getElementById("start-cell").value;
        const end = document.getElementById("end-cell").value;
        const range = getRangeCells(start, end);
        const uniqueRows = [];
        const seen = new Set();

        range.forEach((row) => {
            const rowValues = row.map((cell) => cell.textContent).join(",");
            if (!seen.has(rowValues)) {
                seen.add(rowValues);
                uniqueRows.push(row);
            }
        });

        // For simplicity, clear the range and repopulate with unique rows
        range.forEach((row, i) => {
            row.forEach((cell) => {
                cell.textContent = uniqueRows[i]
                    ? uniqueRows[i][row.indexOf(cell)].textContent
                    : "";
            });
        });
    });

document.getElementById("find-replace-btn").addEventListener("click", () => {
    const start = document.getElementById("start-cell").value;
    const end = document.getElementById("end-cell").value;
    const findText = document.getElementById("find-text").value;
    const replaceText = document.getElementById("replace-text").value;
    const range = getRangeCells(start, end);

    range.forEach((row) => {
        row.forEach((cell) => {
            cell.textContent = cell.textContent.replace(
                new RegExp(findText, "g"),
                replaceText
            );
        });
    });
});

// Apply TRIM, UPPER, LOWER via console for now, e.g., applyDataQualityFunction('TRIM', selectedCell);
function setCellAsNumeric(cell) {
    cell.classList.add("numeric");
    cell.addEventListener("input", (e) => {
        if (!/^\d*\.?\d*$/.test(e.target.textContent)) {
            e.target.textContent = "";
            alert("Only numbers allowed");
        }
    });
}

// Example usage: setCellAsNumeric(selectedCell) in console
document.getElementById("load-sample-data").addEventListener("click", () => {
    for (let i = 1; i <= 5; i++) {
        const cell = document.querySelector(
            `td[data-row="1"][data-col="${String.fromCharCode(64 + i)}"]`
        );
        cell.textContent = i;
    }
});
function saveSpreadsheet() {
    const data = {};
    document.querySelectorAll("td").forEach((cell) => {
        const row = cell.dataset.row;
        const col = cell.dataset.col;
        data[`${col}${row}`] = cell.textContent;
    });
    localStorage.setItem("spreadsheet", JSON.stringify(data));
}

function loadSpreadsheet() {
    const data = JSON.parse(localStorage.getItem("spreadsheet") || "{}");
    for (const [key, value] of Object.entries(data)) {
        const col = key.match(/[A-Z]+/)[0];
        const row = key.match(/\d+/)[0];
        const cell = document.querySelector(
            `td[data-row="${row}"][data-col="${col}"]`
        );
        if (cell) cell.textContent = value;
    }
}

document.getElementById("save-btn").addEventListener("click", saveSpreadsheet);
document.getElementById("load-btn").addEventListener("click", loadSpreadsheet);
document.getElementById("generate-chart").addEventListener("click", () => {
    const start = document.getElementById("start-cell").value;
    const end = document.getElementById("end-cell").value;
    const range = getRangeValues(start, end);

    const ctx = document.getElementById("chart").getContext("2d");
    new Chart(ctx, {
        type: "bar",
        data: {
            labels: range.map((_, i) => `Label ${i + 1}`),
            datasets: [
                {
                    label: "Data",
                    data: range,
                    backgroundColor: "rgba(0, 123, 255, 0.5)",
                },
            ],
        },
    });
});