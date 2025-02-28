document.addEventListener("DOMContentLoaded", () => {
    const spreadsheet = document.getElementById("spreadsheet");
    let table = createTable(10, 10); // Initial 10x10 grid
    spreadsheet.appendChild(table);

    let isSelecting = false, isDragging = false;
    let startCell = null, endCell = null, dragTarget = null;
    let selectedMinCol, selectedMaxCol, selectedMinRow, selectedMaxRow;
    const dependencies = new Map(); // Track cell dependencies

    // **Create Table**
    function createTable(rows, cols) {
        const t = document.createElement("table");
        const headerRow = document.createElement("tr");
        headerRow.appendChild(document.createElement("th"));
        for (let i = 1; i <= cols; i++) {
            const th = document.createElement("th");
            th.textContent = String.fromCharCode(64 + i);
            headerRow.appendChild(th);
        }
        t.appendChild(headerRow);

        for (let i = 1; i <= rows; i++) {
            const tr = document.createElement("tr");
            const th = document.createElement("th");
            th.textContent = i;
            tr.appendChild(th);
            for (let j = 1; j <= cols; j++) {
                const td = document.createElement("td");
                td.setAttribute("contenteditable", "true");
                td.dataset.row = i;
                td.dataset.col = String.fromCharCode(64 + j);
                tr.appendChild(td);
            }
            t.appendChild(tr);
        }
        return t;
    }

    // **Column to Number**
    function colToNum(col) {
        return col.charCodeAt(0) - 64;
    }

    // **Number to Column**
    function numToCol(num) {
        return String.fromCharCode(64 + num);
    }

    // **Selection Handling**
    table.addEventListener("mousedown", (e) => {
        if (e.target.tagName !== "TD") return;
        if (e.target.classList.contains("selected")) {
            isDragging = true;
            dragTarget = e.target;
        } else {
            isSelecting = true;
            startCell = e.target;
            endCell = e.target;
            selectCells(startCell, endCell);
        }
    });

    table.addEventListener("mousemove", (e) => {
        if (e.target.tagName !== "TD") return;
        if (isSelecting) {
            endCell = e.target;
            selectCells(startCell, endCell);
        } else if (isDragging) {
            dragTarget = e.target;
            highlightDragTarget();
        }
    });

    table.addEventListener("mouseup", () => {
        if (isSelecting) {
            isSelecting = false;
            const startCol = colToNum(startCell.dataset.col);
            const startRow = parseInt(startCell.dataset.row);
            const endCol = colToNum(endCell.dataset.col);
            const endRow = parseInt(endCell.dataset.row);
            selectedMinCol = Math.min(startCol, endCol);
            selectedMaxCol = Math.max(startCol, endCol);
            selectedMinRow = Math.min(startRow, endRow);
            selectedMaxRow = Math.max(startRow, endRow);
        }
        if (isDragging) {
            isDragging = false;
            copySelectionToTarget();
            document.querySelectorAll("td").forEach(cell => cell.classList.remove("drag-target"));
        }
    });

    function selectCells(start, end) {
        document.querySelectorAll("td").forEach(cell => cell.classList.remove("selected"));
        const startCol = colToNum(start.dataset.col);
        const startRow = parseInt(start.dataset.row);
        const endCol = colToNum(end.dataset.col);
        const endRow = parseInt(end.dataset.row);
        const minCol = Math.min(startCol, endCol);
        const maxCol = Math.max(startCol, endCol);
        const minRow = Math.min(startRow, endRow);
        const maxRow = Math.max(startRow, endRow);
        for (let i = minRow; i <= maxRow; i++) {
            for (let j = minCol; j <= maxCol; j++) {
                const cell = document.querySelector(`td[data-row="${i}"][data-col="${numToCol(j)}"]`);
                if (cell) cell.classList.add("selected");
            }
        }
    }

    // **Drag Functionality**
    function highlightDragTarget() {
        document.querySelectorAll("td").forEach(cell => cell.classList.remove("drag-target"));
        if (!dragTarget) return;
        const targetCol = colToNum(dragTarget.dataset.col);
        const targetRow = parseInt(dragTarget.dataset.row);
        const selectionWidth = selectedMaxCol - selectedMinCol + 1;
        const selectionHeight = selectedMaxRow - selectedMinRow + 1;
        const maxCols = table.rows[0].cells.length - 1;
        const maxRows = table.rows.length - 1;
        for (let i = 0; i < selectionHeight; i++) {
            for (let j = 0; j < selectionWidth; j++) {
                const col = targetCol + j;
                const row = targetRow + i;
                if (col > maxCols || row > maxRows) continue;
                const cell = document.querySelector(`td[data-row="${row}"][data-col="${numToCol(col)}"]`);
                if (cell) cell.classList.add("drag-target");
            }
        }
    }

    function copySelectionToTarget() {
        const targetCol = colToNum(dragTarget.dataset.col);
        const targetRow = parseInt(dragTarget.dataset.row);
        const selectionWidth = selectedMaxCol - selectedMinCol + 1;
        const selectionHeight = selectedMaxRow - selectedMinRow + 1;
        for (let i = 0; i < selectionHeight; i++) {
            for (let j = 0; j < selectionWidth; j++) {
                const sourceCol = selectedMinCol + j;
                const sourceRow = selectedMinRow + i;
                const targetColNum = targetCol + j;
                const targetRowNum = targetRow + i;
                const sourceCell = document.querySelector(`td[data-row="${sourceRow}"][data-col="${numToCol(sourceCol)}"]`);
                const targetCell = document.querySelector(`td[data-row="${targetRowNum}"][data-col="${numToCol(targetColNum)}"]`);
                if (sourceCell && targetCell) {
                    const value = sourceCell.textContent;
                    if (value.startsWith("=")) {
                        const adjustedFormula = adjustFormula(value, j, i);
                        targetCell.textContent = adjustedFormula;
                        evaluateCell(targetCell);
                    } else {
                        targetCell.textContent = value;
                    }
                }
            }
        }
    }

    function adjustFormula(formula, colOffset, rowOffset) {
        return formula.replace(/([A-Z]+)(\d+)/g, (match, col, row) => {
            const newCol = colToNum(col) + colOffset;
            const newRow = parseInt(row) + rowOffset;
            return `${numToCol(newCol)}${newRow}`;
        });
    }

    // **Cell Formatting**
    document.getElementById("bold-btn").addEventListener("click", () => {
        document.querySelectorAll("td.selected").forEach(cell => {
            cell.style.fontWeight = cell.style.fontWeight === "bold" ? "normal" : "bold";
        });
    });

    document.getElementById("italic-btn").addEventListener("click", () => {
        document.querySelectorAll("td.selected").forEach(cell => {
            cell.style.fontStyle = cell.style.fontStyle === "italic" ? "normal" : "italic";
        });
    });

    document.getElementById("font-size").addEventListener("change", (e) => {
        document.querySelectorAll("td.selected").forEach(cell => {
            cell.style.fontSize = `${e.target.value}px`;
        });
    });

    document.getElementById("font-color").addEventListener("change", (e) => {
        document.querySelectorAll("td.selected").forEach(cell => {
            cell.style.color = e.target.value;
        });
    });

    // **Row/Column Management**
    function updateTable() {
        document.querySelectorAll("td").forEach(cell => {
            if (cell.textContent.startsWith("=")) evaluateCell(cell);
        });
    }

    document.getElementById("add-row-btn").addEventListener("click", () => {
        const rows = table.rows.length - 1;
        const cols = table.rows[0].cells.length - 1;
        const tr = document.createElement("tr");
        const th = document.createElement("th");
        th.textContent = rows + 1;
        tr.appendChild(th);
        for (let j = 1; j <= cols; j++) {
            const td = document.createElement("td");
            td.setAttribute("contenteditable", "true");
            td.dataset.row = rows + 1;
            td.dataset.col = numToCol(j);
            tr.appendChild(td);
        }
        table.appendChild(tr);
    });

    document.getElementById("delete-row-btn").addEventListener("click", () => {
        if (table.rows.length > 2) {
            table.deleteRow(table.rows.length - 1);
            updateTable();
        }
    });

    document.getElementById("add-col-btn").addEventListener("click", () => {
        const cols = table.rows[0].cells.length - 1;
        table.rows[0].appendChild(document.createElement("th")).textContent = numToCol(cols + 1);
        for (let i = 1; i < table.rows.length; i++) {
            const td = document.createElement("td");
            td.setAttribute("contenteditable", "true");
            td.dataset.row = i;
            td.dataset.col = numToCol(cols + 1);
            table.rows[i].appendChild(td);
        }
    });

    document.getElementById("delete-col-btn").addEventListener("click", () => {
        if (table.rows[0].cells.length > 2) {
            for (let i = 0; i < table.rows.length; i++) {
                table.rows[i].deleteCell(table.rows[i].cells.length - 1);
            }
            updateTable();
        }
    });

    // **Formula Input and Evaluation**
    const formulaInput = document.getElementById("formula-input");
    table.addEventListener("click", (e) => {
        if (e.target.tagName === "TD") {
            formulaInput.value = e.target.textContent;
        }
    });

    formulaInput.addEventListener("change", (e) => {
        const selected = document.querySelector("td.selected");
        if (selected) {
            selected.textContent = e.target.value;
            evaluateCell(selected);
            updateDependencies(selected);
        }
    });

    table.addEventListener("input", (e) => {
        if (e.target.tagName === "TD") {
            evaluateCell(e.target);
            updateDependencies(e.target);
        }
    });

    function evaluateCell(cell) {
        let value = cell.textContent.trim();
        if (!value.startsWith("=")) return;
        const formula = value.slice(1).toUpperCase();
        const rangeMatch = formula.match(/\((.*?)\)/);
        let result = "Invalid formula";

        if (rangeMatch) {
            const [start, end] = rangeMatch[1].split(":");
            const values = getRangeValues(start, end, cell);
            const func = formula.split("(")[0];
            switch (func) {
                case "SUM": result = values.reduce((a, b) => a + b, 0); break;
                case "AVERAGE": result = values.reduce((a, b) => a + b, 0) / values.length; break;
                case "MAX": result = Math.max(...values); break;
                case "MIN": result = Math.min(...values); break;
                case "COUNT": result = values.filter(v => !isNaN(v)).length; break;
                default: result = "Invalid function";
            }
        }
        cell.textContent = isNaN(result) ? result : result;
    }

    function getRangeValues(start, end, sourceCell) {
        const startCol = colToNum(start.match(/[A-Z]+/)[0]);
        const startRow = parseInt(start.match(/\d+/)[0]);
        const endCol = colToNum(end.match(/[A-Z]+/)[0]);
        const endRow = parseInt(end.match(/\d+/)[0]);
        const values = [];
        const sourceKey = `${sourceCell.dataset.col}${sourceCell.dataset.row}`;
        dependencies.set(sourceKey, []);

        for (let i = Math.min(startRow, endRow); i <= Math.max(startRow, endRow); i++) {
            for (let j = Math.min(startCol, endCol); j <= Math.max(startCol, endCol); j++) {
                const cell = document.querySelector(`td[data-row="${i}"][data-col="${numToCol(j)}"]`);
                if (cell) {
                    const val = parseFloat(cell.textContent) || 0;
                    values.push(val);
                    dependencies.get(sourceKey).push(`${numToCol(j)}${i}`);
                }
            }
        }
        return values;
    }

    function updateDependencies(changedCell) {
        const changedKey = `${changedCell.dataset.col}${changedCell.dataset.row}`;
        dependencies.forEach((deps, key) => {
            if (deps.includes(changedKey)) {
                const cell = document.querySelector(`td[data-row="${key.match(/\d+/)[0]}"][data-col="${key.match(/[A-Z]+/)[0]}"]`);
                if (cell) evaluateCell(cell);
            }
        });
    }

    // **Data Quality Functions**
    function applyDataQualityFunction(func) {
        document.querySelectorAll("td.selected").forEach(cell => {
            let value = cell.textContent;
            switch (func) {
                case "TRIM": cell.textContent = value.trim(); break;
                case "UPPER": cell.textContent = value.toUpperCase(); break;
                case "LOWER": cell.textContent = value.toLowerCase(); break;
            }
        });
    }

    // Apply via console: applyDataQualityFunction("TRIM")
    document.getElementById("remove-duplicates-btn").addEventListener("click", () => {
        const start = document.getElementById("start-cell").value;
        const end = document.getElementById("end-cell").value;
        const range = getRangeCells(start, end);
        const seen = new Set();
        const uniqueRows = [];

        range.forEach(row => {
            const rowKey = row.map(cell => cell.textContent).join(",");
            if (!seen.has(rowKey)) {
                seen.add(rowKey);
                uniqueRows.push(row.map(cell => cell.textContent));
            }
        });

        range.forEach((row, i) => {
            row.forEach((cell, j) => {
                cell.textContent = i < uniqueRows.length ? uniqueRows[i][j] : "";
            });
        });
    });

    function getRangeCells(start, end) {
        const startCol = colToNum(start.match(/[A-Z]+/)[0]);
        const startRow = parseInt(start.match(/\d+/)[0]);
        const endCol = colToNum(end.match(/[A-Z]+/)[0]);
        const endRow = parseInt(end.match(/\d+/)[0]);
        const cells = [];
        for (let i = Math.min(startRow, endRow); i <= Math.max(startRow, endRow); i++) {
            const row = [];
            for (let j = Math.min(startCol, endCol); j <= Math.max(startCol, endCol); j++) {
                const cell = document.querySelector(`td[data-row="${i}"][data-col="${numToCol(j)}"]`);
                if (cell) row.push(cell);
            }
            cells.push(row);
        }
        return cells;
    }

    document.getElementById("find-replace-btn").addEventListener("click", () => {
        const start = document.getElementById("start-cell").value;
        const end = document.getElementById("end-cell").value;
        const findText = document.getElementById("find-text").value;
        const replaceText = document.getElementById("replace-text").value;
        const range = getRangeCells(start, end);
        range.forEach(row => {
            row.forEach(cell => {
                cell.textContent = cell.textContent.replace(new RegExp(findText, "g"), replaceText);
            });
        });
    });

    // **Data Validation**
    table.addEventListener("input", (e) => {
        if (e.target.classList.contains("numeric") && !/^\d*\.?\d*$/.test(e.target.textContent)) {
            e.target.textContent = "";
            alert("Only numbers allowed in numeric cells");
        }
        if (e.target.classList.contains("date") && !isValidDate(e.target.textContent)) {
            e.target.textContent = "";
            alert("Invalid date format (use YYYY-MM-DD)");
        }
    });

    function setCellType(type) {
        document.querySelectorAll("td.selected").forEach(cell => {
            cell.classList.remove("numeric", "date");
            if (type === "numeric") cell.classList.add("numeric");
            if (type === "date") cell.classList.add("date");
        });
    }

    function isValidDate(dateStr) {
        const regex = /^\d{4}-\d{2}-\d{2}$/;
        if (!regex.test(dateStr)) return false;
        const date = new Date(dateStr);
        return date.toISOString().startsWith(dateStr);
    }

    // Example: setCellType("numeric") via console

    // **Testing and Sample Data**
    document.getElementById("load-sample-data").addEventListener("click", () => {
        for (let i = 1; i <= 5; i++) {
            const cell = document.querySelector(`td[data-row="1"][data-col="${numToCol(i)}"]`);
            cell.textContent = i * 10;
            setCellType("numeric");
        }
        document.querySelector('td[data-row="2"][data-col="A"]').textContent = "=SUM(A1:E1)";
        evaluateCell(document.querySelector('td[data-row="2"][data-col="A"]'));
    });

    // **Save/Load**
    document.getElementById("save-btn").addEventListener("click", () => {
        const data = {};
        document.querySelectorAll("td").forEach(cell => {
            const key = `${cell.dataset.col}${cell.dataset.row}`;
            data[key] = { value: cell.textContent, styles: getCellStyles(cell) };
        });
        localStorage.setItem("spreadsheet", JSON.stringify(data));
        alert("Spreadsheet saved!");
    });

    document.getElementById("load-btn").addEventListener("click", () => {
        const data = JSON.parse(localStorage.getItem("spreadsheet") || "{}");
        for (const [key, { value, styles }] of Object.entries(data)) {
            const col = key.match(/[A-Z]+/)[0];
            const row = key.match(/\d+/)[0];
            const cell = document.querySelector(`td[data-row="${row}"][data-col="${col}"]`);
            if (cell) {
                cell.textContent = value;
                Object.assign(cell.style, styles);
                if (value.startsWith("=")) evaluateCell(cell);
            }
        }
        alert("Spreadsheet loaded!");
    });

    function getCellStyles(cell) {
        return {
            fontWeight: cell.style.fontWeight,
            fontStyle: cell.style.fontStyle,
            fontSize: cell.style.fontSize,
            color: cell.style.color
        };
    }

    // **Chart Generation**
    document.getElementById("generate-chart").addEventListener("click", () => {
        const start = document.getElementById("start-cell").value;
        const end = document.getElementById("end-cell").value;
        const values = getRangeValues(start, end, document.querySelector("td.selected") || table.rows[1].cells[1]);
        const ctx = document.getElementById("chart").getContext("2d");
        if (window.myChart) window.myChart.destroy();
        document.getElementById("chart").style.display = "block";
        window.myChart = new Chart(ctx, {
            type: "bar",
            data: {
                labels: values.map((_, i) => `Cell ${i + 1}`),
                datasets: [{ label: "Data", data: values, backgroundColor: "rgba(0, 123, 255, 0.5)" }]
            },
            options: { scales: { y: { beginAtZero: true } } }
        });
    });
});