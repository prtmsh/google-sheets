document.addEventListener("DOMContentLoaded", () => {
    const spreadsheet = document.getElementById("spreadsheet");
    let table = createTable(100, 26); // 100 rows, A-Z columns
    spreadsheet.appendChild(table);

    let isSelecting = false, isDragging = false, isResizing = false;
    let startCell = null, endCell = null, dragTarget = null, resizeTarget = null;
    let selectedRange = null;
    const dependencies = new Map();
    const undoStack = [];
    const redoStack = [];
    let clipboard = null;

    // **Create Table**
    function createTable(rows, cols) {
        const t = document.createElement("table");
        const headerRow = document.createElement("tr");
        headerRow.appendChild(document.createElement("th"));
        for (let i = 1; i <= cols; i++) {
            const th = document.createElement("th");
            th.textContent = String.fromCharCode(64 + i);
            th.classList.add("resizable");
            const resizeHandle = document.createElement("div");
            resizeHandle.classList.add("resize-handle");
            th.appendChild(resizeHandle);
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
                td.setAttribute("tabindex", "0"); // Accessibility
                td.dataset.row = i;
                td.dataset.col = String.fromCharCode(64 + j);
                tr.appendChild(td);
            }
            t.appendChild(tr);
        }
        return t;
    }

    // **Utility Functions**
    function colToNum(col) { return col.charCodeAt(0) - 64; }
    function numToCol(num) { return String.fromCharCode(64 + num); }

    function getRangeCells(start, end) {
        const startMatch = start.match(/(\$?[A-Z]+)(\$?\d+)/);
        const endMatch = end.match(/(\$?[A-Z]+)(\$?\d+)/);
        const startCol = colToNum(startMatch[1].replace("$", ""));
        const startRow = parseInt(startMatch[2].replace("$", ""));
        const endCol = colToNum(endMatch[1].replace("$", ""));
        const endRow = parseInt(endMatch[2].replace("$", ""));
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

    function saveState() {
        const state = {};
        document.querySelectorAll("td").forEach(cell => {
            const key = `${cell.dataset.col}${cell.dataset.row}`;
            state[key] = {
                value: cell.textContent,
                styles: getCellStyles(cell),
                type: cell.dataset.type || "text",
                options: cell.dataset.options || ""
            };
        });
        return state;
    }

    function pushUndo() {
        undoStack.push(saveState());
        redoStack.length = 0;
        if (undoStack.length > 50) undoStack.shift();
    }

    function applyState(state) {
        for (const [key, { value, styles, type, options }] of Object.entries(state)) {
            const col = key.match(/[A-Z]+/)[0];
            const row = key.match(/\d+/)[0];
            const cell = document.querySelector(`td[data-row="${row}"][data-col="${col}"]`);
            if (cell) {
                cell.textContent = value;
                Object.assign(cell.style, styles);
                cell.dataset.type = type;
                cell.dataset.options = options;
                if (value.startsWith("=")) evaluateCell(cell);
            }
        }
    }

    // **Selection Handling**
    table.addEventListener("mousedown", (e) => {
        if (e.target.tagName === "DIV" && e.target.classList.contains("resize-handle")) {
            isResizing = true;
            resizeTarget = e.target.parentElement;
            return;
        }
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
        if (e.target.tagName !== "TD" && !isResizing) return;
        if (isSelecting) {
            endCell = e.target;
            selectCells(startCell, endCell);
        } else if (isDragging) {
            dragTarget = e.target;
            highlightDragTarget();
        } else if (isResizing) {
            const th = resizeTarget;
            const newWidth = e.clientX - th.getBoundingClientRect().left;
            if (newWidth > 50) {
                const colIdx = colToNum(th.textContent);
                document.querySelectorAll(`td[data-col="${th.textContent}"]`).forEach(cell => {
                    cell.style.width = `${newWidth}px`;
                });
                th.style.width = `${newWidth}px`;
            }
        }
    });

    table.addEventListener("mouseup", () => {
        if (isSelecting) {
            isSelecting = false;
            selectedRange = {
                minCol: colToNum(startCell.dataset.col),
                maxCol: colToNum(endCell.dataset.col),
                minRow: parseInt(startCell.dataset.row),
                maxRow: parseInt(endCell.dataset.row)
            };
            selectedRange = {
                minCol: Math.min(selectedRange.minCol, selectedRange.maxCol),
                maxCol: Math.max(selectedRange.minCol, selectedRange.maxCol),
                minRow: Math.min(selectedRange.minRow, selectedRange.maxRow),
                maxRow: Math.max(selectedRange.minRow, selectedRange.maxRow)
            };
        }
        if (isDragging) {
            isDragging = false;
            pushUndo();
            copySelectionToTarget();
            document.querySelectorAll("td").forEach(cell => cell.classList.remove("drag-target"));
        }
        if (isResizing) isResizing = false;
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
        if (!dragTarget || !selectedRange) return;
        const targetCol = colToNum(dragTarget.dataset.col);
        const targetRow = parseInt(dragTarget.dataset.row);
        const selectionWidth = selectedRange.maxCol - selectedRange.minCol + 1;
        const selectionHeight = selectedRange.maxRow - selectedRange.minRow + 1;
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
        if (!selectedRange || !dragTarget) return;
        const targetCol = colToNum(dragTarget.dataset.col);
        const targetRow = parseInt(dragTarget.dataset.row);
        const selectionWidth = selectedRange.maxCol - selectedRange.minCol + 1;
        const selectionHeight = selectedRange.maxRow - selectedRange.minRow + 1;
        for (let i = 0; i < selectionHeight; i++) {
            for (let j = 0; j < selectionWidth; j++) {
                const sourceCol = selectedRange.minCol + j;
                const sourceRow = selectedRange.minRow + i;
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
                    Object.assign(targetCell.style, getCellStyles(sourceCell));
                    targetCell.dataset.type = sourceCell.dataset.type;
                    targetCell.dataset.options = sourceCell.dataset.options;
                }
            }
        }
    }

    function adjustFormula(formula, colOffset, rowOffset) {
        return formula.replace(/(\$?[A-Z]+)(\$?\d+)/g, (match, col, row) => {
            const isColFixed = col.startsWith("$");
            const isRowFixed = row.startsWith("$");
            col = col.replace("$", "");
            row = row.replace("$", "");
            const newCol = isColFixed ? colToNum(col) : colToNum(col) + colOffset;
            const newRow = isRowFixed ? parseInt(row) : parseInt(row) + rowOffset;
            if (newCol < 1 || newRow < 1) return "#REF!";
            return `${isColFixed ? "$" : ""}${numToCol(newCol)}${isRowFixed ? "$" : ""}${newRow}`;
        });
    }

    // **Cell Formatting**
    function applyFormatting(property, value) {
        pushUndo();
        document.querySelectorAll("td.selected").forEach(cell => {
            cell.style[property] = value;
        });
    }

    document.getElementById("bold-btn").addEventListener("click", () => {
        const selected = document.querySelector("td.selected");
        applyFormatting("fontWeight", selected && selected.style.fontWeight === "bold" ? "normal" : "bold");
    });

    document.getElementById("italic-btn").addEventListener("click", () => {
        const selected = document.querySelector("td.selected");
        applyFormatting("fontStyle", selected && selected.style.fontStyle === "italic" ? "normal" : "italic");
    });

    document.getElementById("font-size").addEventListener("change", (e) => applyFormatting("fontSize", `${e.target.value}px`));
    document.getElementById("font-color").addEventListener("change", (e) => applyFormatting("color", e.target.value));
    document.getElementById("bg-color").addEventListener("change", (e) => applyFormatting("backgroundColor", e.target.value));
    document.getElementById("align-left-btn").addEventListener("click", () => applyFormatting("textAlign", "left"));
    document.getElementById("align-center-btn").addEventListener("click", () => applyFormatting("textAlign", "center"));
    document.getElementById("align-right-btn").addEventListener("click", () => applyFormatting("textAlign", "right"));

    // **Row/Column Management**
    function updateTable() {
        document.querySelectorAll("td").forEach(cell => {
            if (cell.textContent.startsWith("=")) evaluateCell(cell);
        });
    }

    document.getElementById("add-row-btn").addEventListener("click", () => {
        pushUndo();
        const rows = table.rows.length - 1;
        const cols = table.rows[0].cells.length - 1;
        const tr = document.createElement("tr");
        const th = document.createElement("th");
        th.textContent = rows + 1;
        tr.appendChild(th);
        for (let j = 1; j <= cols; j++) {
            const td = document.createElement("td");
            td.setAttribute("contenteditable", "true");
            td.setAttribute("tabindex", "0");
            td.dataset.row = rows + 1;
            td.dataset.col = numToCol(j);
            tr.appendChild(td);
        }
        table.appendChild(tr);
        updateTable();
    });

    document.getElementById("delete-row-btn").addEventListener("click", () => {
        if (table.rows.length > 2) {
            pushUndo();
            table.deleteRow(table.rows.length - 1);
            updateTable();
        }
    });

    document.getElementById("add-col-btn").addEventListener("click", () => {
        pushUndo();
        const cols = table.rows[0].cells.length - 1;
        const th = document.createElement("th");
        th.textContent = numToCol(cols + 1);
        th.classList.add("resizable");
        const resizeHandle = document.createElement("div");
        resizeHandle.classList.add("resize-handle");
        th.appendChild(resizeHandle);
        table.rows[0].appendChild(th);
        for (let i = 1; i < table.rows.length; i++) {
            const td = document.createElement("td");
            td.setAttribute("contenteditable", "true");
            td.setAttribute("tabindex", "0");
            td.dataset.row = i;
            td.dataset.col = numToCol(cols + 1);
            table.rows[i].appendChild(td);
        }
        updateTable();
    });

    document.getElementById("delete-col-btn").addEventListener("click", () => {
        if (table.rows[0].cells.length > 2) {
            pushUndo();
            for (let i = 0; i < table.rows.length; i++) {
                table.rows[i].deleteCell(table.rows[i].cells.length - 1);
            }
            updateTable();
        }
    });

    // **Formula Handling**
    const formulaInput = document.getElementById("formula-input");
    table.addEventListener("click", (e) => {
        if (e.target.tagName === "TD") {
            formulaInput.value = e.target.textContent;
        }
    });

    formulaInput.addEventListener("change", (e) => {
        const selected = document.querySelector("td.selected");
        if (selected) {
            pushUndo();
            selected.textContent = e.target.value;
            evaluateCell(selected);
            updateDependencies(selected);
        }
    });

    table.addEventListener("input", (e) => {
        if (e.target.tagName === "TD") {
            const type = e.target.dataset.type || "text";
            const value = e.target.textContent;
            if (type === "number" && !/^\d*\.?\d*$/.test(value)) {
                e.target.textContent = "";
                alert("Only numbers allowed in numeric cells");
                return;
            }
            if (type === "date" && value && !isValidDate(value)) {
                e.target.textContent = "";
                alert("Invalid date (use YYYY-MM-DD, e.g., 2023-12-31)");
                return;
            }
            if (type === "dropdown" && e.target.dataset.options) {
                const options = e.target.dataset.options.split(",");
                if (value && !options.includes(value)) {
                    e.target.textContent = "";
                    alert("Value must be one of: " + options.join(", "));
                    return;
                }
            }
            pushUndo();
            evaluateCell(e.target);
            updateDependencies(e.target);
        }
    });

    function evaluateCell(cell, visited = new Set()) {
        let value = cell.textContent.trim();
        if (!value.startsWith("=")) return;
        const formula = value.slice(1);
        const sourceKey = `${cell.dataset.col}${cell.dataset.row}`;

        if (visited.has(sourceKey)) {
            cell.textContent = "#CIRCULAR!";
            return;
        }
        visited.add(sourceKey);

        dependencies.set(sourceKey, []);

        try {
            const parsed = formula.replace(/(\$?[A-Z]+)(\$?\d+)/g, (match) => {
                const col = match.match(/\$?[A-Z]+/)[0].replace("$", "");
                const row = match.match(/\$?\d+/)[0].replace("$", "");
                const refCell = document.querySelector(`td[data-row="${row}"][data-col="${col}"]`);
                const refValue = refCell ? (parseFloat(refCell.textContent) || refCell.textContent || 0) : "#REF!";
                dependencies.get(sourceKey).push(`${col}${row}`);
                return refValue;
            });

            const funcMatch = formula.match(/^([A-Z]+)\((.*?)\)$/);
            if (funcMatch) {
                const [, func, args] = funcMatch;
                const rangeMatch = args.match(/(\$?[A-Z]+\$?\d+):(\$?[A-Z]+\$?\d+)/);
                if (rangeMatch) {
                    const values = getRangeValues(rangeMatch[1], rangeMatch[2]);
                    switch (func.toUpperCase()) {
                        case "SUM": cell.textContent = values.reduce((a, b) => a + b, 0); break;
                        case "AVERAGE": cell.textContent = values.length ? values.reduce((a, b) => a + b, 0) / values.length : 0; break;
                        case "MAX": cell.textContent = Math.max(...values); break;
                        case "MIN": cell.textContent = Math.min(...values); break;
                        case "COUNT": cell.textContent = values.filter(v => !isNaN(v) && v !== "").length; break;
                        case "MEDIAN": cell.textContent = math.median(values); break;
                        case "STDEV": cell.textContent = math.std(values); break;
                        case "IF": {
                            const [cond, val1, val2] = args.split(",");
                            const conditionResult = math.evaluate(cond.replace(/(\$?[A-Z]+\$?\d+)/g, (m) => {
                                const col = m.match(/\$?[A-Z]+/)[0].replace("$", "");
                                const row = m.match(/\$?\d+/)[0].replace("$", "");
                                const cell = document.querySelector(`td[data-row="${row}"][data-col="${col}"]`);
                                return cell ? (parseFloat(cell.textContent) || 0) : "#REF!";
                            }));
                            cell.textContent = conditionResult ? val1.trim() : val2.trim();
                            break;
                        }
                        default: cell.textContent = "#NAME?";
                    }
                } else {
                    cell.textContent = math.evaluate(parsed);
                }
            } else {
                cell.textContent = math.evaluate(parsed);
            }
            if (!isFinite(cell.textContent) || isNaN(cell.textContent)) throw new Error("#DIV/0!");
        } catch (e) {
            cell.textContent = e.message.startsWith("#") ? e.message : "#ERROR";
        }
    }

    function getRangeValues(start, end) {
        const cells = getRangeCells(start, end);
        return cells.flat().map(cell => parseFloat(cell.textContent) || 0);
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
        if (!selectedRange) {
            alert("Please select a range first.");
            return;
        }
        pushUndo();
        document.querySelectorAll("td.selected").forEach(cell => {
            let value = cell.textContent;
            switch (func) {
                case "TRIM": cell.textContent = value.trim(); break;
                case "UPPER": cell.textContent = value.toUpperCase(); break;
                case "LOWER": cell.textContent = value.toLowerCase(); break;
            }
        });
    }

    document.getElementById("trim-btn").addEventListener("click", () => applyDataQualityFunction("TRIM"));
    document.getElementById("upper-btn").addEventListener("click", () => applyDataQualityFunction("UPPER"));
    document.getElementById("lower-btn").addEventListener("click", () => applyDataQualityFunction("LOWER"));

    document.getElementById("remove-duplicates-btn").addEventListener("click", () => {
        if (!selectedRange) {
            alert("Please select a range first.");
            return;
        }
        pushUndo();
        const range = getRangeCells(`${numToCol(selectedRange.minCol)}${selectedRange.minRow}`, `${numToCol(selectedRange.maxCol)}${selectedRange.maxRow}`);
        const seen = new Set();
        const uniqueRows = [];
        range.forEach(row => {
            const rowKey = row.map(cell => cell.textContent || "").join(",");
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

    document.getElementById("find-replace-btn").addEventListener("click", () => {
        if (!selectedRange) {
            alert("Please select a range first.");
            return;
        }
        const findText = document.getElementById("find-text").value;
        const replaceText = document.getElementById("replace-text").value;
        if (!findText) {
            alert("Please enter text to find.");
            return;
        }
        pushUndo();
        const range = getRangeCells(`${numToCol(selectedRange.minCol)}${selectedRange.minRow}`, `${numToCol(selectedRange.maxCol)}${selectedRange.maxRow}`);
        range.forEach(row => {
            row.forEach(cell => {
                cell.textContent = cell.textContent.replace(new RegExp(findText, "g"), replaceText);
            });
        });
    });

    // **Additional Mathematical Functions**
    document.getElementById("median-btn").addEventListener("click", () => {
        if (!selectedRange) {
            alert("Please select a range first.");
            return;
        }
        pushUndo();
        const values = getRangeValues(`${numToCol(selectedRange.minCol)}${selectedRange.minRow}`, `${numToCol(selectedRange.maxCol)}${selectedRange.maxRow}`);
        const targetCell = document.querySelector("td.selected");
        if (targetCell) {
            targetCell.textContent = `=MEDIAN(${numToCol(selectedRange.minCol)}${selectedRange.minRow}:${numToCol(selectedRange.maxCol)}${selectedRange.maxRow})`;
            evaluateCell(targetCell);
        }
    });

    document.getElementById("stdev-btn").addEventListener("click", () => {
        if (!selectedRange) {
            alert("Please select a range first.");
            return;
        }
        pushUndo();
        const values = getRangeValues(`${numToCol(selectedRange.minCol)}${selectedRange.minRow}`, `${numToCol(selectedRange.maxCol)}${selectedRange.maxRow}`);
        const targetCell = document.querySelector("td.selected");
        if (targetCell) {
            targetCell.textContent = `=STDEV(${numToCol(selectedRange.minCol)}${selectedRange.minRow}:${numToCol(selectedRange.maxCol)}${selectedRange.maxRow})`;
            evaluateCell(targetCell);
        }
    });

    // **Data Validation and Cell Types**
    const cellTypeSelect = document.getElementById("cell-type");
    const dropdownOptionsInput = document.getElementById("dropdown-options");
    cellTypeSelect.addEventListener("change", (e) => {
        const type = e.target.value;
        dropdownOptionsInput.style.display = type === "dropdown" ? "inline" : "none";
        if (!selectedRange) return;
        pushUndo();
        document.querySelectorAll("td.selected").forEach(cell => {
            cell.dataset.type = type;
            cell.dataset.options = type === "dropdown" ? dropdownOptionsInput.value : "";
            const value = cell.textContent;
            if (type === "number" && !/^\d*\.?\d*$/.test(value)) cell.textContent = "";
            if (type === "date" && value && !isValidDate(value)) cell.textContent = "";
            if (type === "dropdown" && dropdownOptionsInput.value) {
                const options = dropdownOptionsInput.value.split(",");
                if (value && !options.includes(value)) cell.textContent = "";
            }
        });
    });

    function isValidDate(dateStr) {
        const regex = /^\d{4}-\d{2}-\d{2}$/;
        if (!regex.test(dateStr)) return false;
        const [year, month, day] = dateStr.split("-").map(Number);
        const date = new Date(year, month - 1, day);
        return date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day;
    }

    // **Save/Load**
    document.getElementById("save-btn").addEventListener("click", () => {
        localStorage.setItem("spreadsheet", JSON.stringify(saveState()));
        alert("Spreadsheet saved!");
    });

    document.getElementById("load-btn").addEventListener("click", () => {
        const data = JSON.parse(localStorage.getItem("spreadsheet") || "{}");
        pushUndo();
        applyState(data);
        alert("Spreadsheet loaded!");
    });

    function getCellStyles(cell) {
        return {
            fontWeight: cell.style.fontWeight,
            fontStyle: cell.style.fontStyle,
            fontSize: cell.style.fontSize,
            color: cell.style.color,
            backgroundColor: cell.style.backgroundColor,
            textAlign: cell.style.textAlign
        };
    }

    // **Undo/Redo**
    document.getElementById("undo-btn").addEventListener("click", () => {
        if (undoStack.length) {
            redoStack.push(saveState());
            applyState(undoStack.pop());
        }
    });

    document.getElementById("redo-btn").addEventListener("click", () => {
        if (redoStack.length) {
            undoStack.push(saveState());
            applyState(redoStack.pop());
        }
    });

    // **CSV Export/Import**
    document.getElementById("export-csv-btn").addEventListener("click", () => {
        const rows = Array.from(table.rows).slice(1).map(row =>
            Array.from(row.cells).slice(1).map(cell => `"${cell.textContent.replace(/"/g, '""')}"`).join(",")
        );
        const csv = rows.join("\n");
        const blob = new Blob([csv], { type: "text/csv" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "spreadsheet.csv";
        a.click();
        URL.revokeObjectURL(url);
    });

    document.getElementById("import-csv-btn").addEventListener("click", () => {
        document.getElementById("csv-file").click();
    });

    document.getElementById("csv-file").addEventListener("change", (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (event) => {
            pushUndo();
            const csv = event.target.result;
            const rows = csv.split("\n").map(row => row.split(",").map(cell => cell.replace(/^"|"$/g, "").replace(/""/g, '"')));
            for (let i = 0; i < rows.length && i < table.rows.length - 1; i++) {
                for (let j = 0; j < rows[i].length && j < table.rows[0].cells.length - 1; j++) {
                    const cell = table.rows[i + 1].cells[j + 1];
                    cell.textContent = rows[i][j];
                    if (cell.textContent.startsWith("=")) evaluateCell(cell);
                }
            }
        };
        reader.readAsText(file);
        e.target.value = "";
    });

    // **Chart Generation**
    document.getElementById("generate-chart").addEventListener("click", () => {
        if (!selectedRange) {
            alert("Please select a range first.");
            return;
        }
        const values = getRangeValues(`${numToCol(selectedRange.minCol)}${selectedRange.minRow}`, `${numToCol(selectedRange.maxCol)}${selectedRange.maxRow}`);
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

    // **Testing**
    document.getElementById("load-sample-data").addEventListener("click", () => {
        pushUndo();
        for (let i = 1; i <= 5; i++) {
            const cell = document.querySelector(`td[data-row="1"][data-col="${numToCol(i)}"]`);
            if (cell) {
                cell.textContent = i * 10;
                cell.dataset.type = "number";
            }
        }
        const sumCell = document.querySelector('td[data-row="2"][data-col="A"]');
        if (sumCell) {
            sumCell.textContent = "=SUM(A1:E1)";
            evaluateCell(sumCell);
        }
    });

    document.getElementById("run-tests").addEventListener("click", () => {
        pushUndo();
        const tests = [
            { cell: "A1", value: "10" },
            { cell: "A2", value: "20" },
            { cell: "A3", value: "30" },
            { cell: "B1", value: "=SUM(A1:A3)" },
            { cell: "B2", value: "=AVERAGE(A1:A3)" },
            { cell: "B3", value: "=MEDIAN(A1:A3)" },
            { cell: "C1", value: "=IF(A1>15, 'High', 'Low')" },
        ];
        tests.forEach(test => {
            const col = test.cell.match(/[A-Z]+/)[0];
            const row = test.cell.match(/\d+/)[0];
            const cell = document.querySelector(`td[data-row="${row}"][data-col="${col}"]`);
            if (cell) {
                cell.textContent = test.value;
                if (test.value.startsWith("=")) evaluateCell(cell);
            }
        });
        alert("Test data loaded: Check A1:A3 for numbers, B1:B3 for functions, C1 for IF.");
    });

    // **Context Menu**
    const contextMenu = document.getElementById("context-menu");
    table.addEventListener("contextmenu", (e) => {
        if (e.target.tagName === "TD") {
            e.preventDefault();
            contextMenu.style.left = `${e.pageX}px`;
            contextMenu.style.top = `${e.pageY}px`;
            contextMenu.style.display = "block";
            selectCells(e.target, e.target);
        }
    });

    contextMenu.addEventListener("click", (e) => {
        const action = e.target.dataset.action;
        if (!action) return;
        const selectedCells = document.querySelectorAll("td.selected");
        if (action === "copy") {
            clipboard = Array.from(selectedCells).map(cell => ({
                value: cell.textContent,
                styles: getCellStyles(cell),
                type: cell.dataset.type,
                options: cell.dataset.options
            }));
        } else if (action === "paste" && clipboard) {
            pushUndo();
            selectedCells.forEach((cell, i) => {
                if (i < clipboard.length) {
                    cell.textContent = clipboard[i].value;
                    Object.assign(cell.style, clipboard[i].styles);
                    cell.dataset.type = clipboard[i].type;
                    cell.dataset.options = clipboard[i].options;
                    if (cell.textContent.startsWith("=")) evaluateCell(cell);
                }
            });
        } else if (action === "cut") {
            pushUndo();
            clipboard = Array.from(selectedCells).map(cell => ({
                value: cell.textContent,
                styles: getCellStyles(cell),
                type: cell.dataset.type,
                options: cell.dataset.options
            }));
            selectedCells.forEach(cell => {
                cell.textContent = "";
                cell.style = {};
                cell.dataset.type = "text";
                cell.dataset.options = "";
            });
        }
        contextMenu.style.display = "none";
    });

    document.addEventListener("click", () => {
        contextMenu.style.display = "none";
    });

    // **Keyboard Shortcuts**
    document.addEventListener("keydown", (e) => {
        if (e.ctrlKey) {
            if (e.key === "z") {
                document.getElementById("undo-btn").click();
                e.preventDefault();
            } else if (e.key === "y") {
                document.getElementById("redo-btn").click();
                e.preventDefault();
            } else if (e.key === "c") {
                document.querySelector("#context-menu [data-action='copy']").click();
                e.preventDefault();
            } else if (e.key === "v") {
                document.querySelector("#context-menu [data-action='paste']").click();
                e.preventDefault();
            } else if (e.key === "x") {
                document.querySelector("#context-menu [data-action='cut']").click();
                e.preventDefault();
            }
        }
    });
});