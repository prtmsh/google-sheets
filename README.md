# Spreadsheet Application

This is a web-based spreadsheet application designed to mimic the user interface and core functionalities of Google Sheets. The application provides features such as cell formatting, formula evaluation, data validation, and additional enhancements like data visualization and file import/export.

## Features

### 1. Spreadsheet Interface
- **UI**: The application replicates the visual design and layout of Google Sheets, including a toolbar with formatting options, a formula bar for input, and a grid-based cell structure.
- **Drag Functions**: Users can drag cell content, formulas, and selections across the grid, mirroring Google Sheets' behavior, with support for adjusting relative references in formulas.
- **Cell Dependencies**: Formulas dynamically update based on cell dependencies, tracked efficiently to ensure changes propagate correctly.
- **Cell Formatting**: Options include bold, italic, font size, text color, background color, and text alignment (left, center, right).
- **Row and Column Management**: Users can add or delete rows and columns, with column resizing available via drag handles in the header.

### 2. Mathematical Functions
The following functions are implemented for formula evaluation:
- `SUM`: Calculates the sum of a range of cells (e.g., `=SUM(A1:A5)`).
- `AVERAGE`: Computes the average of a range of cells.
- `MAX`: Returns the maximum value in a range.
- `MIN`: Returns the minimum value in a range.
- `COUNT`: Counts the number of cells with numerical values in a range.

### 3. Data Quality Functions
These functions enhance data consistency:
- `TRIM`: Removes leading and trailing whitespace from selected cells.
- `UPPER`: Converts text in selected cells to uppercase.
- `LOWER`: Converts text in selected cells to lowercase.
- `REMOVE_DUPLICATES`: Removes duplicate rows within a selected range.
- `FIND_AND_REPLACE`: Allows users to find and replace text across a selected range.

### 4. Data Entry and Validation
- **Supported Data Types**: Users can input numbers, text, dates (in `YYYY-MM-DD` format), and dropdown selections.
- **Validation**: Ensures data adheres to specified types (e.g., numbers only in numeric cells, valid dates, or predefined dropdown options).

### 5. Bonus Features
- **Additional Mathematical Functions**: 
  - `MEDIAN`: Calculates the median of a range.
  - `STDEV`: Computes the standard deviation of a range.
  - `IF`: Evaluates a condition and returns one of two values (e.g., `=IF(A1>10, "High", "Low")`).
- **Cell Referencing**: Supports relative (e.g., `A1`) and absolute references (e.g., `$A$1`) in formulas.
- **Save/Load**: Stores spreadsheet state in local storage for persistence between sessions.
- **CSV Import/Export**: Allows exporting the grid to a CSV file and importing data from CSV files.
- **Data Visualization**: Generates bar charts from selected ranges using Chart.js.

## Tech Stack and Data Structures

### Tech Stack
- **HTML/CSS/JavaScript**: Core web technologies ensure broad compatibility and straightforward deployment in browsers.
- **Math.js**: Handles formula parsing and evaluation, providing robust support for mathematical operations and custom functions.
- **Chart.js**: Enables bar chart generation, offering an easy-to-use API for data visualization.
- **Font Awesome**: Supplies icons for the toolbar, enhancing the UI with recognizable symbols.

### Data Structures
- **HTML `<table>` Element**: Represents the spreadsheet grid, leveraging its native row-column structure for easy cell access and rendering.
- **Dependencies Map (`Map`)**: Tracks formula dependencies (e.g., which cells a formula references), enabling efficient updates when referenced cells change.
- **Undo/Redo Stacks (`Array`)**: Stores snapshots of the spreadsheet state, supporting undo (`Ctrl+Z`) and redo (`Ctrl+Y`) operations.
- **Selected Range Object**: Holds the coordinates of the currently selected cell range (e.g., `{minCol, maxCol, minRow, maxRow}`), facilitating multi-cell operations.

### Reasoning
- **HTML `<table>`**: Its grid layout aligns perfectly with a spreadsheet’s structure, simplifying DOM manipulation and styling.
- **Math.js**: Chosen for its comprehensive math capabilities and ability to parse complex expressions, reducing custom implementation effort.
- **Chart.js**: Selected for its lightweight footprint and effectiveness in visualizing data, enhancing user experience without complexity.
- **Map for Dependencies**: Offers O(1) lookup and update performance, critical for real-time formula recalculation in large grids.
- **Arrays for Undo/Redo**: Simple and efficient for stack-based operations, with `push` and `pop` aligning naturally with state management needs.

## Running the Application
Access the deployed version here: https://prtmsh.github.io/google-sheets/

Running Locally? Follow these steps:
1. **Clone the repo**: Obtain the source code by cloning the repository.
2. **Open `index.html`**: Launch `index.html` in a web browser (Chrome or Firefox recommended for optimal compatibility).
3. **Start Using**: The application loads immediately and is ready for interaction.

**Note**: Ensure local storage is enabled in your browser for save/load functionality to work.