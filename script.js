document.addEventListener('DOMContentLoaded', function () {
  let spreadsheetCount = 0;
  let currentHot;
  const container = document.getElementById('spreadsheet');
  const bottomBar = document.getElementById('bottomBar');
  const suggestionsContainer = document.getElementById('suggestions');

  function generateColumnNames(num) {
    let result = '';
    while (num >= 0) {
      result = String.fromCharCode(num % 26 + 65) + result;
      num = Math.floor(num / 26) - 1;
    }
    return result;
  }

  function createSpreadsheet() {
    const initialData = [
      ['', '', '', '', ''],
      ['', '', '', '', ''],
      ['', '', '', '', ''],
      ['', '', '', '', ''],
      ['', '', '', '', '']
    ];

    const hot = new Handsontable(container, {
      data: initialData,
      rowHeaders: true,
      colHeaders: function() {
        return Array.from({ length: hot.countCols() }, (_, index) => generateColumnNames(index));
      },
      licenseKey: 'non-commercial-and-evaluation',
      stretchH: 'all',
      height: '100%', // Ensures the table fills the container's height
      width: '100%', // Ensures the table fills the container's width
      contextMenu: true,
      manualColumnResize: true,
      manualRowResize: true,
      formulas: true,
      columnSorting: true,
      allowInsertRow: true,
      allowInsertColumn: true,
      allowRemoveRow: true,
      allowRemoveColumn: true,
      cell: [{
        className: 'htNumeric'
      }],
      afterChange: function (changes, source) {
        if (source === 'edit') {
          // Detect formula changes and apply them
          changes.forEach(([row, col, oldVal, newVal]) => {
            if (newVal && newVal.startsWith('=')) {
              const formula = newVal.substring(1).trim();
              if (formula) {
                try {
                  // Evaluate the formula and set the result in the cell
                  const result = evalFormula(formula, hot, row, col);
                  hot.setDataAtCell(row, col, result);
                } catch (e) {
                  console.error('Formula error:', e);
                }
              }
            }
          });
        }
      }
    });

    currentHot = hot;

    const tab = document.createElement('div');
    tab.classList.add('tab');
    tab.innerText = `Sheet ${++spreadsheetCount}`;
    const closeBtn = document.createElement('span');
    closeBtn.innerText = '×';
    closeBtn.classList.add('close-btn');
    tab.appendChild(closeBtn);

    closeBtn.addEventListener('click', function() {
      bottomBar.removeChild(tab);
      hot.destroy();
    });

    tab.addEventListener('click', function() {
      document.querySelectorAll('.tab').forEach(tab => tab.classList.remove('active'));
      tab.classList.add('active');
    });

    bottomBar.appendChild(tab);
    tab.classList.add('active');
  }

  // Add new spreadsheet tab
  document.getElementById('newSpreadsheetBtn').addEventListener('click', function() {
    createSpreadsheet();
  });

  // Handle suggestions based on formula being typed
  container.addEventListener('keydown', function (e) {
    const activeCell = currentHot.getActiveCell();
    const activeRow = activeCell ? activeCell[0] : null;
    const activeCol = activeCell ? activeCell[1] : null;
    const cellData = currentHot.getDataAtCell(activeRow, activeCol);

    if (cellData.startsWith('=')) {
      const formulaText = cellData.substring(1).toUpperCase();

      if (formulaText.startsWith('SUM')) {
        const suggestions = generateSumSuggestions(activeRow, activeCol);
        showSuggestions(suggestions);
      } else if (formulaText.startsWith('AVERAGE')) {
        const suggestions = generateAverageSuggestions(activeRow, activeCol);
        showSuggestions(suggestions);
      }
      // Add more functions as needed
    } else {
      suggestionsContainer.style.display = 'none';
    }
  });

  function generateSumSuggestions(row, col) {
    const surroundingCells = currentHot.getData();
    const numRows = surroundingCells.length;
    const numCols = surroundingCells[0].length;
    const sumRange = [];

    for (let i = row - 1; i <= row + 1 && i >= 0 && i < numRows; i++) {
      for (let j = col - 1; j <= col + 1 && j >= 0 && j < numCols; j++) {
        if (typeof surroundingCells[i][j] === 'number') {
          sumRange.push(`${generateColumnNames(j)}${i + 1}`);
        }
      }
    }

    return sumRange.length > 0 ? [`=SUM(${sumRange.join(', ')})`] : [];
  }

  function generateAverageSuggestions(row, col) {
    const surroundingCells = currentHot.getData();
    const numRows = surroundingCells.length;
    const numCols = surroundingCells[0].length;
    const avgRange = [];

    for (let i = row - 1; i <= row + 1 && i >= 0 && i < numRows; i++) {
      for (let j = col - 1; j <= col + 1 && j >= 0 && j < numCols; j++) {
        if (typeof surroundingCells[i][j] === 'number') {
          avgRange.push(`${generateColumnNames(j)}${i + 1}`);
        }
      }
    }

    return avgRange.length > 0 ? [`=AVERAGE(${avgRange.join(', ')})`] : [];
  }

  function showSuggestions(suggestions) {
    suggestionsContainer.innerHTML = '';
    if (suggestions.length === 0) {
      suggestionsContainer.style.display = 'none';
      return;
    }

    suggestions.forEach(suggestion => {
      const suggestionDiv = document.createElement('div');
      suggestionDiv.classList.add('suggestion');
      suggestionDiv.innerText = suggestion;
      suggestionDiv.addEventListener('click', function () {
        const activeCell = currentHot.getActiveCell();
        const activeRow = activeCell ? activeCell[0] : null;
        const activeCol = activeCell ? activeCell[1] : null;
        currentHot.setDataAtCell(activeRow, activeCol, '=' + suggestion);
        suggestionsContainer.style.display = 'none';
      });
      suggestionsContainer.appendChild(suggestionDiv);
    });

    const activeCell = currentHot.getActiveCell();
    const rect = currentHot.getCellRect(activeCell[0], activeCell[1]);
    suggestionsContainer.style.top = rect.bottom + 'px';
    suggestionsContainer.style.left = rect.left + 'px';
    suggestionsContainer.style.display = 'block';
  }

});
