const excelFileInput = document.getElementById("excelFile");
const companyRowSelect = document.getElementById("companyRow");
const companyColumnSelect = document.getElementById("companyColumn");
const submitButton = document.getElementById("submitButton");
const downloadButton = document.getElementById("downloadButton");
const newCompanyInput = document.getElementById("newCompany");
const addCompanyButton = document.getElementById("addCompanyButton");

let workbook,
  worksheet,
  companyList = [];

// Function to fill diagonal cells with red for matching companies
function fillDiagonalCells() {
  companyList.forEach((company, index) => {
    const cell = worksheet.getCell(
      `${indexToColumnLetter(index + 1)}${index + 1}`
    );
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFF0000" }, // Red color
    };
  });
}

// Function to populate the data table
function populateDataTable(selectedRowCompany, selectedColumnCompany) {
  const dataBody = document.getElementById("dataBody");
  const fragment = document.createDocumentFragment();
  const rows = {};
  const existingRows = Array.from(dataBody.rows);

  const headerRow = worksheet.getRow(1);
  const columnCount = headerRow.values.length;

  for (let rowIndex = 2; rowIndex <= worksheet.rowCount; rowIndex++) {
    const row = worksheet.getRow(rowIndex);
    const rowCompany = row.getCell(1).value;

    // If a specific row company is selected and it doesn't match, skip this row
    if (selectedRowCompany && rowCompany !== selectedRowCompany) {
      continue;
    }

    for (let columnIndex = 2; columnIndex <= columnCount; columnIndex++) {
      const amount = row.getCell(columnIndex).value;
      const columnCompany = headerRow.getCell(columnIndex).value;

      // If a specific column company is selected and it doesn't match, skip this column
      if (selectedColumnCompany && columnCompany !== selectedColumnCompany) {
        continue;
      }

      // Check if there is a valid amount and row and column companies are not the same
      if (
        amount !== null &&
        amount !== undefined &&
        String(amount).trim() !== "" &&
        !isNaN(parseFloat(String(amount).trim())) &&
        rowCompany !== columnCompany
      ) {
        const key = `${columnCompany}-${rowCompany}`;
        if (!rows[key]) {
          rows[key] = {
            columnCompany,
            rowCompany,
            amounts: [amount],
          };
        } else {
          rows[key].amounts.push(amount);
        }
      }
    }
  }

  // Remove any existing rows with the same company and amount
  existingRows.forEach((row) => {
    row.remove();
  });

  // Add the new rows to the fragment
  Object.keys(rows).forEach((key) => {
    const row = rows[key];
    const rowElement = document.createElement("tr");
    rowElement.innerHTML = `
      <td>${row.columnCompany}</td>
      <td>${row.rowCompany}</td>
      <td>${row.amounts.join(", ")}</td>
    `;
    fragment.appendChild(rowElement);
  });

  // If no data is displayed, add a message
  if (fragment.childNodes.length === 0) {
    const messageRow = document.createElement("tr");
    messageRow.innerHTML = `
      <td colspan="3">No data available for the selected companies.</td>
    `;
    fragment.appendChild(messageRow);
  }

  // Batch DOM updates
  dataBody.innerHTML = "";
  dataBody.appendChild(fragment);
}

// Event listeners for dropdowns to update the data table dynamically
companyRowSelect.addEventListener("change", function () {
  const rowCompany = companyRowSelect.value;
  const columnCompany = companyColumnSelect.value;
  populateDataTable(rowCompany, columnCompany);
});

companyColumnSelect.addEventListener("change", function () {
  const rowCompany = companyRowSelect.value;
  const columnCompany = companyColumnSelect.value;
  populateDataTable(rowCompany, columnCompany);
});

// Function to read and parse the uploaded Excel file
excelFileInput.addEventListener("change", async function (event) {
  const file = event.target.files[0];
  const reader = new FileReader();

  reader.onload = async function (event) {
    const data = new Uint8Array(event.target.result);
    workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(data);
    worksheet = workbook.worksheets[0];

    // Get companies from the first row (header)
    const headers = worksheet.getRow(1).values.slice(1); // Get values, slice to remove empty first element
    companyList = headers;

    // Populate both select boxes
    populateSelect(companyRowSelect, headers);
    populateSelect(companyColumnSelect, headers);

    // Fill diagonal cells for matching companies
    fillDiagonalCells();
  };

  reader.readAsArrayBuffer(file);
});

// Function to populate dropdowns
function populateSelect(selectElement, options) {
  selectElement.innerHTML =
    '<option value="" disabled selected>Select a company</option>';
  options.forEach((option) => {
    const opt = document.createElement("option");
    opt.value = option;
    opt.textContent = option;
    selectElement.appendChild(opt);
  });
}

// Prevent the same company from being selected in both dropdowns
companyRowSelect.addEventListener("change", disableSameCompany);
companyColumnSelect.addEventListener("change", disableSameCompany);

function disableSameCompany() {
  const rowCompany = companyRowSelect.value;
  const columnCompany = companyColumnSelect.value;

  if (rowCompany === columnCompany) {
    alert("You cannot select the same company in both row and column.");
    companyColumnSelect.value = ""; // Reset the column dropdown if the same is selected
    companyRowSelect.value = ""; // Reset the column dropdown if the same is selected
  }
}

// Function to add a new row to the data table
function addNewRowToTable(rowCompany, columnCompany, amount) {
  const dataBody = document.getElementById("dataBody");
  const newRow = document.createElement("tr");
  newRow.classList.add("new-row");

  const rowCell = document.createElement("td");
  rowCell.textContent = rowCompany;

  const columnCell = document.createElement("td");
  columnCell.textContent = columnCompany;

  const amountCell = document.createElement("td");
  amountCell.textContent = amount;

  newRow.appendChild(rowCell);
  newRow.appendChild(columnCell);
  newRow.appendChild(amountCell);

  // Insert new row at the top
  dataBody.insertBefore(newRow, dataBody.firstChild);
}

// Add new company to both row and column
async function addCompany() {
  const newCompany = newCompanyInput.value.trim();

  if (!newCompany) {
    alert("Please enter a company name.");
    return;
  }

  // Check if the company already exists
  if (companyList.includes(newCompany)) {
    alert("This company already exists.");
    return;
  }

  // Add the new company to the companyList
  companyList.push(newCompany);
  const newIndex = companyList.length;

  // Set the new company name in the first column (row header)
  const rowCell = worksheet.getCell(`A${newIndex}`);
  rowCell.value = newCompany;

  // Convert the index to a column letter
  const colLetter = indexToColumnLetter(newIndex); // Use the new function here
  const headerCell = worksheet.getCell(`${colLetter}1`);

  // Ensure headerCell is initialized
  if (!headerCell.row) {
    worksheet.addRow([newCompany]); // This adds a new row with the new company as the first cell
  } else {
    headerCell.value = newCompany; // Set the new header
  }

  // Initialize new cells in the new column
  for (let i = 2; i <= newIndex; i++) {
    const newRow = worksheet.getRow(i);
    if (!newRow) {
      worksheet.addRow([]); // Add a new row if it doesn't exist
    }
    const newCell = newRow.getCell(colLetter);
    newCell.value = ""; // Set the new cell value to empty
  }

  // Set the diagonal cell to "-"
  const diagonalCell = worksheet.getCell(`${colLetter}${newIndex}`);
  diagonalCell.value = "-";

  // Populate both select boxes again
  populateSelect(companyRowSelect, companyList);
  populateSelect(companyColumnSelect, companyList);

  // Clear the input field
  newCompanyInput.value = "";

  // Fill diagonal cells for matching companies
  fillDiagonalCells();
}

// Function to convert a 1-based index to an Excel column letter
function indexToColumnLetter(index) {
  let letter = "";
  while (index > 0) {
    const modulo = (index - 1) % 26;
    letter = String.fromCharCode(65 + modulo) + letter;
    index = Math.floor((index - modulo) / 26);
  }
  return letter;
}

addCompanyButton.addEventListener("click", addCompany);

submitButton.addEventListener("click", async function () {
  const rowCompany = companyRowSelect.value;
  const columnCompany = companyColumnSelect.value;
  const amount = parseFloat(document.getElementById("amount").value);
  const comment = document.getElementById("comment").value.trim(); // Get the comment from the textarea

  if (!rowCompany || !columnCompany || isNaN(amount)) {
    alert("Please select both companies and enter a valid amount.");
    return;
  }

  if (rowCompany === columnCompany) {
    alert("You cannot map a company to itself.");
    return;
  }

  // Find row and column indices of the selected companies
  const rowIndex = companyList.indexOf(rowCompany) + 1; // Add 1 for Excel 1-based index
  const columnIndex = companyList.indexOf(columnCompany) + 1; // Add 1 for Excel 1-based index

  // Ensure the row exists
  const row = worksheet.getRow(rowIndex);
  if (!row) {
    worksheet.addRow([]); // Add a new row if it doesn't exist
  }

  // Get the current value in the cell
  const cell = row.getCell(indexToColumnLetter(columnIndex));
  const existingValue = cell.value;
  const existingComment = cell.note;

  // Check if the existing value in the cell is the same
  if (existingValue !== null) {
    const confirmOverride = confirm(
      `The current value is ${existingValue}. Do you want to override it with ${amount}?`
    );

    if (!confirmOverride) {
      return; // If the user chooses not to override, exit the function
    }
  }

  // Check if the existing comment in the cell is the same
  if (existingComment !== null && comment !== existingComment) {
    if (existingComment === undefined) {
      // If the existing comment is undefined, don't show the confirm box
      cell.note = comment;
    } else {
      const confirmOverrideComment = confirm(
        `The current comment is "${existingComment}". Do you want to override it with "${comment}"?`
      );

      if (!confirmOverrideComment) {
        return; // If the user chooses not to override, exit the function
      } else {
        cell.note = comment;
      }
    }
  } else {
    // If the comment is empty, don't show the confirm box
    if (comment) {
      cell.note = comment;
    }
  }

  // Set the amount in the correct cell in the Excel sheet
  cell.value = amount;

  // Update the displayed list
  updateDataTable(columnCompany, rowCompany, amount); // Pass only the required data to the update function

  // Clear inputs after submission
  companyRowSelect.value = "";
  companyColumnSelect.value = "";
  document.getElementById("amount").value = "";
  document.getElementById("comment").value = ""; // Clear the comment textarea
});

// Function to update the displayed data table
function updateDataTable(columnCompany, rowCompany, amount) {
  const dataBody = document.getElementById("dataBody");
  const rows = Array.from(dataBody.rows);

  let updated = false;

  rows.forEach((row) => {
    const rowCells = row.children;
    if (
      rowCells.length > 2 &&
      rowCells[1].textContent === rowCompany &&
      rowCells[0].textContent === columnCompany
    ) {
      // Update the amount cell
      rowCells[2].textContent = amount;
      updated = true; // Mark as updated
    }
  });

  // If not updated, add a new row
  if (!updated) {
    const newRow = document.createElement("tr");
    newRow.classList.add("new-row"); // Add the "new-row" class
    newRow.innerHTML = `
      <td>${columnCompany}</td>
      <td>${rowCompany}</td>
      <td>${amount}</td>
    `;
    dataBody.insertBefore(newRow, dataBody.firstChild); // Add the new row to the top
  }

  // Remove the "No data available" message if it exists
  const noDataRow = dataBody.querySelector("tr td[colspan='3']");
  if (noDataRow) {
    noDataRow.parentNode.remove();
  }
}

// Download the updated Excel file when "Download" button is clicked
downloadButton.addEventListener("click", async function () {
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/octet-stream" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "UpdatedCompanyMatrix.xlsx";
  link.click();
});
