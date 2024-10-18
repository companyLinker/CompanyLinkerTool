const excelFileInput = document.getElementById("excelFile");
const companyRowSelect = document.getElementById("companyRow");
const companyColumnSelect = document.getElementById("companyColumn");
const submitButton = document.getElementById("submitButton");
const downloadButton = document.getElementById("downloadButton");
const newCompanyInput = document.getElementById("newCompany");
const addCompanyButton = document.getElementById("addCompanyButton");
const affiliatedRadio = document.getElementById("affiliatedRadio");
const nonAffiliatedRadio = document.getElementById("nonAffiliatedRadio");
const employeeFileInput = document.getElementById("employeeFile");
const employeeSelect = document.getElementById("employeeSelect");
const commentTextarea = document.getElementById("comment");

let workbook,
  worksheet,
  companyList = [],
  rowCompanies = [];

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

// Function to populate the employee select
function populateEmployeeSelect(employees) {
  // Clear existing options
  employeeSelect.innerHTML = '<option value="" disabled selected>Select an employee</option>';
  
  // Filter out empty entries and trim whitespace
  const filteredEmployees = employees
    .filter(employee => employee) // Remove empty entries
    .map(employee => employee.trim()); // Trim whitespace

  // Populate the select with filtered employee names
  filteredEmployees.forEach((employee, index) => {
    const option = document.createElement("option");
    option.value = employee; // Use the employee name as the value
    option.textContent = employee; // Display the employee name
    employeeSelect.appendChild(option);
  });

  employeeSelect.disabled = filteredEmployees.length === 0; // Disable if no employees are available
}

// Function to handle employee file upload
employeeFileInput.addEventListener("change", async function (event) {
  const file = event.target.files[0];
  const reader = new FileReader();

  reader.onload = async function (event) {
    const data = new Uint8Array(event.target.result);
    const workbook = new ExcelJS.Workbook();
    
    // Load the workbook
    await workbook.xlsx.load(data);
    
    // Assuming employee names are in the first sheet
    const worksheet = workbook.worksheets[0];

    // Extract employee names from the first column (you may adjust the column index)
    const employees = [];
    worksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
      // Assuming employee names are in the first column (index 1)
      const employeeName = row.getCell(1).value;
      if (typeof employeeName === 'string' && employeeName.trim() !== '') {
        employees.push(employeeName.trim());
      }
    });

    // Populate the employee select dropdown
    populateEmployeeSelect(employees);
  };

  reader.readAsArrayBuffer(file); // Read the file as an ArrayBuffer
});


// Function to handle employee select change
employeeSelect.addEventListener("change", function () {
  const selectedEmployeeIndex = employeeSelect.selectedIndex;
  if (selectedEmployeeIndex > 0) {
    commentTextarea.disabled = false;
  } else {
    commentTextarea.disabled = true;
  }
});

// Function to format the comment with the selected employee
function formatComment(comment) {
  const selectedEmployee = employeeSelect.value;
  
  // Check if an option is selected
  if (selectedEmployee === "") {
    return comment; // Return the comment without the employee name
  }

  const currentDate = new Date();
  const formattedDate = currentDate.toLocaleString();
  
  return `"${comment}" by @${selectedEmployee} on ${formattedDate}`;
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

    // Get companies from the first column (rows)
    rowCompanies = []; // Reset rowCompanies
    for (let rowIndex = 2; rowIndex <= worksheet.rowCount; rowIndex++) {
      const row = worksheet.getRow(rowIndex);
      const rowCompany = row.getCell(1).value;
      rowCompanies.push(rowCompany);
    }

    // Populate both select boxes
    populateSelect(
      companyRowSelect,
      rowCompanies.filter((company) => companyList.includes(company))
    );
    populateSelect(companyColumnSelect, headers);

    // Fill diagonal cells for matching companies
    fillDiagonalCells();

    // Update the dropdowns based on the affiliation type
    updateCompanyDropdowns();
  };

  reader.readAsArrayBuffer(file);
});

// Function to update dropdown options based on affiliation type
function updateCompanyDropdowns() {
  const isAffiliated = affiliatedRadio.checked;

  // Determine available companies for row and column selections
  let availableRowCompanies, availableColumnCompanies;

  // Find the index of the first blank row that separates affiliated and non-affiliated companies
  let blankRowIndex = rowCompanies.findIndex((company) => !company);

  if (isAffiliated) {
    // Affiliated: Show only companies that exist in both rows and columns, excluding the blank row
    availableRowCompanies = rowCompanies
      .slice(0, blankRowIndex)
      .filter((company) => companyList.includes(company));
    availableColumnCompanies = companyList; // Companies that can be columns
  } else {
    // Non-Affiliated: Show all companies for rows, including both affiliated and non-affiliated
    availableRowCompanies = rowCompanies.filter(
      (company) => company !== null && company !== undefined
    );
    availableColumnCompanies = companyList; // Limit columns to those in the header
  }

  // Populate dropdowns
  populateSelect(companyRowSelect, availableRowCompanies);
  populateSelect(companyColumnSelect, availableColumnCompanies);

  // console.log("Row Companies:", availableRowCompanies);
  // console.log("Column Companies:", availableColumnCompanies);
}
// Event listeners for radio buttons
affiliatedRadio.addEventListener("change", updateCompanyDropdowns);
nonAffiliatedRadio.addEventListener("change", updateCompanyDropdowns);

// Initial population of dropdowns
updateCompanyDropdowns();

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
  if (companyList.includes(newCompany) || rowCompanies.includes(newCompany)) {
    alert("This company already exists.");
    return;
  }

  // Find the index of the second blank row (separator) - Improved Blank Check
  let blankRowIndex = rowCompanies.findIndex((company, index) => {
    return (
      (company === null ||
        company === undefined ||
        company.toString().trim() === "") &&
      index < rowCompanies.length - 1
    );
  });

  if (blankRowIndex === -1) {
    // If no second blank row, treat the end as the separator
    blankRowIndex = rowCompanies.length;
  }

  // console.log("rowCompanies before adding:", rowCompanies);
  // console.log("blankRowIndex:", blankRowIndex);

  if (affiliatedRadio.checked) {
    // Affiliated mode:
    const columnIndex = companyList.length + 1;
    const rowIndex = columnIndex;

    // Add to column header
    worksheet.getCell(`${indexToColumnLetter(columnIndex)}1`).value =
      newCompany;
    companyList.push(newCompany);

    // Insert a new row before the blank row
    worksheet.spliceRows(blankRowIndex + 2, 0, [newCompany]);
    rowCompanies.splice(blankRowIndex, 0, newCompany);
  } else {
    // Non-affiliated mode: Add after the blank row
    rowCompanies.splice(blankRowIndex + 1, 0, newCompany);
    worksheet.addRow([newCompany], blankRowIndex + 2);
  }

  // Update the worksheet.rowCount property if necessary
  worksheet.rowCount = Math.max(worksheet.rowCount, rowCompanies.length + 1);

  // Recalculate indices
  companyList = worksheet.getRow(1).values.slice(1);
  rowCompanies = [];
  for (let i = 2; i <= worksheet.rowCount; i++) {
    rowCompanies.push(worksheet.getRow(i).getCell(1).value);
  }

  // Populate both select boxes again
  updateCompanyDropdowns();

  // Clear the input field
  newCompanyInput.value = "";

  // Fill diagonal cells for matching companies
  fillDiagonalCells();
}

// Function to convert a 1-based index to an Excel column letter
function indexToColumnLetter(index) {
  const letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  let columnLetter = "";

  while (index > 0) {
    const remainder = (index - 1) % 26;
    columnLetter = letters[remainder] + columnLetter;
    index = Math.floor((index - 1) / 26);
  }

  return columnLetter;
}

addCompanyButton.addEventListener("click", addCompany);

submitButton.addEventListener("click", async function () {
  const rowCompany = companyRowSelect.value;
  const columnCompany = companyColumnSelect.value;
  const amount = parseFloat(document.getElementById("amount").value);
  const comment = commentTextarea.value.trim();
  const selectedEmployee = employeeSelect.value;

  if (!rowCompany || !columnCompany || isNaN(amount)) {
    alert("Please select both companies and enter a valid amount.");
    return;
  }

  if (rowCompany === columnCompany) {
    alert("You cannot map a company to itself.");
    return;
  }

  // Find row and column indices of the selected companies
  let rowIndex = rowCompanies.indexOf(rowCompany) + 2; // Add 1 for Excel 1-based index
  const columnIndex = companyList.indexOf(columnCompany) + 1; // Add 1 for Excel 1-based index

  // Ensure the row exists
  const row = worksheet.getRow(rowIndex);
  if (!row) {
    // If the row doesn't exist, add a new row to the sheet
    worksheet.addRow([rowCompany]);
    rowCompanies.push(rowCompany); // Update the rowCompanies array
    rowIndex = rowCompanies.indexOf(rowCompany) + 1; // Update the rowIndex
  }

  // Get the current value in the cell
  const cell = worksheet.getCell(`${indexToColumnLetter(columnIndex)}${rowIndex}`);
  const existingValue = cell.value;
  const existingComment = cell.note;

  // Check if the existing value in the cell is the same
  if (existingValue !== null) {
    const confirmOverride = confirm(`The current value is ${existingValue}. Do you want to override it with ${amount}?`);

    if (!confirmOverride) {
      return; // If the user chooses not to override, exit the function
    }
  }

  // Check if the existing comment in the cell is the same
  // if (comment && selectedEmployee) {
  //   const formattedComment = formatComment(comment, selectedEmployee);
  //   if (existingComment !== null && existingComment !== formattedComment) {
  //     const confirmOverrideComment = confirm(`The current comment is "${existingComment}". Do you want to override it with "${comment}"?`);

  //     if (!confirmOverrideComment) {
  //       return; // If the user chooses not to override, exit the function
  //     } else {
  //       cell.note = formattedComment;
  //     }
  //   } else {
  //     cell.note = formattedComment;
  //   }
  // } else {
  //   // If the comment is empty, remove any existing comment
  //   cell.note = null;
  // }

  // Set the amount in the correct cell in the Excel sheet
  cell.value = amount;

  // Update the displayed list
  updateDataTable(columnCompany, rowCompany, amount); // Pass only the required data to the update function

  // Clear inputs after submission
  companyRowSelect.value = "";
  companyColumnSelect.value = "";
  document.getElementById("amount").value = "";
  commentTextarea.value = ""; // Clear the comment textarea
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
