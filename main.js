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
  }
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
    const newCell = worksheet.getCell(`${colLetter}${i}`);
    // Ensure the row exists
    if (!worksheet.getRow(i).hasValues) {
      worksheet.getRow(i).values = []; // Initialize the row if it doesn't exist
    }
    newCell.value = ""; // Set the new cell value to empty
  }

  // Set the diagonal cell to "-"
  const diagonalCell = worksheet.getCell(`${colLetter}${newIndex + 1}`);
  if (!diagonalCell.row) {
    worksheet.addRow(); // Add a new row if it doesn't exist
  }
  diagonalCell.value = "-";

  // Populate both select boxes again
  populateSelect(companyRowSelect, companyList);
  populateSelect(companyColumnSelect, companyList);

  // Clear the input field
  newCompanyInput.value = "";
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

// Store the entered value in the corresponding cell of the Excel sheet
submitButton.addEventListener("click", async function () {
  const rowCompany = companyRowSelect.value;
  const columnCompany = companyColumnSelect.value;
  const amount = parseFloat(document.getElementById("amount").value);

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

  // Check for valid indices
  if (rowIndex < 1 || columnIndex < 1) {
    alert("Invalid selection, please check the company names.");
    return;
  }

  // Get the current value in the cell
  const cell = worksheet.getCell(
    `${String.fromCharCode(64 + columnIndex)}${rowIndex}`
  );
  const existingValue = cell.value;

  // If the cell already has a value, confirm with the user
  if (existingValue !== null) {
    const confirmOverride = confirm(
      `The current value is ${existingValue}. Do you want to override it with ${amount}?`
    );

    if (!confirmOverride) {
      return; // If the user chooses not to override, exit the function
    }
  }

  // Set the amount in the correct cell in the Excel sheet
  cell.value = amount;

  // Clear inputs after submission
  companyRowSelect.value = "";
  companyColumnSelect.value = "";
  document.getElementById("amount").value = "";
});

// Download the updated Excel file when "Download" button is clicked
downloadButton.addEventListener("click", async function () {
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/octet-stream" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "UpdatedCompanyMatrix.xlsx";
  link.click();
});
