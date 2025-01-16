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
const companyListFileInput = document.getElementById("companyListFile");
const commentTextarea = document.getElementById("comment");

let workbook,
  worksheet,
  companyList = [],
  rowCompanies = [],
  isGoogleSheetData = false;

async function fillDiagonalCells(sheetData, spreadsheetId) {
  if (isGoogleSheetData) {
    try {
      // Fetch the spreadsheet metadata to get the sheets
      const spreadsheetResponse = await gapi.client.sheets.spreadsheets.get({
        spreadsheetId: spreadsheetId,
      });

      const sheets = spreadsheetResponse.result.sheets;

      // Get the currently selected sheet name from the dropdown
      const sheetSelect = document.getElementById("sheetSelect");
      const sheetName =
        sheetSelect.value || (sheets[0] && sheets[0].properties.title);

      if (!sheetName) {
        console.error("No sheet name available");
        return;
      }

      // Find the sheet with the selected name
      const selectedSheet = sheets.find(
        (sheet) => sheet.properties.title === sheetName
      );

      if (!selectedSheet) {
        console.error("Selected sheet not found");
        return;
      }

      // Fetch the sheet data to identify "Total" column and rows
      const response = await gapi.client.sheets.spreadsheets.values.get({
        spreadsheetId: spreadsheetId,
        range: sheetName,
      });

      const range = response.result;
      const headers = range.values[0];
      const totalColumnIndex = headers.indexOf("Total");

      const sheetId = selectedSheet.properties.sheetId;
      const requests = []; // Array to hold the requests for batchUpdate

      // Filter out "Total" and special rows from sheetData
      const filteredCompanies = sheetData.filter(
        (company) =>
          company !== "Total" &&
          company !== "Total Affiliated" &&
          company !== "Total Non-Affiliated" &&
          company !== "Grand Total"
      );

      // Loop through the diagonal elements and color specific indices
      for (let index = 0; index < filteredCompanies.length; index++) {
        const company = filteredCompanies[index];

        // Find the actual column index, accounting for skipped columns
        const actualColumnIndex = sheetData.indexOf(company) + 1;

        const request = {
          repeatCell: {
            range: {
              sheetId: sheetId, // Use the selected sheet's ID
              startRowIndex: actualColumnIndex,
              endRowIndex: actualColumnIndex + 1,
              startColumnIndex: actualColumnIndex,
              endColumnIndex: actualColumnIndex + 1,
            },
            cell: {
              userEnteredFormat: {
                backgroundColor: {
                  red: 1, // Full red
                  green: 0,
                  blue: 0,
                },
              },
            },
            fields: "userEnteredFormat(backgroundColor)",
          },
        };
        requests.push(request);
      }

      // Make the batchUpdate request to apply the formatting
      if (requests.length > 0) {
        await gapi.client.sheets.spreadsheets.batchUpdate({
          spreadsheetId: spreadsheetId,
          resource: {
            requests: requests,
          },
        });
        console.log(
          "Diagonal cells filled with red color, skipping Total and special rows."
        );
      }
    } catch (error) {
      console.error("Error applying diagonal cell formatting:", error);
    }
  } else {
    // Existing Excel logic remains the same
    const filteredCompanies = companyList.filter(
      (company) =>
        company !== "Total" &&
        company !== "Total Affiliated" &&
        company !== "Total Non-Affiliated" &&
        company !== "Grand Total"
    );

    filteredCompanies.forEach((company, index) => {
      const actualColumnIndex = companyList.indexOf(company);
      const cell = worksheet.getCell(
        `${indexToColumnLetter(actualColumnIndex + 1)}${actualColumnIndex + 1}`
      );
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFF0000" }, // Red color
      };
    });
  }
}

// Function to populate the employee select
function populateEmployeeSelect(employees) {
  // Clear existing options
  employeeSelect.innerHTML =
    '<option value="" disabled selected>Select an employee</option>';

  // Filter out empty entries and trim whitespace
  const filteredEmployees = employees
    .filter((employee) => employee) // Remove empty entries
    .map((employee) => employee.trim()); // Trim whitespace

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
    worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
      // Assuming employee names are in the first column (index 1)
      const employeeName = row.getCell(1).value;
      if (typeof employeeName === "string" && employeeName.trim() !== "") {
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

async function fetchUserEmail() {
  if (!gapi.client.people) {
    console.error("People API not loaded or initialized.");
    return null;
  }

  const response = await gapi.client.people.people.get({
    resourceName: "people/me",
    personFields: "emailAddresses",
  });

  return response.result.emailAddresses[0].value;
}

// Function to format the comment with the selected employee
function formatComment(comment, selectedEmployee) {
  if (isGoogleSheetData) {
    const currentDate = new Date();
    const formattedDate = currentDate.toLocaleString();

    return `"${comment}" by ${userEmail} on ${formattedDate}`;
  } else {
    const selectedEmployee = employeeSelect.value;

    // Check if an option is selected
    if (selectedEmployee === "") {
      return comment; // Return the comment without the employee name
    }

    const currentDate = new Date();
    const formattedDate = currentDate.toLocaleString();

    return `"${comment}" by @${selectedEmployee} on ${formattedDate}`;
  }
}

// Function to populate the data table
async function populateDataTable(selectedRowCompany, selectedColumnCompany) {
  const dataBody = document.getElementById("dataBody");
  const paginationContainer = document.getElementById("pagination");
  const itemsPerPage = 50; // Configurable number of items per page
  let currentPage = 1;
  const rows = {};

  // Create pagination container if it doesn't exist
  if (!paginationContainer) {
    const paginationDiv = document.createElement("div");
    paginationDiv.id = "paginationContainer";
    paginationDiv.className = "pagination-container";
    dataBody.parentNode.insertBefore(paginationDiv, dataBody.nextSibling);
  }

  // Performance optimization: Use a generator function for data retrieval
  async function* dataGenerator() {
    if (isGoogleSheetData) {
      const sheetUrl = document.getElementById("googleSheetUrl").value;
      const sheetIdMatch = sheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);

      if (!sheetIdMatch) {
        return;
      }

      const spreadsheetId = sheetIdMatch[1];

      try {
        const spreadsheetResponse = await gapi.client.sheets.spreadsheets.get({
          spreadsheetId: spreadsheetId,
        });

        const sheets = spreadsheetResponse.result.sheets;
        if (!sheets || sheets.length === 0) {
          return;
        }

        const sheetName = sheetSelect.value || sheets[0].properties.title;
        const response = await gapi.client.sheets.spreadsheets.values.get({
          spreadsheetId: spreadsheetId,
          range: sheetName,
        });

        const range = response.result;
        if (!range || !range.values || range.values.length === 0) {
          return;
        }

        const firstSheetData = range.values;

        const companyList = firstSheetData[0]
          .slice(1)
          .filter(
            (company) =>
              company &&
              ![
                "Total Affiliated",
                "Total Non-Affiliated",
                "Total",
                "Grand Total",
              ].includes(company) &&
              !String(company).includes("Total")
          );

        const rowCompanies = firstSheetData
          .slice(1)
          .map((row) => row[0])
          .filter(
            (company) =>
              company &&
              ![
                "Total Affiliated",
                "Total Non-Affiliated",
                "Total",
                "Grand Total",
              ].includes(company) &&
              !String(company).includes("Total")
          );

        for (let rowIndex = 1; rowIndex < firstSheetData.length; rowIndex++) {
          const row = firstSheetData[rowIndex];
          const rowCompany = row[0];

          // Skip Total rows and apply row company filter
          if (
            !rowCompany ||
            [
              "Total Affiliated",
              "Total Non-Affiliated",
              "Total",
              "Grand Total",
            ].includes(rowCompany) ||
            String(rowCompany).includes("Total") ||
            (selectedRowCompany && rowCompany !== selectedRowCompany)
          ) {
            continue;
          }

          for (let columnIndex = 1; columnIndex < row.length; columnIndex++) {
            const columnCompany = firstSheetData[0][columnIndex];

            // Skip Total columns and apply column company filter
            if (
              !columnCompany ||
              [
                "Total Affiliated",
                "Total Non-Affiliated",
                "Total",
                "Grand Total",
              ].includes(columnCompany) ||
              String(columnCompany).includes("Total") ||
              (selectedColumnCompany && columnCompany !== selectedColumnCompany)
            ) {
              continue;
            }

            let amount = row[columnIndex];

            // Normalize amount
            if (typeof amount === "string") {
              amount = amount.replace(/[$,]/g, "");
              if (amount.startsWith("(") && amount.endsWith(")")) {
                amount = `(${amount.slice(1, -1)})`;
              }
            }

            // Validate amount
            if (
              amount !== null &&
              amount !== undefined &&
              String(amount).trim() !== "" &&
              rowCompany !== columnCompany
            ) {
              const key = `${columnCompany}-${rowCompany}`;
              yield { columnCompany, rowCompany, amount, key };
            }
          }
        }
      } catch (err) {
        console.error("Error fetching data:", err);
      }
    } else {
      // Excel JS library logic
      const headerRow = worksheet.getRow(1);
      const columnCount = headerRow.values.length;

      for (let rowIndex = 2; rowIndex <= worksheet.rowCount; rowIndex++) {
        const row = worksheet.getRow(rowIndex);
        const rowCompany = row.getCell(1).value;

        // Skip rows based on filters
        if (selectedRowCompany && rowCompany !== selectedRowCompany) {
          continue;
        }

        for (let columnIndex = 2; columnIndex <= columnCount; columnIndex++) {
          let amount = row.getCell(columnIndex).value;
          const columnCompany = headerRow.getCell(columnIndex).value;

          // Skip columns based on filters
          if (
            selectedColumnCompany &&
            columnCompany !== selectedColumnCompany
          ) {
            continue;
          }

          // Normalize amount
          if (typeof amount === "string") {
            amount = amount.replace(/[$,]/g, "");
            if (amount.startsWith("(") && amount.endsWith(")")) {
              amount = `(${amount.slice(1, -1)})`;
            }
          }

          // Validate amount
          if (
            amount !== null &&
            amount !== undefined &&
            String(amount).trim() !== "" &&
            !isNaN(parseFloat(String(amount).trim())) &&
            rowCompany !== columnCompany
          ) {
            const key = `${columnCompany}-${rowCompany}`;
            yield { columnCompany, rowCompany, amount, key };
          }
        }
      }
    }
  }

  // Collect and aggregate data
  const aggregatedRows = {};
  for await (const item of dataGenerator()) {
    if (!aggregatedRows[item.key]) {
      aggregatedRows[item.key] = {
        columnCompany: item.columnCompany,
        rowCompany: item.rowCompany,
        amounts: [item.amount],
      };
    } else {
      aggregatedRows[item.key].amounts.push(item.amount);
    }
  }

  // Pagination function
  function renderPage(page) {
    // Store aggregatedRows in a global or accessible variable
    window.currentAggregatedRows = aggregatedRows;

    dataBody.innerHTML = ""; // Clear existing rows
    const keys = Object.keys(aggregatedRows);
    const totalItems = keys.length;
    const totalPages = Math.ceil(totalItems / itemsPerPage);

    const startIndex = (page - 1) * itemsPerPage;
    const endIndex = startIndex + itemsPerPage;

    // Render current page items
    const pageKeys = keys.slice(startIndex, endIndex);
    pageKeys.forEach((key) => {
      const row = aggregatedRows[key];
      const totalAmount = row.amounts.join(", ");

      const rowElement = document.createElement("tr");
      rowElement.innerHTML = `
      <td>${row.columnCompany}</td>
      <td>${row.rowCompany}</td>
      <td>${totalAmount}</td>
    `;
      dataBody.appendChild(rowElement);
    });

    // If no data, show message
    if (pageKeys.length === 0) {
      const messageRow = document.createElement("tr");
      messageRow.innerHTML = `
        <td colspan="3" class="text-center">No data available</td>
      `;
      dataBody.appendChild(messageRow);
    }

    // Update pagination controls
    renderPaginationControls(page, totalPages);
  }

  // Pagination controls rendering
  function renderPaginationControls(currentPage, totalPages) {
    const paginationContainer = document.getElementById("pagination");
    paginationContainer.innerHTML = "";

    // Helper function to create page button
    function createPageButton(pageNum) {
      const pageButton = document.createElement("button");
      pageButton.textContent = pageNum;
      pageButton.classList.add("btn", "btn-outline-primary", "mx-1");
      pageButton.disabled = pageNum === currentPage;
      pageButton.addEventListener("click", () => renderPage(pageNum));
      return pageButton;
    }

    // Helper function to create ellipsis
    function createEllipsis() {
      const ellipsis = document.createElement("span");
      ellipsis.textContent = "...";
      ellipsis.classList.add("mx-1");
      return ellipsis;
    }

    // Previous button
    if (currentPage > 1) {
      const prevButton = document.createElement("button");
      prevButton.textContent = "Previous";
      prevButton.classList.add("btn", "btn-outline-secondary", "mx-1");
      prevButton.addEventListener("click", () => renderPage(currentPage - 1));
      paginationContainer.appendChild(prevButton);
    }

    // Pagination logic with smart ellipsis
    function generatePaginationButtons() {
      // Always show first page
      if (currentPage > 3) {
        paginationContainer.appendChild(createPageButton(1));

        // Add first ellipsis if there's a gap
        if (currentPage > 4) {
          paginationContainer.appendChild(createEllipsis());
        }
      }

      // Calculate range of page buttons to show
      let startPage = Math.max(1, currentPage - 1);
      let endPage = Math.min(totalPages, currentPage + 1);

      // Adjust range to always show 3 buttons around current page
      if (currentPage === 1) {
        endPage = Math.min(3, totalPages);
      } else if (currentPage === totalPages) {
        startPage = Math.max(1, totalPages - 2);
      }

      // Add page buttons
      for (let i = startPage; i <= endPage; i++) {
        paginationContainer.appendChild(createPageButton(i));
      }

      // Add last ellipsis and last page
      if (currentPage < totalPages - 2) {
        if (currentPage < totalPages - 3) {
          paginationContainer.appendChild(createEllipsis());
        }
        paginationContainer.appendChild(createPageButton(totalPages));
      }
    }

    // Generate pagination buttons
    generatePaginationButtons();

    // Next button
    if (currentPage < totalPages) {
      const nextButton = document.createElement("button");
      nextButton.textContent = "Next";
      nextButton.classList.add("btn", "btn-outline-secondary", "mx-1");
      nextButton.addEventListener("click", () => renderPage(currentPage + 1));
      paginationContainer.appendChild(nextButton);
    }
  }

  // Initial render
  renderPage(currentPage);
}

document.getElementById("downloadLog").addEventListener("click", function () {
  // Use the globally stored aggregatedRows
  const aggregatedRows = window.currentAggregatedRows || {};

  // Create a CSV content string
  let csvContent = [];

  // Add table headers
  const headers = [
    "Company from Quickbooks (Column)",
    "Company from COA (Row)",
    "Amount",
  ];
  csvContent.push(headers);

  // Add all rows from aggregatedRows
  Object.values(aggregatedRows).forEach((row) => {
    const rowData = [
      row.columnCompany,
      row.rowCompany,
      row.amounts.join(","),
    ].map((value) => sanitizeCSVValue(value));
    csvContent.push(rowData);
  });

  // Convert to CSV string
  const csvString = csvContent.map((row) => row.join(",")).join("\n");

  // Create a Blob with UTF-8 encoding
  const blob = new Blob(["\uFEFF" + csvString], {
    type: "text/csv;charset=utf-8;",
  });

  // Create a download link
  const link = document.createElement("a");
  const url = URL.createObjectURL(blob);
  link.setAttribute("href", url);

  // Generate filename with current date and time
  const currentDate = new Date();
  const hours = String(currentDate.getHours()).padStart(2, "0");
  const minutes = String(currentDate.getMinutes()).padStart(2, "0");
  const seconds = String(currentDate.getSeconds()).padStart(2, "0");
  const formattedDateTime = `${currentDate.getFullYear()}-${String(
    currentDate.getMonth() + 1
  ).padStart(2, "0")}-${String(currentDate.getDate()).padStart(
    2,
    "0"
  )}_${hours}-${minutes}-${seconds}`;
  const formattedDateTimeWithHyphens = formattedDateTime.replace(/:/g, "-");
  link.setAttribute(
    "download",
    `transaction_log_${formattedDateTimeWithHyphens}.csv`
  );

  // Trigger download
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  // Clean up
  URL.revokeObjectURL(url);
});

// Helper function to sanitize CSV values
function sanitizeCSVValue(value) {
  if (value == null) return '""';

  // Convert to string and trim
  let sanitizedValue = String(value).trim();

  // Escape double quotes by doubling them
  sanitizedValue = sanitizedValue.replace(/"/g, '""');

  // Normalize Unicode characters
  sanitizedValue = sanitizedValue.normalize("NFC");

  // If the value contains a comma, newline, or double quote, wrap in quotes
  if (/[",\n\r]/.test(sanitizedValue)) {
    sanitizedValue = `"${sanitizedValue}"`;
  }

  return sanitizedValue;
}

// Event listeners for dropdowns to update the data table dynamically
companyRowSelect.addEventListener("change", function () {
  const rowCompany = companyRowSelect.value;
  const columnCompany = companyColumnSelect.value;
  populateDataTable(rowCompany, columnCompany);
  toggleInputs(rowCompany, columnCompany);
});

companyColumnSelect.addEventListener("change", function () {
  const rowCompany = companyRowSelect.value;
  const columnCompany = companyColumnSelect.value;
  populateDataTable(rowCompany, columnCompany);
  toggleInputs(rowCompany, columnCompany);
});

function toggleInputs(rowCompany, columnCompany) {
  const amountInput = document.getElementById("amount");

  if (
    rowCompany !== "Select a company" &&
    columnCompany !== "Select a company" &&
    rowCompany !== columnCompany &&
    rowCompany !== "" &&
    columnCompany !== ""
  ) {
    commentTextarea.disabled = false;
    amountInput.disabled = false;
  } else {
    commentTextarea.disabled = true;
    amountInput.disabled = true;
  }
}

var colOptions = { searchable: true };
let colSelect = NiceSelect.bind(companyColumnSelect, colOptions);

var rowOptions = { searchable: true };
let rowSelect = NiceSelect.bind(companyRowSelect, rowOptions);

// Function to read and parse the uploaded Excel file
excelFileInput.addEventListener("change", async function (event) {
  isGoogleSheetData = false;
  downloadButton.style.display = "block";
  excelFile.setAttribute("disabled", false);
  document.getElementById("employeeFileWrap").style.display = "block";
  document.getElementById("employeeSelectWrap").style.display = "block";
  // Create loading overlay
  const loadingOverlay = document.createElement("div");
  loadingOverlay.innerHTML = `
    <div style="
      position: fixed;
      top: 0;
      left: 0;
      width: 100%;
      height: 100%;
      background: rgba(0,0,0,0.5);
      display: flex;
      justify-content: center;
      align-items: center;
      z-index: 9999;
    ">
      <div style="
        background: white;
        padding: 20px;
        border-radius: 10px;
        text-align: center;
        min-width: 300px;
      ">
        <h3>Loading Excel File</h3>
        <div id="progressContainer" style="width: 100%; background: #e0e0e0; border-radius: 5px; margin-top: 10px;">
          <div id="progressBar" style="width: 0%; height: 20px; background: #4CAF50; border-radius: 5px; transition: width 0.1s;"></div>
        </div>
        <p id="progressText">Preparing to read file...</p>
      </div>
    </div>
  `;
  document.body.appendChild(loadingOverlay);

  const progressBar = loadingOverlay.querySelector("#progressBar");
  const progressText = loadingOverlay.querySelector("#progressText");

  const file = event.target.files[0];
  const reader = new FileReader();

  // Function to update progress
  const updateProgress = (message, percentage) => {
    progressText.textContent = message;
    progressBar.style.width = `${percentage}%`;
  };

  reader.onload = async function (event) {
    try {
      // Update progress - Reading file
      updateProgress("Reading file...", 20);

      const data = new Uint8Array(event.target.result);

      // Update progress - Loading workbook
      updateProgress("Loading workbook...", 40);

      workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(data);
      worksheet = workbook.worksheets[0];

      // Update progress - Processing headers
      updateProgress("Processing headers...", 60);

      // Get companies from the first row (header)
      const headers = worksheet.getRow(1).values.slice(1);
      companyList = headers;

      // Update progress - Processing row companies
      updateProgress("Processing row companies...", 70);

      // Get companies from the first column (rows)
      rowCompanies = []; // Reset rowCompanies
      for (let rowIndex = 2; rowIndex <= worksheet.rowCount; rowIndex++) {
        const row = worksheet.getRow(rowIndex);
        const rowCompany = row.getCell(1).value;
        rowCompanies.push(rowCompany);
      }

      // Update progress - Populating select boxes
      updateProgress("Populating select boxes...", 80);

      // Populate both select boxes
      populateSelect(
        companyRowSelect,
        rowCompanies.filter((company) => companyList.includes(company))
      );
      populateSelect(companyColumnSelect, headers);

      colSelect.update();
      rowSelect.update();

      // Update progress - Updating dropdowns
      updateProgress("Updating dropdowns...", 90);

      // Update the dropdowns based on the affiliation type
      updateCompanyDropdowns();

      // Update progress - Finalizing
      updateProgress("Almost done...", 95);

      // Fill diagonal cells for matching companies
      await fillDiagonalCells(companyList);

      // Update progress - Complete
      updateProgress("Processing complete", 100);

      // Wait a moment to show complete state
      await new Promise((resolve) => setTimeout(resolve, 500));

      // Remove loading overlay
      document.body.removeChild(loadingOverlay);

      // Call the function to find and highlight nullifiable transactions
      await findNullifiableTransactionsExcel(worksheet);
    } catch (error) {
      // Error handling
      updateProgress("Error occurred", 100);
      progressBar.style.backgroundColor = "red";

      console.error("Error processing Excel file:", error);

      // Show error message
      progressText.textContent = `Error: ${error.message}`;

      // Remove loading overlay after a short delay
      setTimeout(() => {
        if (document.body.contains(loadingOverlay)) {
          document.body.removeChild(loadingOverlay);
        }
      }, 2000);
    }
  };

  // Handle file reading errors
  reader.onerror = function (error) {
    updateProgress("File reading error", 100);
    progressBar.style.backgroundColor = "red";
    progressText.textContent = `Error reading file: ${error}`;

    // Remove loading overlay after a short delay
    setTimeout(() => {
      if (document.body.contains(loadingOverlay)) {
        document.body.removeChild(loadingOverlay);
      }
    }, 2000);
  };

  // Start reading the file
  updateProgress("Initializing file read...", 10);
  reader.readAsArrayBuffer(file);
});

let nullifiablePairs = [];
let headers = [];
let sheets = [];
let globalNullifiablePairs = [];

async function findNullifiableTransactionsExcel(worksheet) {
  // Function to parse amount similar to Google Sheets version
  function parseAmount(amount) {
    if (amount === null || amount === undefined || amount === "") {
      return null;
    }

    if (typeof amount === "number" && !isNaN(amount)) {
      return amount;
    }

    if (typeof amount === "string") {
      amount = amount.trim();

      // Check if the amount is in parentheses
      if (amount.startsWith("(") && amount.endsWith(")")) {
        // Remove parentheses and parse as negative
        amount = "-" + amount.slice(1, -1).replace(/[$,]/g, "");
      } else {
        amount = amount.replace(/[$,]/g, ""); // Remove currency symbols
      }

      const parsedNum = parseFloat(amount);
      return !isNaN(parsedNum) ? parsedNum : null;
    }

    return null;
  }

  // Function to extract transactions from Excel worksheet
  function extractTransactions(worksheet) {
    const transactionMap = new Map();
    const headerRow = worksheet?.getRow(1);
    const headerCompanies = headerRow?.values.slice(1); // Skip first empty cell

    // Iterate through rows starting from row 2
    for (let rowIndex = 2; rowIndex <= worksheet?.rowCount; rowIndex++) {
      const row = worksheet?.getRow(rowIndex);
      const fromCompany = row.getCell(1).value;

      // Iterate through columns starting from column 2
      for (
        let colIndex = 2;
        colIndex <= headerCompanies.length + 1;
        colIndex++
      ) {
        const toCompany = headerRow.getCell(colIndex).value;
        const amountCell = row.getCell(colIndex);

        // Extract raw value before parsing
        let rawValue = amountCell.value;

        // Try multiple methods to get the value
        if (amountCell.text) {
          rawValue = amountCell.text;
        } else if (amountCell.result) {
          rawValue = amountCell.result;
        }

        const amount = parseAmount(rawValue);

        // Skip null, undefined, or zero amounts and self-transactions
        if (amount === null || amount === 0 || fromCompany === toCompany)
          continue;

        if (!transactionMap.has(fromCompany)) {
          transactionMap.set(fromCompany, {});
        }
        transactionMap.get(fromCompany)[toCompany] = amount;
      }
    }

    return transactionMap;
  }

  // Function to find nullifiable pairs
  function findNullifiablePairs(transactionMap) {
    const nullifiablePairsLocal = [];
    const nonNullifiablePairs = [];
    const processedPairs = new Set(); // To avoid duplicate processing

    // Iterate through all companies
    for (const [companyA, transactionsA] of transactionMap.entries()) {
      for (const [companyB, amountAtoB] of Object.entries(transactionsA)) {
        // Skip if amount is null or zero
        if (amountAtoB === null || amountAtoB === 0) continue;

        // Create a unique key for the pair to avoid duplicate processing
        const pairKey = `${companyA}-${companyB}`;
        const reversePairKey = `${companyB}-${companyA}`;

        // Skip if this pair has been processed
        if (processedPairs.has(pairKey) || processedPairs.has(reversePairKey))
          continue;

        // Check if reverse transaction exists
        if (
          transactionMap.has(companyB) &&
          transactionMap.get(companyB)[companyA] !== undefined
        ) {
          const amountBtoA = transactionMap.get(companyB)[companyA];

          // Mark these pairs as processed to avoid duplicate checking
          processedPairs.add(pairKey);
          processedPairs.add(reversePairKey);

          // Check if amounts are equal in magnitude but opposite in sign
          if (
            amountBtoA !== null &&
            Math.abs(amountAtoB) === Math.abs(amountBtoA) &&
            Math.sign(amountAtoB) !== Math.sign(amountBtoA)
          ) {
            nullifiablePairsLocal.push({
              companyA,
              companyB,
              amountAtoB,
              amountBtoA,
            });
          } else {
            // If not nullifiable, add to non-nullifiable pairs
            nonNullifiablePairs.push({
              companyA,
              companyB,
              amountAtoB,
              amountBtoA,
            });
          }
        } else {
          // No reverse transaction found, add to non-nullifiable pairs
          nonNullifiablePairs.push({
            companyA,
            companyB,
            amountAtoB,
            amountBtoA: null,
          });
        }
      }
    }
    // Add this section at the end of the highlighting logic, before the console.log statements
    // Color diagonal cells red
    companyList.forEach((company, index) => {
      const cell = worksheet?.getCell(
        `${indexToColumnLetter(index + 1)}${index + 1}`
      );

      // Preserve original border style
      const originalBorder = cell?.border || {
        top: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };

      // Apply red fill while preserving original border
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFF0000" }, // Red color
      };

      // Explicitly set style with original border
      cell.border = originalBorder;
      cell.style = {
        ...cell.style,
        fill: {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFF0000" },
        },
        border: originalBorder,
      };
    });

    globalNullifiablePairs = nullifiablePairsLocal;

    return {
      nullifiablePairs: nullifiablePairsLocal,
      nonNullifiablePairs,
    };
  }

  // Extract transactions from the worksheet
  const transactionMap = extractTransactions(worksheet);

  // Find nullifiable and non-nullifiable pairs
  const { nullifiablePairs, nonNullifiablePairs } =
    findNullifiablePairs(transactionMap);

  // Highlight nullifiable and non-nullifiable pairs
  const headerRow = worksheet.getRow(1);
  const headerCompanies = headerRow.values.slice(1);

  // Function to get the original border style of a cell
  function getCellBorderStyle(cell) {
    const originalBorder = cell.border || {};

    // Default border if no border exists
    const defaultBorder = {
      top: { style: "thin", color: { argb: "FF000000" } },
      left: { style: "thin", color: { argb: "FF000000" } },
      bottom: { style: "thin", color: { argb: "FF000000" } },
      right: { style: "thin", color: { argb: "FF000000" } },
    };

    return {
      top: originalBorder.top || defaultBorder.top,
      left: originalBorder.left || defaultBorder.left,
      bottom: originalBorder.bottom || defaultBorder.bottom,
      right: originalBorder.right || defaultBorder.right,
    };
  }

  // Highlight nullifiable pairs in yellow
  nullifiablePairs.forEach(({ companyA, companyB }) => {
    const rowIndexA = headerCompanies.indexOf(companyA) + 1;
    const colIndexA = headerCompanies.indexOf(companyB) + 1;
    const rowIndexB = headerCompanies.indexOf(companyB) + 1;
    const colIndexB = headerCompanies.indexOf(companyA) + 1;

    // Color A to B cell
    const cellAtoB = worksheet.getCell(
      `${indexToColumnLetter(colIndexA)}${rowIndexA}`
    );

    // Preserve original border style
    const originalBorderAtoB = getCellBorderStyle(cellAtoB);

    // Reset and apply fill
    cellAtoB.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFFF00" }, // Yellow
    };

    // Explicitly set style with original border
    cellAtoB.border = originalBorderAtoB;
    cellAtoB.style = {
      ...cellAtoB.style,
      fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFFF00" },
      },
      border: originalBorderAtoB,
    };

    // Color B to A cell
    const cellBtoA = worksheet.getCell(
      `${indexToColumnLetter(colIndexB)}${rowIndexB}`
    );

    // Preserve original border style
    const originalBorderBtoA = getCellBorderStyle(cellBtoA);

    // Reset and apply fill
    cellBtoA.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFFF00" }, // Yellow
    };

    // Explicitly set style with original border
    cellBtoA.border = originalBorderBtoA;
    cellBtoA.style = {
      ...cellBtoA.style,
      fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFFF00" },
      },
      border: originalBorderBtoA,
    };
  });

  // Highlight non-nullifiable pairs with a light color
  nonNullifiablePairs.forEach(({ companyA, companyB }) => {
    const rowIndexA = headerCompanies.indexOf(companyA) + 1;
    const colIndexA = headerCompanies.indexOf(companyB) + 1;

    // Color the cell with a light gray background
    const cellAtoB = worksheet.getCell(
      `${indexToColumnLetter(colIndexA)}${rowIndexA}`
    );

    // Check if the cell is blank or contains only spaces
    if (
      cellAtoB.value === null ||
      (typeof cellAtoB.value === "string" && cellAtoB.value.trim() === "")
    ) {
      return; // Skip blank cells
    }

    // Preserve original border style
    const originalBorderAtoB = getCellBorderStyle(cellAtoB);

    // Reset and apply fill
    cellAtoB.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFFF" }, // Very light gray
    };

    // Explicitly set style with original border
    cellAtoB.border = originalBorderAtoB;
    cellAtoB.style = {
      ...cellAtoB.style,
      fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFFF" },
      },
      border: originalBorderAtoB,
    };

    // If there's a reverse transaction, color that cell too
    const rowIndexB = headerCompanies.indexOf(companyB) + 1;
    const colIndexB = headerCompanies.indexOf(companyA) + 1;
    const cellBtoA = worksheet.getCell(
      `${indexToColumnLetter(colIndexB)}${rowIndexB}`
    );

    // Check if the cell is blank or contains only spaces
    if (
      cellBtoA.value === null ||
      (typeof cellBtoA.value === "string" && cellBtoA.value.trim() === "")
    ) {
      return; // Skip blank cells
    }

    // Preserve original border style
    const originalBorderBtoA = getCellBorderStyle(cellBtoA);

    // Reset and apply fill
    cellBtoA.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFFF" }, // Very light gray
    };

    // Explicitly set style with original border
    cellBtoA.border = originalBorderBtoA;
    cellBtoA.style = {
      ...cellBtoA.style,
      fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFFF" },
      },
      border: originalBorderBtoA,
    };
  });

  // Optional: Log the results
  console.log("Nullifiable Pairs:", nullifiablePairs);
  console.log("Non-Nullifiable Pairs:", nonNullifiablePairs);

  // Return the pairs if needed
  return { nullifiablePairs, nonNullifiablePairs };
}

// Function to find and highlight nullifiable transactions
async function findNullifiableTransactions() {
  let transactionMap;
  let localHeaders;

  function parseAmount(amount) {
    if (amount === null || amount === undefined || amount === "") {
      return null;
    }

    if (typeof amount === "number" && !isNaN(amount)) {
      return amount;
    }

    if (typeof amount === "string") {
      amount = amount.trim();

      // Check if the amount is in parentheses
      if (amount.startsWith("(") && amount.endsWith(")")) {
        // Remove parentheses and parse as negative
        amount = "-" + amount.slice(1, -1).replace(/[$,]/g, "");
      } else {
        amount = amount.replace(/[$,]/g, ""); // Remove currency symbols
      }

      const parsedNum = parseFloat(amount);
      return !isNaN(parsedNum) ? parsedNum : null;
    }

    return null;
  }

  function findNullifiablePairs(transactionMap) {
    if (!transactionMap || transactionMap.size === 0) {
      return {
        nullifiablePairs: [],
        nonNullifiablePairs: [],
      };
    }
    const nullifiablePairsLocal = [];
    const nonNullifiablePairs = [];
    const processedPairs = new Set(); // To avoid duplicate processing

    // Iterate through all companies
    for (const [companyA, transactionsA] of transactionMap.entries()) {
      for (const [companyB, amountAtoB] of Object.entries(transactionsA)) {
        // Create unique pair keys
        const pairKey = `${companyA}-${companyB}`;
        const reversePairKey = `${companyB}-${companyA}`;

        // Skip if this pair has been processed
        if (processedPairs.has(pairKey) || processedPairs.has(reversePairKey)) {
          continue;
        }

        // Check if reverse transaction exists
        if (
          transactionMap.has(companyB) &&
          transactionMap.get(companyB)[companyA] !== undefined
        ) {
          const amountBtoA = transactionMap.get(companyB)[companyA];

          // Mark these pairs as processed
          processedPairs.add(pairKey);
          processedPairs.add(reversePairKey);

          // If either amount is 0 or null, it's non-nullifiable
          if (
            amountAtoB === 0 ||
            amountBtoA === 0 ||
            amountAtoB === null ||
            amountBtoA === null
          ) {
            nonNullifiablePairs.push({
              companyA,
              companyB,
              amountAtoB,
              amountBtoA,
            });
            continue;
          }

          // Check if amounts are equal in magnitude but opposite in sign
          if (
            Math.abs(amountAtoB) === Math.abs(amountBtoA) &&
            Math.sign(amountAtoB) !== Math.sign(amountBtoA)
          ) {
            nullifiablePairsLocal.push({
              companyA,
              companyB,
              amountAtoB,
              amountBtoA,
            });
          } else {
            // If not nullifiable, add to non-nullifiable pairs
            nonNullifiablePairs.push({
              companyA,
              companyB,
              amountAtoB,
              amountBtoA,
            });
          }
        }
      }
    }

    return {
      nullifiablePairs: nullifiablePairsLocal,
      nonNullifiablePairs,
    };
  }

  function extractTransactions(values, headers) {
    const transactionMap = new Map();

    for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
      const fromCompany = values[rowIndex][0];

      for (let colIndex = 1; colIndex < values[rowIndex].length; colIndex++) {
        const toCompany = headers[colIndex - 1];
        const amount = parseAmount(values[rowIndex][colIndex]);

        // Skip null, undefined, or zero amounts and self-transactions
        if (amount === null || fromCompany === toCompany) {
          continue;
        }

        if (!transactionMap.has(fromCompany)) {
          transactionMap.set(fromCompany, {});
        }
        transactionMap.get(fromCompany)[toCompany] = amount;
      }
    }

    return transactionMap;
  }

  // Google Sheets processing
  if (isGoogleSheetData) {
    const sheetUrl = document.getElementById("googleSheetUrl").value;
    const sheetIdMatch = sheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);

    if (!sheetIdMatch) {
      console.error("Invalid Google Sheet URL");
      return;
    }

    const spreadsheetId = sheetIdMatch[1];

    try {
      const spreadsheetResponse = await gapi.client.sheets.spreadsheets.get({
        spreadsheetId: spreadsheetId,
      });

      sheets = spreadsheetResponse.result.sheets;
      if (!sheets || sheets.length === 0) {
        console.error("No sheets found in the spreadsheet.");
        return;
      }

      const sheetSelect = document.getElementById("sheetSelect");
      const sheetName = sheetSelect.value || sheets[0].properties.title;
      const response = await gapi.client.sheets.spreadsheets.values.get({
        spreadsheetId: spreadsheetId,
        range: sheetName,
      });

      const range = response.result;
      if (!range || !range.values || range.values.length === 0) {
        console.error("No values found.");
        return;
      }

      const values = range.values;
      headers = values[0].slice(1);
      transactionMap = extractTransactions(values, headers);
    } catch (error) {
      console.error("Error fetching Google Sheet data:", error);
      return;
    }
  }

  // Find nullifiable and non-nullifiable pairs
  const { nullifiablePairs: localNullifiablePairs, nonNullifiablePairs } =
    findNullifiablePairs(transactionMap);

  // Assign to global variable
  nullifiablePairs = localNullifiablePairs;

  // Detailed logging
  console.log("Transaction Map:", transactionMap);
  console.log("Nullifiable Pairs:", nullifiablePairs);
  console.log("Non-Nullifiable Pairs:", nonNullifiablePairs);

  // Highlighting logic for nullifiable and non-nullifiable pairs
  if (isGoogleSheetData) {
    const sheetUrl = document.getElementById("googleSheetUrl").value;
    const sheetIdMatch = sheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
    const spreadsheetId = sheetIdMatch[1];

    try {
      const spreadsheetResponse = await gapi.client.sheets.spreadsheets.get({
        spreadsheetId: spreadsheetId,
      });

      const sheets = spreadsheetResponse.result.sheets;
      const sheetSelect = document.getElementById("sheetSelect");
      const sheetName = sheetSelect.value || sheets[0].properties.title;
      const selectedSheet = sheets.find(
        (sheet) => sheet.properties.title === sheetName
      );
      const sheetId = selectedSheet.properties.sheetId;
      // Combine all pairs for comprehensive highlighting
      const allPairRequests = [
        ...nullifiablePairs.flatMap(({ companyA, companyB }) => [
          {
            repeatCell: {
              range: {
                sheetId: sheetId,
                startRowIndex: headers.findIndex((c) => c === companyA) + 1,
                endRowIndex: headers.findIndex((c) => c === companyA) + 2,
                startColumnIndex: headers.findIndex((c) => c === companyB) + 1,
                endColumnIndex: headers.findIndex((c) => c === companyB) + 2,
              },
              cell: {
                userEnteredFormat: {
                  backgroundColor: {
                    red: 1,
                    green: 1,
                    blue: 0,
                    alpha: 0.5, // Yellow color for nullifiable pairs
                  },
                },
              },
              fields: "userEnteredFormat(backgroundColor)",
            },
          },
          {
            repeatCell: {
              range: {
                sheetId: sheetId,
                startRowIndex: headers.findIndex((c) => c === companyB) + 1,
                endRowIndex: headers.findIndex((c) => c === companyB) + 2,
                startColumnIndex: headers.findIndex((c) => c === companyA) + 1,
                endColumnIndex: headers.findIndex((c) => c === companyA) + 2,
              },
              cell: {
                userEnteredFormat: {
                  backgroundColor: {
                    red: 1,
                    green: 1,
                    blue: 0,
                    alpha: 0.5, // Yellow color for nullifiable pairs
                  },
                },
              },
              fields: "userEnteredFormat(backgroundColor)",
            },
          },
        ]),
        ...nonNullifiablePairs.flatMap(({ companyA, companyB }) => [
          {
            repeatCell: {
              range: {
                sheetId: sheetId,
                startRowIndex: headers.findIndex((c) => c === companyA) + 1,
                endRowIndex: headers.findIndex((c) => c === companyA) + 2,
                startColumnIndex: headers.findIndex((c) => c === companyB) + 1,
                endColumnIndex: headers.findIndex((c) => c === companyB) + 2,
              },
              cell: {
                userEnteredFormat: {
                  backgroundColor: null, // Reset background
                  textFormat: {
                    foregroundColor: null, // Reset text color
                  },
                },
              },
              fields: "userEnteredFormat(backgroundColor,textFormat)",
            },
          },
          {
            repeatCell: {
              range: {
                sheetId: sheetId,
                startRowIndex: headers.findIndex((c) => c === companyB) + 1,
                endRowIndex: headers.findIndex((c) => c === companyB) + 2,
                startColumnIndex: headers.findIndex((c) => c === companyA) + 1,
                endColumnIndex: headers.findIndex((c) => c === companyA) + 2,
              },
              cell: {
                userEnteredFormat: {
                  backgroundColor: null, // Reset background
                  textFormat: {
                    foregroundColor: null, // Reset text color
                  },
                },
              },
              fields: "userEnteredFormat(backgroundColor,textFormat)",
            },
          },
        ]),
      ];

      // Execute batch update for all pairs
      if (allPairRequests.length > 0) {
        await gapi.client.sheets.spreadsheets.batchUpdate({
          spreadsheetId: spreadsheetId,
          resource: { requests: allPairRequests },
        });
      }
    } catch (error) {
      console.error("Error highlighting transactions:", error);
    }
  }
}

// Call this function after populating the data
findNullifiableTransactions();

const NULLIFY_PASSWORD = "123";

// Event listener for nullify button
document.getElementById("nullifyBtn").addEventListener("click", function () {
  // Prompt for password
  const enteredPassword = prompt("Enter the password to nullify amounts:");

  // Check if password matches
  if (enteredPassword === NULLIFY_PASSWORD) {
    // Call the function to nullify amounts
    nullifyAmounts();
  } else {
    // Show error for incorrect password
    alert("Incorrect password. Nullification cancelled.");
  }
});

function nullifyAmounts() {
  // Google Sheets nullification
  if (isGoogleSheetData) {
    nullifyGoogleSheetAmounts();
  }
  // Excel nullification
  else {
    nullifyExcelAmounts();
  }
}

function nullifyGoogleSheetAmounts() {
  // Get the spreadsheet URL
  const sheetUrl = document.getElementById("googleSheetUrl").value;
  const sheetIdMatch = sheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);

  if (!sheetIdMatch) {
    console.error("Invalid Google Sheet URL");
    return;
  }

  const spreadsheetId = sheetIdMatch[1];

  // Add totals to the selected sheet
  addTotalsToGoogleSheet(spreadsheetId);

  // Use the nullifiablePairs from previous processing
  if (!nullifiablePairs || nullifiablePairs.length === 0) {
    alert("No nullifiable amounts found.");
    return;
  }

  // Get the spreadsheet metadata to retrieve sheet names
  gapi.client.sheets.spreadsheets
    .get({
      spreadsheetId: spreadsheetId,
    })
    .then(async (spreadsheetResponse) => {
      const sheets = spreadsheetResponse.result.sheets;
      const sheetSelect = document.getElementById("sheetSelect");
      const sheetName = sheetSelect.value || sheets[0].properties.title;
      const selectedSheet = sheets.find(
        (sheet) => sheet.properties.title === sheetName
      );
      const sheetId = selectedSheet.properties.sheetId;

      // Fetch the current sheet data to get headers
      const response = await gapi.client.sheets.spreadsheets.values.get({
        spreadsheetId: spreadsheetId,
        range: sheetName,
      });

      const range = response.result;
      if (!range || !range.values || range.values.length === 0) {
        throw new Error("No data found in the selected sheet");
      }

      // Use the headers from the current sheet
      const headers = range.values[0].slice(1);

      // Prepare batch update requests to set nullifiable amounts to zero
      const nullifyRequests = nullifiablePairs.flatMap(
        ({ companyA, companyB }) => [
          {
            updateCells: {
              range: {
                sheetId: sheetId, // Use the selected sheet's ID
                startRowIndex: headers.findIndex((c) => c === companyA) + 1,
                endRowIndex: headers.findIndex((c) => c === companyA) + 2,
                startColumnIndex: headers.findIndex((c) => c === companyB) + 1,
                endColumnIndex: headers.findIndex((c) => c === companyB) + 2,
              },
              rows: [{ values: [{ userEnteredValue: { numberValue: 0 } }] }],
              fields: "userEnteredValue",
            },
          },
          {
            updateCells: {
              range: {
                sheetId: sheetId, // Use the selected sheet's ID
                startRowIndex: headers.findIndex((c) => c === companyB) + 1,
                endRowIndex: headers.findIndex((c) => c === companyB) + 2,
                startColumnIndex: headers.findIndex((c) => c === companyA) + 1,
                endColumnIndex: headers.findIndex((c) => c === companyA) + 2,
              },
              rows: [{ values: [{ userEnteredValue: { numberValue: 0 } }] }],
              fields: "userEnteredValue",
            },
          },
        ]
      );

      // Execute batch update to set nullifiable amounts to zero
      return gapi.client.sheets.spreadsheets.batchUpdate({
        spreadsheetId: spreadsheetId,
        resource: { requests: nullifyRequests },
      });
    })
    .then(async () => {
      // Refetch the updated sheet data and re-run nullifiable transactions check
      await findNullifiableTransactions();
      alert("Amounts nullified successfully!");
    })
    .catch((error) => {
      console.error("Error nullifying amounts:", error);
      alert("Failed to nullify amounts.");
    });
}

function nullifyExcelAmounts() {
  // Ensure we have nullifiable pairs from previous processing
  if (!globalNullifiablePairs || globalNullifiablePairs.length === 0) {
    alert("No nullifiable amounts found.");
    return;
  }

  // Create loading overlay with performance-optimized UI
  const loadingOverlay = document.createElement("div");
  loadingOverlay.innerHTML = `
    <div style="
      position: fixed;
      top: 0;
      left: 0;
      width: 100%;
      height: 100%;
      background: rgba(0,0,0,0.5);
      display: flex;
      justify-content: center;
      align-items: center;
      z-index: 9999;
    ">
      <div style="
        background: white;
        padding: 20px;
        border-radius: 10px;
        text-align: center;
        min-width: 300px;
      ">
        <h3>Nullifying Amounts</h3>
        <div id="progressContainer" style="width: 100%; background: #e0e0e0; border-radius: 5px; margin-top: 10px;">
          <div id="progressBar" style="width: 0%; height: 20px; background: #4CAF50; border-radius: 5px; transition: width 0.1s;"></div>
        </div>
        <p id="progressText">Preparing to nullify...</p>
      </div>
    </div>
  `;
  document.body.appendChild(loadingOverlay);

  const progressBar = loadingOverlay.querySelector("#progressBar");
  const progressText = loadingOverlay.querySelector("#progressText");

  // Advanced Nullification Function with Web Workers
  function nullifyAmountsAdvanced() {
    return new Promise((resolve, reject) => {
      // Optimization: Create a map for faster lookup
      const nullifiablePairsMap = new Map(
        globalNullifiablePairs.map((pair) => [
          `${pair.companyA}-${pair.companyB}`,
          true,
        ])
      );

      // Parallel processing strategy
      const worksheetRows = worksheet.rowCount;
      const worksheetColumns = worksheet.columnCount;

      // Performance tracking
      let processedCells = 0;
      const totalCells = (worksheetRows - 1) * (worksheetColumns - 1);

      // Batch update strategy
      function processRowBatch(startRow, batchSize) {
        const endRow = Math.min(startRow + batchSize, worksheetRows);

        for (let rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
          const row = worksheet.getRow(rowIndex);
          const fromCompany = row.getCell(1).value;

          for (let colIndex = 2; colIndex <= worksheetColumns; colIndex++) {
            const toCompany = worksheet.getRow(1).getCell(colIndex).value;

            // Ultra-fast lookup
            if (
              nullifiablePairsMap.has(`${fromCompany}-${toCompany}`) ||
              nullifiablePairsMap.has(`${toCompany}-${fromCompany}`)
            ) {
              const amountCell = row.getCell(colIndex);
              amountCell.value = 0;
            }

            // Update progress
            processedCells++;
            if (processedCells % 100 === 0) {
              const progress = Math.round((processedCells / totalCells) * 100);
              progressBar.style.width = `${progress}%`;
              progressText.textContent = `Nullifying: ${progress}%`;
            }
          }
        }
      }

      // Asynchronous batch processing
      function processBatches() {
        const batchSize = 50; // Adjust based on performance
        let currentRow = 2;

        function processBatch() {
          processRowBatch(currentRow, batchSize);
          currentRow += batchSize;

          if (currentRow < worksheetRows) {
            // Use setTimeout for non-blocking
            setTimeout(processBatch, 0);
          } else {
            resolve();
          }
        }

        processBatch();
      }

      // Start processing
      processBatches();
    });
  }

  // Nullification workflow
  nullifyAmountsAdvanced()
    .then(() => {
      // Remove loading overlay
      document.body.removeChild(loadingOverlay);
      alert("Amounts nullified successfully!");
    })
    .catch((error) => {
      console.error("Nullification Error:", error);

      // Ensure overlay removal
      if (document.body.contains(loadingOverlay)) {
        document.body.removeChild(loadingOverlay);
      }
    });
}

// Function to update dropdown options based on affiliation type
function updateCompanyDropdowns() {
  const isAffiliated = affiliatedRadio.checked;

  // Determine available companies for row and column selections
  let availableRowCompanies, availableColumnCompanies;

  // Find the index of the first blank row that separates affiliated and non-affiliated companies
  let blankRowIndex = rowCompanies.findIndex((company) => !company);

  if (isAffiliated) {
    // Affiliated: Show only companies that exist in both rows and columns, excluding the blank row
    availableRowCompanies = rowCompanies.slice(0, rowCompanies.length).filter(
      (company) =>
        company !== "Total Affiliated" &&
        company !== "Total Non-Affiliated" &&
        company !== "Grand Total" && // Exclude Grand Total
        company !== null &&
        company !== undefined &&
        companyList.includes(company)
    );
    availableColumnCompanies = companyList.filter(
      (company) => company !== "Total" // Exclude Total column
    ); // Companies that can be columns
    console.log(availableRowCompanies, ":::rowCompanies");
  } else {
    // Non-Affiliated: Show all companies for rows, excluding Total rows and Grand Total
    availableRowCompanies = rowCompanies.filter(
      (company) =>
        company !== "Total Affiliated" &&
        company !== "Total Non-Affiliated" &&
        company !== "Grand Total" && // Exclude Grand Total
        company !== null &&
        company !== undefined &&
        company !== ""
    );
    availableColumnCompanies = companyList.filter(
      (company) => company !== "Total" // Exclude Total column
    ); // Limit columns to those in the header
  }

  // Populate dropdowns
  populateSelect(companyRowSelect, availableRowCompanies);
  populateSelect(companyColumnSelect, availableColumnCompanies);

  console.log(availableRowCompanies, ":::availableRowCompanies");
  console.log(availableColumnCompanies, ":::availableColumnCompanies");

  colSelect.update();
  rowSelect.update();
}
// Event listeners for radio buttons
affiliatedRadio.addEventListener("change", updateCompanyDropdowns);
nonAffiliatedRadio.addEventListener("change", updateCompanyDropdowns);

// Initial population of dropdowns
updateCompanyDropdowns();

// Function to populate dropdowns
function populateSelect(selectElement, options) {
  console.log(options, ":::options");
  selectElement.innerHTML =
    '<option value="" selected>Select a company</option>';
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
    colSelect.update();
    rowSelect.update();
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

class ProgressTracker {
  constructor(options = {}) {
    // Dynamic delay calculation based on operation complexity
    const calculateDelay = (baseDelay) => {
      // Consider file size, number of companies, and operation type
      const fileSize = options.fileSize || 0;
      const companiesCount = options.companiesCount || 0;

      // Base delay calculation with exponential backoff
      let dynamicDelay = baseDelay;

      // Adjust delay based on file size (in KB)
      if (fileSize > 0) {
        dynamicDelay += Math.log(fileSize) * 100;
      }

      // Adjust delay based on number of companies
      if (companiesCount > 0) {
        dynamicDelay += Math.sqrt(companiesCount) * 50;
      }

      // Ensure minimum and maximum delay
      return Math.max(50, Math.min(dynamicDelay, 100));
    };

    this.options = {
      container: document.body,
      showDelay: calculateDelay(50), // Dynamic show delay
      hideDelay: calculateDelay(50), // Dynamic hide delay
      fileSize: options.fileSize || 0,
      companiesCount: options.companiesCount || 0,
      ...options,
    };

    this.startTime = null;
    this.endTime = null;
    this.overlayElement = null;
    this.timeoutId = null;
  }

  create() {
    // Create a lightweight, dynamically sized overlay
    this.overlayElement = document.createElement("div");
    this.overlayElement.style.cssText = `
      position: fixed;
      top: 0;
      left: 0;
      width: 100%;
      height: 100%;
      background: rgba(0,0,0,0.5);
      display: flex;
      justify-content: center;
      align-items: center;
      z-index: 9999;
      opacity: 0;
      transition: opacity 0.3s ease;
    `;

    const progressContainer = document.createElement("div");
    progressContainer.style.cssText = `
      background: white;
      padding: 20px;
      border-radius: 10px;
      text-align: center;
      min-width: 300px;
      max-width: 500px;
      box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    `;

    this.progressBar = document.createElement("div");
    this.progressBar.style.cssText = `
      width: 100%;
      height: 10px;
      background: #e0e0e0;
      border-radius: 5px;
      overflow: hidden;
      margin-top: 10px;
    `;

    this.progressFill = document.createElement("div");
    this.progressFill.style.cssText = `
      width: 0%;
      height: 100%;
      background: #4CAF50;
      transition: width 0.3s ease;
    `;

    this.messageElement = document.createElement("p");
    this.messageElement.style.margin = "10px 0";

    this.detailElement = document.createElement("p");
    this.detailElement.style.fontSize = "0.8em";
    this.detailElement.style.color = "#666";

    this.progressBar.appendChild(this.progressFill);
    progressContainer.appendChild(this.messageElement);
    progressContainer.appendChild(this.progressBar);
    progressContainer.appendChild(this.detailElement);

    this.overlayElement.appendChild(progressContainer);
  }

  start() {
    // Only create if not already created
    if (!this.overlayElement) {
      this.create();
    }

    this.startTime = performance.now();

    // Delayed show to prevent flicker for quick operations
    this.timeoutId = setTimeout(() => {
      this.options.container.appendChild(this.overlayElement);
      // Force reflow to enable transition
      this.overlayElement.offsetWidth;
      this.overlayElement.style.opacity = "1";
    }, this.options.showDelay);

    return this;
  }

  update(options = {}) {
    if (!this.overlayElement) return this;

    const { progress = 0, message = "", detail = "" } = options;

    // Update progress bar
    requestAnimationFrame(() => {
      this.progressFill.style.width = `${Math.min(
        100,
        Math.max(0, progress)
      )}%`;

      if (message) {
        this.messageElement.textContent = message;
      }

      if (detail) {
        this.detailElement.textContent = detail;
      }
    });

    return this;
  }

  success(message = "Operation completed successfully") {
    this.update({
      progress: 100,
      message: message,
      detail: "",
    });

    this.end(true);
    return this;
  }

  error(message = "Operation failed") {
    this.progressFill.style.background = "#FF6B6B";
    this.update({
      progress: 100,
      message: message,
      detail: "",
    });

    this.end(false);
    return this;
  }

  end(success = true) {
    // Clear any pending show timeout
    if (this.timeoutId) {
      clearTimeout(this.timeoutId);
    }

    this.endTime = performance.now();

    // Fade out and remove
    if (this.overlayElement) {
      this.overlayElement.style.opacity = "0";

      setTimeout(() => {
        if (this.overlayElement && this.overlayElement.parentNode) {
          this.overlayElement.parentNode.removeChild(this.overlayElement);
        }
        this.overlayElement = null;
      }, this.options.hideDelay);
    }

    return success;
  }
}

// Add new company to both row and column
async function addCompany() {
  const file = companyListFileInput.files[0];

  if (file) {
    try {
      // Calculate file size and estimate companies
      const fileSize = file ? file.size / 1024 : 0; // Size in KB
      const estimatedCompaniesCount =
        fileSize > 0
          ? Math.ceil(fileSize / 10) // Rough estimate: 1 company per 10 KB
          : 0;

      // Create progress tracker with dynamic delays
      const progressTracker = new ProgressTracker({
        fileSize: fileSize,
        companiesCount: estimatedCompaniesCount,
      });

      // Start tracking with dynamic configuration
      progressTracker.start().update({
        progress: 10,
        message: "Preparing to add companies",
        detail: `File size: ${fileSize.toFixed(
          2
        )} KB, Estimated companies: ${estimatedCompaniesCount}`,
      });
      const reader = new FileReader();
      reader.onload = async (e) => {
        const data = new Uint8Array(e.target.result);
        const companies = [];

        progressTracker.update({
          progress: 30,
          message: "Processing file",
          detail: `Parsing ${file.name}, Size: ${fileSize.toFixed(2)} KB`,
        });

        // Check file extension
        const fileName = file.name.toLowerCase();
        const fileExtension = fileName.split(".").pop();

        if (fileExtension === "xlsx") {
          // Excel file processing
          const newWorkbook = new ExcelJS.Workbook();
          await newWorkbook.xlsx.load(data);

          const newWorksheet = newWorkbook.worksheets[0];
          newWorksheet.eachRow((row, rowNumber) => {
            if (rowNumber >= 1) {
              const companyName = row.getCell(1).value;
              if (companyName) {
                // Normalize the company name to handle different types of input
                const normalizedName = normalizeCompanyName(companyName);
                if (normalizedName) {
                  companies.push(normalizedName);
                }
              }
            }
          });
        } else if (fileExtension === "csv") {
          // CSV file processing with improved encoding handling
          const csvString = decodeCSVContent(data);
          const rows = csvString.split(/\r?\n/);

          rows.forEach((row, index) => {
            // Skip header row if needed, or adjust based on your CSV structure
            if (index > 0 || row.trim() !== "") {
              // More robust CSV parsing
              const companyName = parseCSVCell(row.split(",")[0]);

              if (companyName) {
                companies.push(companyName);
              }
            }
          });
        } else {
          // Unsupported file type
          alert("Please upload an Excel (.xlsx) or CSV (.csv) file.");
          return;
        }

        progressTracker.update({
          progress: 40,
          message: "Finalizing",
          detail: `Added ${companies.length} companies`,
        });

        // Filter out companies that match existing companies case-insensitively
        const existingCompanies = companies.filter(
          (company) =>
            companyList.includes(company) || rowCompanies.includes(company)
        );

        // If there are existing companies, show an alert
        if (existingCompanies.length > 0) {
          alert(
            `The following companies already exist: ${existingCompanies.join(
              ", "
            )}`
          );
        }

        // Filter out existing companies to get unique new companies
        const uniqueNewCompanies = companies.filter(
          (company) =>
            !companyList.includes(company) && !rowCompanies.includes(company)
        );

        // Check if there are any unique new companies
        if (uniqueNewCompanies.length === 0) {
          progressTracker.error("No new companies to add.");

          setTimeout(() => {
            alert("No new companies to add.");
          }, 100);

          return;
        }

        // if (companies.length === 0) {
        //   alert("No new companies to add.");
        //   return;
        // }

        // Use the existing logic for adding companies
        if (isGoogleSheetData) {
          try {
            const sheetUrl = document.getElementById("googleSheetUrl").value;

            const sheetIdMatch = sheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);

            const spreadsheetId = sheetIdMatch[1];
            // Get the sheet metadata to retrieve sheet names
            const spreadsheetResponse =
              await gapi.client.sheets.spreadsheets.get({
                spreadsheetId: spreadsheetId,
              });

            const sheets = spreadsheetResponse.result.sheets;
            if (!sheets || sheets.length === 0) {
              alert("No sheets found in the spreadsheet.");
              return;
            }

            const sheetSelect = document.getElementById("sheetSelect");
            const sheetName = sheetSelect.value || sheets[0].properties.title;
            console.log(sheetName);

            const selectedSheet = sheets.find(
              (sheet) => sheet.properties.title === sheetName
            );

            // Fetch the current sheet data
            const response = await gapi.client.sheets.spreadsheets.values.get({
              spreadsheetId: spreadsheetId,
              range: sheetName,
            });

            const range = response.result;

            // Find the blank row index
            let blankRowIndex = rowCompanies.findIndex((company, index) => {
              return (
                (company === null ||
                  company === undefined ||
                  company.toString().trim() === "") &&
                index < rowCompanies.length - 1
              );
            });

            if (blankRowIndex === -1) {
              blankRowIndex = rowCompanies.length;
            }

            // Prepare the request for updating the sheet
            const requests = [];
            progressTracker.update({
              progress: 50,
              message: "Finalizing",
              detail: `Added ${companies.length} companies`,
            });
            if (affiliatedRadio.checked) {
              // Affiliated mode: Add to column headers and rows
              const columnIndex = companyList.length + 1;

              // Find the index of 'Total Affiliated' row
              let totalAffiliatedRowIndex = -1;

              // Safely check for Total Affiliated row
              if (range && range.values && Array.isArray(range.values)) {
                totalAffiliatedRowIndex = range.values.findIndex(
                  (row) => row && row[0] === "Total Affiliated"
                );
              }

              // Find the index of the Total column
              let totalColumnIndex = -1;
              if (
                range &&
                range.values &&
                Array.isArray(range.values) &&
                range.values[0]
              ) {
                totalColumnIndex = range.values[0].indexOf("Total");
              }
              console.log(totalColumnIndex);
              // Determine the insertion point for new companies
              let insertionColumnIndex;
              if (totalColumnIndex !== -1) {
                // If Total column exists, insert before it
                insertionColumnIndex = totalColumnIndex;
              } else {
                // If Total column doesn't exist, use the original column index
                insertionColumnIndex = columnIndex;
              }

              // Determine the insertion point for new companies in rows
              let insertionRowIndex;
              if (totalAffiliatedRowIndex !== -1) {
                // If 'Total Affiliated' exists, insert before it
                insertionRowIndex = totalAffiliatedRowIndex;
              } else {
                // If 'Total Affiliated' doesn't exist, use the original blankRowIndex
                insertionRowIndex = blankRowIndex + 1;
              }

              progressTracker.update({
                progress: 60,
                message: "Finalizing",
                detail: `Added ${companies.length} companies`,
              });
              // Add header request
              requests.push({
                insertDimension: {
                  range: {
                    sheetId: selectedSheet.properties.sheetId,
                    dimension: "COLUMNS",
                    startIndex: insertionColumnIndex,
                    endIndex:
                      insertionColumnIndex +
                      (uniqueNewCompanies ? uniqueNewCompanies.length : 0),
                  },
                  inheritFromBefore: false,
                },
              });

              // Set the new column header values
              const headerValues = (uniqueNewCompanies || []).map(
                (company) => ({
                  userEnteredValue: { stringValue: company },
                  userEnteredFormat: {
                    textFormat: {
                      foregroundColor: {
                        red: 0.0, // Black text
                        green: 0.0,
                        blue: 0.0,
                      },
                      bold: false,
                    },
                  },
                })
              );

              requests.push({
                updateCells: {
                  rows: [{ values: headerValues }],
                  fields: "userEnteredValue,userEnteredFormat",
                  start: {
                    sheetId: selectedSheet.properties.sheetId,
                    rowIndex: 0, // Header row
                    columnIndex: insertionColumnIndex,
                  },
                },
              });

              // Add row request
              requests.push({
                insertDimension: {
                  range: {
                    sheetId: selectedSheet.properties.sheetId,
                    dimension: "ROWS",
                    startIndex: insertionRowIndex,
                    endIndex:
                      insertionRowIndex +
                      (uniqueNewCompanies ? uniqueNewCompanies.length : 0),
                  },
                  inheritFromBefore: false,
                },
              });

              // Set the new row values
              const rowValues = (uniqueNewCompanies || []).map((company) => ({
                values: [
                  {
                    userEnteredValue: { stringValue: company },
                    userEnteredFormat: {
                      textFormat: {
                        foregroundColor: {
                          red: 0.0, // Black text
                          green: 0.0,
                          blue: 0.0,
                        },
                        bold: false,
                      },
                    },
                  },
                ],
              }));

              requests.push({
                updateCells: {
                  rows: rowValues,
                  fields: "userEnteredValue,userEnteredFormat",
                  start: {
                    sheetId: selectedSheet.properties.sheetId,
                    rowIndex: insertionRowIndex,
                    columnIndex: 0,
                  },
                },
              });
              progressTracker.update({
                progress: 70,
                message: "Finalizing",
                detail: `Added ${companies.length} companies`,
              });
            } else {
              // Non-affiliated mode: Add rows at the end of the second list
              if (
                !range.values ||
                range.values.length === 0 ||
                (range.values.length === 1 &&
                  range.values[0][0] === "Total Non-Affiliated")
              ) {
                alert("Please add an affiliated company first.");
                return;
              }
              let totalNonAffiliatedRowIndex;
              if (range && range.values && Array.isArray(range.values)) {
                totalNonAffiliatedRowIndex = range.values.findIndex(
                  (row) => row[0] === "Total Non-Affiliated"
                );
              }

              // Find the last non-empty row index before Total Non-Affiliated
              let lastNonEmptyRowIndex = range.values.length - 1;
              for (let i = range.values.length - 1; i >= 0; i--) {
                if (
                  range.values[i][0] !== null &&
                  range.values[i][0] !== undefined &&
                  range.values[i][0].toString().trim() !== "" &&
                  (totalNonAffiliatedRowIndex === -1 ||
                    i < totalNonAffiliatedRowIndex)
                ) {
                  lastNonEmptyRowIndex = i;
                  break;
                }
              }

              function findFirstBlankRow(uniqueNewCompanies) {
                for (
                  let index = 0;
                  index < uniqueNewCompanies.length;
                  index++
                ) {
                  const company = uniqueNewCompanies[index];
                  // Check if the current row is blank
                  const isCurrentRowBlank =
                    company === null ||
                    company === undefined ||
                    (typeof company === "string" && company.trim() === "");

                  // If a blank row is found, return its index
                  if (isCurrentRowBlank) {
                    return index; // Return the index of the first blank row
                  }
                }
                return -1; // Return -1 if no blank row is found
              }
              const firstBlankRowIndex = findFirstBlankRow(rowCompanies);

              // Determine the insertion point for new companies
              let insertionRowIndex;
              if (totalNonAffiliatedRowIndex !== -1) {
                // If 'Total Non-Affiliated' exists, insert before it
                insertionRowIndex = totalNonAffiliatedRowIndex;
              } else {
                // If 'Total Non-Affiliated' doesn't exist, insert at the end of non-affiliated companies
                if (firstBlankRowIndex === -1) {
                  insertionRowIndex = lastNonEmptyRowIndex + 2;
                } else {
                  insertionRowIndex = lastNonEmptyRowIndex + 1;
                }
              }

              // Prepare row values
              const rowValues = uniqueNewCompanies.map((company) => ({
                values: [
                  {
                    userEnteredValue: { stringValue: company },
                  },
                ],
              }));

              progressTracker.update({
                progress: 60,
                message: "Finalizing",
                detail: `Added ${companies.length} companies`,
              });
              // Insert new rows for new companies
              requests.push({
                insertDimension: {
                  range: {
                    sheetId: selectedSheet.properties.sheetId,
                    dimension: "ROWS",
                    startIndex: insertionRowIndex,
                    endIndex: insertionRowIndex + uniqueNewCompanies.length,
                  },
                  inheritFromBefore: false,
                },
              });

              // Update cells
              requests.push({
                updateCells: {
                  rows: rowValues,
                  fields: "userEnteredValue",
                  start: {
                    sheetId: selectedSheet.properties.sheetId,
                    rowIndex: insertionRowIndex,
                    columnIndex: 0,
                  },
                },
              });
              progressTracker.update({
                progress: 70,
                message: "Finalizing",
                detail: `Added ${companies.length} companies`,
              });
            }

            progressTracker.update({
              progress: 80,
              message: "Finalizing",
              detail: `Added ${companies.length} companies`,
            });
            // Execute batch update
            await gapi.client.sheets.spreadsheets.batchUpdate({
              spreadsheetId: spreadsheetId,
              resource: { requests: requests },
            });

            // Refetch the updated sheet data
            const updatedResponse =
              await gapi.client.sheets.spreadsheets.values.get({
                spreadsheetId: spreadsheetId,
                range: sheetName,
              });

            // Update local data structures
            const updatedData = updatedResponse.result.values;
            companyList = updatedData[0].slice(1);
            rowCompanies = updatedData.slice(1).map((row) => row[0]);

            // Update dropdowns and UI
            updateCompanyDropdowns();
            newCompanyInput.value = "";
            await fillDiagonalCells(companyList, spreadsheetId);
            await addTotalsToGoogleSheet(spreadsheetId);
            progressTracker.success(
              `Successfully processed ${companies.length} companies`
            );
          } catch (error) {
            console.error("Error adding companies to Google Sheet:", error);
            progressTracker.error("Failed to process the file.");
          }
        } else {
          // Existing Excel file upload logic remains unchanged
          for (const newCompany of uniqueNewCompanies) {
            if (
              companyList.includes(newCompany.toLowerCase()) ||
              rowCompanies.includes(newCompany.toLowerCase())
            ) {
              alert(`Company ${newCompany} already exists. Skipping.`);
              continue; // Skip to the next company if it already exists
            }

            // Determine where to add the new company
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

            if (affiliatedRadio.checked) {
              // Affiliated mode: Add to column headers and rows
              const columnIndex = companyList.length + 1;

              // Add to column header if it doesn't already exist
              if (!companyList.includes(newCompany)) {
                worksheet.getCell(
                  `${indexToColumnLetter(columnIndex)}1`
                ).value = newCompany;
                companyList.push(newCompany);
              }

              // Insert a new row before the blank row without replacing existing data
              worksheet.spliceRows(blankRowIndex + 2, 0, [newCompany]);
              rowCompanies.splice(blankRowIndex, 0, newCompany);
            } else {
              // Non-affiliated mode: Add after the last entry of the second list
              let lastNonEmptyRowIndex = rowCompanies.length; // Start with the length of rowCompanies

              // Find the last non-empty row in the second list
              for (let i = rowCompanies.length - 1; i >= 0; i--) {
                if (
                  rowCompanies[i] !== null &&
                  rowCompanies[i] !== undefined &&
                  rowCompanies[i].toString().trim() !== ""
                ) {
                  lastNonEmptyRowIndex = i + 1; // Set to the next index
                  break;
                }
              }

              // Insert a new row after the last non-empty row
              worksheet.spliceRows(lastNonEmptyRowIndex + 2, 0, [newCompany]);
              rowCompanies.splice(lastNonEmptyRowIndex + 1, 0, newCompany);
            }
          }
          if (worksheet && typeof worksheet.rowCount !== "undefined") {
            worksheet.rowCount = Math.max(
              worksheet.rowCount,
              rowCompanies.length + 1
            );
          } else {
            console.warn(
              "Worksheet is undefined or does not have a rowCount property"
            );
          }
          // console.log(worksheet.rowCount, ":::worksheet.rowCount");
          companyList = worksheet.getRow(1).values.slice(1);
          rowCompanies = [];
          for (let i = 2; i <= worksheet.rowCount; i++) {
            rowCompanies.push(worksheet.getRow(i).getCell(1).value);
          }

          updateCompanyDropdowns();
          fillDiagonalCells();
        }
      };

      reader.readAsArrayBuffer(file);
    } catch (error) {
      console.error("Error processing file:", error);
      alert("Failed to process the file.");
    }
    // Clear file input
    companyListFileInput.value = "";
  } else {
    const newCompany = newCompanyInput.value.trim();

    if (
      companyList.includes(newCompany.toLowerCase()) ||
      rowCompanies.includes(newCompany.toLowerCase())
    ) {
      alert(`Company ${newCompany} already exists. Skipping.`);
      return;
    }

    if (!newCompany) {
      alert("Please enter a company name.");
      return;
    }

    // Check if the company already exists
    if (companyList.includes(newCompany) || rowCompanies.includes(newCompany)) {
      alert("This company already exists.");
      return;
    }

    // If using Google Sheet data
    if (isGoogleSheetData) {
      try {
        const progressTracker = new ProgressTracker();

        // Start tracking
        progressTracker.start().update({
          progress: 10,
          message: "Preparing to add company",
          detail: `Adding ${newCompany}`,
        });
        // Get the sheet URL and extract spreadsheet ID
        const sheetUrl = document.getElementById("googleSheetUrl").value;
        const sheetIdMatch = sheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);

        if (!sheetIdMatch) {
          alert("Invalid Google Sheet URL.");
          return;
        }

        const spreadsheetId = sheetIdMatch[1];

        // Get the sheet metadata to retrieve sheet names
        const spreadsheetResponse = await gapi.client.sheets.spreadsheets.get({
          spreadsheetId: spreadsheetId,
        });

        const sheets = spreadsheetResponse.result.sheets;
        if (!sheets || sheets.length === 0) {
          alert("No sheets found in the spreadsheet.");
          return;
        }

        // Use the first sheet by default (you can modify this logic if needed)
        const sheetName = sheetSelect.value || sheets[0].properties.title;
        const selectedSheet = sheets.find(
          (sheet) => sheet.properties.title === sheetName
        );
        if (!selectedSheet) {
          alert("Selected sheet not found.");
          return;
        }

        // Fetch the current sheet data
        const response = await gapi.client.sheets.spreadsheets.values.get({
          spreadsheetId: spreadsheetId,
          range: sheetName,
        });

        const range = response.result;

        // Find the blank row index
        let blankRowIndex = rowCompanies.findIndex((company, index) => {
          return (
            (company === null ||
              company === undefined ||
              company.toString().trim() === "") &&
            index < rowCompanies.length - 1
          );
        });

        if (blankRowIndex === -1) {
          blankRowIndex = rowCompanies.length;
        }

        // Prepare the request for updating the sheet
        const requests = [];
        progressTracker.update({
          progress: 30,
          message: "Preparing to add company",
          detail: `Adding ${newCompany}`,
        });
        if (affiliatedRadio.checked) {
          // Affiliated mode: Add to column headers and rows
          const columnIndex = companyList.length + 1;

          let totalAffiliatedRowIndex = -1;
          let totalColumnIndex = -1;
          // Safely check for Total Affiliated row
          if (range && range.values && Array.isArray(range.values)) {
            totalAffiliatedRowIndex = range.values.findIndex(
              (row) => row && row[0] === "Total Affiliated"
            );

            // Find the index of the Total column
            totalColumnIndex = range.values[0].findIndex(
              (header) => header === "Total"
            );
          }

          // Determine the insertion point for new companies
          let insertionRowIndex;
          if (totalAffiliatedRowIndex !== -1) {
            // If 'Total Affiliated' exists, insert before it
            insertionRowIndex = totalAffiliatedRowIndex;
          } else {
            // If 'Total Affiliated' doesn't exist, use the original blankRowIndex
            insertionRowIndex = blankRowIndex + 1;
          }

          // Determine the insertion column index
          let insertionColumnIndex;
          if (totalColumnIndex !== -1) {
            // If Total column exists, insert before it
            insertionColumnIndex = totalColumnIndex;
          } else {
            // If Total column doesn't exist, use the original column index
            insertionColumnIndex = columnIndex;
          }

          progressTracker.update({
            progress: 50,
            message: "Finalizing",
            detail: `Added ${newCompany}`,
          });
          // Add header request
          requests.push({
            insertDimension: {
              range: {
                sheetId: selectedSheet.properties.sheetId,
                dimension: "COLUMNS",
                startIndex: insertionColumnIndex, // Insert at the end of the current columns
                endIndex: insertionColumnIndex + 1,
              },
              inheritFromBefore: false,
            },
          });

          // Set the new column header value
          requests.push({
            updateCells: {
              rows: [
                {
                  values: [
                    {
                      userEnteredValue: { stringValue: newCompany },
                      userEnteredFormat: {
                        backgroundColor: {
                          red: 1.0, // White background
                          green: 1.0,
                          blue: 1.0,
                        },
                        textFormat: {
                          foregroundColor: {
                            red: 0.0, // Black text
                            green: 0.0,
                            blue: 0.0,
                          },
                          bold: false,
                        },
                      },
                    },
                  ],
                },
              ],
              fields: "userEnteredValue,userEnteredFormat",
              start: {
                sheetId: selectedSheet.properties.sheetId,
                rowIndex: 0, // Header row
                columnIndex: insertionColumnIndex, // New column index
              },
            },
          });

          // Add row request
          requests.push({
            insertDimension: {
              range: {
                sheetId: selectedSheet.properties.sheetId,
                dimension: "ROWS",
                startIndex: insertionRowIndex,
                endIndex: insertionRowIndex + 1,
              },
              inheritFromBefore: false,
            },
          });

          // Set the new row value
          requests.push({
            updateCells: {
              rows: [
                {
                  values: [
                    {
                      userEnteredValue: { stringValue: newCompany },
                    },
                  ],
                },
              ],
              fields: "userEnteredValue",
              start: {
                sheetId: selectedSheet.properties.sheetId,
                rowIndex: insertionRowIndex,
                columnIndex: 0,
              },
            },
          });
          progressTracker.update({
            progress: 70,
            message: "Finalizing",
            detail: `Added ${newCompany}`,
          });
        } else {
          // Check if the sheet is empty or there are no affiliated companies
          if (
            !range.values ||
            range.values.length === 0 ||
            (range.values.length === 1 &&
              range.values[0][0] === "Total Non-Affiliated")
          ) {
            alert("Please add an affiliated company first.");
            return;
          }

          // Non-affiliated mode: Add row at the end of the second list
          let lastNonEmptyRowIndex = range.values.length - 1; // Start with the length of rowCompanies

          // Find the last non-empty row in the second list

          function findFirstBlankRow(uniqueNewCompanies) {
            for (let index = 0; index < uniqueNewCompanies.length; index++) {
              const company = uniqueNewCompanies[index];
              // Check if the current row is blank
              const isCurrentRowBlank =
                company === null ||
                company === undefined ||
                (typeof company === "string" && company.trim() === "");

              // If a blank row is found, return its index
              if (isCurrentRowBlank) {
                return index; // Return the index of the first blank row
              }
            }
            return -1; // Return -1 if no blank row is found
          }

          // Use the function to find the first blank row
          const firstBlankRowIndex = findFirstBlankRow(rowCompanies);

          // Find the index of the "Total Non-Affiliated" row
          let totalNonAffiliatedRowIndex;
          if (range && range.values && Array.isArray(range.values)) {
            totalNonAffiliatedRowIndex = range.values.findIndex(
              (row) => row[0] === "Total Non-Affiliated"
            );
          }

          console.log(firstBlankRowIndex, ":::firstBlankRowIndex");

          // Determine the insertion point
          let insertionRowIndex;
          if (totalNonAffiliatedRowIndex !== -1) {
            // If "Total Non-Affiliated" exists, insert before it
            insertionRowIndex = totalNonAffiliatedRowIndex;
          } else {
            // If "Total Non-Affiliated" doesn't exist, use the last non-empty row
            if (firstBlankRowIndex === -1) {
              insertionRowIndex = lastNonEmptyRowIndex + 2;
            } else {
              insertionRowIndex = lastNonEmptyRowIndex + 1;
            }
          }
          console.log(insertionRowIndex, ":::insertionRowIndex");

          console.log(firstBlankRowIndex, ":::firstBlankRowIndex");

          for (let i = range.values.length - 1; i >= 0; i--) {
            if (
              range.values[i][0] !== null &&
              range.values[i][0] !== undefined &&
              range.values[i][0].toString().trim() !== "" &&
              (totalNonAffiliatedRowIndex === -1 ||
                i < totalNonAffiliatedRowIndex)
            ) {
              lastNonEmptyRowIndex = i;
              break;
            }
          }
          progressTracker.update({
            progress: 50,
            message: "Finalizing",
            detail: `Added ${newCompany}`,
          });
          // Insert a new row at the determined insertion point
          requests.push({
            insertDimension: {
              range: {
                sheetId: selectedSheet.properties.sheetId,
                dimension: "ROWS",
                startIndex: insertionRowIndex,
                endIndex: insertionRowIndex + 1,
              },
              inheritFromBefore: false,
            },
          });

          // Set the new row value
          requests.push({
            updateCells: {
              rows: [
                {
                  values: [
                    {
                      userEnteredValue: { stringValue: newCompany },
                    },
                  ],
                },
              ],
              fields: "userEnteredValue",
              start: {
                sheetId: selectedSheet.properties.sheetId,
                rowIndex: insertionRowIndex,
                columnIndex: 0,
              },
            },
          });
          progressTracker.update({
            progress: 70,
            message: "Finalizing",
            detail: `Added ${newCompany}`,
          });
        }
        progressTracker.update({
          progress: 90,
          message: "Finalizing",
          detail: `Added ${newCompany}`,
        });
        // Execute batch update
        await gapi.client.sheets.spreadsheets.batchUpdate({
          spreadsheetId: spreadsheetId,
          resource: { requests: requests },
        });

        // Refetch the updated sheet data
        const updatedResponse =
          await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: spreadsheetId,
            range: sheetName,
          });

        // Update local data structures
        const updatedData = updatedResponse.result.values;
        companyList = updatedData[0].slice(1);
        rowCompanies = updatedData.slice(1).map((row) => row[0]);

        // Update dropdowns and UI
        updateCompanyDropdowns();
        newCompanyInput.value = "";
        await fillDiagonalCells(companyList, spreadsheetId);
        await addTotalsToGoogleSheet(spreadsheetId);
        progressTracker.success(`Successfully added ${newCompany}`);
      } catch (error) {
        console.error("Error adding company to Google Sheet:", error);
        progressTracker.error("Failed to add the company.");
      }
    } else {
      // Existing Excel file upload logic remains unchanged
      // Check if the company already exists
      if (
        companyList.includes(newCompany) ||
        rowCompanies.includes(newCompany)
      ) {
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

      if (affiliatedRadio.checked) {
        // Affiliated mode:
        const columnIndex = companyList.length + 1;

        // Add to column header
        worksheet.getCell(`${indexToColumnLetter(columnIndex)}1`).value =
          newCompany;
        companyList.push(newCompany);

        // Insert a new row before the blank row
        worksheet.spliceRows(blankRowIndex + 2, 0, [newCompany]);
        rowCompanies.splice(blankRowIndex, 0, newCompany);
      } else {
        // Non-affiliated mode: Add after the last entry of the second list
        let lastNonEmptyRowIndex = rowCompanies.length; // Start with the length of rowCompanies

        // Find the last non-empty row in the second list
        for (let i = rowCompanies.length - 1; i >= 0; i--) {
          if (
            rowCompanies[i] !== null &&
            rowCompanies[i] !== undefined &&
            rowCompanies[i].toString().trim() !== ""
          ) {
            lastNonEmptyRowIndex = i + 1; // Set to the next index
            break;
          }
        }

        // Insert a new row after the last non-empty row
        worksheet.spliceRows(lastNonEmptyRowIndex + 2, 0, [newCompany]);
        rowCompanies.splice(lastNonEmptyRowIndex + 1, 0, newCompany);
      }

      // Update the worksheet.rowCount property if necessary
      worksheet.rowCount = Math.max(
        worksheet.rowCount,
        rowCompanies.length + 1
      );

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
      await fillDiagonalCells(companyList);
    }
    // Clear file input
    companyListFileInput.value = "";
  }
}

// Helper function to normalize company names
function normalizeCompanyName(name) {
  // Handle different input types (string, object, etc.)
  let companyName = " ";

  if (typeof name === "string") {
    companyName = name;
  } else if (name && name.richText) {
    // Handle rich text cells in Excel
    companyName = name.richText.map((text) => text.text).join(" ");
  } else if (name && name.text) {
    companyName = name.text;
  } else if (name && typeof name.toString === "function") {
    companyName = name.toString();
  }

  // Trim and remove any problematic characters
  return companyName
    .trim()
    .replace(/[\u200B-\u200D\uFEFF]/g, " ") // Remove zero-width characters
    .replace(/�/g, " ") // Remove replacement characters
    .normalize("NFC"); // Normalize Unicode representation
}

// Helper function to decode CSV content with multiple encodings
function decodeCSVContent(data) {
  const encodings = [
    "UTF-8",
    "Windows-1252", // Most common for Western European languages
    "ISO-8859-1",
    "UTF-16",
  ];

  for (const encoding of encodings) {
    try {
      const decoder = new TextDecoder(encoding);
      const decodedString = decoder.decode(data);

      // Additional validation to ensure meaningful content
      if (decodedString && decodedString.trim().length > 0) {
        return decodedString;
      }
    } catch (error) {
      console.warn(`Failed to decode with ${encoding} encoding`);
    }
  }

  // Fallback to UTF-8 if all else fails
  return new TextDecoder("UTF-8").decode(data);
}

// Helper function to parse CSV cell with more robust handling
function parseCSVCell(cell) {
  if (!cell) return "";

  // Remove surrounding quotes
  let cleanCell = cell.trim().replace(/^["']|["']$/g, "");

  // Normalize the cell content
  return normalizeCompanyName(cleanCell);
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

  // Only proceed with Google Sheets if necessary
  if (isGoogleSheetData) {
    const sheetUrl = document.getElementById("googleSheetUrl").value;
    const sheetIdMatch = sheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!sheetIdMatch) {
      alert("Invalid Google Sheet URL.");
      return;
    }
    const spreadsheetId = sheetIdMatch[1];
    const sheetSelect = document.getElementById("sheetSelect");
    const sheetName = sheetSelect.value || sheets[0].properties.title;

    try {
      // Fetch spreadsheet and sheet data in parallel, only fetch the necessary data
      const [spreadsheetResponse, valuesResponse] = await Promise.all([
        gapi.client.sheets.spreadsheets.get({
          spreadsheetId: spreadsheetId,
        }),
        gapi.client.sheets.spreadsheets.values.get({
          spreadsheetId: spreadsheetId,
          range: sheetName,
        }),
      ]);

      // Extract values and metadata in a single step
      const sheet = spreadsheetResponse.result.sheets.find(
        (s) => s.properties.title === sheetName
      );
      const sheetId = sheet.properties.sheetId;
      const values = valuesResponse.result.values;

      rowCompanies = values.slice(1).map((row) => row[0]);

      // Find row and column indices
      const rowIndex = rowCompanies.indexOf(rowCompany) + 2; // 1-based index
      const columnIndex = companyList.indexOf(columnCompany) + 2; // 1-based index
      const cellRange = `${indexToColumnLetter(columnIndex)}${rowIndex}`;
      const fullRange = `${sheetName}!${cellRange}`;

      // Prepare the update requests
      const updateRequests = [];

      // Handle existing value check and prompt override
      const existingValueResponse =
        await gapi.client.sheets.spreadsheets.values.get({
          spreadsheetId: spreadsheetId,
          range: fullRange,
        });

      const existingValue = existingValueResponse.result.values
        ? existingValueResponse.result.values[0][0]
        : null;

      if (existingValue !== null && existingValue !== amount) {
        const confirmOverride = confirm(
          `The current value is ${existingValue}. Do you want to override it with ${amount}?`
        );
        if (!confirmOverride) {
          return;
        }
      }

      // Add value update request
      updateRequests.push({
        updateCells: {
          range: {
            sheetId: sheetId,
            startRowIndex: rowIndex - 1,
            endRowIndex: rowIndex,
            startColumnIndex: columnIndex - 1,
            endColumnIndex: columnIndex,
          },
          rows: [
            {
              values: [
                {
                  userEnteredValue: {
                    numberValue: amount,
                  },
                },
              ],
            },
          ],
          fields: "userEnteredValue",
        },
      });

      // Handle comment update if provided
      if (comment) {
        const formattedComment = formatComment(comment, selectedEmployee);

        // Check for existing comment and prompt override
        const commentResponse = await gapi.client.sheets.spreadsheets.get({
          spreadsheetId: spreadsheetId,
          includeGridData: true,
          ranges: [fullRange],
        });

        // Safely extract existing note with multiple fallback checks
        let existingNote = null;
        try {
          existingNote =
            commentResponse.result.sheets?.[0]?.data?.[0]?.rowData?.[0]
              ?.values?.[0]?.note || null;
        } catch (extractionError) {
          console.warn("Could not extract existing note:", extractionError);
        }

        if (existingNote) {
          const confirmOverrideComment = confirm(
            `An existing comment exists: "${existingNote}". Do you want to override it with "${formattedComment}"?`
          );
          if (!confirmOverrideComment) {
            return;
          }
        }

        // Add comment update request
        updateRequests.push({
          repeatCell: {
            range: {
              sheetId: sheetId,
              startRowIndex: rowIndex - 1,
              endRowIndex: rowIndex,
              startColumnIndex: columnIndex - 1,
              endColumnIndex: columnIndex,
            },
            cell: {
              note: formattedComment,
            },
            fields: "note",
          },
        });
      }

      // Perform batch update (value and comment) together
      await gapi.client.sheets.spreadsheets.batchUpdate({
        spreadsheetId: spreadsheetId,
        resource: {
          requests: updateRequests,
        },
      });

      // Update the displayed data and fill diagonal cells
      updateDataTable(columnCompany, rowCompany, amount);
      await fillDiagonalCells(companyList, spreadsheetId);
    } catch (error) {
      console.error("Error updating Google Sheet:", error);
      alert("Failed to update the sheet. " + error.message);
    }
  } else {
    // Existing Excel file logic
    if (isGoogleSheetData) {
      columnIndex = companyList.indexOf(columnCompany) + 2; // Add 1 for Excel 1-based index
    } else {
      columnIndex = companyList.indexOf(columnCompany) + 1; // Add 1 for Excel 1-based index
    }
    let rowIndex = rowCompanies.indexOf(rowCompany) + 2;
    // Ensure the row exists
    const row = worksheet.getRow(rowIndex);
    if (!row) {
      // If the row doesn't exist, add a new row to the sheet
      worksheet.addRow([rowCompany]);
      rowCompanies.push(rowCompany); // Update the rowCompanies array
      rowIndex = rowCompanies.indexOf(rowCompany) + 1; // Update the rowIndex
    }

    // Get the current value in the cell
    const cell = worksheet.getCell(
      `${indexToColumnLetter(columnIndex)}${rowIndex}`
    );
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
    if (selectedEmployee) {
      if (comment) {
        const formattedComment = formatComment(comment, selectedEmployee);

        if (existingComment !== null && existingComment !== formattedComment) {
          const confirmOverrideComment = confirm(
            `The current comment is "${existingComment}". Do you want to override it with "${formattedComment}"?`
          );

          if (!confirmOverrideComment) {
            return; // If the user chooses not to override, exit the function
          }
        }

        // Set the formatted comment
        cell.note = formattedComment;
      }
    }

    // Set the amount in the correct cell in the Excel sheet
    cell.value = amount;

    // Update the displayed list
    updateDataTable(columnCompany, rowCompany, amount);
    await fillDiagonalCells(companyList);
    await findNullifiableTransactionsExcel(worksheet);
  }

  // Clear inputs after submission
  document.getElementById("amount").value = "";
  commentTextarea.value = ""; // Clear the comment textarea
  rowSelect.clear();
  if (isGoogleSheetData) {
    await findNullifiableTransactions();
  }
});

// Function to update the displayed data table
function updateDataTable(columnCompany, rowCompany, amount) {
  const dataBody = document.getElementById("dataBody");
  const rows = Array.from(dataBody.rows);
  let updated = false;
  console.log(columnCompany);
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

// GOOGLE SHEET START
const CLIENT_ID =
  "115660540991-17v3opc0ja64ivqt8rrrd5kt4fogjto7.apps.googleusercontent.com";
const API_KEY = "AIzaSyA9EniwLTLORTX_B2RKcrKHNUujpmLMuyw";

// Discovery doc URL for APIs used by the quickstart
const DISCOVERY_DOC =
  "https://sheets.googleapis.com/$discovery/rest?version=v4";

const DISCOVERY_DOC_PEOPLE =
  "https://people.googleapis.com/$discovery/rest?version=v1";

// Authorization scopes required by the API; multiple scopes can be
// included, separated by spaces.
const SCOPES =
  "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/userinfo.email";

let gapiInited = false;
let gisInited = false;
let tokenClient;
let userEmail = "";

// document.getElementById("authorize_button").style.visibility = "hidden";
// document.getElementById("signout_button").style.visibility = "hidden";

/**
 * Callback after api.js is loaded.
 */
function gapiLoaded() {
  gapi.load("client", initGoogleAPIs); // This function is called when gapi is loaded
}

async function initGoogleAPIs() {
  await gapi.client.init({
    apiKey: API_KEY,
    discoveryDocs: [DISCOVERY_DOC, DISCOVERY_DOC_PEOPLE],
  });
  gapiInited = true;
  maybeEnableButtons();
}

/**
 * Callback after the API client is loaded. Loads the
 * discovery doc to initialize the API.
 */

/**
 * Callback after Google Identity Services are loaded.
 */
function gisLoaded() {
  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: CLIENT_ID,
    scope: SCOPES,
    callback: "",
  });
  gisInited = true;
  maybeEnableButtons();
}

document.addEventListener("DOMContentLoaded", function () {
  gapiLoaded();
  gisLoaded();
});

/**
 * Enables user interaction after all libraries are loaded.
 */
function maybeEnableButtons() {
  if (gapiInited && gisInited) {
    document.getElementById("authorize_button").style.visibility = "visible";
  }
}

/**
 *  Sign in the user upon button click.
 */

async function handleAuthClick() {
  isGoogleSheetData = true;

  downloadButton.style.display = "none";

  excelFile.setAttribute("disabled", true);
  // employeeFile.setAttribute("disabled", true);

  document.getElementById("employeeFileWrap").style.display = "none";
  document.getElementById("employeeSelectWrap").style.display = "none";

  tokenClient.callback = async (resp) => {
    if (resp.error !== undefined) {
      throw resp;
    }
    document.getElementById("authorize_button").innerText = "Refresh";
    // Fetch data after successful authentication
    await fetchDataFromSheet();
    // Wait for People API to be ready
    await new Promise((resolve) => {
      const interval = setInterval(() => {
        if (gapi.client.people) {
          clearInterval(interval);
          resolve();
        }
      }, 100); // Check every 100ms
    });

    // Fetch user's email after successful authentication
    userEmail = await fetchUserEmail();
  };

  if (gapi.client.getToken() === null) {
    // Prompt the user to select a Google Account and ask for consent to share their data
    tokenClient.requestAccessToken({ prompt: "consent" });
  } else {
    // Skip display of account chooser and consent dialog for an existing session.
    tokenClient.requestAccessToken({ prompt: "" });
  }
}

function findBlankRowIndex(values) {
  // If values is empty or undefined, return 0
  if (!values || values.length === 0) {
    return 0;
  }

  // Find the first truly blank row (all cells empty or undefined)
  const blankRowIndex = values.findIndex((row) => {
    // Check if the row is completely empty or contains only an empty string
    return (
      !row ||
      row.length === 0 ||
      (row.length === 1 &&
        (row[0] === "" || row[0] === undefined || row[0] === null))
    );
  });

  // If no blank row found, return the length of values
  return blankRowIndex === -1 ? values.length : blankRowIndex;
}

function findNonAffiliatedCompanies(values) {
  console.log("Original Values:", values);

  // Find the index of 'Total Affiliated' row
  const totalAffiliatedRowIndex = values.findIndex(
    (row) => row[0] === "Total Affiliated"
  );

  console.log("Total Affiliated Row Index:", totalAffiliatedRowIndex);

  // If 'Total Affiliated' row is not found, return empty array
  if (totalAffiliatedRowIndex === -1) {
    console.log("No 'Total Affiliated' row found");
    return [];
  }

  // Find non-affiliated companies starting from the row AFTER 'Total Affiliated'
  const nonAffiliatedCompanies = values
    .slice(totalAffiliatedRowIndex + 1)
    .filter((row) => {
      // Ensure the row exists and is not a total or empty row
      return (
        row &&
        row.length > 0 &&
        row[0] !== null &&
        row[0] !== undefined &&
        row[0].toString().trim() !== "" &&
        row[0] !== "Total Non-Affiliated" &&
        row[0] !== "Grand Total" &&
        row[0] !== "Total Affiliated"
      );
    });

  console.log("Non-Affiliated Companies:", nonAffiliatedCompanies);
  console.log("Non-Affiliated Companies Count:", nonAffiliatedCompanies.length);

  return nonAffiliatedCompanies;
}

async function addTotalsToGoogleSheet(spreadsheetId) {
  try {
    // Get the spreadsheet metadata to retrieve sheet names
    const spreadsheetResponse = await gapi.client.sheets.spreadsheets.get({
      spreadsheetId: spreadsheetId,
    });

    const sheets = spreadsheetResponse.result.sheets;
    if (!sheets || sheets.length === 0) {
      return;
    }

    // Use the first sheet by default (you can modify this logic if needed)
    const sheetSelect = document.getElementById("sheetSelect");
    const sheetName = sheetSelect.value || sheets[0].properties.title;
    const selectedSheet = sheets.find(
      (sheet) => sheet.properties.title === sheetName
    );
    // const sheetId = sheets[0].properties.sheetId;
    const sheetId = selectedSheet.properties.sheetId;

    // Fetch the current sheet data
    const response = await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId: spreadsheetId,
      range: sheetName,
    });

    const range = response.result;
    if (!range || !range.values || range.values.length === 0) {
      return;
    }

    const values = range.values;
    const headers = values[0];
    const numHeaders = headers.length;

    // Find the index of the first blank row that separates affiliated and non-affiliated companies
    let blankRowIndex = findBlankRowIndex(values) + 1;

    console.log(blankRowIndex, ":::blankRowIndex");

    // Find the index of 'Total Affiliated' row
    let totalAffiliatedRowIndex = values.findIndex(
      (row) => row[0] === "Total Affiliated"
    );

    // Find the index of 'Total Non-Affiliated' row
    let totalNonAffiliatedRowIndex = values.findIndex(
      (row) => row[0] === "Total Non-Affiliated"
    );

    // Find the index of 'Grand Total' row
    let grandTotalRowIndex = values.findIndex(
      (row) => row[0] === "Grand Total"
    );

    console.log(totalAffiliatedRowIndex, ":::totalAffiliatedRowIndex");
    // Check if 'Total Affiliated' row exists
    if (totalAffiliatedRowIndex === -1) {
      // Insert 'Total Affiliated' row
      await gapi.client.sheets.spreadsheets.batchUpdate({
        spreadsheetId: spreadsheetId,
        resource: {
          requests: [
            {
              insertDimension: {
                range: {
                  sheetId: sheetId,
                  dimension: "ROWS",
                  startIndex: blankRowIndex - 1,
                  endIndex: blankRowIndex,
                },
                inheritFromBefore: false,
              },
            },
            {
              updateCells: {
                rows: [
                  {
                    values: [
                      {
                        userEnteredValue: { stringValue: "Total Affiliated" },
                      },
                    ],
                  },
                ],
                fields: "userEnteredValue",
                start: {
                  sheetId: sheetId,
                  rowIndex: blankRowIndex - 1,
                  columnIndex: 0,
                },
              },
            },
          ],
        },
      });

      // Update the index of 'Total Affiliated' row
      totalAffiliatedRowIndex = blankRowIndex - 1;
    }

    // Use the improved function to find non-affiliated companies
    const nonAffiliatedCompanies = findNonAffiliatedCompanies(values);

    // Only proceed with adding/updating Total Non-Affiliated and Grand Total if companies exist
    if (nonAffiliatedCompanies.length > 0) {
      // Prepare requests array
      const requests = [];

      // Check if 'Total Non-Affiliated' row exists
      if (totalNonAffiliatedRowIndex === -1) {
        // If row doesn't exist, insert it
        const lastNonEmptyRowIndex =
          totalAffiliatedRowIndex + nonAffiliatedCompanies.length + 2;

        requests.push(
          // Insert dimension request
          {
            insertDimension: {
              range: {
                sheetId: sheetId,
                dimension: "ROWS",
                startIndex: lastNonEmptyRowIndex,
                endIndex: lastNonEmptyRowIndex + 1,
              },
              inheritFromBefore: false,
            },
          },
          // Update cell request
          {
            updateCells: {
              rows: [
                {
                  values: [
                    {
                      userEnteredValue: {
                        stringValue: "Total Non-Affiliated",
                      },
                    },
                  ],
                },
              ],
              fields: "userEnteredValue",
              start: {
                sheetId: sheetId,
                rowIndex: lastNonEmptyRowIndex,
                columnIndex: 0,
              },
            },
          }
        );

        // Update the index of 'Total Non-Affiliated' row
        totalNonAffiliatedRowIndex = lastNonEmptyRowIndex;
      }

      // Check if 'Grand Total' row exists
      if (grandTotalRowIndex === -1) {
        // Insert 'Grand Total' row
        requests.push(
          {
            insertDimension: {
              range: {
                sheetId: sheetId,
                dimension: "ROWS",
                startIndex: totalNonAffiliatedRowIndex + 1,
                endIndex: totalNonAffiliatedRowIndex + 2,
              },
              inheritFromBefore: false,
            },
          },
          {
            updateCells: {
              rows: [
                {
                  values: [
                    {
                      userEnteredValue: { stringValue: "Grand Total" },
                    },
                  ],
                },
              ],
              fields: "userEnteredValue",
              start: {
                sheetId: sheetId,
                rowIndex: totalNonAffiliatedRowIndex + 1,
                columnIndex: 0,
              },
            },
          }
        );

        // Update the index of 'Grand Total' row
        grandTotalRowIndex = totalNonAffiliatedRowIndex + 1;
      }

      // Execute batch update if there are any requests
      if (requests.length > 0) {
        await gapi.client.sheets.spreadsheets.batchUpdate({
          spreadsheetId: spreadsheetId,
          resource: { requests: requests },
        });
      }
    }

    // Calculate totals for affiliated companies
    const affiliatedTotals = [];
    for (let colIndex = 1; colIndex < numHeaders; colIndex++) {
      const sumFormula = `=SUM('${sheetName}'!${indexToColumnLetter(
        colIndex + 1
      )}2:${indexToColumnLetter(colIndex + 1)}${totalAffiliatedRowIndex})`;
      affiliatedTotals.push({
        userEnteredValue: {
          formulaValue: sumFormula,
        },
        userEnteredFormat: {
          numberFormat: {
            type: "NUMBER",
            pattern: "0.00",
          },
        },
      });
    }

    // Calculate totals for non-affiliated companies

    // Calculate totals for non-affiliated companies
    const nonAffiliatedTotals = [];
    for (let colIndex = 1; colIndex < numHeaders; colIndex++) {
      const sumFormula = `=SUM('${sheetName}'!${indexToColumnLetter(
        colIndex + 1
      )}${totalAffiliatedRowIndex + 2}:${indexToColumnLetter(
        colIndex + 1
      )}${totalNonAffiliatedRowIndex})`;
      nonAffiliatedTotals.push({
        userEnteredValue: {
          formulaValue: sumFormula,
        },
        userEnteredFormat: {
          numberFormat: {
            type: "NUMBER",
            pattern: "0.00",
          },
        },
      });
    }

    // Calculate grand totals
    const grandTotals = [];
    for (let colIndex = 1; colIndex < numHeaders; colIndex++) {
      const sumFormula = `=SUM('${sheetName}'!${indexToColumnLetter(
        colIndex + 1
      )}${totalAffiliatedRowIndex + 1}, '${sheetName}'!${indexToColumnLetter(
        colIndex + 1
      )}${totalNonAffiliatedRowIndex + 1})`;
      grandTotals.push({
        userEnteredValue: {
          formulaValue: sumFormula,
        },
        userEnteredFormat: {
          numberFormat: {
            type: "NUMBER",
            pattern: "0.00",
          },
        },
      });
    }

    // Define the requests variable
    const requests = [];

    // Update 'Total Affiliated' row with calculated totals
    if (totalAffiliatedRowIndex >= 0) {
      requests.push({
        updateCells: {
          rows: [
            {
              values: affiliatedTotals,
            },
          ],
          fields: "userEnteredValue,userEnteredFormat",
          start: {
            sheetId: sheetId,
            rowIndex: totalAffiliatedRowIndex,
            columnIndex: 1,
          },
        },
      });
    }

    // Update 'Total Non-Affiliated' row with calculated totals
    if (totalNonAffiliatedRowIndex >= 0) {
      requests.push({
        updateCells: {
          rows: [
            {
              values: nonAffiliatedTotals,
            },
          ],
          fields: "userEnteredValue,userEnteredFormat",
          start: {
            sheetId: sheetId,
            rowIndex: totalNonAffiliatedRowIndex,
            columnIndex: 1,
          },
        },
      });
    }

    // Update 'Grand Total' row with calculated totals
    if (grandTotalRowIndex >= 0) {
      requests.push({
        updateCells: {
          rows: [
            {
              values: grandTotals,
            },
          ],
          fields: "userEnteredValue,userEnteredFormat",
          start: {
            sheetId: sheetId,
            rowIndex: grandTotalRowIndex,
            columnIndex: 1,
          },
        },
      });
    }

    // Make a batch update request
    await gapi.client.sheets.spreadsheets.batchUpdate({
      spreadsheetId: spreadsheetId,
      resource: { requests: requests },
    });

    // Add a total column for row-wise totals
    // Find the index of the "Total" column
    const headerRow = values[0];
    const totalColumnIndex = headerRow.indexOf("Total");

    let rowWiseTotalColumnIndex;
    if (totalColumnIndex === -1) {
      // If "Total" column doesn't exist, add it
      rowWiseTotalColumnIndex = values[0].length;

      // Add column header for row-wise total
      requests.push({
        insertDimension: {
          range: {
            sheetId: sheetId,
            dimension: "COLUMNS",
            startIndex: rowWiseTotalColumnIndex,
            endIndex: rowWiseTotalColumnIndex + 1,
          },
          inheritFromBefore: false,
        },
      });

      // Update header with "Total"
      requests.push({
        updateCells: {
          rows: [
            {
              values: [
                {
                  userEnteredValue: { stringValue: "Total" },
                  userEnteredFormat: {
                    textFormat: {
                      bold: true,
                    },
                  },
                },
              ],
            },
          ],
          fields: "userEnteredValue,userEnteredFormat",
          start: {
            sheetId: sheetId,
            rowIndex: 0,
            columnIndex: rowWiseTotalColumnIndex,
          },
        },
      });
    } else {
      // If "Total" column exists, use its index
      rowWiseTotalColumnIndex = totalColumnIndex;

      // Clear existing values in the Total column
      requests.push({
        updateCells: {
          rows: values.slice(1).map(() => ({
            values: [{ userEnteredValue: { numberValue: 0 } }],
          })),
          fields: "userEnteredValue",
          start: {
            sheetId: sheetId,
            rowIndex: 1,
            columnIndex: rowWiseTotalColumnIndex,
          },
        },
      });
    }

    // Update the total column in the sheet
    const totalColumnFormula = [];
    for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
      const rowTotalFormula = `=SUM('${sheetName}'!${indexToColumnLetter(2)}${
        rowIndex + 1
      }:${indexToColumnLetter(rowWiseTotalColumnIndex)}${rowIndex + 1})`; // Adjusting for 1-based index
      totalColumnFormula.push({
        userEnteredValue: { formulaValue: rowTotalFormula },
      });
    }

    requests.push({
      updateCells: {
        rows: totalColumnFormula.map((formula) => ({
          values: [formula],
        })),
        fields: "userEnteredValue",
        start: {
          sheetId: sheetId,
          rowIndex: 1,
          columnIndex: rowWiseTotalColumnIndex,
        },
      },
    });

    // Make a batch update request
    await gapi.client.sheets.spreadsheets.batchUpdate({
      spreadsheetId: spreadsheetId,
      resource: { requests: requests },
    });
  } catch (error) {
    console.error("Error adding totals to Google Sheet:", error);
  }
}

/**
 * Fetch data from the Google Sheet using the provided URL.
 */
// Function to populate sheet selection dropdown
function populateSheetSelect(sheets) {
  const sheetSelect = document.getElementById("sheetSelect");

  // Clear existing options except the default
  sheetSelect.innerHTML =
    '<option value="" disabled selected>Select a sheet</option>';

  // Populate with available sheets
  sheets.forEach((sheet, index) => {
    const option = document.createElement("option");
    option.value = sheet.properties.title;
    option.textContent = sheet.properties.title;
    sheetSelect.appendChild(option);
  });

  // Show the sheet selection container
  document.querySelector(".sheet-select-form-floating").style.display = "block";
}

async function fetchDataFromSheet() {
  const sheetUrl = document.getElementById("googleSheetUrl").value;
  const sheetIdMatch = sheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);

  if (!sheetIdMatch) {
    return;
  }

  const spreadsheetId = sheetIdMatch[1];

  try {
    // Step 1: Get the spreadsheet metadata to retrieve sheet names
    const spreadsheetResponse = await gapi.client.sheets.spreadsheets.get({
      spreadsheetId: spreadsheetId,
    });

    const sheets = spreadsheetResponse.result.sheets;
    if (!sheets || sheets.length === 0) {
      alert("No sheets found in the spreadsheet.");
      return;
    }

    // If multiple sheets exist, show sheet selection
    if (sheets.length > 1) {
      populateSheetSelect(sheets);
      return; // Wait for user to select a sheet
    }

    // If only one sheet, proceed with fetching data
    // Use the first sheet's title, but now with more robust selection
    const sheetSelect = document.getElementById("sheetSelect");
    const sheetName = sheetSelect.value || sheets[0].properties.title;

    // Find the sheet with the selected/default name
    const selectedSheet = sheets.find(
      (sheet) => sheet.properties.title === sheetName
    );

    if (!selectedSheet) {
      alert("Selected sheet not found.");
      return;
    }

    // Fetch data for the selected/default sheet
    await fetchSelectedSheet(spreadsheetId, sheetName);
  } catch (err) {
    console.error("Error fetching spreadsheet metadata:", err);
    alert("Failed to fetch spreadsheet details.");
  }
}

// New function to fetch data for a selected sheet
async function fetchSelectedSheet(spreadsheetId, sheetName) {
  try {
    // Step 1: Fetch data from the selected sheet
    const response = await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId: spreadsheetId,
      range: sheetName,
    });

    const range = response.result;

    // Modify the data handling to support empty sheets
    if (!range || !range.values) {
      // If no data, set up an empty sheet structure
      companyList = [];
      rowCompanies = [];

      // Clear existing selects
      companyRowSelect.innerHTML =
        '<option value="" selected>Select a company</option>';
      companyColumnSelect.innerHTML =
        '<option value="" selected>Select a company</option>';

      // Reinitialize Nice Select
      if (colSelect) colSelect.destroy();
      if (rowSelect) rowSelect.destroy();

      colSelect = NiceSelect.bind(companyColumnSelect, { searchable: true });
      rowSelect = NiceSelect.bind(companyRowSelect, { searchable: true });

      // Keep the sheet selection visible
      document.querySelector(".sheet-select-form-floating").style.display =
        "block";
    }

    // If there's data, proceed with processing
    const firstSheetData = range.values;

    // Store data in the desired format
    companyList = firstSheetData[0].slice(1); // First row (header)
    rowCompanies = firstSheetData.slice(1).map((row) => row[0]); // First column (row companies)

    // Output the results to the console for debugging
    console.log("Company List:", companyList);
    console.log("Row Companies:", rowCompanies);

    populateSelect(companyRowSelect, rowCompanies);
    populateSelect(companyColumnSelect, companyList);

    // Reinitialize Nice Select
    if (colSelect) colSelect.destroy();
    if (rowSelect) rowSelect.destroy();

    colSelect = NiceSelect.bind(companyColumnSelect, { searchable: true });
    rowSelect = NiceSelect.bind(companyRowSelect, { searchable: true });

    // Fill diagonal cells for matching companies
    await fillDiagonalCells(companyList, spreadsheetId);

    // Update dropdowns based on the affiliation type
    updateCompanyDropdowns();
    await addTotalsToGoogleSheet(spreadsheetId);

    await findNullifiableTransactions();
  } catch (err) {
    console.error("Error fetching data:", err);
  }
}

// Add event listener for sheet selection
document
  .getElementById("sheetSelect")
  .addEventListener("change", async function () {
    const spreadsheetId = document
      .getElementById("googleSheetUrl")
      .value.match(/\/d\/([a-zA-Z0-9-_]+)/)[1];
    const selectedSheetName = this.value;

    // Fetch data for the selected sheet
    await fetchSelectedSheet(spreadsheetId, selectedSheetName);
  });

// Add event listener for the authorize button
document
  .getElementById("authorize_button")
  .addEventListener("click", handleAuthClick);

newCompanyInput.addEventListener("input", function () {
  const inputValue = this.value.trim().toLowerCase();
  let nearbyCompanies = [];

  if (isGoogleSheetData && rowCompanies.length > 0) {
    nearbyCompanies = rowCompanies.filter(
      (company) =>
        company &&
        company.toLowerCase().includes(inputValue) &&
        company.toLowerCase() !== "total affiliated" &&
        company.toLowerCase() !== "total non-affiliated" &&
        company.toLowerCase() !== "grand total"
    );
  } else if (worksheet && worksheet.rowCount > 1) {
    for (let i = 2; i <= worksheet.rowCount; i++) {
      const companyName = worksheet.getRow(i).getCell(1).value;
      if (companyName && companyName.toLowerCase().includes(inputValue)) {
        nearbyCompanies.push(companyName);
      }
    }
  }

  // Create a div to display the nearby companies
  let nearbyCompaniesDiv = document.getElementById("nearby-companies");
  if (!nearbyCompaniesDiv) {
    nearbyCompaniesDiv = document.createElement("div");
    nearbyCompaniesDiv.id = "nearby-companies";
    nearbyCompaniesDiv.style.position = "absolute";
    nearbyCompaniesDiv.style.background = "white";
    nearbyCompaniesDiv.style.border = "1px solid #ccc";
    nearbyCompaniesDiv.style.padding = "10px";
    nearbyCompaniesDiv.style.zIndex = "1000";
    this.parentNode.appendChild(nearbyCompaniesDiv);
  }

  // Display the nearby companies
  nearbyCompaniesDiv.innerHTML = ""; // Clear previous entries
  nearbyCompanies.forEach((company) => {
    const companyDiv = document.createElement("div");
    companyDiv.textContent = company;
    companyDiv.style.cursor = "pointer";
    companyDiv.addEventListener("click", function () {
      // Simply close the dropdown without adding to input
      nearbyCompaniesDiv.style.display = "none";
    });
    nearbyCompaniesDiv.appendChild(companyDiv);
  });

  // Hide the nearby companies div if the input is empty or no nearby companies found
  nearbyCompaniesDiv.style.display =
    inputValue === "" || nearbyCompanies.length === 0 ? "none" : "block";
});

// Add click event listener to the document to close the nearby companies div
document.addEventListener("click", function (event) {
  const nearbyCompaniesDiv = document.getElementById("nearby-companies");
  const newCompanyInput = document.getElementById("newCompany");

  if (
    nearbyCompaniesDiv &&
    nearbyCompaniesDiv.style.display !== "none" &&
    !nearbyCompaniesDiv.contains(event.target) &&
    event.target !== newCompanyInput
  ) {
    nearbyCompaniesDiv.style.display = "none";
  }
});
// GOOGLE SHEET END
