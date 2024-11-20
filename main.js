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
  rowCompanies = [],
  isGoogleSheetData = false;

async function getSheetId(spreadsheetId) {
  try {
    const response = await gapi.client.sheets.spreadsheets.get({
      spreadsheetId: spreadsheetId,
    });

    // Log the sheets to find the correct sheetId
    console.log("Sheets in the spreadsheet:", response.result.sheets);

    // Assume you want the first sheet, you can modify this logic as needed
    const sheet = response.result.sheets[0]; // Get the first sheet
    return sheet.properties.sheetId; // Return the sheetId
  } catch (error) {
    console.error("Error fetching sheet ID:", error);
    throw error; // Rethrow the error for handling elsewhere
  }
}

async function fillDiagonalCells(sheetData, spreadsheetId) {
  if (isGoogleSheetData) {
    const sheetId = await getSheetId(spreadsheetId); // Get the correct sheet ID
    const requests = []; // Array to hold the requests for batchUpdate
    const range = sheetData.length; // Assuming square matrix for diagonal

    // Loop through the diagonal elements and color specific indices
    for (let index = 0; index <= range; index++) {
      // Skip coloring for index 1 (Facebook)
      if (index === 0) continue;

      const request = {
        repeatCell: {
          range: {
            sheetId: sheetId, // Use the correct sheet ID
            startRowIndex: index,
            endRowIndex: index + 1,
            startColumnIndex: index,
            endColumnIndex: index + 1,
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
    try {
      await gapi.client.sheets.spreadsheets.batchUpdate({
        spreadsheetId: spreadsheetId,
        resource: {
          requests: requests,
        },
      });
      console.log("Diagonal cells filled with red color, skipping Facebook.");
    } catch (error) {
      console.error("Error applying cell formatting:", error);
    }
  } else {
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
  const fragment = document.createDocumentFragment();
  const rows = {};
  const existingRows = Array.from(dataBody.rows);

  // Check if the data is coming from a Google Sheet
  if (isGoogleSheetData) {
    const sheetUrl = document.getElementById("googleSheetUrl").value;
    const sheetIdMatch = sheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);

    if (!sheetIdMatch) {
      document.getElementById("content").innerText =
        "Invalid Google Sheet URL.";
      return;
    }

    const spreadsheetId = sheetIdMatch[1]; // Extracted spreadsheet ID

    try {
      // Step 1: Get the spreadsheet metadata to retrieve sheet names
      const spreadsheetResponse = await gapi.client.sheets.spreadsheets.get({
        spreadsheetId: spreadsheetId,
      });

      const sheets = spreadsheetResponse.result.sheets;
      if (!sheets || sheets.length === 0) {
        document.getElementById("content").innerText =
          "No sheets found in the spreadsheet.";
        return;
      }

      // Step 2: Select the first sheet name (or any other logic to choose a sheet)
      const sheetName = sheets[0].properties.title; // Get the title of the first sheet
      console.log("Using sheet name:", sheetName);

      // Step 3: Fetch data from the selected sheet
      const response = await gapi.client.sheets.spreadsheets.values.get({
        spreadsheetId: spreadsheetId,
        range: sheetName, // Use the dynamic sheet name
      });

      const range = response.result;
      if (!range || !range.values || range.values.length === 0) {
        document.getElementById("content").innerText = "No values found.";
        return;
      }

      // Store data in the desired format
      const allSheetsData = [range]; // Wrapping it in an array to mimic your original structure
      const firstSheetData = allSheetsData[0].values;

      // Declare new variables instead of reassigning constants
      const companyList = firstSheetData[0].slice(1); // First row (header)
      const rowCompanies = firstSheetData.slice(1).map((row) => row[0]); // First column (row companies)

      // Output the results to the console for debugging
      console.log("Company List:", companyList);
      console.log("Row Companies:", rowCompanies);

      // Populate rows based on the fetched data
      for (let rowIndex = 1; rowIndex < firstSheetData.length; rowIndex++) {
        const row = firstSheetData[rowIndex];
        const rowCompany = row[0];

        // If a specific row company is selected and it doesn't match, skip this row
        if (selectedRowCompany && rowCompany !== selectedRowCompany) {
          continue;
        }

        for (let columnIndex = 1; columnIndex < row.length; columnIndex++) {
          let amount = row[columnIndex];
          const columnCompany = companyList[columnIndex - 1]; // Adjust index for column company

          // Keep the value as is if it is formatted in parentheses
          if (typeof amount === "string") {
            // Remove any dollar signs and commas
            amount = amount.replace(/[$,]/g, ""); // Remove $, and commas

            // Check if the value is formatted as (value)
            if (amount.startsWith("(") && amount.endsWith(")")) {
              // Keep the amount as a string with parentheses
              amount = `(${amount.slice(1, -1)})`; // Retain the format (value)
            }
          }

          // If a specific column company is selected and it doesn't match, skip this column
          if (
            selectedColumnCompany &&
            columnCompany !== selectedColumnCompany
          ) {
            continue;
          }

          // Check if there is a valid amount and row and column companies are not the same
          if (
            amount !== null &&
            amount !== undefined &&
            amount.trim() !== "" && // Check for empty string
            rowCompany !== columnCompany // Ensure row and column companies are not the same
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
    } catch (err) {
      document.getElementById("content").innerText = err.message;
      return;
    }
  } else {
    // Existing Excel JS library logic goes here
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
        let amount = row.getCell(columnIndex).value;
        const columnCompany = headerRow.getCell(columnIndex).value;

        // Keep the value as is if it is formatted in parentheses
        if (typeof amount === "string") {
          amount = amount.replace(/[$,]/g, ""); // Remove $, and commas
          if (amount.startsWith("(") && amount.endsWith(")")) {
            // Keep the amount as a string with parentheses
            amount = `(${amount.slice(1, -1)})`; // Retain the format (value)
          }
        }

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
  }

  // Remove any existing rows with the same company and amount
  existingRows.forEach((row) => {
    row.remove();
  });

  // Add the new rows to the fragment
  Object.keys(rows).forEach((key) => {
    const row = rows[key];
    const totalAmount = row.amounts.join(", "); // Join amounts as a string for display

    // Define rowElement here
    const rowElement = document.createElement("tr");
    rowElement.innerHTML = `
      <td>${row.columnCompany}</td>
      <td>${row.rowCompany}</td>
      <td>${totalAmount}</td> <!-- Display the total amount -->
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

var colOptions = { searchable: true };
let colSelect = NiceSelect.bind(companyColumnSelect, colOptions);

var rowOptions = { searchable: true };
let rowSelect = NiceSelect.bind(companyRowSelect, rowOptions);

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

    colSelect.update();
    rowSelect.update();

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
  console.log(rowCompanies);
  // Find the index of the first blank row that separates affiliated and non-affiliated companies
  let blankRowIndex = rowCompanies.findIndex((company) => !company);
  console.log(blankRowIndex, ":::blankRowIndex");
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

  colSelect.update();
  rowSelect.update();

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

// Add new company to both row and column
async function addCompany() {
  const companyListFileInput = document.getElementById("companyListFile");
  const file = companyListFileInput.files[0];

  if (file) {
    // File upload logic
    try {
      const reader = new FileReader();
      reader.onload = async (e) => {
        const data = new Uint8Array(e.target.result);

        // Create a new workbook instead of overwriting the global one
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(data);

        const worksheet = workbook.worksheets[0];
        const companies = [];

        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber >= 1) {
            const companyName = row.getCell(1).value;
            if (companyName && typeof companyName === "string") {
              companies.push(companyName.trim());
            }
          }
        });

        // Process each company from the file
        for (const newCompany of companies) {
          // Check if the company already exists
          if (
            companyList.includes(newCompany) ||
            rowCompanies.includes(newCompany)
          ) {
            alert(`Company ${newCompany} already exists. Skipping.`);
            continue;
          }

          // Use the existing logic for adding companies
          if (isGoogleSheetData) {
            try {
              // Get the sheet URL and extract spreadsheet ID
              const sheetUrl = document.getElementById("googleSheetUrl").value;
              const sheetIdMatch = sheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);

              if (!sheetIdMatch) {
                alert("Invalid Google Sheet URL.");
                continue;
              }

              const spreadsheetId = sheetIdMatch[1];

              // Get the sheet metadata to retrieve sheet names
              const spreadsheetResponse =
                await gapi.client.sheets.spreadsheets.get({
                  spreadsheetId: spreadsheetId,
                });

              const sheets = spreadsheetResponse.result.sheets;
              if (!sheets || sheets.length === 0) {
                alert("No sheets found in the spreadsheet.");
                continue;
              }

              // Use the first sheet by default (you can modify this logic if needed)
              const sheetName = sheets[0].properties.title;

              // Fetch the current sheet data
              const response = await gapi.client.sheets.spreadsheets.values.get(
                {
                  spreadsheetId: spreadsheetId,
                  range: sheetName,
                }
              );

              const range = response.result;
              if (!range || !range.values || range.values.length === 0) {
                alert("No values found in the sheet.");
                continue;
              }

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

              if (affiliatedRadio.checked) {
                // Affiliated mode: Add to column headers and rows
                const columnIndex = companyList.length + 1;

                // Add header request
                requests.push({
                  insertDimension: {
                    range: {
                      sheetId: sheets[0].properties.sheetId,
                      dimension: "COLUMNS",
                      startIndex: columnIndex,
                      endIndex: columnIndex + 1,
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
                          },
                        ],
                      },
                    ],
                    fields: "userEnteredValue",
                    start: {
                      sheetId: sheets[0].properties.sheetId,
                      rowIndex: 0,
                      columnIndex: columnIndex,
                    },
                  },
                });

                // Add row request
                requests.push({
                  insertDimension: {
                    range: {
                      sheetId: sheets[0].properties.sheetId,
                      dimension: "ROWS",
                      startIndex: blankRowIndex + 1,
                      endIndex: blankRowIndex + 2,
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
                      sheetId: sheets[0].properties.sheetId,
                      rowIndex: blankRowIndex + 1,
                      columnIndex: 0,
                    },
                  },
                });
              } else {
                // Non-affiliated mode: Add row at the end of the second list
                let lastNonEmptyRowIndex = rowCompanies.length;

                // Find the last non-empty row in the second list
                for (let i = rowCompanies.length - 1; i >= 0; i--) {
                  if (
                    rowCompanies[i] !== null &&
                    rowCompanies[i] !== undefined &&
                    rowCompanies[i].toString().trim() !== ""
                  ) {
                    lastNonEmptyRowIndex = i + 1;
                    break;
                  }
                }

                // Insert a new row after the last non-empty row
                requests.push({
                  insertDimension: {
                    range: {
                      sheetId: sheets[0].properties.sheetId,
                      dimension: "ROWS",
                      startIndex: lastNonEmptyRowIndex + 1,
                      endIndex: lastNonEmptyRowIndex + 2,
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
                      sheetId: sheets[0].properties.sheetId,
                      rowIndex: lastNonEmptyRowIndex + 1,
                      columnIndex: 0,
                    },
                  },
                });
              }

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
              fillDiagonalCells(companyList, spreadsheetId);
            } catch (error) {
              console.error("Error adding company to Google Sheet:", error);
              alert("Failed to add company: " + error.message);
            }
          } else {
            // Existing Excel file upload logic remains unchanged
            // Check if the company already exists
            if (
              companyList.includes(newCompany) ||
              rowCompanies.includes(newCompany)
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
        }
        worksheet.rowCount = Math.max(
          worksheet.rowCount,
          rowCompanies.length + 1
        );
        console.log(worksheet.rowCount, ":::worksheet.rowCount");
        companyList = worksheet.getRow(1).values.slice(1);
        rowCompanies = [];
        for (let i = 2; i <= worksheet.rowCount; i++) {
          rowCompanies.push(worksheet.getRow(i).getCell(1).value);
        }
        updateCompanyDropdowns();
        fillDiagonalCells();
        // Clear file input
        companyListFileInput.value = "";
      };

      reader.readAsArrayBuffer(file);
    } catch (error) {
      console.error("Error processing file:", error);
      alert("Failed to process the file.");
    }
  } else {
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

    // If using Google Sheet data
    if (isGoogleSheetData) {
      try {
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
        const sheetName = sheets[0].properties.title;

        // Fetch the current sheet data
        const response = await gapi.client.sheets.spreadsheets.values.get({
          spreadsheetId: spreadsheetId,
          range: sheetName,
        });

        const range = response.result;
        if (!range || !range.values || range.values.length === 0) {
          alert("No values found in the sheet.");
          return;
        }

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

        if (affiliatedRadio.checked) {
          // Affiliated mode: Add to column headers and rows
          const columnIndex = companyList.length + 1;

          // Add header request
          requests.push({
            insertDimension: {
              range: {
                sheetId: sheets[0].properties.sheetId,
                dimension: "COLUMNS",
                startIndex: columnIndex, // Insert at the end of the current columns
                endIndex: columnIndex + 1,
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
                    },
                  ],
                },
              ],
              fields: "userEnteredValue",
              start: {
                sheetId: sheets[0].properties.sheetId,
                rowIndex: 0, // Header row
                columnIndex: columnIndex, // New column index
              },
            },
          });

          // Add row request
          requests.push({
            insertDimension: {
              range: {
                sheetId: sheets[0].properties.sheetId,
                dimension: "ROWS",
                startIndex: blankRowIndex + 1,
                endIndex: blankRowIndex + 2,
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
                sheetId: sheets[0].properties.sheetId,
                rowIndex: blankRowIndex + 1,
                columnIndex: 0,
              },
            },
          });
        } else {
          // Non-affiliated mode: Add row at the end of the second list
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
          requests.push({
            insertDimension: {
              range: {
                sheetId: sheets[0].properties.sheetId,
                dimension: "ROWS",
                startIndex: lastNonEmptyRowIndex + 1, // Insert after the last non-empty row
                endIndex: lastNonEmptyRowIndex + 2, // Insert one new row
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
                sheetId: sheets[0].properties.sheetId,
                rowIndex: lastNonEmptyRowIndex + 1, // Set the value in the newly inserted row
                columnIndex: 0,
              },
            },
          });
        }

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
        fillDiagonalCells(companyList, spreadsheetId);
      } catch (error) {
        console.error("Error adding company to Google Sheet:", error);
        alert("Failed to add company: " + error.message);
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
      fillDiagonalCells();
    }
  }
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
  colSelect.update();
  rowSelect.update();
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
  let columnIndex;

  if (isGoogleSheetData) {
    // Dynamically extract spreadsheetId and sheetName for Google Sheets
    const sheetUrl = document.getElementById("googleSheetUrl").value;
    const sheetIdMatch = sheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);

    if (!sheetIdMatch) {
      alert("Invalid Google Sheet URL.");
      return;
    }

    const spreadsheetId = sheetIdMatch[1];

    try {
      // Get the spreadsheet metadata to retrieve sheet names
      const spreadsheetResponse = await gapi.client.sheets.spreadsheets.get({
        spreadsheetId: spreadsheetId,
      });

      const sheets = spreadsheetResponse.result.sheets;
      if (!sheets || sheets.length === 0) {
        alert("No sheets found in the spreadsheet.");
        return;
      }

      // Use the first sheet by default (you can modify this logic if needed)
      const sheetName = sheets[0].properties.title;
      const sheetId = sheets[0].properties.sheetId;

      columnIndex = companyList.indexOf(columnCompany) + 2; // Add 1 for Excel 1-based index

      // Prepare the range for the specific cell
      const cellRange = `${indexToColumnLetter(columnIndex)}${rowIndex}`;
      const fullRange = `${sheetName}!${cellRange}`;

      // Get existing value
      const response = await gapi.client.sheets.spreadsheets.values.get({
        spreadsheetId: spreadsheetId,
        range: fullRange,
      });

      const existingValue = response.result.values
        ? response.result.values[0][0]
        : null;

      // Check if the existing value in the cell is the same
      if (existingValue !== null) {
        const confirmOverride = confirm(
          `The current value is ${existingValue}. Do you want to override it with ${amount}?`
        );

        if (!confirmOverride) {
          return; // If the user chooses not to override, exit the function
        }
      }

      // Prepare the value to update
      const updateResponse =
        await gapi.client.sheets.spreadsheets.values.update({
          spreadsheetId: spreadsheetId,
          range: fullRange,
          valueInputOption: "RAW",
          resource: {
            values: [[amount]],
          },
        });

      // Add comment if employee and comment text are provided
      if (comment) {
        // Format the comment with employee and timestamp
        const formattedComment = formatComment(comment, selectedEmployee);

        // First, check for existing comments
        try {
          // Retrieve sheet metadata to check existing comments
          const spreadsheetResponse = await gapi.client.sheets.spreadsheets.get(
            {
              spreadsheetId: spreadsheetId,
              includeGridData: true,
              ranges: [fullRange],
            }
          );

          // Extract existing note/comment from the cell
          const sheet = spreadsheetResponse.result.sheets[0];
          const rowData = sheet.data[0].rowData[0];
          const cellData = rowData.values[0];
          const existingNote = cellData.note;

          // Check if existing comment is different
          if (existingNote) {
            const confirmOverrideComment = confirm(
              `An existing comment exists: "${existingNote}". Do you want to override it with "${formattedComment}"?`
            );

            if (!confirmOverrideComment) {
              return; // If the user chooses not to override, exit the function
            }
          }

          // Proceed with adding/updating the comment
          const commentResponse =
            await gapi.client.sheets.spreadsheets.batchUpdate({
              spreadsheetId: spreadsheetId,
              resource: {
                requests: [
                  {
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
                  },
                ],
              },
            });

          console.log("Comment added successfully", commentResponse);
        } catch (commentError) {
          console.error("Error checking/adding comment:", commentError);
          alert(`Could not add comment. Error: ${commentError.message}`);
        }
      }

      // Update the displayed list
      updateDataTable(columnCompany, rowCompany, amount);
    } catch (error) {
      console.error("Error updating Google Sheet:", error);
      alert("Failed to update the sheet. " + error.message);
      return;
    }
  } else {
    // Existing Excel file logic
    if (isGoogleSheetData) {
      columnIndex = companyList.indexOf(columnCompany) + 2; // Add 1 for Excel 1-based index
    } else {
      columnIndex = companyList.indexOf(columnCompany) + 1; // Add 1 for Excel 1-based index
    }

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
  }

  // Clear inputs after submission
  document.getElementById("amount").value = "";
  commentTextarea.value = ""; // Clear the comment textarea
  rowSelect.clear();
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

// GOOGLE SHEET START
const CLIENT_ID = process.env.NEXT_PUBLIC_CLIENT_ID;
const API_KEY = process.env.NEXT_PUBLIC_API_KEY;

console.log(CLIENT_ID)

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

document.getElementById("authorize_button").style.visibility = "hidden";
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
document.getElementById("authorize_button").addEventListener("click", handleAuthClick())


/**
 * Fetch data from the Google Sheet using the provided URL.
 */
async function fetchDataFromSheet() {
  const sheetUrl = document.getElementById("googleSheetUrl").value;
  const sheetIdMatch = sheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);

  if (!sheetIdMatch) {
    document.getElementById("content").innerText = "Invalid Google Sheet URL.";
    return;
  }

  const spreadsheetId = sheetIdMatch[1]; // Extracted spreadsheet ID

  try {
    // Step 1: Get the spreadsheet metadata to retrieve sheet names
    const spreadsheetResponse = await gapi.client.sheets.spreadsheets.get({
      spreadsheetId: spreadsheetId,
    });

    const sheets = spreadsheetResponse.result.sheets;
    if (!sheets || sheets.length === 0) {
      document.getElementById("content").innerText =
        "No sheets found in the spreadsheet.";
      return;
    }

    // Step 2: Select the first sheet name (or any other logic to choose a sheet)
    const sheetName = sheets[0].properties.title; // Get the title of the first sheet
    console.log("Using sheet name:", sheetName);

    // Step 3: Fetch data from the selected sheet
    const response = await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId: spreadsheetId,
      range: sheetName, // Use the dynamic sheet name
    });

    console.log(response);

    const range = response.result;
    if (!range || !range.values || range.values.length === 0) {
      document.getElementById("content").innerText = "No values found.";
      return;
    }

    // Store data in the desired format
    const allSheetsData = [range]; // Wrapping it in an array to mimic your original structure
    const firstSheetData = allSheetsData[0].values;

    // Declare new variables instead of reassigning constants
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
  } catch (err) {
    document.getElementById("content").innerText = err.message;
    console.error("Error fetching data:", err);
  }
}

// Add event listener for the authorize button
document
  .getElementById("authorize_button")
  .addEventListener("click", handleAuthClick);
// GOOGLE SHEET END
