// global vairables declaration
var CUSTOMER_COL, INWARE_COL, PACKAGETYPE_COL, CONTAINER_COL, PACKAGENUM_COL, EXPRESS_COL, 
    WEIGHT_COL, VOLUME_COL, OVERLENGTH_COL, PAID_COL, PAIDAMOUNT_COL, FREIGHT_COL, DUTY_COL, 
    GST_COL, ADDITION_COL, INSURANCE_COL, AMOUNT_COL;

var file;

// Listen for file drop event
function handleFileDrop(event) {
    event.preventDefault();
  
    file = event.dataTransfer.files[0];
    var fileExtension = file.name.split('.').pop();
  
    if (fileExtension === 'xlsx') {
        readExcelFile(file);
    } else {
        alert('Please drop a valid XLSX file!');
    }
}

// handle uploaded file
function handleFileUpload() {
    var fileInput = document.getElementById('fileInput');
    fileInput.addEventListener('change', function() {
        file = fileInput.files[0];
        var fileExtension = file.name.split('.').pop();
        if (fileExtension === 'xlsx') {
            readExcelFile(file);
        } else {
            alert('Please upload a valid XLSX file!');
        }
    });
    fileInput.click();
}

function getColNum(jsonData) {
    for (var i = 0; i < jsonData.length; i++) {
        switch (jsonData[0][i]) {
            case '客户编号': CUSTOMER_COL = i;break;
            case '入仓号': INWARE_COL = i;break;
            case '包裹类型': PACKAGETYPE_COL = i;break;
            case '集装箱号': CONTAINER_COL = i;break;
            case '总件数': PACKAGENUM_COL = i;break;
            case '快递号': EXPRESS_COL = i;break;
            case '总重量(kg)': WEIGHT_COL = i;break;
            case '总体积重': VOLUME_COL = i;break;
            case '是否超长': OVERLENGTH_COL = i;break;
            case '是否付款': PAID_COL = i;break;
            case '实付金额(C$)': PAIDAMOUNT_COL = i;break;
            case 'freight(C$)': FREIGHT_COL = i;break;
            case 'Duty(C$)': DUTY_COL = i;break;
            case 'GST(C$)': GST_COL = i;break;
            case '附加服务': ADDITION_COL = i;break;
            case '保险': INSURANCE_COL = i;break;
            case '总金额(C$)': AMOUNT_COL = i;break;
        }
      }
}
  
// Read XLSX file data and pre-process the data
function readExcelFile(file) {
    var reader = new FileReader();
  
    reader.onload = function(event) {
      var arrayBuffer = event.target.result;
      var workbook = XLSX.read(arrayBuffer, { type: 'array' });
  
      var sheetName = workbook.SheetNames[0];
      var worksheet = workbook.Sheets[sheetName];
      var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
      // get the column numbers
      getColNum(jsonData);

      // Remove some useless columns from the JSON data
      for (var i = 0; i < jsonData.length; i++) {
            jsonData[i].splice(ADDITION_COL, 1);
            jsonData[i].splice(PAIDAMOUNT_COL, 1);
            jsonData[i].splice(PAID_COL, 1);
            jsonData[i].splice(PACKAGETYPE_COL, 1);
      }

      // get the updated column number after removal
      getColNum(jsonData);

      // adjust the price
      for (var i = 0; i < jsonData.length; i++) {
            if (jsonData[i][FREIGHT_COL] == 3) jsonData[i][FREIGHT_COL]=5.99;
            if (jsonData[i][AMOUNT_COL] == 3) jsonData[i][AMOUNT_COL]=5.99;
      }

      // Extract the header row from the jsonData
      var headerRow = jsonData[0];

      // Sort the data based on the values in the first column (except the header row)
      var sortedData = jsonData.slice(1).sort(function(a, b) {
            if (a[0]<b[0]) return -1;
            if (a[0]>b[0]) return 1;
            return 0;
      });

      // Combine the sorted data with the header row
      var sortedJsonData = [headerRow].concat(sortedData);

      // automatically creat organize file
      var jsonDataOrganize = [];

      var headerRowOrg = [jsonData[0][CUSTOMER_COL], jsonData[0][INWARE_COL], jsonData[0][PACKAGENUM_COL], jsonData[0][VOLUME_COL], "入库"];
      jsonDataOrganize.push(headerRowOrg);

      for (var i = 1; i < jsonData.length; i++) {
        var rowDataOrg = sortedJsonData[i];
        if (!isFiveLetterString(rowDataOrg[CUSTOMER_COL])) continue;
        var organizedRow = [rowDataOrg[CUSTOMER_COL], rowDataOrg[INWARE_COL], rowDataOrg[PACKAGENUM_COL], rowDataOrg[VOLUME_COL], false];
        jsonDataOrganize.push(organizedRow);
      }
      var worksheetOrgnize = XLSX.utils.aoa_to_sheet(jsonDataOrganize);
      var workbookOrgnize = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbookOrgnize, worksheetOrgnize, 'Sheet1');
      XLSX.writeFile(workbookOrgnize, "理货表-"+getCurrentTime()+'-'+file.name);

      displayTable(sortedJsonData);
      onlyShowPage(2);
    }
  
    reader.readAsArrayBuffer(file);
}

function isFiveLetterString(str) {
    var regex = /^[A-Z]{5}$/;
    return regex.test(str);
}

// Display data in a table
function displayTable(data) {
    var table = document.getElementById('tableData');
    table.innerHTML = '';
  
    var thead = document.createElement('thead');
    var tbody = document.createElement('tbody');
  
    // Create table header
    var headerRow = document.createElement('tr');
    for (var header of data[0]) {
      var th = document.createElement('th');
      th.textContent = header;
      headerRow.appendChild(th);
    }
    thead.appendChild(headerRow);

    var sortedData = data.slice(1);

    var curCustomer = sortedData[0][CUSTOMER_COL];
    var sumCustPackage = 0;
    var sumCustAmount = 0;
    var sumPackage = 0;
    var sumAmount = 0;

    // Create data rows
    for (var rowData of sortedData) {
        // valid check
        if (!isFiveLetterString(rowData[CUSTOMER_COL])) continue;

        if (rowData[CUSTOMER_COL] !== curCustomer) { // A summary row needs to be inserted
            var dataRow = document.createElement('tr');

            var td = document.createElement('td');
            td.textContent = curCustomer + "-" + "Summary";
            td.setAttribute("colspan", PACKAGENUM_COL-CUSTOMER_COL);
            td.style.color = "blue";
            td.style.textAlign = "center"; 
            dataRow.appendChild(td);

            var td = document.createElement('td');
            td.textContent = sumCustPackage;
            td.style.color = "blue";
            td.style.textAlign = "right"; 
            dataRow.appendChild(td);

            var td = document.createElement('td');
            td.textContent = sumCustAmount.toFixed(2);
            td.setAttribute("colspan", AMOUNT_COL-PACKAGENUM_COL);
            td.style.color = "blue";
            td.style.textAlign = "right"; 
            dataRow.appendChild(td);

            tbody.appendChild(dataRow);

            curCustomer = rowData[CUSTOMER_COL];
            sumCustPackage = 0;
            sumCustAmount = 0;
        } 

        var dataRow = document.createElement('tr');

        for (var i = 0; i < rowData.length; i++) {
            var td = document.createElement('td');

            if (i == FREIGHT_COL || i == AMOUNT_COL) {
                td.textContent = rowData[i].toFixed(2);
            } else {
                td.textContent = rowData[i];
            }
            td.style.textAlign = "right"; 

            dataRow.appendChild(td);
        }

        tbody.appendChild(dataRow);

        sumCustPackage += rowData[PACKAGENUM_COL];
        sumCustAmount += rowData[AMOUNT_COL];
        sumPackage += rowData[PACKAGENUM_COL];
        sumAmount += rowData[AMOUNT_COL];
    }

    // A summary row for last customer needs to be inserted
    var dataRow = document.createElement('tr');

    var td = document.createElement('td');
    td.textContent = curCustomer + "-" + "Summary";
    td.setAttribute("colspan", PACKAGENUM_COL-CUSTOMER_COL);
    td.style.color = "blue";
    td.style.textAlign = "center"; 
    dataRow.appendChild(td);

    var td = document.createElement('td');
    td.textContent = sumCustPackage;
    td.style.color = "blue";
    td.style.textAlign = "right"; 
    dataRow.appendChild(td);

    var td = document.createElement('td');
    td.textContent = sumCustAmount.toFixed(2);
    td.setAttribute("colspan", AMOUNT_COL-PACKAGENUM_COL);
    td.style.color = "blue";
    td.style.textAlign = "right"; 
    dataRow.appendChild(td);

    tbody.appendChild(dataRow);

    // add summary row for all customers
    var dataRow = document.createElement('tr');

    var td = document.createElement('td');
    td.textContent = "@ALL-Summary";
    td.setAttribute("colspan", PACKAGENUM_COL-CUSTOMER_COL);
    td.style.color = "red";
    td.style.textAlign = "center"; 
    dataRow.appendChild(td);

    var td = document.createElement('td');
    td.textContent = sumPackage;
    td.style.color = "red";
    td.style.textAlign = "right"; 
    dataRow.appendChild(td);

    var td = document.createElement('td');
    td.textContent = sumAmount.toFixed(2);
    td.setAttribute("colspan", AMOUNT_COL-PACKAGENUM_COL);
    td.style.color = "red";
    td.style.textAlign = "right"; 
    dataRow.appendChild(td);


    tbody.appendChild(dataRow);
  
    table.appendChild(thead);
    table.appendChild(tbody);

    // Make the table header sticky
    var tableHeader = document.querySelector('#tableData thead');
    tableHeader.style.position = 'sticky';
    tableHeader.style.backgroundColor = "grey"
    tableHeader.style.top = '0';
}

// show the given page, hide other pages
function onlyShowPage(pagenum) {
    var page1 = document.getElementById('page1');
    var page2 = document.getElementById('page2');
    var page3 = document.getElementById('page3');

    switch (pagenum) {
        case 1:
            page1.classList.remove('hidden');
            page2.classList.add('hidden');
            page3.classList.add('hidden');
            break;
        case 2:
            page2.classList.remove('hidden');
            page1.classList.add('hidden');
            page3.classList.add('hidden');
            break;
        case 3:
            page3.classList.remove('hidden');
            page1.classList.add('hidden');
            page2.classList.add('hidden');
            break;
    }
}
  
  // Add drop event listener
var dropZone = document.getElementById('dropZone');
dropZone.addEventListener('dragover', function(event) {
    event.preventDefault();
});
  
dropZone.addEventListener('drop', handleFileDrop);

function downloadTable() {
    // get table data
    const table = document.getElementById('tableData'); 
    const workbook = XLSX.utils.table_to_book(table, { sheet: 'Sheet1', cellStyles: true });
    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  
    const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);

    link.download = prompt('File name', "ER清单-"+getCurrentTime()+'-'+file.name); 

    link.click();
}

// Helper function to add leading zero to single-digit numbers
function getCurrentTime() {
    var currentDate = new Date();
    var currentYear = currentDate.getFullYear();
    var currentMonth = currentDate.getMonth() + 1; // Month is zero-based, so add 1
    var currentDay = currentDate.getDate();
    var currentHour = currentDate.getHours();
    var currentMinute = currentDate.getMinutes();
    var currentSecond = currentDate.getSeconds();

    var formattedTime = currentYear + addLeadingZero(currentMonth) + addLeadingZero(currentDay) +'-'+ addLeadingZero(currentHour) + addLeadingZero(currentMinute) + addLeadingZero(currentSecond);
    
    return formattedTime;
}

function addLeadingZero(number) {
    return number < 10 ? '0' + number : number;
}

function searchCustomer() {
    var input = document.getElementById('customerID');
    var customerID = input.value.toUpperCase();
    
    var table = document.getElementById('tableData'); 
    var rows = table.getElementsByTagName('tr');
    
    for (var i = 0; i < rows.length; i++) {
      var customerCol = rows[i].getElementsByTagName('td')[CUSTOMER_COL]; 
  
      if (customerCol) {
        var textValue = customerCol.textContent || customerCol.innerText;

        let isMatch = true;

        for (let i=0; i<customerID.length; i++) {
            if (textValue.toUpperCase().charAt(i) != customerID.charAt(i)) {
                isMatch = false;
                break;
            }
        }

        if (isMatch) {
            rows[i].style.display = '';
        } else {
            rows[i].style.display = 'none';
        }
      }
    }
}
  
// handle updated file
function handleOrgFileUpdate() {
    var fileUpdate = document.getElementById('OrgfileUpdate');
    fileUpdate.addEventListener('change', function() {
        file = fileUpdate.files[0];
        var fileExtension = file.name.split('.').pop();
        if (fileExtension === 'xlsx') {
            readOrgExcelFile(file);
        } else {
            alert('Please update a valid XLSX file!');
        }
    });
    fileUpdate.click();
}

function readOrgExcelFile(file) {
    var readerOrg = new FileReader();
  
    readerOrg.onload = function(event) {
        var arrayBufferOrg = event.target.result;
        var workbookOrg = XLSX.read(arrayBufferOrg, { type: 'array' });
  
        var sheetNameOrg = workbookOrg.SheetNames[0];
        var worksheetOrg = workbookOrg.Sheets[sheetNameOrg];
        var jsonDataOrg = XLSX.utils.sheet_to_json(worksheetOrg, { header: 1 });
    
        onlyShowPage(3);
        displayOrgTable(jsonDataOrg);
    }
  
    readerOrg.readAsArrayBuffer(file);
}
  
// Display data in a organize table
function displayOrgTable(data) {
    var table = document.getElementById('tableOrgData');
    table.innerHTML = '';
  
    var thead = document.createElement('thead');
    var tbody = document.createElement('tbody');
  
    // Create table header
    var headerRow = document.createElement('tr');
    for (var header of data[0]) {
      var th = document.createElement('th');
      th.textContent = header;
      headerRow.appendChild(th);
    }
    thead.appendChild(headerRow);

    var sortedOrgData = data.slice(1);

    // Create data rows
    for (var rowData of sortedOrgData) {

        var dataRow = document.createElement('tr');

        for (var i = 0; i < rowData.length; i++) {
            var td = document.createElement('td');

            // Check if it is the "入库" column
            if (data[0][i] === '入库') {
                var select = createDropdown(); // Create dropdown
                select.value = rowData[i]; // Set the value of dropdown to the cell content
                td.appendChild(select);
            } else {
                td.textContent = rowData[i];
            }

            dataRow.appendChild(td);
        }

        tbody.appendChild(dataRow);
    }

    table.appendChild(thead);
    table.appendChild(tbody);

    // Make the table header sticky
    var tableOrgHeader = document.querySelector('#tableOrgData thead');
    tableOrgHeader.style.position = 'sticky';
    tableOrgHeader.style.backgroundColor = "grey"
    tableOrgHeader.style.top = '0';
} 

// Create dropdown
function createDropdown() {
    var select = document.createElement('select');
  
    // Create options
    var option1 = document.createElement('option');
    option1.value = 'true';
    option1.textContent = 'true';
    select.appendChild(option1);
  
    var option2 = document.createElement('option');
    option2.value = 'false';
    option2.textContent = 'false';
    select.appendChild(option2);
  
    return select;
}

function handleOrgFileSave() {
    // Get table data
    const table = document.getElementById('tableOrgData');
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.table_to_sheet(table);

    // Get header row
    const headerRow = table.querySelector('thead tr');
    headerRow.querySelectorAll('th').forEach(function (th, columnIndex) {
    const cellObject = { t: 's', v: th.textContent };
    const cellAddress = XLSX.utils.encode_cell({ r: 0, c: columnIndex });
    worksheet[cellAddress] = cellObject;
    });

    // Get data rows (excluding the first row)
    const dataRows = table.querySelectorAll('tbody tr:not(:first-child)');
    dataRows.forEach(function (row, rowIndex) {
    row.querySelectorAll('td').forEach(function (cell, columnIndex) {
        const dropdown = cell.querySelector('select');
        const value = dropdown ? dropdown.value === 'true' : cell.textContent;
        const cellObject = { t: dropdown ? 'b' : 's', v: value, w: value ? 'TRUE' : 'FALSE' };
        const cellAddress = XLSX.utils.encode_cell({ r: rowIndex + 1, c: columnIndex });
        worksheet[cellAddress] = cellObject;
    });
    });

    // Process dropdown in the 5th column
    const dropdowns = table.querySelectorAll('tbody td:nth-child(5) select');
    dropdowns.forEach(function (dropdown, rowIndex) {
        const cell = dropdown.parentNode;
        const columnIndex = 4; // Index of the 5th column (0-based)
        const value = dropdown.value === 'true';
        const cellObject = { t: 'b', v: value, w: value ? 'TRUE' : 'FALSE' };
        const cellAddress = XLSX.utils.encode_cell({ r: rowIndex + 1, c: columnIndex });
        worksheet[cellAddress] = cellObject;
    });

    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);

    var originalString = file.name;
    var newDate = getCurrentTime();

    // Extract the date part from the original string
    var dateRegex = /(\d{8})(\d{6})/;
    var match = originalString.match(dateRegex);
    var originalDate = match[1] + match[2];

    // Replace the date with the new date
    var replacedString = originalString.replace(dateRegex, newDate);

    link.download = prompt('File name', replacedString);

    link.click();
}

// exit without saving any change
function handleOrgFileCancel() {
    onlyShowPage(1);
}