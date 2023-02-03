function preview(input) {
    // Get the input element with the id 'input'
    var input = document.getElementById('input');
  
    // Add an event listener to the input element
    input.addEventListener('change', function () {
      // Read the data from the uploaded XLSX file
      readXlsxFile(input.files[0]).then(function (data) {
        var i = 0;
        // Loop through each row of data
        data.map((row, index) => {
          // If this is the first row, create the table header
          if (i == 0) {
            let table = document.getElementById('tbl-data');
            generateTableHead(table, row);
          }
          // If this is not the first row, add a new row to the table
          if (i > 0) {
            let table = document.getElementById('tbl-data');
            generateTableRows(table, row);
          }
        });
      });
    });
  
    // Function to generate the table header
    function generateTableHead(table, data) {
      let thead = table.createTHead();
      let row = thead.insertRow();
      // Loop through each header column
      for (let key of data) {
        let th = document.createElement('th');
        let text = document.createTextNode(key);
        th.appendChild(text);
        row.appendChild(th);
      }
    }
  
    // Function to generate the table rows
    function generateTableRows(table, data) {
      let newRow = table.insertRow(-1);
      // Loop through each cell in the row
      data.map((row, index) => {
        let newCell = newRow.insertCell();
        let newText = document.createTextNode(row);
        newCell.appendChild(newText);
      });
    }
  }