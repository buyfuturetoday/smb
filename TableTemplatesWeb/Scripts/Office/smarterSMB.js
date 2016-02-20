// Add any initialization logic to this function.
/*
Office.initialize = function (reason) {

    // Checks for the DOM to load.
    $(document).ready(function () {
        $("#getDatabtn").click(function () { getData("selectedData"); });

        // Checks if setSelectedDataAsync is supported and adds appropriate click handler
        if (Office.context.document.setSelectedDataAsync) {
            $("#setDatabtn").click(function () { setData("Sample data"); });
        }
        else {
            $("#setDatabtn").remove();
        }
    });
}
*/

// Writes data to current selection.
function setData(dataToInsert) {
    Office.context.document.setSelectedDataAsync(dataToInsert);
}
/*
// Reads data from current selection.
function getData(elementIdToUpdate) {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
    function (result) {
        if (result.status == "succeeded") {
            document.getElementById(elementIdToUpdate).value = result.value;
        }
    });
}
*/
function writeToPage(text) {
    document.getElementById('outputs').innerText = text;
}
function bindData() {
    Office.context.document.bindings.addFromNamedItemAsync("Leads!Table1", "table", { id: 'bindingdata' }, function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        } else {
            writeToPage('Added binding with type: ' + asyncResult.value.type + ' and id: ' +
                asyncResult.value.id);
        }
    });
}

function readBoundData() {
    Office.select("bindings#bindingdata").getDataAsync({ coercionType: "text" },
        function (asyncResult) {
            if (asyncResult.status == "failed") {
                writeToPage('Error: ' + asyncResult.error.message);
            } else {
                writeToPage('Selected data: ' + asyncResult.value);
            }
            showBindingRowCount();
            //deleteAllRowsFromTable();
        });
}
function addEvent() {
    Office.select("bindings#bindingdata").addHandlerAsync("bindingDataChanged", myHandler, function (asyncResult) {
        if (asyncResult.status == "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        } else {
            writeToPage('Added event handler');
        }
    });
}
function myHandler(eventArgs) {
    eventArgs.binding.getDataAsync({ coerciontype: "table" }, function (asyncResult) {

        if (asyncResult.status == "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        } else {
            writeToPage('Bound data: ' + asyncResult.value);
        }
    });
}

function deleteAllRowsFromTable() {
    Office.context.document.bindings.getByIdAsync("bindingdata", function (asyncResult) {
        var binding = asyncResult.value;
        binding.deleteAllDataValuesAsync();
    });
}

function showBindingRowCount() {
    Office.context.document.bindings.getByIdAsync("bindingdata", function (asyncResult) {
        writeToPage("Rows: " + asyncResult.value.rowCount);
    });
}
///////////////////////////////// Working Code Starts here /////////////////////////
// Writes data to current selection.
function setData(tableType, data) {
    Office.context.document.setSelectedDataAsync(createTablefromTemplate(tableType, data));
    // Office.context.document.setSelectedDataAsync(RefreshStatusSheet(tableType,data));

}

// Writes data to current selection.
function setDataStatus(tableType) {
    Office.context.document.bindings.addFromNamedItemAsync("Sheet2!Table1", createTablefromTemplate(tableType));
    //Sheet1!$1:$1048576
}

// Reads data from current selection.
function getData(elementIdToUpdate) {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
    function (result) {
        if (result.status == "succeeded") {
            document.getElementById(elementIdToUpdate).value = result.value;
        }
    });
}




function bindTableForStatusSheet() {
    /* Select the table from the previous step and click Run Code. 
This will create a binding to the table */

    //Bind to the table in the document from user current selection
    Office.context.document.bindings.addFromSelectionAsync(
//Office.context.document.bindings.addFromNamedItemAsync("Sheet1",
  Office.BindingType.Table,
  { id: "StatusSheetBinding" },
  function (asyncResult) {
      if (asyncResult.status == "failed") {
          app.showNotification("Action failed with error: " + asyncResult.error.message);
      } else {
          $("#bindStatusTableButton").css("display", "none");

          app.showNotification("Added new binding with type: " + asyncResult.value.type +
          " and id: " + asyncResult.value.id);
      }
  });
}


function RefreshStatusSheet(tableType, data) {
    /* Click Run Code to replace the 3rd row with new data */

    //Creating the row to update 
    //var table = new Office.TableData();
    //table.rows = ["Seattle", "WA"], ["Seattle", "WA"];
    var rowToUpdate = 0;
    //app.showNotification("Hello",data.sync.length);
    var table = new Office.TableData();
    var len = data.sync.length;
    var incr = 0;
    var rowsdata;

    //table.headers = [['Sl no', 'Company Name', 'Contact Person', 'Designation ', 'Contact Number', 'Email ID', 'Type of Call', 'Total Call duration (Hrs mts)', 'Status', 'Called on', 'Time', 'Next Meeting', 'Time', 'Listen to Call Recordings', 'Sales Personal', 'Address', 'Remarks']];
    // table.rows = [[data.sync[0]['slno'], data.sync[0]['Company Name'], data.sync[0]['Contact Person'], data.sync[0]['Designation'], data.sync[0]['Contact Number'], data.sync[0]['Email ID'], data.sync[0]['Type of Call'], data.sync[0]['Total Call duration (Hrs mts)'], data.sync[0]['Status'], data.sync[0]['Called on'], data.sync[0]['Time'], data.sync[0]['Next Meeting'], data.sync[0]['Time'], data.sync[0]['Listen to Call Recordings'], data.sync[0]['Sales Personal'], data.sync[0]['Address'], data.sync[0]['Remarks']], [data.sync[0]['slno'], data.sync[0]['Company Name'], data.sync[0]['Contact Person'], data.sync[0]['Designation'], data.sync[0]['Contact Number'], data.sync[0]['Email ID'], data.sync[0]['Type of Call'], data.sync[0]['Total Call duration (Hrs mts)'], data.sync[0]['Status'], data.sync[0]['Called on'], data.sync[0]['Time'], data.sync[0]['Next Meeting'], data.sync[0]['Time'], data.sync[0]['Listen to Call Recordings'], data.sync[0]['Sales Personal'], data.sync[0]['Address'], data.sync[0]['Remarks']], [data.sync[0]['slno'], data.sync[0]['Company Name'], data.sync[0]['Contact Person'], data.sync[0]['Designation'], data.sync[0]['Contact Number'], data.sync[0]['Email ID'], data.sync[0]['Type of Call'], data.sync[0]['Total Call duration (Hrs mts)'], data.sync[0]['Status'], data.sync[0]['Called on'], data.sync[0]['Time'], data.sync[0]['Next Meeting'], data.sync[0]['Time'], data.sync[0]['Listen to Call Recordings'], data.sync[0]['Sales Personal'], data.sync[0]['Address'], data.sync[0]['Remarks']], [data.sync[0]['slno'], data.sync[0]['Company Name'], data.sync[0]['Contact Person'], data.sync[0]['Designation'], data.sync[0]['Contact Number'], data.sync[0]['Email ID'], data.sync[0]['Type of Call'], data.sync[0]['Total Call duration (Hrs mts)'], data.sync[0]['Status'], data.sync[0]['Called on'], data.sync[0]['Time'], data.sync[0]['Next Meeting'], data.sync[0]['Time'], data.sync[0]['Listen to Call Recordings'], data.sync[0]['Sales Personal'], data.sync[0]['Address'], data.sync[0]['Remarks']]];
    //  table.rows = [[data.sync[0]['slno'], data.sync[0]['Company Name'], data.sync[0]['Contact Person'], data.sync[0]['Designation'], data.sync[0]['Contact Number'], data.sync[0]['Email ID'], data.sync[0]['Type of Call'], data.sync[0]['Total Call duration (Hrs mts)'], data.sync[0]['Status'], data.sync[0]['Called on'], data.sync[0]['Time'], data.sync[0]['Next Meeting'], data.sync[0]['Time'], data.sync[0]['Listen to Call Recordings'], data.sync[0]['Sales Personal'], data.sync[0]['Address'], data.sync[0]['Remarks']]];
    //table.rows = [[data.sync[1]['slno'], data.sync[1]['Company Name'], data.sync[1]['Contact Person'], data.sync[1]['Designation'], data.sync[1]['Contact Number'], data.sync[1]['Email ID'], data.sync[1]['Type of Call'], data.sync[1]['Total Call duration (Hrs mts)'], data.sync[1]['Status'], data.sync[1]['Called on'], data.sync[1]['Time'], data.sync[1]['Next Meeting'], data.sync[1]['Time'], data.sync[1]['Listen to Call Recordings'], data.sync[1]['Sales Personal'], data.sync[1]['Address'], data.sync[1]['Remarks']]];
    //table.rows = [[data.sync[2]['slno'], data.sync[2]['Company Name'], data.sync[2]['Contact Person'], data.sync[2]['Designation'], data.sync[2]['Contact Number'], data.sync[2]['Email ID'], data.sync[2]['Type of Call'], data.sync[2]['Total Call duration (Hrs mts)'], data.sync[2]['Status'], data.sync[2]['Called on'], data.sync[2]['Time'], data.sync[2]['Next Meeting'], data.sync[2]['Time'], data.sync[2]['Listen to Call Recordings'], data.sync[2]['Sales Personal'], data.sync[2]['Address'], data.sync[2]['Remarks']]];
    //table.rows = [[data.sync[3]['slno'], data.sync[3]['Company Name'], data.sync[3]['Contact Person'], data.sync[3]['Designation'], data.sync[3]['Contact Number'], data.sync[3]['Email ID'], data.sync[3]['Type of Call'], data.sync[3]['Total Call duration (Hrs mts)'], data.sync[3]['Status'], data.sync[3]['Called on'], data.sync[3]['Time'], data.sync[3]['Next Meeting'], data.sync[3]['Time'], data.sync[3]['Listen to Call Recordings'], data.sync[3]['Sales Personal'], data.sync[3]['Address'], data.sync[3]['Remarks']]];

    // return table;

    //app.showNotification("Status Sheet rows." + len);
    for (incr = 0; incr < len; incr++) {
        var table = new Office.TableData();

        table.rows = [[data.sync[incr]['slno'], data.sync[incr]['Company Name'], data.sync[incr]['Contact Person'], data.sync[incr]['Designation'], data.sync[incr]['Contact Number'], data.sync[incr]['Email ID'], data.sync[incr]['Type of Call'], data.sync[incr]['Total Call duration (Hrs mts)'], data.sync[incr]['Status'], data.sync[incr]['Called on'], data.sync[incr]['Time'], data.sync[incr]['Next Meeting'], data.sync[incr]['Time'], data.sync[incr]['Listen to Call Recordings'], data.sync[incr]['Sales Personal'], data.sync[incr]['Address'], data.sync[incr]['Remarks']]];
        // table.rows = [[data.sync[incr]['slno']]];

        //Getting the table binding and setting data in 3rd row.
        Office.select("bindings#StatusSheetBinding", onBindingNotFound).setDataAsync(
          table,
          { coercionType: Office.CoercionType.Table, startRow: rowToUpdate },
          function (asyncResult) {
              if (asyncResult.status == "failed") {
                  app.showNotification("Action failed with error: " + asyncResult.error.message);
              } else {
                  $("#bindStatusTableButton").css("display", "none");

                  app.showNotification("Status Sheet is updated with the latest records.");
              }
          }
        );

        //Show error message in case the binding object wasn't found
        function onBindingNotFound() {
            $("#bindStatusTableButton").css("display", "block");
            app.showNotification("The binding object was not found. " +
            "Please return to previous step to create the binding");
        }
        rowToUpdate++;
    }
    /*
    //Getting the table binding and setting data in 3rd row.
    Office.select("bindings#AgentBinding", onBindingNotFound).setDataAsync(
      table,
      { coercionType: Office.CoercionType.Table, startRow: rowToUpdate },
      function (asyncResult) {
          if (asyncResult.status == "failed") {
              app.showNotification("Action failed with error: " + asyncResult.error.message);
          } else {
              $("#bindStatusTableButton").css("display", "none");

              app.showNotification("Status Sheet is updated with the latest records.");
          }
      }
    );

    //Show error message in case the binding object wasn't found
    function onBindingNotFound() {
        $("#bindStatusTableButton").css("display", "block");
        app.showNotification("The binding object was not found. " +
        "Please return to previous step to create the binding");
    }*/
}


function RefreshStatusSheet_org(tableType, data) {
    /* Click Run Code to replace the 3rd row with new data */

    //Creating the row to update 
    //var table = new Office.TableData();
    //table.rows = ["Seattle", "WA"], ["Seattle", "WA"];
    var rowToUpdate = 1;

    var table = new Office.TableData();
    var len = data.length;
    var incr = 0;
    var rowsdata;

    table.headers = [['Sl no', 'Company Name', 'Contact Person', 'Designation ', 'Contact Number', 'Email ID', 'Type of Call', 'Total Call duration (Hrs mts)', 'Status', 'Called on', 'Time', 'Next Meeting', 'Time', 'Listen to Call Recordings', 'Sales Personal', 'Address', 'Remarks']];
    //   table.rows = [[data.sync[0]['slno'], data.sync[0]['Company Name'], data.sync[0]['Contact Person'], data.sync[0]['Designation'], data.sync[0]['Contact Number'], data.sync[0]['Email ID'], data.sync[0]['Type of Call'], data.sync[0]['Total Call duration (Hrs mts)'], data.sync[0]['Status'], data.sync[0]['Called on'], data.sync[0]['Time'], data.sync[0]['Next Meeting'], data.sync[0]['Time'], data.sync[0]['Listen to Call Recordings'], data.sync[0]['Sales Personal'], data.sync[0]['Address'], data.sync[0]['Remarks']], [data.sync[0]['slno'], data.sync[0]['Company Name'], data.sync[0]['Contact Person'], data.sync[0]['Designation'], data.sync[0]['Contact Number'], data.sync[0]['Email ID'], data.sync[0]['Type of Call'], data.sync[0]['Total Call duration (Hrs mts)'], data.sync[0]['Status'], data.sync[0]['Called on'], data.sync[0]['Time'], data.sync[0]['Next Meeting'], data.sync[0]['Time'], data.sync[0]['Listen to Call Recordings'], data.sync[0]['Sales Personal'], data.sync[0]['Address'], data.sync[0]['Remarks']], [data.sync[0]['slno'], data.sync[0]['Company Name'], data.sync[0]['Contact Person'], data.sync[0]['Designation'], data.sync[0]['Contact Number'], data.sync[0]['Email ID'], data.sync[0]['Type of Call'], data.sync[0]['Total Call duration (Hrs mts)'], data.sync[0]['Status'], data.sync[0]['Called on'], data.sync[0]['Time'], data.sync[0]['Next Meeting'], data.sync[0]['Time'], data.sync[0]['Listen to Call Recordings'], data.sync[0]['Sales Personal'], data.sync[0]['Address'], data.sync[0]['Remarks']], [data.sync[0]['slno'], data.sync[0]['Company Name'], data.sync[0]['Contact Person'], data.sync[0]['Designation'], data.sync[0]['Contact Number'], data.sync[0]['Email ID'], data.sync[0]['Type of Call'], data.sync[0]['Total Call duration (Hrs mts)'], data.sync[0]['Status'], data.sync[0]['Called on'], data.sync[0]['Time'], data.sync[0]['Next Meeting'], data.sync[0]['Time'], data.sync[0]['Listen to Call Recordings'], data.sync[0]['Sales Personal'], data.sync[0]['Address'], data.sync[0]['Remarks']]];
    // return table;
    for (incr = 0; incr < len; incr++) {
        rowsdata += [data.sync[incr]['slno'], data.sync[incr]['Company Name'], data.sync[incr]['Contact Person'], data.sync[incr]['Designation'], data.sync[incr]['Contact Number'], data.sync[incr]['Email ID'], data.sync[incr]['Type of Call'], data.sync[incr]['Total Call duration (Hrs mts)'], data.sync[incr]['Status'], data.sync[incr]['Called on'], data.sync[incr]['Time'], data.sync[incr]['Next Meeting'], data.sync[incr]['Time'], data.sync[incr]['Listen to Call Recordings'], data.sync[incr]['Sales Personal'], data.sync[incr]['Address'], data.sync[incr]['Remarks']] + ',';

    }
    rowsdata += [data.sync[incr]['slno'], data.sync[incr]['Company Name'], data.sync[incr]['Contact Person'], data.sync[incr]['Designation'], data.sync[incr]['Contact Number'], data.sync[incr]['Email ID'], data.sync[incr]['Type of Call'], data.sync[incr]['Total Call duration (Hrs mts)'], data.sync[incr]['Status'], data.sync[incr]['Called on'], data.sync[incr]['Time'], data.sync[incr]['Next Meeting'], data.sync[incr]['Time'], data.sync[incr]['Listen to Call Recordings'], data.sync[incr]['Sales Personal'], data.sync[incr]['Address'], data.sync[incr]['Remarks']];
    //rowsdata = rowsdata.slice(0, -1);
    table.rows = [rowsdata];

    //Getting the table binding and setting data in 3rd row.
    Office.select("bindings#StatusSheetBinding", onBindingNotFound).setDataAsync(
      table,
      { coercionType: Office.CoercionType.Table, startRow: rowToUpdate },
      function (asyncResult) {
          if (asyncResult.status == "failed") {
              app.showNotification("Action failed with error: " + asyncResult.error.message);
          } else {
              $("#bindStatusTableButton").css("display", "none");

              app.showNotification("Status Sheet is updated with the latest records.");
          }
      }
    );

    //Show error message in case the binding object wasn't found
    function onBindingNotFound() {
        $("#bindStatusTableButton").css("display", "block");
        app.showNotification("The binding object was not found. " +
        "Please return to previous step to create the binding");
    }
}
function createTablefromTemplate(tableType, data) {
    var tableData = new Office.TableData();
    switch (tableType) {
        case "task":
            tableData.headers = [['Sl no', 'Company Name', 'Contact Person', 'Designation ', 'Contact Number', 'Email ID', 'Type of Call', 'Total Call duration (Hrs mts)', 'Status', 'Called on', 'Time', 'Next Meeting', 'Time', 'Listen to Call Recordings', 'Sales Personal', 'Address', 'Remarks']];
            tableData.rows = [[data.sync[0]['slno'], data.sync[0]['Company Name'], data.sync[0]['Contact Person'], data.sync[0]['Designation'], data.sync[0]['Contact Number'], data.sync[0]['Email ID'], data.sync[0]['Type of Call'], data.sync[0]['Total Call duration (Hrs mts)'], data.sync[0]['Status'], data.sync[0]['Called on'], data.sync[0]['Time'], data.sync[0]['Next Meeting'], data.sync[0]['Time'], data.sync[0]['Listen to Call Recordings'], data.sync[0]['Sales Personal'], data.sync[0]['Address'], data.sync[0]['Remarks']]];
            return tableData;
            break;

        case "task1":
            tableData.headers = [['Task Title', 'Start Date', 'Due Date', 'Status', 'Category', 'Assigned To']];
            tableData.rows = [['Sample Task #1', '=Today()', '=Today()+1', 'Not Started', 'Work', 'Me'], ['Sample Task #2', '=Today()', '=Today()', 'Not Started', 'Work', 'Me'], ['Sample Task #3', '=Today()', '=Today()', 'Not Started', 'Work', 'Me']];
            return tableData;
            break;

        case "contact":
            tableData.headers = [['First Name', 'Last Name', 'Title', 'Phone Number', 'Email', 'Street', 'City', 'State', 'Zip']];
            tableData.rows = [['John', 'Doe', 'Generic Person', '212-555-1212', 'john@doe.com', '1234 Main', 'Oakland', 'CA', '94601']];
            return tableData;
            break;

        case "expense":
            tableData.headers = [['Expense Date', 'Category', 'Client/Project', 'Amount']];
            tableData.rows = [['=Today()', 'Meals', 'Custom App for Office', '$55.00']];
            return tableData;
            break;

        case "issue":
            tableData.headers = [['Issue', 'Status', 'Priority', 'Area', 'Due Date', 'Date Resoled', 'Notes']];
            tableData.rows = [['Severely limited Office.js API', 'In-Progress', 'Normal', 'Development', '=today()+7', '', 'I wish the API was more extensive.']];
            return tableData;
            break;


    }
}


function bindTablebyName() {
    /* Select the table from the previous step and click Run Code. 
This will create a binding to the table */

    //Bind to the table in the document from user current selection
    Office.context.document.bindings.addFromSelectionAsync(
//Office.context.document.bindings.addFromNamedItemAsync("Sheet1",
  Office.BindingType.Table,
  { id: "MyTableBinding" },
  function (asyncResult) {
      if (asyncResult.status == "failed") {
          app.showNotification("Action failed with error: " + asyncResult.error.message);
      } else {
          $("#bindLeadsTableButton").css("display", "none");

          app.showNotification("Added new binding with type: " + asyncResult.value.type +
          " and id: " + asyncResult.value.id);
      }
  });
}



function readDataExcel() {
    var table = new Office.TableData();

    //Office.context.document.getSelectedDataAsync(
    Office.select("bindings#MyTableBinding", onBindingNotFound).getSelectedDataAsync(
       Office.CoercionType.Matrix,
       function (asyncResult) {
           if (asyncResult.status == "failed") {
               app.showNotification("Action failed with error: " + asyncResult.error.message);
           } else {
               app.showNotification("Selected data: " + asyncResult.value);
           }
       }
       );

}
function updateColumn() {

    var randomh = Math.random();

    $.ajax({
        type: "GET",
        //url: "http://buyfuturetoday.com/smb/SyncSalesAgent.php?email=" + userName,
        //url: "http://localhost:81/demo/check_login.json",
        url: "http://localhost:81/demo/sync_lead.json?x=" + randomh + "&email=" + userName + "",
        dataType: "json",
        success: function (data) {
            //syncAgents(data);
            //setData("task", data);
            updateLeadSheet("task", data);
            //setDataStatus("task");
        },
        error: function (xhr) {
            app.showNotification('Error:', xhr.responseText);
        }
    });

}
function updateLeadSheet(tableType, data) {

    /* Click Run Code to replace the 3rd row with new data */

    //Creating the row to update 
    //var table = new Office.TableData();
    //table.rows = ["Seattle", "WA"], ["Seattle", "WA"];
    var rowToUpdate = 1;
    //app.showNotification("Hello",data.sync.length);
    var table = new Office.TableData();
    var len = data.sync.length;
    var incr = 0;
    var rowsdata;

    //table.headers = [['Sl no', 'Company Name', 'Contact Person', 'Designation ', 'Contact Number', 'Email ID', 'Type of Call', 'Total Call duration (Hrs mts)', 'Status', 'Called on', 'Time', 'Next Meeting', 'Time', 'Listen to Call Recordings', 'Sales Personal', 'Address', 'Remarks']];
    // table.rows = [[data.sync[0]['slno'], data.sync[0]['Company Name'], data.sync[0]['Contact Person'], data.sync[0]['Designation'], data.sync[0]['Contact Number'], data.sync[0]['Email ID'], data.sync[0]['Type of Call'], data.sync[0]['Total Call duration (Hrs mts)'], data.sync[0]['Status'], data.sync[0]['Called on'], data.sync[0]['Time'], data.sync[0]['Next Meeting'], data.sync[0]['Time'], data.sync[0]['Listen to Call Recordings'], data.sync[0]['Sales Personal'], data.sync[0]['Address'], data.sync[0]['Remarks']], [data.sync[0]['slno'], data.sync[0]['Company Name'], data.sync[0]['Contact Person'], data.sync[0]['Designation'], data.sync[0]['Contact Number'], data.sync[0]['Email ID'], data.sync[0]['Type of Call'], data.sync[0]['Total Call duration (Hrs mts)'], data.sync[0]['Status'], data.sync[0]['Called on'], data.sync[0]['Time'], data.sync[0]['Next Meeting'], data.sync[0]['Time'], data.sync[0]['Listen to Call Recordings'], data.sync[0]['Sales Personal'], data.sync[0]['Address'], data.sync[0]['Remarks']], [data.sync[0]['slno'], data.sync[0]['Company Name'], data.sync[0]['Contact Person'], data.sync[0]['Designation'], data.sync[0]['Contact Number'], data.sync[0]['Email ID'], data.sync[0]['Type of Call'], data.sync[0]['Total Call duration (Hrs mts)'], data.sync[0]['Status'], data.sync[0]['Called on'], data.sync[0]['Time'], data.sync[0]['Next Meeting'], data.sync[0]['Time'], data.sync[0]['Listen to Call Recordings'], data.sync[0]['Sales Personal'], data.sync[0]['Address'], data.sync[0]['Remarks']], [data.sync[0]['slno'], data.sync[0]['Company Name'], data.sync[0]['Contact Person'], data.sync[0]['Designation'], data.sync[0]['Contact Number'], data.sync[0]['Email ID'], data.sync[0]['Type of Call'], data.sync[0]['Total Call duration (Hrs mts)'], data.sync[0]['Status'], data.sync[0]['Called on'], data.sync[0]['Time'], data.sync[0]['Next Meeting'], data.sync[0]['Time'], data.sync[0]['Listen to Call Recordings'], data.sync[0]['Sales Personal'], data.sync[0]['Address'], data.sync[0]['Remarks']]];
    //  table.rows = [[data.sync[0]['slno'], data.sync[0]['Company Name'], data.sync[0]['Contact Person'], data.sync[0]['Designation'], data.sync[0]['Contact Number'], data.sync[0]['Email ID'], data.sync[0]['Type of Call'], data.sync[0]['Total Call duration (Hrs mts)'], data.sync[0]['Status'], data.sync[0]['Called on'], data.sync[0]['Time'], data.sync[0]['Next Meeting'], data.sync[0]['Time'], data.sync[0]['Listen to Call Recordings'], data.sync[0]['Sales Personal'], data.sync[0]['Address'], data.sync[0]['Remarks']]];
    //table.rows = [[data.sync[1]['slno'], data.sync[1]['Company Name'], data.sync[1]['Contact Person'], data.sync[1]['Designation'], data.sync[1]['Contact Number'], data.sync[1]['Email ID'], data.sync[1]['Type of Call'], data.sync[1]['Total Call duration (Hrs mts)'], data.sync[1]['Status'], data.sync[1]['Called on'], data.sync[1]['Time'], data.sync[1]['Next Meeting'], data.sync[1]['Time'], data.sync[1]['Listen to Call Recordings'], data.sync[1]['Sales Personal'], data.sync[1]['Address'], data.sync[1]['Remarks']]];
    //table.rows = [[data.sync[2]['slno'], data.sync[2]['Company Name'], data.sync[2]['Contact Person'], data.sync[2]['Designation'], data.sync[2]['Contact Number'], data.sync[2]['Email ID'], data.sync[2]['Type of Call'], data.sync[2]['Total Call duration (Hrs mts)'], data.sync[2]['Status'], data.sync[2]['Called on'], data.sync[2]['Time'], data.sync[2]['Next Meeting'], data.sync[2]['Time'], data.sync[2]['Listen to Call Recordings'], data.sync[2]['Sales Personal'], data.sync[2]['Address'], data.sync[2]['Remarks']]];
    //table.rows = [[data.sync[3]['slno'], data.sync[3]['Company Name'], data.sync[3]['Contact Person'], data.sync[3]['Designation'], data.sync[3]['Contact Number'], data.sync[3]['Email ID'], data.sync[3]['Type of Call'], data.sync[3]['Total Call duration (Hrs mts)'], data.sync[3]['Status'], data.sync[3]['Called on'], data.sync[3]['Time'], data.sync[3]['Next Meeting'], data.sync[3]['Time'], data.sync[3]['Listen to Call Recordings'], data.sync[3]['Sales Personal'], data.sync[3]['Address'], data.sync[3]['Remarks']]];

    // return table;

    //app.showNotification("Status Sheet rows." + len);
    for (incr = 0; incr < len; incr++) {
        var table = new Office.TableData();

        table.rows = [[data.sync[incr]['slno'], data.sync[incr]['Company Name'], data.sync[incr]['Designation'], data.sync[incr]['Contact Number'], data.sync[incr]['Email id'], data.sync[incr]['Status'], data.sync[incr]['Date to Be called'], data.sync[incr]['Concern sales agent'], data.sync[incr]['Location'], data.sync[incr]['Submit'], data.sync[incr]['Submit']]];
        // table.rows = [[data.sync[incr]['slno']]];

        //Getting the table binding and setting data in 3rd row.
        Office.select("bindings#MyTableBinding", onBindingNotFound).setDataAsync(
          table,
          { coercionType: Office.CoercionType.Table, startRow: rowToUpdate },
          function (asyncResult) {
              if (asyncResult.status == "failed") {
                  app.showNotification("Action failed with error: " + asyncResult.error.message);
              } else {
                  $("#bindLeadsTableButton").css("display", "none");

                  app.showNotification("Leads Sheet is updated with the latest records.");
              }
          }
        );

        //Show error message in case the binding object wasn't found
        function onBindingNotFound() {
            $("#bindLeadsTableButton").css("display", "block");
            app.showNotification("The binding object was not found. " +
            "Please return to previous step to create the binding");
        }
        rowToUpdate++;
    }
}

function updateColumn_old() {
    /* Click Run Code to replace the 3rd row with new data */

    //Creating the row to update 
    var table = new Office.TableData();
    table.headers = ['Sl no', 'Company Name', 'Contact Person', 'Designation', 'Email ID', 'Status', 'Date to Be called', 'Concern sales agent', 'Location', 'Submit'];
    //table.rows = ["Seattle", "WA"], ["Seattle", "WA"];
    table.rows = [["1", "Moserbaer India Ltd  ( Entertainment Division )", "Sharath", "Sales Promotion Mgr", "", "Cold Call", "30-10-2015", "Raghu", "", "submit"],
["2", "LPS Bossard Pvt Ltd", "Prashant R Salikeri", "Sales / BD Mgr", "", "Contacted", "31-10-2015", "Vijay", "", "submit"],
["3", "Tyco Electronics  -  M / A - COM Division", "Madhusudhan.H.R", "Head / VP / GM / National Mgr -Sales", "", "Not Contacted", "01-11-2015", "Ravi", "", "submit"],
["4", "FE Global India Pvt  LTD", "Vikas Kaul", "Sales / BD Mgr", "", "Not Interested", "02-11-2015", "Shree", "", "submit"],
["5", "Enventure Technologies Inc", "Raghavendra . L", "Head / VP / GM / National Mgr -Sales", "", "Junk", "03-11-2015", "Donald", "", "submit"],
["6", "Hindustan Lever Ltd", "Alexander George", "Sales / BD Mgr", "", "Need Analysis", "04-11-2015", "Micheal", "", "submit"],
["7", "General Logistics System A / S ,  Denmark", "Nataraj", "Sales / BD Mgr", "", "Value Proposition ", "05-11-2015", "Madhu", "", "submit"],
["8", "Hewlett - Packard India Sales Pvt Ltd", "Philip", "Sales / BD Mgr", "", "Decision Maker", "06-11-2015", "Vijay", "", "submit"],
["9", "IGARASHI MOTORS SALES P .  Ltd .", "Asgar Ali", "Sales / BD Mgr", "", "Negotiation", "07-11-2015", "Ravi", "", "submit"],
["10", "ICICI Bank Ltd", "Vincent", "Sales Exec. / Officer", "", "Purchase Order", "08-11-2015", "Shree", "", "submit"],
["11", "Deutsche Bank AG .", "Ajay Bhamare", "Sales Head", "", "Invoice ", "09-11-2015", "Donald", "", "submit"],
["12", "SNAPON TOOLS PRIVATE LIMITED", "Randhir Kumar", "Sales / BD Mgr", "", "Payment Pending", "10-11-2015", "Micheal", "", "submit"],
["13", "", "Dinesh.B.K", "Sales / BD Mgr", "", "Partial Payment ", "11-11-2015", "Madhu", "", "submit"],
["14", "Programming Research Software Technology", "Madhu G Rao", "Sales / BD Mgr", "", "Customer Complain", "12-11-2015", "Vijay", "", "submit"],
["15", "", "Gautom K Das", "Sales / BD Mgr", "", "", "13-11-2015", "Ravi", "", "submit"]];

    var rowToUpdate = 1;

    //Getting the table binding and setting data in 3rd row.
    Office.select("bindings#MyTableBinding", onBindingNotFound).setDataAsync(
      table,
      { coercionType: Office.CoercionType.Table, startRow: rowToUpdate },
      function (asyncResult) {
          if (asyncResult.status == "failed") {
              app.showNotification("Action failed with error: " + asyncResult.error.message);
          } else {
              $("#bindLeadsTableButton").css("display", "none");

              app.showNotification("Successfully submitted to server.");
          }
      }
    );

    //Show error message in case the binding object wasn't found
    function onBindingNotFound() {
        $("#bindLeadsTableButton").css("display", "block");

        app.showNotification("The binding object was not found. " +
        "Please return to previous step to create the binding");
    }
}

function AgentsRowBindTable() {
    /* Select the table from the previous step and click Run Code. 
This will create a binding to the table */

    //Bind to the table in the document from user current selection
    Office.context.document.bindings.addFromSelectionAsync(
//Office.context.document.bindings.addFromNamedItemAsync("Sheet1",
  Office.BindingType.Table,
  { id: "AgentBinding" },
  function (asyncResult) {
      if (asyncResult.status == "failed") {
          app.showNotification("Action failed with error: " + asyncResult.error.message);
      } else {
          $("#AgentsButton").css("display", "none");

          app.showNotification("Added new binding with type: " + asyncResult.value.type +
          " and id: " + asyncResult.value.id);
      }
  });

}


function SyncAgentStatusSheet(tableType, data) {
    /* Click Run Code to replace the 3rd row with new data */

    //Creating the row to update 
    //var table = new Office.TableData();
    //table.rows = ["Seattle", "WA"], ["Seattle", "WA"];
    var rowToUpdate = 0;
    //app.showNotification("Hello",data.sync.length);
    var table = new Office.TableData();
    var len = data.sync.length;
    var incr = 0;
    var rowsdata;

    //table.headers = [['Sl no', 'Company Name', 'Contact Person', 'Designation ', 'Contact Number', 'Email ID', 'Type of Call', 'Total Call duration (Hrs mts)', 'Status', 'Called on', 'Time', 'Next Meeting', 'Time', 'Listen to Call Recordings', 'Sales Personal', 'Address', 'Remarks']];
    // table.rows = [[data.sync[0]['slno'], data.sync[0]['Company Name'], data.sync[0]['Contact Person'], data.sync[0]['Designation'], data.sync[0]['Contact Number'], data.sync[0]['Email ID'], data.sync[0]['Type of Call'], data.sync[0]['Total Call duration (Hrs mts)'], data.sync[0]['Status'], data.sync[0]['Called on'], data.sync[0]['Time'], data.sync[0]['Next Meeting'], data.sync[0]['Time'], data.sync[0]['Listen to Call Recordings'], data.sync[0]['Sales Personal'], data.sync[0]['Address'], data.sync[0]['Remarks']], [data.sync[0]['slno'], data.sync[0]['Company Name'], data.sync[0]['Contact Person'], data.sync[0]['Designation'], data.sync[0]['Contact Number'], data.sync[0]['Email ID'], data.sync[0]['Type of Call'], data.sync[0]['Total Call duration (Hrs mts)'], data.sync[0]['Status'], data.sync[0]['Called on'], data.sync[0]['Time'], data.sync[0]['Next Meeting'], data.sync[0]['Time'], data.sync[0]['Listen to Call Recordings'], data.sync[0]['Sales Personal'], data.sync[0]['Address'], data.sync[0]['Remarks']], [data.sync[0]['slno'], data.sync[0]['Company Name'], data.sync[0]['Contact Person'], data.sync[0]['Designation'], data.sync[0]['Contact Number'], data.sync[0]['Email ID'], data.sync[0]['Type of Call'], data.sync[0]['Total Call duration (Hrs mts)'], data.sync[0]['Status'], data.sync[0]['Called on'], data.sync[0]['Time'], data.sync[0]['Next Meeting'], data.sync[0]['Time'], data.sync[0]['Listen to Call Recordings'], data.sync[0]['Sales Personal'], data.sync[0]['Address'], data.sync[0]['Remarks']], [data.sync[0]['slno'], data.sync[0]['Company Name'], data.sync[0]['Contact Person'], data.sync[0]['Designation'], data.sync[0]['Contact Number'], data.sync[0]['Email ID'], data.sync[0]['Type of Call'], data.sync[0]['Total Call duration (Hrs mts)'], data.sync[0]['Status'], data.sync[0]['Called on'], data.sync[0]['Time'], data.sync[0]['Next Meeting'], data.sync[0]['Time'], data.sync[0]['Listen to Call Recordings'], data.sync[0]['Sales Personal'], data.sync[0]['Address'], data.sync[0]['Remarks']]];
    //  table.rows = [[data.sync[0]['slno'], data.sync[0]['Company Name'], data.sync[0]['Contact Person'], data.sync[0]['Designation'], data.sync[0]['Contact Number'], data.sync[0]['Email ID'], data.sync[0]['Type of Call'], data.sync[0]['Total Call duration (Hrs mts)'], data.sync[0]['Status'], data.sync[0]['Called on'], data.sync[0]['Time'], data.sync[0]['Next Meeting'], data.sync[0]['Time'], data.sync[0]['Listen to Call Recordings'], data.sync[0]['Sales Personal'], data.sync[0]['Address'], data.sync[0]['Remarks']]];
    //table.rows = [[data.sync[1]['slno'], data.sync[1]['Company Name'], data.sync[1]['Contact Person'], data.sync[1]['Designation'], data.sync[1]['Contact Number'], data.sync[1]['Email ID'], data.sync[1]['Type of Call'], data.sync[1]['Total Call duration (Hrs mts)'], data.sync[1]['Status'], data.sync[1]['Called on'], data.sync[1]['Time'], data.sync[1]['Next Meeting'], data.sync[1]['Time'], data.sync[1]['Listen to Call Recordings'], data.sync[1]['Sales Personal'], data.sync[1]['Address'], data.sync[1]['Remarks']]];
    //table.rows = [[data.sync[2]['slno'], data.sync[2]['Company Name'], data.sync[2]['Contact Person'], data.sync[2]['Designation'], data.sync[2]['Contact Number'], data.sync[2]['Email ID'], data.sync[2]['Type of Call'], data.sync[2]['Total Call duration (Hrs mts)'], data.sync[2]['Status'], data.sync[2]['Called on'], data.sync[2]['Time'], data.sync[2]['Next Meeting'], data.sync[2]['Time'], data.sync[2]['Listen to Call Recordings'], data.sync[2]['Sales Personal'], data.sync[2]['Address'], data.sync[2]['Remarks']]];
    //table.rows = [[data.sync[3]['slno'], data.sync[3]['Company Name'], data.sync[3]['Contact Person'], data.sync[3]['Designation'], data.sync[3]['Contact Number'], data.sync[3]['Email ID'], data.sync[3]['Type of Call'], data.sync[3]['Total Call duration (Hrs mts)'], data.sync[3]['Status'], data.sync[3]['Called on'], data.sync[3]['Time'], data.sync[3]['Next Meeting'], data.sync[3]['Time'], data.sync[3]['Listen to Call Recordings'], data.sync[3]['Sales Personal'], data.sync[3]['Address'], data.sync[3]['Remarks']]];

    // return table;

    app.showNotification("Status Sheet rows." + len);
    for (incr = 0; incr < len; incr++) {
        var table = new Office.TableData();

        // table.rows = [[data.sync[incr]['slno'], data.sync[incr]['Company Name'], data.sync[incr]['Contact Person'], data.sync[incr]['Designation'], data.sync[incr]['Contact Number'], data.sync[incr]['Email ID'], data.sync[incr]['Type of Call'], data.sync[incr]['Total Call duration (Hrs mts)'], data.sync[incr]['Status'], data.sync[incr]['Called on'], data.sync[incr]['Time'], data.sync[incr]['Next Meeting'], data.sync[incr]['Time'], data.sync[incr]['Listen to Call Recordings'], data.sync[incr]['Sales Personal'], data.sync[incr]['Address'], data.sync[incr]['Remarks']]];
        table.rows = [[data.sync[incr]['slno']]];

        //Getting the table binding and setting data in 3rd row.
        Office.select("bindings#AgentBinding", onBindingNotFound).setDataAsync(
          table,
          { coercionType: Office.CoercionType.Table, startRow: rowToUpdate },
          function (asyncResult) {
              if (asyncResult.status == "failed") {
                  app.showNotification("Action failed with error: " + asyncResult.error.message);
              } else {
                  $("#bindStatusTableButton").css("display", "none");

                  app.showNotification("Status Sheet is updated with the latest records.");
              }
          }
        );

        //Show error message in case the binding object wasn't found
        function onBindingNotFound() {
            $("#bindStatusTableButton").css("display", "block");
            app.showNotification("The binding object was not found. " +
            "Please return to previous step to create the binding");
        }
        rowToUpdate++;
    }
    /*
    //Getting the table binding and setting data in 3rd row.
    Office.select("bindings#AgentBinding", onBindingNotFound).setDataAsync(
      table,
      { coercionType: Office.CoercionType.Table, startRow: rowToUpdate },
      function (asyncResult) {
          if (asyncResult.status == "failed") {
              app.showNotification("Action failed with error: " + asyncResult.error.message);
          } else {
              $("#bindStatusTableButton").css("display", "none");

              app.showNotification("Status Sheet is updated with the latest records.");
          }
      }
    );

    //Show error message in case the binding object wasn't found
    function onBindingNotFound() {
        $("#bindStatusTableButton").css("display", "block");
        app.showNotification("The binding object was not found. " +
        "Please return to previous step to create the binding");
    }*/
}


function syncSalesAgent() {
    var randomh = Math.random();
    $.ajax({
        type: "GET",
        //url: "http://buyfuturetoday.com/smb/SyncSalesAgent.php?email=" + userName,
        //url: "http://localhost:81/demo/check_login.json",
        url: "http://localhost:81/demo/sync_agents.json?x=" + randomh + "&email=" + userName + "",
        dataType: "json",
        success: function (data) {
            //syncAgents(data);
            //setData("task", data);
            SyncAgentStatusSheet("task", data);
            //setDataStatus("task");
            app.showNotification('SUccess:');
        },
        error: function (xhr) {
            app.showNotification('Error:', xhr.responseText);
        }
    });
}


function RefreshTable(data) {

    Office.context.document.bindings.getByIdAsync('StatusSheetBinding', function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {

            var binding = asyncResult.value;

            if (binding.hasHeaders == false) {
                debug("I can't refresh the table that doesn't have headers. Turn on the table header and try again.");
            }
            else {
                binding.deleteAllDataValuesAsync(function (deleteResult) {
                    if (deleteResult.status == "succeeded") {
                        binding.addRowsAsync(data, function () {
                            debug("Returned " + data.length + " results.");
                        });

                    }
                });
            }
        }
    });
}