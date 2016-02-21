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


function p(i) {
    return Math.floor(i / 10) + "" + i % 10;
}
function trunc(i) {
    var j = Math.round(i * 100);
    return Math.floor(j / 100) + (j % 100 > 0 ? "." + p(j % 100) : "");
}
function calculate(date1, date2) {
    var date1 = new Date(date1);
    var date2 = new Date(date2);
    var sec = date2.getTime() - date1.getTime();
    if (isNaN(sec)) {
       // alert("Input data is incorrect!");
        return;
    }
    if (sec < 0) {
        //alert("The second date ocurred earlier than the first one!");
        return;
    }

    var second = 1000, minute = 60 * second, hour = 60 * minute, day = 24 * hour;


    var days = Math.floor(sec / day);
    sec -= days * day;
    var hours = Math.floor(sec / hour);
    sec -= hours * hour;
    var minutes = Math.floor(sec / minute);
    sec -= minutes * minute;
    var seconds = Math.floor(sec / second);
    //var final = days + " day" + (days != 1 ? "s" : "") + ", " + hours + " hour" + (hours != 1 ? "s" : "") + ", " + minutes + " minute" + (minutes != 1 ? "s" : "") + ", " + seconds + " second" + (seconds != 1 ? "s" : "");
    return hours + " : " + minutes + " : " + seconds;
}

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
function loginError(text) {
    document.getElementById('errorBox').innerText = text;
}

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


function RefreshStatusSheetPost(rowid, data) {
    //deleteAllRowsFromStatusSheet();
    /* Click Run Code to replace the 3rd row with new data */

    //Creating the row to update 
    //var table = new Office.TableData();
    //table.rows = ["Seattle", "WA"], ["Seattle", "WA"];
    var rowToUpdate = rowid + 2;
    //app.showNotification("Hello",data.sync.length);
    var table = new Office.TableData();
    //var len = data.sync.length;
    var incr = 0;
    var rowsdata;
    //writeToPage("I am in success" + rowid);
    //    table.rows = ['SUBMITTET', 'SUBMITTET', 'SUBMITTET', 'SUBMITTET', 'SUBMITTET', 'SUBMITTET', 'SUBMITTET', 'SUBMITTET', 'SUBMITTET', 'SUBMITTET', 'SUBMITTET', 'SUBMITTET', 'SUBMITTET', 'SUBMITTET', 'SUBMITTET', 'SUBMITTET'];
    //    table.rows = [data[0], data[1], data[2], data[3], data[4], data[5], data[6], data[7], data[8], data[9], data[10], data[11], data[12], data[13], 'SUBMITTED', 'SUBMITTED'];
    //    table.rows = [[data[0], data[1], data[2], data[3], data[4], data[5], data[6], data[7], data[8], data[9], data[10], data[11], data[12], data[13], data[14], data[15]]];
    table.rows = [[data[0], data[1], data[2], data[3], data[4], data[5], data[6], data[7], data[8], data[9], data[10], data[11], data[12], data[13], data[14], data[15], data[16], data[17], data[18], data[19], data[20], data[21], data[22], data[23]]];
    //table.rows = [['Sl no', 'Company Name', 'Contact Person', 'Designation ', 'Contact Number', 'Email ID', 'Type of Call', 'Total Call duration (Hrs mts)', 'Status', 'Called on', 'Time', 'Next Meeting', 'Time', 'Listen to Call Recordings', 'Sales Personal', 'Address', 'Remarks']];

    // table.rows = [[data.sync[incr]['slno']]];

    //Getting the table binding and setting data in 3rd row.
    Office.select("bindings#MyTableBinding", onBindingNotFound).setDataAsync(
          table,
          { coercionType: Office.CoercionType.Table, startRow: rowToUpdate },
          function (asyncResult) {
              if (asyncResult.status == "failed") {
                  //app.showNotification("Action failed with error: " + asyncResult.error.message);
              } else {
                  $("#bindLeadsTableButton").css("display", "none");

                  //app.showNotification("Leads Sheet is updated with the latest records.");
              }
          }
        );

    //Show error message in case the binding object wasn't found
    function onBindingNotFound() {
        $("#bindLeadsTableButton").css("display", "block");
        //app.showNotification("The binding object was not found. " +        "Please return to previous step to create the binding");
    }

}


function syncStatusPost(rowid, postdata, excelvalue) {
    var randomh = Math.random();
    // writeToPage(":: "+rowid +"::"+postdata+"::"+excelvalue);
    //writeToPage(excelvalue[0]);
    //RefreshStatusSheetPost(rowid, excelvalue);
    //writeToPage("I am in success" + JSON.stringify(postdata));


    $.ajax({
        type: "POST",
        //url: "http://buyfuturetoday.com/smb/SyncSalesAgent.php?email=" + userName,
        //url: "http://localhost:81/demo/check_login.json",
        //url: "http://localhost:81/demo/sync_one.json?x=" + randomh + "&email=" + userName + "",
        //        url: "http://localhost:81/demo/sync_lead.php?x=" + randomh + "",
        //url: "http://localhost:81/demo/sync_lead.php",
        url: "https://smarter-biz.com/google/calendar/create?x=" + randomh + "",
        dataType: "json",
        data: JSON.stringify(postdata),
        contentType: "application/json",
        async: true,
        cache: false,
        //data: [postdata],
        //data: JSON.stringify(postdata),
        //data: postdata,
        //        data: { json: JSON.stringify({ "name": ["uday", "sffsfs"], "email": "" }) },
        success: function (data) {
            //writeToPage("I am in success" + JSON.stringify(data));

            //syncAgents(data);
            //setData("task", data);
            //data.length
            //writeToPage((data + '').length);
            //var resultlength = (data + '').length;
            /*if (data.hasOwnProperty('errors')) {
               // excelvalue[21] = JSON.stringify(data);
                excelvalue[22] = "SYNC FAILED";
                excelvalue[23] = "SYNC FAILED";
                //testingAjax(rowid, postdata, excelvalue);
                RefreshStatusSheetPost(rowid, excelvalue);

            } else*/ if (data.hasOwnProperty('success')) {
            //if (resultlength) {
                if (data.success == "success") {
                    //excelvalue[14] = "SUBMITTED";
                    //excelvalue[15] = "SUBMITTED";
                    
                   // excelvalue[21] = JSON.stringify(data);
                    excelvalue[22] = "SUBMITTED";
                    excelvalue[23] = "SUBMITTED";
                    //testingAjax(rowid, postdata, excelvalue);
                    RefreshStatusSheetPost(rowid, excelvalue);
                    //RefreshStatusSheet('task', excelvalue);
                    //writeToPage("I am in success"+data.success);
                    //writeToPage("Data Refreshed Successfully.");
                } else {
                    //            } else if (data.success == "failer" || data.errors) {
                    //excelvalue[14] = "SYNC FAILED";
                    //excelvalue[15] = "SYNC FAILED";
//                    excelvalue[21] = JSON.stringify(data.errors);
                    excelvalue[21] = data.errors;
                    excelvalue[22] = "SYNC FAILED";
                    excelvalue[23] = "SYNC FAILED";
                    //testingAjax(rowid, postdata, excelvalue);
                    RefreshStatusSheetPost(rowid, excelvalue);
                    //RefreshStatusSheet('task', excelvalue);
                    //writeToPage("I am in failur");
                }
            } else {
//                excelvalue[21] = JSON.stringify(data.errors);
                excelvalue[21] = data.errors;
                excelvalue[22] = "SYNC FAILED";
                excelvalue[23] = "SYNC FAILED";

                // writeToPage(":" + data.error);
              //  excelvalue[21] = JSON.stringify(data);
//                excelvalue[24] = data.errors;
                //testingAjax(rowid, postdata, excelvalue);
                RefreshStatusSheetPost(rowid, excelvalue);
            }

            //RefreshStatusSheetPost(rowid, postdata);
            //setDataStatus("task");
        },
        error: function (xhr) {
            // app.showNotification('Error:', xhr.responseText);
            //writeToPage(JSON.stringify(xhr));
            //writeToPage("Sync Failed." + JSON.stringify(postdata));
            //writeToPage("Sync Failed");

        }
    });
    
}

function readBoundData(userid) {

    Office.select("bindings#bindingdata").getDataAsync({ coercionType: "matrix", valueFormat: "formatted" },
        function (asyncResult) {
            if (asyncResult.status == "failed") {
                writeToPage('Error: ' + asyncResult.error.message);
            } else {
                var unsubmitcount = 0;
                var submitcount = 0;
                var resultconcat = "";
                //writeToPage('Selected data: ' + asyncResult.value.length);

                var leadSheetRowSize = asyncResult.value.length;
                for (i = 1; i <= leadSheetRowSize; i++) {
                    //var checkvalue = String(asyncResult.value[i][15]);
                    //  var checkvalue = String(asyncResult.value[i][23]);
                    var checkvalue = "NEW";
                    //if (checkvalue == "new" || checkvalue == "edited" || checkvalue == "sync failed") {

                    if (checkvalue == "NEW" || checkvalue == "EDITED" || checkvalue == "SYNC FAILED" || checkvalue == "") {
                        //if(false){
                        // if (asyncResult.value[3][15] == "NEW" || asyncResult.value[3][15] == "EDITED" || asyncResult.value[3][15] == "SYNC FAILED" || asyncResult.value[3][15] == "") {
                        //unsubmitcount = unsubmitcount + 1;
                        /*
                        var datapost = "'" + escape(asyncResult.value[i][0]) + "','" + escape(asyncResult.value[i][1]) + "','" + escape(asyncResult.value[i][2]) + "','" + escape(asyncResult.value[i][3]) + "','" + escape(asyncResult.value[i][4]) + "','" + escape(asyncResult.value[i][5]) + "','" + escape(asyncResult.value[i][6]) + "','" + escape(asyncResult.value[i][7]) + "','" + escape(asyncResult.value[i][8]) + "','" + escape(asyncResult.value[i][9]) + "','" + escape(asyncResult.value[i][10]) + "','" + escape(asyncResult.value[i][11]) + "','" + escape(asyncResult.value[i][12]) + "','" + escape(asyncResult.value[i][13]) + "'";
                        
                        var postdatavalue = '{' +
                            '"to": "", "time_zone": "", "location": "", "event_start_date": "",' +
                            '"description": "' + datapost + '", "subject": "", "user_id": "",' +
                            '"event_end_date": "", "from": ""}';
                        */

                        //                            var datapost = ['" + asyncResult.value[i][0] + asyncResult.value[i][9] + "','" + asyncResult.value[i][10] + "','" + asyncResult.value[i][11] + "','" + asyncResult.value[i][12] + "','" + asyncResult.value[i][13] + "'];
                        //                            var datapost = "'" + asyncResult.value[i][0] + asyncResult.value[i][9] + "','" + asyncResult.value[i][10] + "','" + asyncResult.value[i][11] + "','" + asyncResult.value[i][12] + "','" + asyncResult.value[i][13] + "','" + asyncResult.value[i][15] + "','" + asyncResult.value[i][16] + "','" + asyncResult.value[i][17] + "','" + asyncResult.value[i][18] + "','" + asyncResult.value[i][19] + "','" + asyncResult.value[i][20] + "','" + asyncResult.value[i][21] + "'";
                        //                            var datapost = [" + asyncResult.value[i][0]+", " + asyncResult.value[i][9] + ", " + asyncResult.value[i][10] + ", " + asyncResult.value[i][11] + ", " + asyncResult.value[i][12] + ", " + asyncResult.value[i][13] + ", " + asyncResult.value[i][15] + ", " + asyncResult.value[i][16] + ", " + asyncResult.value[i][17] + ", " + asyncResult.value[i][18] + ", " + asyncResult.value[i][19] + ", " + asyncResult.value[i][20] + ", " + asyncResult.value[i][21] + "];
                        //var datapost = [asyncResult.value[i][0], asyncResult.value[i][9], asyncResult.value[i][10], asyncResult.value[i][11], asyncResult.value[i][12], asyncResult.value[i][13], asyncResult.value[i][15], asyncResult.value[i][16], asyncResult.value[i][17], asyncResult.value[i][18], asyncResult.value[i][19], asyncResult.value[i][20], asyncResult.value[i][21]];
//                        var datapost = asyncResult.value[i][0] + ',' + asyncResult.value[i][9] + ',' + asyncResult.value[i][10] + ',' + asyncResult.value[i][11] + ',' + asyncResult.value[i][12] + ',' + asyncResult.value[i][13] + ',' + asyncResult.value[i][15] + ',' + asyncResult.value[i][16] + ',' + asyncResult.value[i][17] + ',' + asyncResult.value[i][18] + ',' + asyncResult.value[i][19] + ',' + asyncResult.value[i][20] + ',' + asyncResult.value[i][21];
/*    //////////// This comment block is Working code for live template before changes after live.
                        var datapost = ',,,,' + asyncResult.value[i][9] + ',' + asyncResult.value[i][10] + ',' + asyncResult.value[i][12] + ',' + asyncResult.value[i][11] + ',' + asyncResult.value[i][23] + ',' + asyncResult.value[i][18] ;

                        //                        var datapost = "fvhjjhhj,0,http://smarter-biz.com/audios/1fe4ab16da66a606f183ffdbd43303cb.mp4"
                        //4423
                        //var datapost = '["dfdfsdf","asfdsgg"]';
                        var postdatavalue = { "user_id": userid, "from": asyncResult.value[i][1], "to": asyncResult.value[i][2], "event_start_date": asyncResult.value[i][3], "event_end_date": asyncResult.value[i][4], "location": asyncResult.value[i][5], "subject": asyncResult.value[i][14], "description": datapost, "repeat_params": asyncResult.value[i][6], "time_zone": asyncResult.value[i][7], "proxy_email": asyncResult.value[i][8] };
*/
                        

                        var datapost = '"";"";"";"";"' + asyncResult.value[i][1] + '";"' + asyncResult.value[i][2] + '";"' + asyncResult.value[i][5] + '";"' + asyncResult.value[i][3] + '";"' + asyncResult.value[i][23] + '";"' + asyncResult.value[i][18] + '"';
                        var appointmentstartdate = asyncResult.value[i][7] + " " + asyncResult.value[i][8];
                        var appointmentenddate = asyncResult.value[i][9] + " " + asyncResult.value[i][10];

                        var postdatavalue = { "user_id": userid, "from": asyncResult.value[i][11], "to": asyncResult.value[i][12], "event_start_date": appointmentstartdate, "event_end_date": appointmentenddate, "location": asyncResult.value[i][13], "subject": asyncResult.value[i][4], "description": datapost, "repeat_params": asyncResult.value[i][14], "time_zone": asyncResult.value[i][15], "proxy_email": asyncResult.value[i][18] };



                        // var postdatavalue = { "user_id": "4423", "from": asyncResult.value[i][1], "to": asyncResult.value[i][2], "event_start_date": asyncResult.value[i][3], "event_end_date": asyncResult.value[i][4], "location": asyncResult.value[i][5], "subject": asyncResult.value[i][14], "description": datapost, "repeat_params": asyncResult.value[i][6], "time_zone": asyncResult.value[i][7], "proxy_email": asyncResult.value[i][8] };

                        //var postdatavalue = { "user_id": "", "from": "", "to": "", "event_start_date": "", "event_end_date": "", "location": "", "subject": asyncResult.value[i][14], "description": ["afa","sfsdfd"], "repeat_params": "", "time_zone": "", "proxy_email": "" };

                        //var postdatavalue = { "to": "testingusersmb4@gmail.com", "time_zone": "Asia/Calcutta", "location": "Bangalore", "event_start_date": "Mon Dec 16 17:30:00 GMT+05:30 2015", "description": "fvhjjhhj,0,http://smarter-biz.com/audios/1fe4ab16da66a606f183ffdbd43303cb.mp4", "subject": "palsglobal", "user_id": "4423", "event_end_date": "Mon Dec 16 20:30:00 GMT+05:30 2015", "from": "bitschips21@gmail.com", "proxy_email": "testingusersmb4@gmail.com" }

                        //var obj  = JSON.parse(postdatavalue);
                        //testingAjax(i, postdatavalue, asyncResult.value[i]);
                        // resultconcat += checkvalue;
                        /*
                        setTimeout(function () {
                            writeToPage("aaa"+ i);
                            syncStatusPost(i, postdatavalue, asyncResult.value[i]);
                            writeToPage("bbb" + i);
                        }, 1000)
                        */
                        //writeToPage("values :: "+JSON.stringify(postdatavalue));
                        syncStatusPost(i, postdatavalue, asyncResult.value[i]);
                    } else if (checkvalue == "SUBMITTED") {
                        //syncStatusPost(asyncResult.value[i][15], postdatavalue, asyncResult.value[i]);
                        //testingAjax(i, postdatavalue, asyncResult.value[i]);
                        //syncStatusPost(i, postdatavalue, asyncResult.value[i]);
                        //submitcount = submitcount + 1;
                    }
                    //i = i + 2;
                    //writeToPage("Data sent to if:" + resultconcat);
                    //writeToPage('unsubmitcount: ' + unsubmitcount + " submitL: " + submitcount);
                }
                //writeToPage('Final : ' + asyncResult.value[3][10]);
            }
            //showBindingRowCount();
            //  deleteAllRowsFromTable();
            // Add updateRecord in Lead Sheet

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

function deleteAllRowsFromTable() { //uday
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


/////////////////////// Start Agents Functions ///////////////

function bindDataAgentsTable() {
    Office.context.document.bindings.addFromNamedItemAsync("Agents!AgentsTable", "table", { id: 'bindingAgentsTabledata' }, function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        } else {
            writeToPage('Added binding with type: ' + asyncResult.value.type + ' and id: ' +
                asyncResult.value.id);
        }
    });
}

function myAgentsHandler(eventArgs) {
    eventArgs.binding.getDataAsync({ coerciontype: "table" }, function (asyncResult) {

        if (asyncResult.status == "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        } else {
            writeToPage('Bound data: ' + asyncResult.value);
        }
    });
}

function RefreshAgentsSheet(tableType, data) {
    //deleteAllRowsFromStatusSheet();
    SyncAgentStatusSheet("task", data);
    app.showNotification("Hello");
}


function addEventAgentsTable() {
    Office.select("bindings#bindingAgentsTabledata").addHandlerAsync("bindingDataChanged", myAgentsHandler, function (asyncResult) {
        if (asyncResult.status == "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        } else {
            writeToPage('Added event handler');
        }
    });
}

function deleteAllRowsFromAgentsTable() {
    //   addEventAgentsTable();
    //testingAjax();
    Office.context.document.bindings.getByIdAsync("bindingAgentsTabledata", function (asyncResult) {
        var binding = asyncResult.value;
        binding.deleteAllDataValuesAsync();
    });
    if (asyncResult.status == "failed") {
        $("#BindAgentsRowTable").css("display", "block");
        //app.showNotification("Action failed with error: " + asyncResult.error.message);
    } else {
        $("#BindAgentsRowTable").css("display", "none");

        //app.showNotification("Added new binding with type: " + asyncResult.value.type + " and id: " + asyncResult.value.id);
    }
}

function RefreshAgentsSheet111(tableType, data) {
    //deleteAllRowsFromStatusSheet();
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
    for (incr = 0; incr < len; incr++) {
        var table = new Office.TableData();

        table.rows = [[data.sync[incr]['slno']]];
        // table.rows = [[data.sync[incr]['slno']]];

        //Getting the table binding and setting data in 3rd row.
        Office.select("bindings#bindingAgentsTabledata", onBindingNotFound).setDataAsync(
          table,
          { coercionType: Office.CoercionType.Table, startRow: rowToUpdate },
          function (asyncResult) {
              if (asyncResult.status == "failed") {
                  //app.showNotification("Action failed with error: " + asyncResult.error.message);
              } else {
                  $("#bindAgentsTableButton").css("display", "none");

                  app.showNotification("Agents are updated with the latest records.");
              }
          }
        );

        //Show error message in case the binding object wasn't found
        function onBindingNotFound() {
            $("#bindAgentsTableButton").css("display", "block");
            //app.showNotification("The binding object was not found. " +            "Please return to previous step to create the binding");
        }
        rowToUpdate++;
    }

}

/////////////////////// End Agents Functions ///////////////


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
          //app.showNotification("Action failed with error: " + asyncResult.error.message);
      } else {
          $("#bindStatusTableButton").css("display", "block");

          //app.showNotification("Added new binding with type: " + asyncResult.value.type + " and id: " + asyncResult.value.id);
      }
  });
}


function deleteAllRowsFromStatusSheet() {
    Office.context.document.bindings.getByIdAsync("StatusSheetBinding", function (asyncResult) {
        var binding = asyncResult.value;
        binding.deleteAllDataValuesAsync();
    });
}


// Return array of string values, or NULL if CSV string not well formed.
function CSVtoArray(text) {
    var re_valid = /^\s*(?:'[^'\\]*(?:\\[\S\s][^'\\]*)*'|"[^"\\]*(?:\\[\S\s][^"\\]*)*"|[^,'"\s\\]*(?:\s+[^,'"\s\\]+)*)\s*(?:,\s*(?:'[^'\\]*(?:\\[\S\s][^'\\]*)*'|"[^"\\]*(?:\\[\S\s][^"\\]*)*"|[^,'"\s\\]*(?:\s+[^,'"\s\\]+)*)\s*)*$/;
    var re_value = /(?!\s*$)\s*(?:'([^'\\]*(?:\\[\S\s][^'\\]*)*)'|"([^"\\]*(?:\\[\S\s][^"\\]*)*)"|([^,'"\s\\]*(?:\s+[^,'"\s\\]+)*))\s*(?:,|$)/g;
    // Return NULL if input string is not well formed CSV string.
    if (!re_valid.test(text)) return null;
    var a = [];                     // Initialize array to receive values.
    text.replace(re_value, // "Walk" the string using replace with callback.
        function (m0, m1, m2, m3) {
            // Remove backslash from \' in single quoted values.
            if (m1 !== undefined) a.push(m1.replace(/\\'/g, "'"));
                // Remove backslash from \" in double quoted values.
            else if (m2 !== undefined) a.push(m2.replace(/\\"/g, '"'));
            else if (m3 !== undefined) a.push(m3);
            return ''; // Return empty string.
        });
    // Handle special case of empty last value.
    if (/,\s*$/.test(text)) a.push('');
    return a;
}


function RefreshStatusSheet_testing_eachLoop(tableType, data) {
    var rowToUpdate = 3;
    var incr = 0;
    var rowsdata;
    var zxy;

    //table.headers = [['Sl no', 'Company Name', 'Contact Person', 'Designation ', 'Contact Number', 'Email ID', 'Type of Call', 'Total Call duration (Hrs mts)', 'Status', 'Called on', 'Time', 'Next Meeting', 'Time', 'Listen to Call Recordings', 'Sales Personal', 'Address', 'Remarks']];
    var table = new Office.TableData();
    $.each(data, function (i, item) {
        if (item.event_type == "call") {
            var descriptionCSV = ",,,,,,,,,";
        } else if (item.event_type == "calendar#event") {
            item.start_time = item.start_time['dateTime'];
            item.end_time = item.end_time['dateTime'];
            var addressColumn = "";
            var descriptionCSV = String(item.description);
        } else {
            var descriptionCSV = ",,,,,,,,,";
        }
        descriptionCSV = descriptionCSV.split(',');

        //       descriptionCSV = CSVtoArray(String(descriptionCSV));
        if (descriptionCSV.length < 10) {
            for (var le = descriptionCSV.length; le <= 10; le++) {
                descriptionCSV.push("");
            }
        }
        //writeToPage(String(descriptionCSV[0]));

        var calledduration = calculate(item.start_time, item.end_time);
        var calledon = calculate(item.start_time, item.end_time);
        var timecolumn = calledon;
        //table.rows = [["", String(descriptionCSV[4]), String(descriptionCSV[5]), String(descriptionCSV[6]), String(descriptionCSV[7]), String(item.email), String(item.event_type), calledduration, String(descriptionCSV[8]), calledon, timecolumn, String(item.created_at), String(item.created_at), String(descriptionCSV[2]), String(descriptionCSV[3]), addressColumn, "", String(descriptionCSV[0]), ""]];
        zxy.push(['Sl no', 'Company Name', 'Contact Person', 'Designation ', 'Contact Number', 'Email ID', 'Type of Call', 'Total Call duration (Hrs mts)', 'Status', 'Called on', 'Time', 'Next Meeting', 'Time', 'Listen to Call Recordings', 'Sales Personal', 'Address', 'Remarks']);

        rowToUpdate++;
    });
    table.rows.push([zxy]);
    Office.select("bindings#StatusSheetBinding", onBindingNotFound).setDataAsync(
      table,
      { coercionType: Office.CoercionType.Table, startRow: rowToUpdate },
      function (asyncResult) {
          if (asyncResult.status == "failed") {
              //app.showNotification("Action failed with error: " + asyncResult.error.message);
          } else {
              $("#bindStatusTableButton").css("display", "none");

              //app.showNotification("Status Sheet is updated with the latest records.");
          }
      }
    );

    //Show error message in case the binding object wasn't found
    function onBindingNotFound() {
        $("#bindStatusTableButton").css("display", "block");
        //app.showNotification("The binding object was not found. " +            "Please return to previous step to create the binding");
    }

}


function RefreshStatusSheet_uday(tableType, data) {
    //deleteAllRowsFromStatusSheet();
    /* Click Run Code to replace the 3rd row with new data */
    //writeToPage(JSON.stringify(data));
    var rowToUpdate = 3;
    var incr = 0;
    var rowsdata;
    var table = new Office.TableData();

    $.each(data, function (i, item) {
        if (item.event_type == "call") {
            var descriptionCSV = ",,,,,,,,,";
        } else if (item.event_type == "calendar#event") {
            item.start_time = item.start_time['dateTime'];
            item.end_time = item.end_time['dateTime'];
            var addressColumn = "";
            var descriptionCSV = String(item.description);
        } else {
            var descriptionCSV = ",,,,,,,,,";
        }

        descriptionCSV = CSVtoArray(String(descriptionCSV));
        if (descriptionCSV.length < 10) {
            for (var le = descriptionCSV.length; le <= 10; le++) {
                descriptionCSV.push("");
            }
        }
        var calledduration = calculate(item.start_time, item.end_time);
        var calledon = calculate(item.start_time, item.end_time);
        var timecolumn = calledon;
        table.rows = [["", String(descriptionCSV[4]), String(descriptionCSV[5]), String(descriptionCSV[6]), String(descriptionCSV[7]), String(item.email), String(item.event_type), calledduration, String(descriptionCSV[8]), calledon, timecolumn, String(item.created_at), String(item.created_at), String(descriptionCSV[2]), String(descriptionCSV[3]), addressColumn, "", String(descriptionCSV[0]), ""]];
        //table.rows = [["", String(descriptionCSV[4]), String(descriptionCSV[5]), String(descriptionCSV[6]), String(descriptionCSV[7]), String(item.email), String(item.event_type), calledduration, String(descriptionCSV[8]), calledon, timecolumn, String(item.created_at), String(item.created_at), String(descriptionCSV[2]), String(descriptionCSV[3]), addressColumn, "", String(descriptionCSV[0]), ""]];

        rowToUpdate++;
        // }
    });

    //Getting the table binding and setting data in 3rd row.
    Office.select("bindings#StatusSheetBinding", onBindingNotFound).setDataAsync(
      table,
      { coercionType: Office.CoercionType.Table, startRow: rowToUpdate },
      function (asyncResult) {
          if (asyncResult.status == "failed") {
              //app.showNotification("Action failed with error: " + asyncResult.error.message);
          } else {
              $("#bindStatusTableButton").css("display", "none");

              //app.showNotification("Status Sheet is updated with the latest records.");
          }
      }
    );

    //Show error message in case the binding object wasn't found
    function onBindingNotFound() {
        $("#bindStatusTableButton").css("display", "block");
        //app.showNotification("The binding object was not found. " +            "Please return to previous step to create the binding");
    }

}
function setDelay() {
    setTimeout(function () {
        //console.log(i);
    }, 1000);
}

function RefreshStatusSheet(tableType, data) {
    //deleteAllRowsFromStatusSheet();
    /* Click Run Code to replace the 3rd row with new data */
    //writeToPage(JSON.stringify(data));
    //Creating the row to update 
    //  UserName: testingusersmb1@gmail.com
    //  Password: testing12345
    // below is the event to clear the status sheet table.
    //    Office.context.document.bindings.getByIdAsync("StatusSheetBinding", function (asyncResult) {
    Office.context.document.bindings.getByIdAsync("bindingStatusTable2data", function (asyncResult) {

        var binding = asyncResult.value;
        binding.deleteAllDataValuesAsync();
    });
    var rowToUpdate = 3;
    var table = new Office.TableData();
    //table.rows = [["Seattle", "WA"],["sfasfasf","sffewfewwer"]];
    table.rows = [];
    var datalength = data.length;
    //writeToPage(data.length);
    for (var i = 0; i < datalength; i++){
        var item = data[i];
        //table.rows += "," + ["aa"+i,"nn"+i];
        //table.rows[i] = ["aa", i];
        if (item.event_type == "call") {
            var descriptionCSV = '"";"";"";"";"";"";"";"";"";';
            //item.description = ",,,,,,,,,";
            //var addressColumn = item.lat + "\\" + item.long;

            //table.rows = [[String(item.name), String(item.email), String(item.lat)]];
        } else if (item.event_type == "calendar#event") {
            //            table.rows = [[String(item.name), String(item.email), String(item.description)]];
            item.start_time = item.start_time['dateTime'];
            item.end_time = item.end_time['dateTime'];
            //item.created_at = item.create_at;
            //item.lat = null;
            //item.long = null;
            var addressColumn = "";
            //item.end['created_at'] = item.end['create_at'];
            var descriptionCSV = String(item.description);
        } else {
            var descriptionCSV = '"";"";"";"";"";"";"";"";"";';
        }
        descriptionCSV = descriptionCSV.split(';');

        //       descriptionCSV = CSVtoArray(String(descriptionCSV));
        if (descriptionCSV.length < 10) {
            for (var le = descriptionCSV.length; le <= 10; le++) {
                descriptionCSV.push("");
            }
        }
        //writeToPage(String(descriptionCSV[0]));

        //        table.rows = [[String(item.name), String(item.email), String(item.phone), String(item.send_map_location), String(item.report_email), String(item.status), String(item.htmlLink), String(item.summary), String(item.description), String(item.location), String(item.creator['email']), String(item.organizer['email']), String(item.start['dateTime']), String(item.start['timeZone']), String(item.end['dateTime']), String(item.end['timeZone']), String(item.attendees[0]['email']), String(item.attendees[0]['displayName']), String(item.attendees[0]['responseStatus'])]];
        //        table.rows = [["",String(item.name), String(item.email), String(item.phone), String(item.send_map_location), String(item.report_email), String(item.status), String(item.htmlLink), String(item.summary), String(item.description), String(item.location), String(item.creator['email']), String(item.organizer['email']), String(item.start_time['dateTime']), String(item.start_time['timeZone']), String(item.end_time['dateTime']), String(item.end_time['timeZone']), String(item.attendees[0]['email']), String(item.attendees[0]['displayName']), String(item.attendees[0]['responseStatus'])]];
        var calledduration = calculate(item.start_time, item.end_time);
        var calledon = calculate(item.start_time, item.end_time);
        var timecolumn = calledon;
        //        table.rows = [["", "", "", "", String(item.email), String(item.event_type), String(item.start_time['dateTime']), String(item.end_time['dateTime']), "", "", "", String(item.create_at), String(item.create_at), "", "", String(item.lat), String(item.long), "", "", ""]];
        //        table.rows = [[String(descriptionCSV[0]), "", "", "", String(item.email), String(item.event_type), calledduration, timecolumn, "", "", "", String(item.create_at), String(item.create_at), "", "", addressColumn, "", "", ""]];

        // Below is the original when went live, changes done by Rajesh and Sheeladitya.
        //SL. NO.,COMPANY NAME,CONTACT PERSON,DESIGNATION ,CONTACT NUMBER,EMAIL ID,TYPE OF CALL,TOTAL CALL DURATION (Min) ,STATUS,CALLED ON ,TIME,NEXT MEETING,MEETING TIME,CALL RECORDING,SALES PERSON,ADDRESS ,WEBSITE ,PERSONAL NOTES,REMARKS,
        //table.rows[i] = ["", String(descriptionCSV[4]), String(descriptionCSV[5]), String(descriptionCSV[6]), String(descriptionCSV[7]), String(item.email), String(item.event_type), calledduration, String(descriptionCSV[8]), calledon, timecolumn, String(item.created_at), String(item.created_at), String(descriptionCSV[2]), String(descriptionCSV[3]), addressColumn, "", String(descriptionCSV[0]), ""];

        // Below is the latest changes done by Rajesh on 4 Dec 2016.
        // SL. NO.,COMPANY NAME,CONTACT PERSON,CONTACT NUMBER,STATUS,PERSONAL NOTES,TOTAL CALL DURATION (Min),EMAIL ID,DESIGNATION2,TYPE OF CALL,TOTAL CALL DURATION (Min)2,CALLED ON,TIME,NEXT MEETING,MEETING TIME,CALL RECORDING,SALES PERSON,ADDRESS,WEBSITE,REMARKS
        table.rows[i] = [i+1, String(descriptionCSV[4]), String(descriptionCSV[5]), String(descriptionCSV[7]), String(descriptionCSV[8]), String(descriptionCSV[0]), calledduration, String(item.email), String(descriptionCSV[6]), String(item.event_type), calledduration, calledon, timecolumn, String(item.created_at), String(item.created_at), String(descriptionCSV[2]), String(descriptionCSV[3]), addressColumn,"","",""];


    }
    
      //writeToPage(JSON.stringify(table.rows));
    //Getting the table binding and setting data in 3rd row.
    Office.select("bindings#StatusSheetBinding", onBindingNotFound).setDataAsync(
      table,
      { coercionType: Office.CoercionType.Table, startRow: rowToUpdate },
      function (asyncResult) {
          if (asyncResult.status == "failed") {
              app.showNotification("Action failed with error: " + asyncResult.error.message);
          } else {
              $("#Binding").css("display", "none");
              $("#bindStatusTableButton").css("display", "none");
              
              //app.showNotification("Status Sheet is updated with the latest records.");
          }
      }
    );

    //Show error message in case the binding object wasn't found
    function onBindingNotFound() {
        $("#Binding").css("display", "block");

        $("#bindStatusTableButton").css("display", "block");
        //app.showNotification("The binding object was not found. " +            "Please return to previous step to create the binding");
    }
}
function RefreshStatusSheetworkingcopy(tableType, data) {
    //deleteAllRowsFromStatusSheet();
    /* Click Run Code to replace the 3rd row with new data */
    //writeToPage(JSON.stringify(data));
    //Creating the row to update 
    //var table = new Office.TableData();
    //table.rows = ["Seattle", "WA"], ["Seattle", "WA"];
    var rowToUpdate = 3;
    //app.showNotification("Hello",data.sync.length);
    // var table = new Office.TableData();
    // var len = data.sync.length;
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
    // for (incr = 0; incr < len; incr++) {
    //writeToPage(JSON.stringify(data));
    $.each(data, function (i, item) {
        var table = new Office.TableData();
        // Below line is for the demo Excel sheet in working condition.
        //         table.rows = [[data.sync[incr]['slno'], data.sync[incr]['Company Name'], data.sync[incr]['Contact Person'], data.sync[incr]['Designation'], data.sync[incr]['Contact Number'], data.sync[incr]['Email ID'], data.sync[incr]['Type of Call'], data.sync[incr]['Total Call duration (Hrs mts)'], data.sync[incr]['Status'], data.sync[incr]['Called on'], data.sync[incr]['Time'], data.sync[incr]['Next Meeting'], data.sync[incr]['Time'], data.sync[incr]['Listen to Call Recordings'], data.sync[incr]['Sales Personal'], data.sync[incr]['Address'], data.sync[incr]['Remarks']]];
        // Below line is for the live Excel sheet with first json resonse from the api.
        //        table.rows = [[String(data.sync[incr]['name']), String(data.sync[incr]['email']), String(data.sync[incr]['phone']), String(data.sync[incr]['send_map_location']), String(data.sync[incr]['report_email']), String(data.sync[incr]['from']), String(data.sync[incr]['to']), String(data.sync[incr]['cc']), String(data.sync[incr]['bcc']), String(data.sync[incr]['event_type']), String(data.sync[incr]['url']), String(data.sync[incr]['jd_url']), String(data.sync[incr]['jobcard_url']), String(data.sync[incr]['live_recording_url']), String(data.sync[incr]['start_time']), String(data.sync[incr]['end_time']), String(data.sync[incr]['subject']), String(data.sync[incr]['message']), String(data.sync[incr]['lat']), String(data.sync[incr]['long']), String(data.sync[incr]['attachments']), String(data.sync[incr]['children']), String(data.sync[incr]['parent']), String(data.sync[incr]['job_card']), String(data.sync[incr]['employee_assignee']), String(data.sync[incr]['status']), String(data.sync[incr]['caller_name']), String(data.sync[incr]['favourite']), String(data.sync[incr]['cmail_url']), String(data.sync[incr]['other_party'])]];
        //table.rows = [[String(data.sync[incr]['name']), String(data.sync[incr]['email']), String(data.sync[incr]['phone']), String(data.sync[incr]['send_map_location']), String(data.sync[incr]['report_email']), String(data.sync[incr]['kind']), String(data.sync[incr]['status']), String(data.sync[incr]['htmlLink']), String(data.sync[incr]['summary']), String(data.sync[incr]['description']), String(data.sync[incr]['location']), String(data.sync[incr]['creator']['email']), String(data.sync[incr]['organizer']['email']), String(data.sync[incr]['start']['dateTime']), String(data.sync[incr]['start']['timeZone']), String(data.sync[incr]['end']['dateTime']), String(data.sync[incr]['end']['timeZone']), String(data.sync[incr]['attendees'][0]['email']), String(data.sync[incr]['attendees'][0]['displayName']), String(data.sync[incr]['attendees'][0]['responseStatus'])]];
        //        table.rows = [[String(data[incr]['name']), String(data[incr]['email']), String(data[incr]['phone']), String(data[incr]['send_map_location']), String(data[incr]['report_email']), String(data[incr]['kind']), String(data[incr]['status']), String(data[incr]['htmlLink']), String(data[incr]['summary']), String(data[incr]['description']), String(data[incr]['location']), String(data[incr]['creator']['email']), String(data[incr]['organizer']['email']), String(data[incr]['start']['dateTime']), String(data[incr]['start']['timeZone']), String(data[incr]['end']['dateTime']), String(data[incr]['end']['timeZone']), String(data[incr]['attendees'][0]['email']), String(data[incr]['attendees'][0]['displayName']), String(data[incr]['attendees'][0]['responseStatus'])]];
        //table.rows = [[String(item.name), String(item.email), String(data.phone), String(data.send_map_location), String(data.report_email), String(data.kind), String(data.status), String(data.htmlLink), String(data.summary), String(data.description), String(data.location), String(data.creator['email']), String(data.organizer['email']), String(data.start['dateTime']), String(data.start['timeZone']), String(data.end['dateTime']), String(data.end['timeZone']), String(data.attendees[0]['email']), String(data.attendees[0]['displayName']), String(data.attendees[0]['responseStatus'])]];
        // Working condition
        //        table.rows = [[String(item.name), String(item.email), String(item.phone), String(item.send_map_location), String(item.report_email), String(item.kind), String(item.status), String(item.htmlLink), String(item.summary), String(item.description), String(item.location), String(item.creator['email']), String(item.organizer['email']), String(item.start['dateTime']), String(item.start['timeZone']), String(item.end['dateTime']), String(item.end['timeZone']), String(item.attendees[0]['email']), String(item.attendees[0]['displayName']), String(item.attendees[0]['responseStatus'])]];
        if (item.event_type == "call") {
            var descriptionCSV = ",,,,,,,,,";
            //item.description = ",,,,,,,,,";
            //var addressColumn = item.lat + "\\" + item.long;

            //table.rows = [[String(item.name), String(item.email), String(item.lat)]];
        } else if (item.event_type == "calendar#event") {
            //            table.rows = [[String(item.name), String(item.email), String(item.description)]];
            item.start_time = item.start_time['dateTime'];
            item.end_time = item.end_time['dateTime'];
            //item.created_at = item.create_at;
            //item.lat = null;
            //item.long = null;
            var addressColumn = "";
            //item.end['created_at'] = item.end['create_at'];
            var descriptionCSV = String(item.description);
        } else {
            var descriptionCSV = ",,,,,,,,,";
        }
        descriptionCSV = descriptionCSV.split(',');

 //       descriptionCSV = CSVtoArray(String(descriptionCSV));
        if (descriptionCSV.length < 10) {
            for (var le = descriptionCSV.length; le <= 10; le++) {
                descriptionCSV.push("");
            }
        }
        //writeToPage(String(descriptionCSV[0]));

        //        table.rows = [[String(item.name), String(item.email), String(item.phone), String(item.send_map_location), String(item.report_email), String(item.status), String(item.htmlLink), String(item.summary), String(item.description), String(item.location), String(item.creator['email']), String(item.organizer['email']), String(item.start['dateTime']), String(item.start['timeZone']), String(item.end['dateTime']), String(item.end['timeZone']), String(item.attendees[0]['email']), String(item.attendees[0]['displayName']), String(item.attendees[0]['responseStatus'])]];
        //        table.rows = [["",String(item.name), String(item.email), String(item.phone), String(item.send_map_location), String(item.report_email), String(item.status), String(item.htmlLink), String(item.summary), String(item.description), String(item.location), String(item.creator['email']), String(item.organizer['email']), String(item.start_time['dateTime']), String(item.start_time['timeZone']), String(item.end_time['dateTime']), String(item.end_time['timeZone']), String(item.attendees[0]['email']), String(item.attendees[0]['displayName']), String(item.attendees[0]['responseStatus'])]];
        var calledduration = calculate(item.start_time, item.end_time);
        var calledon = calculate(item.start_time, item.end_time);
        var timecolumn = calledon;
        //        table.rows = [["", "", "", "", String(item.email), String(item.event_type), String(item.start_time['dateTime']), String(item.end_time['dateTime']), "", "", "", String(item.create_at), String(item.create_at), "", "", String(item.lat), String(item.long), "", "", ""]];
        //        table.rows = [[String(descriptionCSV[0]), "", "", "", String(item.email), String(item.event_type), calledduration, timecolumn, "", "", "", String(item.create_at), String(item.create_at), "", "", addressColumn, "", "", ""]];
        table.rows = [["", String(descriptionCSV[4]), String(descriptionCSV[5]), String(descriptionCSV[6]), String(descriptionCSV[7]), String(item.email), String(item.event_type), calledduration, String(descriptionCSV[8]), calledon, timecolumn, String(item.created_at), String(item.created_at), String(descriptionCSV[2]), String(descriptionCSV[3]), addressColumn, "", String(descriptionCSV[0]), ""]];

        //table.rows = [[String(item.name), String(item.email), String(item.phone), String(item.send_map_location), String(item.report_email), String(item.kind), String(item.status), String(item.htmlLink), String(item.summary), String(item.description), String(item.location), String(item.creator['email']), String(item.organizer['email']), String(item.start['dateTime']), String(item.start['timeZone']), String(item.end['dateTime']), String(item.end['timeZone']), String(item.attendees[0]['email']), String(item.attendees[0]['displayName']), String(item.attendees[0]['responseStatus'])]];

       // var descReturn = CSVtoArray(descriptionCSV);
        //writeToPage(descReturn);
        //table.rows = [[String(item.name), String(item.email), String(item.lat), descriptionCSV]];

        /*
        String(item.email)		    String(item.email)
        String(item.event_type)		String(item.event_type)
        String(item.start_time)		String(item.start['dateTime'])
        String(item.end_time)		String(item.end['dateTime'])
        String(item.end['lat'])		null
        String(item.end['long'])	null
        String(item.end['created_at'])   String(item.end['create_at'])
        String(item.description)	null


        */
        // table.rows = [[data.sync[incr]['slno']]];

        //Getting the table binding and setting data in 3rd row.
        Office.select("bindings#StatusSheetBinding", onBindingNotFound).setDataAsync(
          table,
          { coercionType: Office.CoercionType.Table, startRow: rowToUpdate },
          function (asyncResult) {
              if (asyncResult.status == "failed") {
                  //app.showNotification("Action failed with error: " + asyncResult.error.message);
              } else {
                  $("#bindStatusTableButton").css("display", "none");

                  //app.showNotification("Status Sheet is updated with the latest records.");
              }
          }
        );

        //Show error message in case the binding object wasn't found
        function onBindingNotFound() {
            $("#bindStatusTableButton").css("display", "block");
            //app.showNotification("The binding object was not found. " +            "Please return to previous step to create the binding");
        }
        rowToUpdate++;
        //setDelay();
        // }
    });
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
              //app.showNotification("Action failed with error: " + asyncResult.error.message);
          } else {
              $("#bindStatusTableButton").css("display", "none");

              //app.showNotification("Status Sheet is updated with the latest records.");
          }
      }
    );

    //Show error message in case the binding object wasn't found
    function onBindingNotFound() {
        $("#bindStatusTableButton").css("display", "block");
        //app.showNotification("The binding object was not found. " +        "Please return to previous step to create the binding");
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
          //app.showNotification("Action failed with error: " + asyncResult.error.message);
      } else {
          $("#bindLeadsTableButton").css("display", "none");

          //app.showNotification("Added new binding with type: " + asyncResult.value.type +  " and id: " + asyncResult.value.id);
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
function updateColumn(userName) {

    var randomh = Math.random();
    var postdatavalue = { email: userName, status: "submit" };
    $.ajax({
        method: "POST",
        //url: "http://buyfuturetoday.com/smb/SyncSalesAgent.php?email=" + userName,
        //url: "http://localhost:81/demo/check_login.json",
        url: "http://localhost:81/demo/sync_lead.php?x=" + randomh + "&email=" + userName + "",
        //url: "http://localhost:81/demo/sync_lead.php?x=" + randomh + "",
        dataType: "json",
        success: function (data) {
            //syncAgents(data);
            //setData("task", data);
            app.showNotification("Leads Sheetadasa is updated with the latest records.");

            updateLeadSheet("task", data);
            //setDataStatus("task");
        },
        error: function (xhr) {
            app.showNotification('Error:', xhr.responseText);
        },
        data: postdatavalue
    });

}

function updateLeadSheet(tableType, data) {
    app.showNotification("Leads Sheet is updated with the latest records.");
    /* Click Run Code to replace the 3rd row with new data */

    //Creating the row to update 
    //var table = new Office.TableData();
    //table.rows = ["Seattle", "WA"], ["Seattle", "WA"];
    var rowToUpdate = 3;
    //app.showNotification("Hello",data.sync.length);
    var table = new Office.TableData();
    var len = data.sync.length;
    var incr = 0;
    var rowsdata;

    //table.headers = [['Sl no', 'Company Name', 'Contact Person', 'Designation ', 'Contact Number', 'Email ID', 'Type of Call', 'Total Call duration (Hrs mts)', 'Status', 'Called on', 'Time', 'Next Meeting', 'Time', 'Listen to Call Recordings', 'Sales Personal', 'Address', 'Remarks']];
    // table.rows = [[data.sync[0]['slno'], data.sync[0]['Company Name'], data.sync[0]['Contact Person'], data.sync[0]['Designation'], data.sync[0]['Contact Number'], data.sync[0]['Email ID'], data.sync[0]['Type of Call'], data.sync[0]['Total Call duration (Hrs mts)'], data.sync[0]['Status'], data.sync[0]['Called on'], data.sync[0]['Time'], data.sync[0]['Next Meeting'], data.sync[0]['Time'], data.sync[0]['Listen to Call Recordings'], data.sync[0]['Sales Personal'], data.sync[0]['Address'], data.sync[0]['Remarks']], [data.sync[0]['slno'], data.sync[0]['Company Name'], data.sync[0]['Contact Person'], data.sync[0]['Designation'], data.sync[0]['Contact Number'], data.sync[0]['Email ID'], data.sync[0]['Type of Call'], data.sync[0]['Total Call duration (Hrs mts)'], data.sync[0]['Status'], data.sync[0]['Called on'], data.sync[0]['Time'], data.sync[0]['Next Meeting'], data.sync[0]['Time'], data.sync[0]['Listen to Call Recordings'], data.sync[0]['Sales Personal'], data.sync[0]['Address'], data.sync[0]['Remarks']], [data.sync[0]['slno'], data.sync[0]['Company Name'], data.sync[0]['Contact Person'], data.sync[0]['Designation'], data.sync[0]['Contact Number'], data.sync[0]['Email ID'], data.sync[0]['Type of Call'], data.sync[0]['Total Call duration (Hrs mts)'], data.sync[0]['Status'], data.sync[0]['Called on'], data.sync[0]['Time'], data.sync[0]['Next Meeting'], data.sync[0]['Time'], data.sync[0]['Listen to Call Recordings'], data.sync[0]['Sales Personal'], data.sync[0]['Address'], data.sync[0]['Remarks']], [data.sync[0]['slno'], data.sync[0]['Company Name'], data.sync[0]['Contact Person'], data.sync[0]['Designation'], data.sync[0]['Contact Number'], data.sync[0]['Email ID'], data.sync[0]['Type of Call'], data.sync[0]['Total Call duration (Hrs mts)'], data.sync[0]['Status'], data.sync[0]['Called on'], data.sync[0]['Time'], data.sync[0]['Next Meeting'], data.sync[0]['Time'], data.sync[0]['Listen to Call Recordings'], data.sync[0]['Sales Personal'], data.sync[0]['Address'], data.sync[0]['Remarks']]];
    table.rows = [[data.sync[0]['slno'], data.sync[0]['Company Name'], data.sync[0]['Contact Person'], data.sync[0]['Designation'], data.sync[0]['Contact Number'], data.sync[0]['Email ID'], data.sync[0]['Type of Call'], data.sync[0]['Total Call duration (Hrs mts)'], data.sync[0]['Status'], data.sync[0]['Called on'], data.sync[0]['Time'], data.sync[0]['Next Meeting'], data.sync[0]['Time'], data.sync[0]['Listen to Call Recordings'], data.sync[0]['Sales Personal'], data.sync[0]['Address'], data.sync[0]['Remarks']]];
    table.rows = [[data.sync[1]['slno'], data.sync[1]['Company Name'], data.sync[1]['Contact Person'], data.sync[1]['Designation'], data.sync[1]['Contact Number'], data.sync[1]['Email ID'], data.sync[1]['Type of Call'], data.sync[1]['Total Call duration (Hrs mts)'], data.sync[1]['Status'], data.sync[1]['Called on'], data.sync[1]['Time'], data.sync[1]['Next Meeting'], data.sync[1]['Time'], data.sync[1]['Listen to Call Recordings'], data.sync[1]['Sales Personal'], data.sync[1]['Address'], data.sync[1]['Remarks']]];
    table.rows = [[data.sync[2]['slno'], data.sync[2]['Company Name'], data.sync[2]['Contact Person'], data.sync[2]['Designation'], data.sync[2]['Contact Number'], data.sync[2]['Email ID'], data.sync[2]['Type of Call'], data.sync[2]['Total Call duration (Hrs mts)'], data.sync[2]['Status'], data.sync[2]['Called on'], data.sync[2]['Time'], data.sync[2]['Next Meeting'], data.sync[2]['Time'], data.sync[2]['Listen to Call Recordings'], data.sync[2]['Sales Personal'], data.sync[2]['Address'], data.sync[2]['Remarks']]];
    table.rows = [[data.sync[3]['slno'], data.sync[3]['Company Name'], data.sync[3]['Contact Person'], data.sync[3]['Designation'], data.sync[3]['Contact Number'], data.sync[3]['Email ID'], data.sync[3]['Type of Call'], data.sync[3]['Total Call duration (Hrs mts)'], data.sync[3]['Status'], data.sync[3]['Called on'], data.sync[3]['Time'], data.sync[3]['Next Meeting'], data.sync[3]['Time'], data.sync[3]['Listen to Call Recordings'], data.sync[3]['Sales Personal'], data.sync[3]['Address'], data.sync[3]['Remarks']]];

    // return table;

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
        //app.showNotification("The binding object was not found. " +        "Please return to previous step to create the binding");
    }
    /*
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
    */
}

function updateLeadSheet_workinglive(tableType, data) {
    app.showNotification("Leads Sheet is updated with the latest records.");
    /* Click Run Code to replace the 3rd row with new data */

    //Creating the row to update 
    //var table = new Office.TableData();
    //table.rows = ["Seattle", "WA"], ["Seattle", "WA"];
    var rowToUpdate = 3;
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
            //app.showNotification("The binding object was not found. " +            "Please return to previous step to create the binding");
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

        //app.showNotification("The binding object was not found. " +        "Please return to previous step to create the binding");
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
          //app.showNotification("Action failed with error: " + asyncResult.error.message);
      } else {
          $("#AgentsButton").css("display", "none");

          //app.showNotification("Added new binding with type: " + asyncResult.value.type + " and id: " + asyncResult.value.id);
      }
  });

}


function SyncAgentStatusSheet(tableType, data) {

    Office.context.document.bindings.getByIdAsync("bindingAgentsTabledata", function (asyncResult) {
        var binding = asyncResult.value;
        binding.deleteAllDataValuesAsync();
    });
    
    /* Click Run Code to replace the 3rd row with new data */
    var rowToUpdate = 1;
    //app.showNotification("Hello",data.sync.length);
    //var len = data.sync.length;
    var incr = 0;
    var rowsdata;
    var table = new Office.TableData();
    table.rows = [];
    var datalength = data.groups.length;
    
    for (var i = 0; i < datalength; i++) {
        var item = data.groups[i];

//        table.rows[i] = [String(item.name)];
        table.rows[i] = [String(item.email)];
        //table.rows = [[JSON.stringify(item.name)]];

    }
    
    //writeToPage(JSON.stringify(data));
    //Getting the table binding and setting data in 3rd row.
    Office.select("bindings#AgentBinding", onBindingNotFound).setDataAsync(
      table,
      { coercionType: Office.CoercionType.Table, startRow: rowToUpdate },
      function (asyncResult) {
          if (asyncResult.status == "failed") {
              //app.showNotification("Action failed with error: " + asyncResult.error.message);
          } else {
              $("#bindStatusTableButton").css("display", "none");

              //app.showNotification("Agents are updated.");
          }
      }
    );

    //Show error message in case the binding object wasn't found
    function onBindingNotFound() {
        $("#bindStatusTableButton").css("display", "block");
        //app.showNotification("The binding object was not found. " +            "Please return to previous step to create the binding");
    }
}

function SyncAgentStatusSheet_working_copy_live(tableType, data) {

    Office.context.document.bindings.getByIdAsync("bindingAgentsTabledata", function (asyncResult) {
        var binding = asyncResult.value;
        binding.deleteAllDataValuesAsync();
    });

    /* Click Run Code to replace the 3rd row with new data */
    var rowToUpdate = 1;
    //app.showNotification("Hello",data.sync.length);
    //var len = data.sync.length;
    var incr = 0;
    var rowsdata;

    $.each(data.groups, function (i, item) {
        var table = new Office.TableData();
        table.rows = [[String(item.name)]];
        //table.rows = [[JSON.stringify(item.name)]];

        //Getting the table binding and setting data in 3rd row.
        Office.select("bindings#AgentBinding", onBindingNotFound).setDataAsync(
          table,
          { coercionType: Office.CoercionType.Table, startRow: rowToUpdate },
          function (asyncResult) {
              if (asyncResult.status == "failed") {
                  //app.showNotification("Action failed with error: " + asyncResult.error.message);
              } else {
                  $("#bindStatusTableButton").css("display", "none");

                  //app.showNotification("Agents are updated.");
              }
          }
        );

        //Show error message in case the binding object wasn't found
        function onBindingNotFound() {
            $("#bindStatusTableButton").css("display", "block");
            //app.showNotification("The binding object was not found. " +            "Please return to previous step to create the binding");
        }
        rowToUpdate++;
    });
}

function SyncAgentStatusSheet_working_demo(tableType, data) {
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

        // table.rows = [[data.sync[incr]['slno'], data.sync[incr]['Company Name'], data.sync[incr]['Contact Person'], data.sync[incr]['Designation'], data.sync[incr]['Contact Number'], data.sync[incr]['Email ID'], data.sync[incr]['Type of Call'], data.sync[incr]['Total Call duration (Hrs mts)'], data.sync[incr]['Status'], data.sync[incr]['Called on'], data.sync[incr]['Time'], data.sync[incr]['Next Meeting'], data.sync[incr]['Time'], data.sync[incr]['Listen to Call Recordings'], data.sync[incr]['Sales Personal'], data.sync[incr]['Address'], data.sync[incr]['Remarks']]];
        table.rows = [[data.sync[incr]['slno']]];

        //Getting the table binding and setting data in 3rd row.
        Office.select("bindings#AgentBinding", onBindingNotFound).setDataAsync(
          table,
          { coercionType: Office.CoercionType.Table, startRow: rowToUpdate },
          function (asyncResult) {
              if (asyncResult.status == "failed") {
                  //app.showNotification("Action failed with error: " + asyncResult.error.message);
              } else {
                  $("#bindStatusTableButton").css("display", "none");

                  //app.showNotification("Agents are updated.");
              }
          }
        );

        //Show error message in case the binding object wasn't found
        function onBindingNotFound() {
            $("#bindStatusTableButton").css("display", "block");
            //app.showNotification("The binding object was not found. " + "Please return to previous step to create the binding");
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


function syncSalesAgent(userid) {
    var randomh = Math.random();
    $.ajax({
        method: "GET",
        //url: "http://buyfuturetoday.com/smb/SyncSalesAgent.php?email=" + userid,
        //url: "http://localhost:81/demo/check_login.json",
        //url: "http://localhost:81/demo/sync_agents.json?x=" + randomh + "&email=" + userid + "",
        //url: "http://localhost:81/demo/sync_agents.php?x=" + randomh + "&email=" + userid + "",
        //url: "https://smartersmb.azurewebsites.net/api/getgroupsofuser?user_id=4411",
        url: "https://smarter-biz.com/api/getgroupsofuser?user_id=" + userid + "",
        dataType: "json",
        //data: JSON.stringify(postdata),
        //contentType: "application/json",
        async: false,
        success: function (data) {
            //syncAgents(data);
            //setData("task", data);
            //deleteAllRowsFromAgentsTable();
            //RefreshAgentsSheet("task", data);
            SyncAgentStatusSheet("task", data);
            //testingAjax();
            //setDataStatus("task");
            //app.showNotification('SUccess:');
            //writeToPage("Agents list updated." + JSON.stringify(data));
        },
        error: function (xhr) {
            app.showNotification('Error:', xhr.responseText);
        }
    });
}

function testingAjax(i, datapost, excelvalue) {
    writeToPage("Hello Test." + i + datapost + excelvalue[14]);
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

///////////////////////////////////////// New Status Sheet  /////////////////////////


function bindDataStatusTable() {
    Office.context.document.bindings.addFromNamedItemAsync("Status!Table2", "table", { id: 'bindingStatusTable2data' }, function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        } else {
            writeToPage('Added binding with type: ' + asyncResult.value.type + ' and id: ' +
                asyncResult.value.id);
        }
    });
}
// Currently Not using to read data in Status Sheet
function readBoundDataStatusTable() {
    Office.select("bindings#bindingStatusTable2data").getDataAsync({ coercionType: "text" },
        function (asyncResult) {
            if (asyncResult.status == "failed") {
                writeToPage('Error: ' + asyncResult.error.message);
            } else {
                writeToPage('Data Clear : ' + asyncResult.value);
            }
            //showBindingRowCount();
            //deleteAllRowsFromStatusTable();
        });
}

function addEventStatusTable() {
    Office.select("bindings#bindingStatusTable2data").addHandlerAsync("bindingDataChanged", myHandler, function (asyncResult) {
        if (asyncResult.status == "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        } else {
            writeToPage('Added event handler');
        }
    });
}

function deleteAllRowsFromStatusTable() {
    Office.context.document.bindings.getByIdAsync("bindingStatusTable2data", function (asyncResult) {
        var binding = asyncResult.value;
        binding.deleteAllDataValuesAsync();
    });
}


function addEventAgentsTable() {
    Office.select("bindings#AgentBinding").addHandlerAsync("bindingDataChanged", myHandler, function (asyncResult) {
        if (asyncResult.status == "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        } else {
            writeToPage('Added event handler');
        }
    });
}

function aaaaaa() {
    // Create a TableData object.
    var myTable = new Office.TableData();
    myTable.headers = ["First Name", "Last Name", "Balance"];
    myTable.rows = [["Brittney", "Booker", "1223.10"],
                    ["Sanjit", "Pandit", "34234.99"],
                    ["Naomi", "Peacock", "-50.78"]];

    // Set the myTable in the document.
    Office.context.document.setSelectedDataAsync(myTable, { coercionType: Office.CoercionType.Table },
      function (asyncResult) {
          if (asyncResult.status == "failed") {
              //showMessage("Action failed with error: " + asyncResult.error.message);
          } else {
              //Create a new table binding for the selected table.
              Office.context.document.bindings.addFromSelectionAsync(Office.CoercionType.Table, { id: "MyTableBindingaaa" },
                function (asyncResult) {
                    if (asyncResult.status == "failed") {
                        //app.showNotification("Action failed with error: " + asyncResult.error.message);
                    } else {
                        //app.showNotification("Added new binding with type: " + asyncResult.value.type + " and id: " + asyncResult.value.id + ". Click next to learn how to navigate to this new binding.");
                    }
                }
              )
          }
      }
    );
}
function changeTab() {
    //Go to binding by ID. Scroll so the binding is off-screen, then click Run Code  MyTableBinding  StatusSheetBinding
    Office.context.document.goToByIdAsync("MyTableBinding", Office.GoToType.Binding,
      function (asyncResult) {
          if (asyncResult.status == "failed") {
              app.showNotification("Action failed with error: " + asyncResult.error.message +
                          ". Try going back to the previous step to set a binding.");
          } else {
              app.showNotification("Navigation successful!");
          }
      }
    );
    app.showNotification("i am last");
}
function goToTable() {
    //    Office.context.document.goToByIdAsync("AgentsTable", Office.GoToType.NamedItem, function (asyncResult) {
//    Office.context.document.goToByIdAsync("AgentsTable", Office.GoToType.Slide, function (asyncResult) {    Office.context.document.goToByIdAsync("AgentsTable", Office.GoToType.NamedItem, function (asyncResult) {
    //    Office.context.document.goToByIdAsync("AgentsTable", Office.GoToType.Index, function (asyncResult) {
      Office.context.document.goToByIdAsync("AgentsTable", Office.GoToType.NamedItem, function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage("Navigation successful");
        }
    });
}