/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            //$('#get-data-from-selection').click(getDataFromSelection);
        });
    };

    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {

                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    app.showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    app.showNotification('Error:', result.error.message);
                }
            }
        );
    }
    var loginMethod = "";
    var userName = "";
    var password = "";
    var jsonData = "";

    $(document).ready(function () {
       // dowloadlink();

        $("#rbtnOffice365").click(function () {
            chkPanelChanged();
        });

        $("#rbtnSMB").click(function () {
            chkPanelChanged();
        });

        $("#btnLogin").click(function () {
            loginSMB();
        });

        $("#btnSyncSalesAgent").click(function () {
            syncStatus($('#userName').val(), $('#userid').val());
        });

        $("#btnSubmit").click(function () {
            createAppointment();
        });
        alreadyLoggedin();
    });


    function dowloadlink() {
        Office.select("bindings#bindingdata").getDataAsync({ coercionType: "matrix" },
           function (asyncResult) {
               if (asyncResult.status == "failed") {
                   //app.showNotification('Error: ');
               } else {
                   //app.showNotification("Successss");
                   $("#downloadlink").hide();
               }
           });
    }


    function chkPanelChanged() {
        if ($("#rbtnOffice365").is(':checked')) {
            $("#loginDiv").show();
            $("#divLoginSMB").hide();
            loginMethod = "o365";
            //var userProfile = Office.context.mailbox.userProfile;
            //var name = userProfile.emailAddress;
            //app.showNotification(name);
            //createTable();
        }
        if ($("#rbtnSMB").is(':checked')) {
            $("#loginDiv").show();
            $("#divLoginSMB").show();
//            loginMethod = "smb";
            loginMethod = "gplus";
        }
    }
    function alreadyLoggedin() {
        if ($("#userName").val()) {
            $("#loginDiv").hide();
            $("#OptionDiv").hide();
            $("#divSyncAgents").show();
            $("#userName").val(userName);
            $("#userid").val(data.id);
        }
        
    }

    
    function loginSMB() {
        var randomh = Math.random();
        userName = $("#txtName").val();
        password = $("#txtPassword").val();
        loginMethod = "gplus";
        $("#errorBox").html("Please wait..");
        //app.showNotification(userName);
        if ($("#txtName").val() == "") {
            $("#txtName").focus();
            $("#errorBox").html("Please enter the User Name");
            return false;
        }
        else if ($("#txtPassword").val() == "") {
            $("#lname").focus();
            $("#errorBox").html("Please Enter the Password");
            return false;
        }
       // var loginURL = "http://buyfuturetoday.com/smb/check_login.php?email=" + userName + "&password=" + password + "&login_method=" + loginMethod;
        var loginURL = "http://smartersmb.azurewebsites.net/check_login.json?x="+randomh+"&email=" + userName + "&password=" + password + "&login_method=" + loginMethod + "&ver=3";
        //var loginURL = "http://localhost:81/demo/check_login.json";
        $.ajax({
            type: "GET",
            url: loginURL,
            dataType: "json",
            success: function (data) {
                
                if (data.id) {
                    app.showNotification("");
                    //addNewSheets();
                    //createLeadsTable();

                    $("#loginDiv").hide();
                    $("#OptionDiv").hide();
                    $("#divSyncAgents").show();
                    $("#usernameTab").show();
                    $("#userName").val(userName);
                    $("#userid").val(data.id);

                    jsonData = data;
                    //app.showNotification("Welcome " + data.name);
                    var nameuser = "Welcome " + data.name;
                    document.getElementById('nameuser').innerText = nameuser;
                    //bindTablebyName();
                }
                else if (data.result == "failure") {
                    //app.showNotification(data.errors);
                    loginError("Please check your Username and Password.");
                } else {
                    //loginError("Please check your internet connection.");
                    //loginError("" + JSON.stringify(data.error));
                    loginError("Please check your Username and Password.");

                }
            },
            error: function (xhr) {

                //app.showNotification('Error:', xhr.responseText);
                loginError("There is some problem connecting to server.");
            }
        });

    }

    function syncStatus(userName,userid) {
        var randomh = Math.random();
        //var userid = $("#userid").val();              
        var month_name = new Array("Jan", "Feb", "Mar",
                    "Apr", "May", "Jun", "Jul", "Aug", "Sep",
                    "Oct", "Nov", "Dec");

        var startdatenow = new Date(($("#startdate").val()));
        var options = {
            weekday: "long", year: "numeric", month: "short",
            day: "numeric", hour: "2-digit", minute: "2-digit"
        };
        var startdatetoapi = month_name[startdatenow.getUTCMonth()] + " " + startdatenow.getDate() + " " + startdatenow.getFullYear();
        var startnow = new Date($("#startdate").val());
        var startnowUtc = new Date(startnow.getTime() + (startnow.getTimezoneOffset() * 60000));

        var endnow = new Date($("#enddate").val());
        var endnowUtc = new Date(endnow.getTime() + (endnow.getTimezoneOffset() * 60000));
        //userid = 4423;
        var aurl = "http://smartersmb.azurewebsites.net/api/getdumpusers_mails?x=" + randomh + "&user_id=" + userid + "&start_date=" + startdatetoapi + "&end_date=" + endnowUtc + "";
        //document.getElementById('outputs').innerText = aurl;

        $.ajax({
            type: "GET",
            //url: "http://buyfuturetoday.com/smb/SyncSalesAgent.php?email=" + userName,
            //url: "http://localhost:81/demo/check_login.json",
            //url: "http://localhost:81/demo/sync_one.json?x=" + randomh + "&email=" + userName + "",
            //url: "http://localhost:81/demo/sync_one.php?x=" + randomh + "&user_id=" + userid + "&start_date=" + startdatetoapi + "&end_date=" + endnowUtc + "",
            url: "http://smartersmb.azurewebsites.net/api/getdumpusers_mails?x=" + randomh + "&user_id=" + userid + "&start_date=" + startdatetoapi + "&end_date=" + endnowUtc + "",
            dataType: "json",
            async: false,
            success: function (data) {
                //syncAgents(data);
                //setData("task", data);
                 //$.each(data, function (i, item) {
                   // document.getElementById('outputs').innerText = item.name;
                    
                //    console.log(item.name[0]);
                //    console.log(item.name[1]);
                //    console.log(item.email[0]);
                //    console.log(item.email[1]);
                    
                    //console.log(item.phone);
                    //    console.log(data[i].name[0]);
                    //    console.log(data[i].name[1]);
                    //    console.log(data[i].email[0]);
                    //    console.log(data[i].email[1]);
                    //    console.log(data[i].phone);
                // });

                RefreshStatusSheet("task", data);
                //document.getElementById('outputs').innerText = JSON.stringify(data);
                //setDataStatus("task");
            },
            error: function (xhr) {
                app.showNotification('Error:', xhr.responseText);
            }
        });
    }


    function syncAgents(userName) {
        var randomh = Math.random();
        
        $.ajax({
            type: "GET",
            //url: "http://buyfuturetoday.com/smb/SyncSalesAgent.php?email=" + userName,
            //url: "http://localhost:81/demo/check_login.json",
            url: "http://localhost:81/demo/sync_agents.php?x=" + randomh + "&email=" + userName + "",
            //url: "http://localhost:81/demo/sync_agents.php?x=" + randomh + "&user_id=" + userid + "&start_date=" + startnowUtc + "&end_date=" + endnowUtc + "",
            dataType: "json",
            success: function (data) {
               // app.showNotification('HHHHH');
                //syncAgents(data);
                //setData("task", data);
                //RefreshStatusSheet("task", data);
                SyncAgentStatusSheet("task", data);
                //RefreshAgentsSheet111("task", data);
                //setDataStatus("task");
                writeToPage(url);
            },
            error: function (xhr) {
                app.showNotification('Error:', xhr.responseText);
            }
        });
    }
    function syncAgents_old(data) {
        var agent0 = data.Agent0;
        var agent1 = data.Agent1;
        var agent2 = data.Agent2;
        var agent3 = data.Agent3;
        Excel.run(function (ctx) {
            var worksheets = ctx.workbook.worksheets;
            var wSalesAgent = worksheets.getItem("SalesAgent");
            wSalesAgent.getRange("A1:A1").values = agent0;
            wSalesAgent.getRange("A2:A2").values = agent1;
            wSalesAgent.getRange("A3:A3").values = agent2;
            wSalesAgent.getRange("A4:A4").values = agent3;
            $("#divSubmit").show();
            worksheets.getItem('Leads').activate();
            return ctx.sync();
        })
		.catch(function (error) {

		});
    }

    function createAppointment() {
        var jsonD = '{"to":"","time_zone":"Asia\/Calcutta","location":"Bangalore","event_start_date":"Thu Nov 26 18:35:00 GMT+05:30 2015","description":"for testing,0,http:\/\/smarter-biz.com\/audios\/1fe4ab16da66a606f183ffdbd43303cb.mp4,Customer Care","subject":"testing","user_id":"4425","event_end_date":"Thu Nov 26 19:35:00 GMT+05:30 2015","from":"testinfh@gmail.com"}';
        var firstUsedRange;
        Excel.run(function (ctx) {
            var worksheets = ctx.workbook.worksheets;
            var worksheet = worksheets.getItem('Leads');
            firstUsedRange = worksheet.getUsedRange();
            firstUsedRange.load("address, values");

            //worksheet.activate();            
            return ctx.sync().then(function () {
                //email = firstUsedRange.m_values[1][5];
            });

        }).catch(function (error) {
            app.showNotification("Error: " + error);
        });
        var test = '';
        //jsonData
        $.ajax({
            type: "POST",
            url: "http://buyfuturetoday.com/smb/createAppointment.php",
            dataType: "json",
            data: jsonD,
            success: function (data) {
                app.showNotification("Created Appointment from: " + data.from + " , Event Start Data and location: " + data.location + " ," + data.event_start_date);
            },
            error: function (xhr) {
                app.showNotification("Error: " + error);
            }
        });

    }

    function addNewSheets() {
        Excel.run(function (ctx) {
            var worksheets = ctx.workbook.worksheets;
            var worksheet = worksheets.getItem('Sheet1');
            worksheet.name = "Leads"
            //worksheet.activate();
            worksheets.add("SalesAgent");
            worksheets.add("Status");
            worksheets.add("Reports");
            return ctx.sync().then(function () {
                var worksheetStatus = worksheets.getItem('Status');
                worksheetStatus.activate();

                return ctx.sync().then(function () {
                    createStatusTable();
                    return ctx.sync().then(function () {
                        worksheet.activate();
                    });
                });
            });
        })
		.catch(function (error) {
		    //app.showNotification("Error: " + error);
		});
    }

    function createLeadsTable() {
        var myTable = new Office.TableData();
        myTable.headers = [["Sl no", "Company Name", "Contact Person", "Designation", "Contact Number", "Email ID", "Status", "Date to Be called", "Concern sales agent", "Address", "Submit Status(New/Existing)"]];
        myTable.rows = [];
        Office.context.document.setSelectedDataAsync(myTable, bindTable);
    }

    function createStatusTable() {
        var myTable = new Office.TableData();
        myTable.headers = [["Sl no", "Company Name", "Contact Person", "Designation", "Contact Number", "Email ID", "Status", "Called on", "Time", "Next Meeting", "Time", "Listen to Call Recordings", "Sales Personal", "Address"]];
        myTable.rows = [];
        Office.context.document.setSelectedDataAsync(myTable, bindTable);
    }


    function bindTable(e) {
        Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Table, { id: 'employeeTable' }, function (asyncResult) {

            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                var error = asyncResult.error;
                app.showNotification("Error", error.name + ": " + error.message);
            } else {
                // bind selection changes to react to cell selection
            }
        });
    }

})();