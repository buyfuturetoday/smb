
window.config = {
    tenant: 'buyfuturetoday.com',
    //clientId: '3c027d1a-2036-4c34-8f15-43190590dba6',
    //clientId: '44435a15-a1c1-4ffa-8121-41e17ad0547a', // smboffice365  buyfuturetoday.com
    clientId: 'fa36bc18-c771-4d53-807a-c2a540f8efe4',
    //tenant: 'udaypalsglobalsolutions.onmicrosoft.com',
    //clientId: 'dfdb1cda-cb69-4ec8-b2c4-4d259f278921',

    postLogoutRedirectUri: window.location.origin,
    endpoints: {
        officeGraph: 'https://graph.microsoft.com',
    },
    cacheLocation: 'localStorage'
};






/* Common app functionality */

var app = (function () {
    "use strict";

    var app = {};

    // Common initialization function (to be called from each page)
    app.initialize = function () {
        $('body').append(
            '<div id="notification-message">' +
                '<div class="padding">' +
                    '<div id="notification-message-close"></div>' +
                    '<div id="notification-message-header"></div>' +
                    '<div id="notification-message-body"></div>' +
                '</div>' +
            '</div>');

        $('#notification-message-close').click(function () {
            $('#notification-message').hide();
        });


        // After initialization, expose a common notification function
        app.showNotification = function (header, text) {
            $('#notification-message-header').text(header);
            $('#notification-message-body').text(text);
            $('#notification-message').slideDown('fast');
        };


        var authContext = new AuthenticationContext(config);

        var $userDisplay = $("#app-user");
        var $signInButton = $(".app-login");
        var $signOutButton = $(".app-logout");
        //Begin - Code Written By Srikanth on 01-02-2016
        var $TokenID = $("#results");

        //End

        // Check For & Handle Redirect From AAD After Login
        var isCallback = authContext.isCallback(window.location.hash);
        authContext.handleWindowCallback();

        if (isCallback && !authContext.getLoginError()) {
            window.location = authContext._getItem(authContext.CONSTANTS.STORAGE.LOGIN_REQUEST);

        }

        // Check Login Status, Update UI
        var user = authContext.getCachedUser();
        
      
        //Begin
        // Code Written by Srikanth on 01-02-2016
        var TokenID = authContext._getItem(authContext.CONSTANTS.STORAGE.IDTOKEN);
        //End

        if (user) {

           
            $userDisplay.val(user.userName);
            $TokenID.html(TokenID);
            $userDisplay.show();
            $signInButton.hide();
            $signOutButton.show();


            if ($("#app-user").val() != "")
            {
                loginO365(user.userName, TokenID);
                $("#downloadlink").hide();
            }

           

           

        } else {
           
            $userDisplay.empty();
            $userDisplay.hide();
            $signInButton.show();
            $signOutButton.hide();
        }

        // Register NavBar Click Handlers
        $signOutButton.click(function () {
            authContext.logOut();
        });
        $signInButton.click(function () {
            authContext.login();
           
        });






    };

    return app;
})();




function loginO365(user,Token) {
    var randomh = Math.random();
    var userName = user;
  var TokenID = Token;
    //password = $("#txtPassword").val();
    loginMethod = "o365";
    $("#errorBox").html("Please wait..");
    //app.showNotification(userName);
    if (userName == "") {
        //$("#txtName").focus();
        $("#errorBox").html("Please enter the User Name");
        return false;
    }
    //else if ($("#txtPassword").val() == "") {
    //    $("#lname").focus();
    //    $("#errorBox").html("Please Enter the Password");
    //    return false;
    //}
    var loginURL = "https://smarter-biz.com/check_login.json?x=" + randomh + "&email=" + userName + "&login_method=" + loginMethod + "&token=" + TokenID + "&ver=3";
   // var loginURL = "https://smarter-biz.com/check_login.json?x=" + randomh + "&email=testingusersmb1@gmail.com&password=testing12345&login_method=gplus&ver=3";
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
                $("#downloadlink").hide();
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
            $(".smbmodelwindow").hide();
            //app.showNotification('Error:', xhr.responseText);
            loginError("The user " + userName + " does not exists in SmarterSMB.");
        }
    });

}

$(document).ready(function () {

    if ($("#userName").val() !="")
    {
        $("#downloadlink").hide();
    }

});