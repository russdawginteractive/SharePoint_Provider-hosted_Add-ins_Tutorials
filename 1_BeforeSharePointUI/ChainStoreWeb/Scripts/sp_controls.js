var hostweburl;

// Load the SharePoint Resources
$(document).ready(function () {
    hostweburl = decodeURIComponent(getQueryStringParameter("SPHostUrl"));

    // The Sharepoint js files URL are in the form:
    // web_url/_layouts/15/resource.js
    var scriptbase = hostweburl + "/_layouts/15/";

    // Load the js file and continue to the sucess handler.
    $.getScript(scriptbase + "SP.UI.Controls.js");
});

// Function to retrieve a query string value.
function getQueryStringParameter(paramToRetrieve) {
    var params = document.URL.split("?")[1].split("&");
    for (var i = 0; i < params.length; i = i + 1) {
        var singleParam = params[i].split("=");
        if (singleParam[0] === paramToRetrieve) {
            return singleParam[1];
        }
    }
}