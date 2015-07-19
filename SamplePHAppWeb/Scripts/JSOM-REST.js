var hostweburl;
var appweburl;   

function GoJSOMREST() {
    hostweburl = decodeURIComponent(getQueryStringParameter("SPHostUrl"));
    appweburl = decodeURIComponent(getQueryStringParameter("SPAppWebUrl"));
    var scriptbase = hostweburl + "/_layouts/15/";

    $.getScript(scriptbase + "SP.Runtime.js",
        function() {
            $.getScript(scriptbase + "SP.js",
                function() {
                    $.getScript(scriptbase + "SP.RequestExecutor.js", execCrossDomainGetListRESTRequest);
                }
            );
        }
    );
}

function execCrossDomainGetListRESTRequest() {
    var executor = new SP.RequestExecutor(appweburl);
    // Get app web lists
    executor.executeAsync({
        url: appweburl + "/_api/web/lists",
        method: "GET",
        headers: { "Accept": "application/json;odata=verbose" },
        success: onRESTAppWebGetListSuccess,
        error: onRESTError
    });

    // Get host web lists using SP.AppContextSite
    executor.executeAsync({
        url: appweburl + "/_api/SP.AppContextSite(@target)/web/lists?@target='" + hostweburl + "'",
        method: "GET",
        headers: { "Accept": "application/json;odata=verbose" },
        success: onRESTHostWebGetListSuccess,
        error: onRESTError
    });
}

function onRESTAppWebGetListSuccess(data) {
    var jsonObject = JSON.parse(data.body);
    var results = jsonObject.d.results;
    for (var i = 0; i < results.length; i++) {
        var x = document.getElementById("ddJSOMRESTAppWebLists");
        var option = document.createElement("option");
        option.text = results[i].Title;
        option.value = results[i].Id;
        x.add(option);
    }
}

function onRESTHostWebGetListSuccess(data) {
    var jsonObject = JSON.parse(data.body);
    var results = jsonObject.d.results;
    for (var i = 0; i < results.length; i++) {
        var x = document.getElementById("ddJSOMRESTHostWebLists");
        var option = document.createElement("option");
        option.text = results[i].Title;
        option.value = results[i].Id;
        x.add(option);
    }
}

function btnJSOMRESTGetAppWebList_Click() {
    var x = document.getElementById("ddJSOMRESTAppWebLists");
    var listname = x.options[x.selectedIndex].text;

    var executor = new SP.RequestExecutor(appweburl);
    // Get app web lists
    executor.executeAsync({
        url: appweburl + "/_api/web/lists/GetByTitle('" + listname + "')/items",
        method: "GET",
        headers: { "Accept": "application/json;odata=verbose" },
        success: onRESTAppWebGetListDataSuccess,
        error: onRESTError
    });
}

function onRESTAppWebGetListDataSuccess(data) {
    var jsonObject = JSON.parse(data.body);
    var results = jsonObject.d.results;
    var itemstring = "";
    for (var i = 0; i < results.length; i++) {
        itemstring = itemstring + results[i].Title + "<br/>";
    }
    document.getElementById("divJSOMRESTAppWebListItems").innerHTML = itemstring;
}

function btnJSOMRESTGetHostWebList_Click() {
    var x = document.getElementById("ddJSOMRESTHostWebLists");
    var listname = x.options[x.selectedIndex].text;

    var executor = new SP.RequestExecutor(appweburl);
    // Get host web lists using the SP.AppContextSite
    executor.executeAsync({
        url: appweburl + "/_api/SP.AppContextSite(@target)/web/lists/GetByTitle('" + listname + "')/items?@target='" + hostweburl + "'",
        method: "GET",
        headers: { "Accept": "application/json;odata=verbose" },
        success: onRESTHostWebGetListDataSuccess,
        error: onRESTError
    });
}

function onRESTHostWebGetListDataSuccess(data) {
    var jsonObject = JSON.parse(data.body);
    var results = jsonObject.d.results;
    var itemstring = "";
    for (var i = 0; i < results.length; i++) {
        itemstring = itemstring + results[i].Title + "<br/>";
    }
    document.getElementById("divJSOMRESTHostWebListItems").innerHTML = itemstring;
}

function btnJSOMRESTAddHostWebListItem_Click() {
    var x = document.getElementById("ddJSOMRESTHostWebLists");
    var listname = x.options[x.selectedIndex].text;
    var itemType = GetItemTypeForListName(listname);
    var item = {
        "__metadata": { "type": itemType },
        "Title": $("#txtJSOMRESTNewHostWebListItem").val()
    };

    var executor = new SP.RequestExecutor(appweburl);
    // Get host web lists using the SP.AppContextSite
    executor.executeAsync({
        url: appweburl + "/_api/SP.AppContextSite(@target)/web/lists/GetByTitle('" + listname + "')/items?@target='" + hostweburl + "'",
        method: "POST",
        body: JSON.stringify(item),
        headers: {
            "accept": "application/json;odata=verbose",
            "content-type": "application/json;odata=verbose",
            "X-RequestDigest": $("#__REQUESTDIGEST").val()
        },
        success: onRESTHostWebAddListItemSuccess,
        error: onRESTError
    });
}

function GetItemTypeForListName(name) {
    return "SP.Data." + name.charAt(0).toUpperCase() + name.split(" ").join("").slice(1) + "ListItem";
}

function onRESTHostWebAddListItemSuccess(data) {
    $("#txtJSOMRESTNewHostWebListItem").val("");
    // Refresh the list items view
    btnJSOMRESTGetHostWebList_Click();
}

function onRESTError(data, errorCode, errorMessage) {
    alert("Could not complete cross-domain call: " + errorMessage);
}