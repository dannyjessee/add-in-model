var appweburl;
var hostweburl;
var appWebList;
var hostWebList;
var appWebListColl;
var hostWebListColl;
var appWebListItems;
var hostWebListItems;

function GoJSOM() {
    hostweburl = decodeURIComponent(getQueryStringParameter("SPHostUrl"));
    appweburl = decodeURIComponent(getQueryStringParameter("SPAppWebUrl"));
    var scriptbase = hostweburl + "/_layouts/15/";

    $.getScript(scriptbase + "SP.Runtime.js",
        function () {
            $.getScript(scriptbase + "SP.js",
                function () {
                    $.getScript(scriptbase + "SP.RequestExecutor.js", execCrossDomainGetListRequest);
                }
            );
        }
    );
}

function execCrossDomainGetListRequest() {
    var clientContext = new SP.ClientContext(appweburl);
    var factory = new SP.ProxyWebRequestExecutorFactory(appweburl);
    clientContext.set_webRequestExecutorFactory(factory);
    var appWeb = clientContext.get_web();
    appWebListColl = appWeb.get_lists();
    clientContext.load(appWebListColl);
    clientContext.executeQueryAsync(onAppWebGetListSuccess, onJSOMError);
    
    var appContextSite = new SP.AppContextSite(clientContext, hostweburl);
    var hostWeb = appContextSite.get_web();
    hostWebListColl = hostWeb.get_lists();
    clientContext.load(hostWebListColl);
    clientContext.executeQueryAsync(onHostWebGetListSuccess, onJSOMError);
}

function onAppWebGetListSuccess(sender, args) {
    var listEnumerator = appWebListColl.getEnumerator();

    while (listEnumerator.moveNext()) {
        var x = document.getElementById("ddJSOMAppWebLists");
        var oList = listEnumerator.get_current();
        var option = document.createElement("option");
        option.text = oList.get_title();
        x.add(option);
    }
}

function onHostWebGetListSuccess(sender, args) {
    var listEnumerator = hostWebListColl.getEnumerator();

    while (listEnumerator.moveNext()) {
        var x = document.getElementById("ddJSOMHostWebLists");
        var oList = listEnumerator.get_current();
        var option = document.createElement("option");
        option.text = oList.get_title();
        x.add(option);
    }
}

function btnJSOMGetAppWebList_Click() {
    var x = document.getElementById("ddJSOMAppWebLists");
    var listname = x.options[x.selectedIndex].text;

    var clientContext = new SP.ClientContext(appweburl);
    var factory = new SP.ProxyWebRequestExecutorFactory(appweburl);
    clientContext.set_webRequestExecutorFactory(factory);
    var appWeb = clientContext.get_web();
    appWebList = appWeb.get_lists().getByTitle(listname);
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml("<View><RowLimit>50</RowLimit></View>");
    appWebListItems = appWebList.getItems(camlQuery);
    clientContext.load(appWebList);
    clientContext.load(appWebListItems);
    clientContext.executeQueryAsync(onAppWebGetListDataSuccess, onJSOMError);
}

function onAppWebGetListDataSuccess() {
    var itemstring = "";
    var listItemEnumerator = appWebListItems.getEnumerator();
    while (listItemEnumerator.moveNext()) {
        var oListItem = listItemEnumerator.get_current();
        itemstring = itemstring + oListItem.get_item("Title") + "<br/>";
    }
    document.getElementById("divJSOMAppWebListItems").innerHTML = itemstring;
}

function btnJSOMGetHostWebList_Click() {
    var x = document.getElementById("ddJSOMHostWebLists");
    var listname = x.options[x.selectedIndex].text;

    var clientContext = new SP.ClientContext(appweburl);
    var factory = new SP.ProxyWebRequestExecutorFactory(appweburl);
    clientContext.set_webRequestExecutorFactory(factory);
    var appContextSite = new SP.AppContextSite(clientContext, hostweburl);
    var hostWeb = appContextSite.get_web();
    hostWebList = hostWeb.get_lists().getByTitle(listname);
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml("<View><RowLimit>50</RowLimit></View>");
    hostWebListItems = hostWebList.getItems(camlQuery);
    clientContext.load(hostWebList);
    clientContext.load(hostWebListItems);
    clientContext.executeQueryAsync(onHostWebGetListDataSuccess, onJSOMError);
}

function onHostWebGetListDataSuccess(data) {
    var itemstring = "";
    var listItemEnumerator = hostWebListItems.getEnumerator();
    while (listItemEnumerator.moveNext()) {
        var oListItem = listItemEnumerator.get_current();
        itemstring = itemstring + oListItem.get_item("Title") + "<br/>";
    }
    document.getElementById("divJSOMHostWebListItems").innerHTML = itemstring;
}

function btnJSOMAddHostWebListItem_Click() {
    var x = document.getElementById("ddJSOMHostWebLists");
    var listname = x.options[x.selectedIndex].text;

    var clientContext = new SP.ClientContext(appweburl);
    var factory = new SP.ProxyWebRequestExecutorFactory(appweburl);
    clientContext.set_webRequestExecutorFactory(factory);
    var appContextSite = new SP.AppContextSite(clientContext, hostweburl);
    var hostWeb = appContextSite.get_web();
    hostWebList = hostWeb.get_lists().getByTitle(listname);
    var itemCreateInfo = new SP.ListItemCreationInformation();
    var hostWebListItem = hostWebList.addItem(itemCreateInfo);
    hostWebListItem.set_item("Title", $("#txtJSOMNewHostWebListItem").val());
    hostWebListItem.update();

    clientContext.executeQueryAsync(onHostWebAddListItemSuccess, onJSOMError);
}

function GetItemTypeForListName(name) {
    return "SP.Data." + name.charAt(0).toUpperCase() + name.split(" ").join("").slice(1) + "ListItem";
}

function onHostWebAddListItemSuccess(data) {
    $("#txtJSOMNewHostWebListItem").val("");
    // Refresh the list items view
    btnJSOMGetHostWebList_Click();
}

function onJSOMError(sender, args) {
    alert("Request failed. " + args.get_message() + "\n" + args.get_stackTrace());
}