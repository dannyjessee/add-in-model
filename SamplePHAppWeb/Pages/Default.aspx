<%@ Page MaintainScrollPositionOnPostback="true" Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="SamplePHAppWeb.Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>CSOM/JSOM/REST Demos</title>
    <script src="//ajax.aspnetcdn.com/ajax/4.0/1/MicrosoftAjax.js" type="text/javascript"></script>
    <script src="//ajax.aspnetcdn.com/ajax/jQuery/jquery-2.1.3.min.js" type="text/javascript"></script>
    <script type="text/javascript">
        function getQueryStringParameter(param) {
            var params = document.URL.split("?")[1].split("&");
            var strParams = "";
            for (var i = 0; i < params.length; i = i + 1) {
                var singleParam = params[i].split("=");
                if (singleParam[0] == param) {
                    return singleParam[1];
                }
            }
        }

        // Add the client chrome control
        $(document).ready(function () {
            hostweburl = decodeURIComponent(getQueryStringParameter("SPHostUrl"));
            var scriptbase = hostweburl + "/_layouts/15/";
            $.getScript(scriptbase + "SP.UI.Controls.js", renderChrome);
        });

        function renderChrome() {
            var options = {
                "appIconUrl": decodeURIComponent(getQueryStringParameter("SPHostLogoUrl")),
                "appTitle": "CSOM/JSOM/REST demos",
                "onCssLoaded": "chromeLoaded()"
            };

            var nav = new SP.UI.Controls.Navigation("chrome_ctrl_placeholder", options);
            nav.setVisible(true);
        }

        function chromeLoaded() {
            $("body").show();
        }
    </script>
    <script src="../Scripts/JSOM.js" type="text/javascript"></script>
    <script src="../Scripts/JSOM-REST.js" type="text/javascript"></script>
</head>
<body style="display:none; overflow-y: scroll">
    <form id="form1" runat="server">
        <div>
            <div id="chrome_ctrl_placeholder"></div>
            <asp:LinkButton ID="LinkButton1" runat="server" OnClick="LinkButton1_Click" Text="CSOM / C#"></asp:LinkButton>
            |
            <asp:LinkButton ID="LinkButton2" runat="server" OnClick="LinkButton2_Click" Text="C# / REST"></asp:LinkButton>
            |
            <asp:LinkButton ID="LinkButton3" runat="server" OnClick="LinkButton3_Click" Text="JSOM"></asp:LinkButton>
            |
            <asp:LinkButton ID="LinkButton4" runat="server" OnClick="LinkButton4_Click" Text="JS / REST"></asp:LinkButton>

            <asp:MultiView ID="MultiView1" runat="server" ActiveViewIndex="0">
                <asp:View ID="View1" runat="server">
                    <h1>CSOM / C#</h1>
                    <asp:Button ID="btnLoadCSOM" runat="server" OnClick="btnLoadCSOM_Click" Text="Get Data" /><br /><br />
                    <asp:CheckBox ID="chkAppOnly" runat="server" /> Use Add-in-only Authorization Policy
                    <h2>Add-in Web</h2>
                    List Name:
                    <asp:DropDownList ID="ddCSOMAppWebLists" runat="server"></asp:DropDownList>
                    <asp:Button ID="btnCSOMGetAppWebList" runat="server" OnClick="btnCSOMGetAppWebList_Click" Text="Get List Data" />
                    <hr />
                    <h3>Add-in Web List Items</h3>
                    <asp:Label ID="lblCSOMAppWebItems" runat="server"></asp:Label><br />
                    <br />
                    <hr />
                    <h2>Host Web</h2>
                    List Name:
                    <asp:DropDownList ID="ddCSOMHostWebLists" runat="server"></asp:DropDownList>
                    <asp:Button ID="btnCSOMGetHostWebList" runat="server" OnClick="btnCSOMGetHostWebList_Click" Text="Get List Data" />
                    <hr />
                    <h3>Host Web List Items</h3>
                    <asp:Label ID="lblCSOMHostWebItems" runat="server"></asp:Label><br />
                    <br />
                    Add New Item (Title):
                    <asp:TextBox ID="txtCSOMNewHostWebListItem" runat="server"></asp:TextBox>
                    <asp:Button ID="btnCSOMNewHostWebListItem" runat="server" OnClick="btnCSOMNewHostWebListItem_Click" Text="Add New Item" /><br />
                    <asp:Label ID="lblExceptionInfo" runat="server"></asp:Label>
                </asp:View>

                <asp:View ID="View2" runat="server">
                    <h1>C# / REST</h1>
                    <asp:Button ID="btnLoadCSOMREST" runat="server" OnClick="btnLoadCSOMREST_Click" Text="Get Data" /><br /><br />
                    <asp:CheckBox ID="chkAppOnlyREST" runat="server" /> Use Add-in-only Authorization Policy
                    <h2>Add-in Web</h2>
                    List Name:
                    <asp:DropDownList ID="ddCSOMRESTAppWebLists" runat="server"></asp:DropDownList>
                    <asp:Button ID="btnCSOMRESTGetAppWebList" runat="server" OnClick="btnCSOMRESTGetAppWebList_Click" Text="Get List Data" />
                    <hr />
                    <h3>Add-in Web List Items</h3>
                    <asp:Label ID="lblCSOMRESTAppWebItems" runat="server"></asp:Label><br />
                    <br />
                    <hr />
                    <h2>Host Web</h2>
                    List Name:
                    <asp:DropDownList ID="ddCSOMRESTHostWebLists" runat="server"></asp:DropDownList>
                    <asp:Button ID="btnCSOMRESTGetHostWebList" runat="server" OnClick="btnCSOMRESTGetHostWebList_Click" Text="Get List Data" />
                    <hr />
                    <h3>Host Web List Items</h3>
                    <asp:Label ID="lblCSOMRESTHostWebItems" runat="server"></asp:Label><br />
                    <br />
                    Add New Item (Title):
                    <asp:TextBox ID="txtCSOMRESTNewHostWebListItem" runat="server"></asp:TextBox>
                    <asp:Button ID="btnCSOMRESTNewHostWebListItem" runat="server" OnClick="btnCSOMRESTNewHostWebListItem_Click" Text="Add New Item" /><br />
                    <asp:Label ID="lblRESTExceptionInfo" runat="server"></asp:Label>
                </asp:View>

                <asp:View ID="View3" runat="server">
                    <h1>JSOM</h1>
                    <asp:Button ID="btnLoadJSOM" runat="server" OnClientClick="GoJSOM(); return false;" Text="Get Data" />
                    <h2>Add-in Web</h2>
                    List Name:
                    <select id="ddJSOMAppWebLists"></select>
                    <asp:Button ID="btnJSOMGetAppWebList" runat="server" OnClientClick="btnJSOMGetAppWebList_Click(); return false;" Text="Get List Data" />
                    <hr />
                    <h3>Add-in Web List Items</h3>
                    <div id="divJSOMAppWebListItems"></div>
                    <br />
                    <br />
                    <hr />
                    <h2>Host Web</h2>
                    List Name:
                    <select id="ddJSOMHostWebLists"></select>
                    <asp:Button ID="btnJSOMGetHostWebList" runat="server" OnClientClick="btnJSOMGetHostWebList_Click(); return false;" Text="Get List Data" />
                    <hr />
                    <h3>Host Web List Items</h3>
                    <div id="divJSOMHostWebListItems"></div>
                    <br />
                    <br />
                    Add New Item (Title):
                    <input type="text" id="txtJSOMNewHostWebListItem">
                    <asp:Button ID="btnJSOMAddHostWebListItem" runat="server" OnClientClick="btnJSOMAddHostWebListItem_Click(); return false;" Text="Add New Item" /><br />
                </asp:View>

                <asp:View ID="View4" runat="server">
                    <h1>JS / REST</h1>
                    <asp:Button ID="btnLoadJSOMREST" runat="server" OnClientClick="GoJSOMREST(); return false;" Text="Get Data" />
                    <h2>Add-in Web</h2>
                    List Name:
                    <select id="ddJSOMRESTAppWebLists"></select>
                    <asp:Button ID="btnJSOMRESTGetAppWebList" runat="server" OnClientClick="btnJSOMRESTGetAppWebList_Click(); return false;" Text="Get List Data" />
                    <hr />
                    <h3>Add-in Web List Items</h3>
                    <div id="divJSOMRESTAppWebListItems"></div>
                    <br />
                    <br />
                    <hr />
                    <h2>Host Web</h2>
                    List Name:
                    <select id="ddJSOMRESTHostWebLists"></select>
                    <asp:Button ID="btnJSOMRESTGetHostWebList" runat="server" OnClientClick="btnJSOMRESTGetHostWebList_Click(); return false;" Text="Get List Data" />
                    <hr />
                    <h3>Host Web List Items</h3>
                    <div id="divJSOMRESTHostWebListItems"></div>
                    <br />
                    <br />
                    Add New Item (Title):
                    <input type="text" id="txtJSOMRESTNewHostWebListItem">
                    <asp:Button ID="btnJSOMRESTAddHostWebListItem" runat="server" OnClientClick="btnJSOMRESTAddHostWebListItem_Click(); return false;" Text="Add New Item" /><br />
                </asp:View>
            </asp:MultiView>
        </div>
    </form>
</body>
</html>