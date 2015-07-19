using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Microsoft.SharePoint.Client;
using System.Net;
using System.Xml;
using System.IO;

namespace SamplePHAppWeb
{
    /// <summary>
    /// Much of the REST API code was adapted from the example code available here:
    /// http://code.msdn.microsoft.com/SharePoint-2013-Perform-335d925b/sourcecode?fileId=77297&pathId=614432390
    /// </summary>
    public partial class Default : System.Web.UI.Page
    {
        XmlNamespaceManager xmlnspm = new XmlNamespaceManager(new NameTable());

        protected void Page_PreInit(object sender, EventArgs e)
        {
            Uri redirectUrl;
            switch (SharePointContextProvider.CheckRedirectionStatus(Context, out redirectUrl))
            {
                case RedirectionStatus.Ok:
                    return;
                case RedirectionStatus.ShouldRedirect:
                    Response.Redirect(redirectUrl.AbsoluteUri, endResponse: true);
                    break;
                case RedirectionStatus.CanNotRedirect:
                    Response.Write("An error occurred while processing your request.");
                    Response.End();
                    break;
            }
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            // Add pertinent namespaces to the namespace manager. 
            xmlnspm.AddNamespace("atom", "http://www.w3.org/2005/Atom");
            xmlnspm.AddNamespace("d", "http://schemas.microsoft.com/ado/2007/08/dataservices");
            xmlnspm.AddNamespace("m", "http://schemas.microsoft.com/ado/2007/08/dataservices/metadata");
        }

        private void PopulateAppWebLists()
        {
            try
            {
                var spContext = SharePointContextProvider.Current.GetSharePointContext(Context);

                ClientContext clientContext = null;
                if (chkAppOnly.Checked)
                {
                    clientContext = spContext.CreateAppOnlyClientContextForSPAppWeb();
                }
                else
                {
                    clientContext = spContext.CreateUserClientContextForSPAppWeb();
                }

                using (clientContext)
                {
                    Web web = clientContext.Web;
                    clientContext.Load(web);
                    clientContext.ExecuteQuery();

                    ListCollection lists = web.Lists;
                    clientContext.Load<ListCollection>(lists);
                    clientContext.ExecuteQuery();

                    ddCSOMAppWebLists.Items.Clear();
                    foreach (List list in lists)
                    {
                        ddCSOMAppWebLists.Items.Add(list.Title);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteException(ex);
            }
        }

        private void PopulateHostWebLists()
        {
            try
            {
                var spContext = SharePointContextProvider.Current.GetSharePointContext(Context);

                ClientContext clientContext = null;
                if (chkAppOnly.Checked)
                {
                    clientContext = spContext.CreateAppOnlyClientContextForSPHost();
                }
                else
                {
                    clientContext = spContext.CreateUserClientContextForSPHost();
                }

                using (clientContext)
                {
                    Web web = clientContext.Web;
                    clientContext.Load(web);
                    clientContext.ExecuteQuery();

                    ListCollection lists = web.Lists;
                    clientContext.Load<ListCollection>(lists);
                    clientContext.ExecuteQuery();

                    ddCSOMHostWebLists.Items.Clear();
                    foreach (List list in lists)
                    {
                        ddCSOMHostWebLists.Items.Add(list.Title);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteException(ex);
            }
        }

        protected void btnCSOMGetAppWebList_Click(object sender, EventArgs e)
        {
            lblCSOMAppWebItems.Text = "";

            try
            {
                var spContext = SharePointContextProvider.Current.GetSharePointContext(Context);

                ClientContext clientContext = null;
                if (chkAppOnly.Checked)
                {
                    clientContext = spContext.CreateAppOnlyClientContextForSPAppWeb();
                }
                else
                {
                    clientContext = spContext.CreateUserClientContextForSPAppWeb();
                }

                using (clientContext)
                {
                    Web web = clientContext.Web;
                    ListCollection lists = web.Lists;
                    List selectedList = lists.GetByTitle(ddCSOMAppWebLists.SelectedValue);

                    CamlQuery camlQuery = CamlQuery.CreateAllItemsQuery(100);

                    Microsoft.SharePoint.Client.ListItemCollection listItems = selectedList.GetItems(camlQuery);
                    clientContext.Load<ListCollection>(lists);
                    clientContext.Load<List>(selectedList);
                    clientContext.Load<Microsoft.SharePoint.Client.ListItemCollection>(listItems);
                    clientContext.ExecuteQuery();

                    if (listItems.Count == 0)
                    {
                        lblCSOMAppWebItems.Text = "(No items in list)";
                    }
                    foreach (Microsoft.SharePoint.Client.ListItem item in listItems)
                    {
                        lblCSOMAppWebItems.Text += item["Title"] + "<br/>";
                    }
                }
            }
            catch (Exception ex)
            {
                WriteException(ex);
            }
        }

        protected void btnCSOMGetHostWebList_Click(object sender, EventArgs e)
        {
            lblCSOMHostWebItems.Text = "";

            try
            {
                var spContext = SharePointContextProvider.Current.GetSharePointContext(Context);

                ClientContext clientContext = null;
                if (chkAppOnly.Checked)
                {
                    clientContext = spContext.CreateAppOnlyClientContextForSPHost();
                }
                else
                {
                    clientContext = spContext.CreateUserClientContextForSPHost();
                }

                using (clientContext)
                {
                    Web web = clientContext.Web;
                    ListCollection lists = web.Lists;
                    List selectedList = lists.GetByTitle(ddCSOMHostWebLists.SelectedValue);

                    CamlQuery camlQuery = CamlQuery.CreateAllItemsQuery(100);

                    Microsoft.SharePoint.Client.ListItemCollection listItems = selectedList.GetItems(camlQuery);
                    clientContext.Load<ListCollection>(lists);
                    clientContext.Load<List>(selectedList);
                    clientContext.Load<Microsoft.SharePoint.Client.ListItemCollection>(listItems);
                    clientContext.ExecuteQuery();

                    if (listItems.Count == 0)
                    {
                        lblCSOMHostWebItems.Text = "(No items in list)";
                    }
                    foreach (Microsoft.SharePoint.Client.ListItem item in listItems)
                    {
                        lblCSOMHostWebItems.Text += item["Title"] + "<br/>";
                    }
                }
            }
            catch (Exception ex)
            {
                WriteException(ex);
            }
        }

        protected void btnCSOMNewHostWebListItem_Click(object sender, EventArgs e)
        {
            try
            {
                var spContext = SharePointContextProvider.Current.GetSharePointContext(Context);

                ClientContext clientContext = null;
                if (chkAppOnly.Checked)
                {
                    clientContext = spContext.CreateAppOnlyClientContextForSPHost();
                }
                else
                {
                    clientContext = spContext.CreateUserClientContextForSPHost();
                }

                using (clientContext)
                {
                    Web web = clientContext.Web;
                    ListCollection lists = web.Lists;
                    List selectedList = lists.GetByTitle(ddCSOMHostWebLists.SelectedValue);

                    ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                    Microsoft.SharePoint.Client.ListItem newItem = selectedList.AddItem(itemCreateInfo);
                    newItem["Title"] = txtCSOMNewHostWebListItem.Text;
                    newItem.Update();

                    clientContext.ExecuteQuery();

                    txtCSOMNewHostWebListItem.Text = "";
                    btnCSOMGetHostWebList_Click(this, null);
                }
            }
            catch (Exception ex)
            {
                WriteException(ex);
            }
        }

        private void WriteException(Exception ex)
        {
            lblExceptionInfo.Text = "<font color='red'><b>" + ex.Message + "</b><br/>" + ex.StackTrace + "</font>";
        }

        protected void LinkButton1_Click(object sender, EventArgs e)
        {
            MultiView1.SetActiveView(View1);
        }

        protected void LinkButton2_Click(object sender, EventArgs e)
        {
            MultiView1.SetActiveView(View2);
        }

        protected void LinkButton3_Click(object sender, EventArgs e)
        {
            MultiView1.SetActiveView(View3);
        }

        protected void LinkButton4_Click(object sender, EventArgs e)
        {
            MultiView1.SetActiveView(View4);
        }

        protected void btnLoadCSOM_Click(object sender, EventArgs e)
        {
            PopulateAppWebLists();
            PopulateHostWebLists();
        }

        protected void btnLoadCSOMREST_Click(object sender, EventArgs e)
        {
            PopulateAppWebListsREST();
            PopulateHostWebListsREST();
        }

        protected void btnCSOMRESTGetAppWebList_Click(object sender, EventArgs e)
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(Context);

            HttpWebRequest itemRequest = (HttpWebRequest)HttpWebRequest.Create(
                spContext.SPAppWebUrl + "/_api/web/lists/GetByTitle('" + ddCSOMRESTAppWebLists.SelectedValue + "')/items");
            itemRequest.Method = "GET";
            itemRequest.Accept = "application/atom+xml";
            itemRequest.ContentType = "application/atom+xml;type=entry";

            if (chkAppOnlyREST.Checked)
            {
                itemRequest.Headers.Add("Authorization", "Bearer " + spContext.AppOnlyAccessTokenForSPAppWeb);
            }
            else
            {
                itemRequest.Headers.Add("Authorization", "Bearer " + spContext.UserAccessTokenForSPAppWeb);
            }

            HttpWebResponse itemResponse = (HttpWebResponse)itemRequest.GetResponse();
            StreamReader itemReader = new StreamReader(itemResponse.GetResponseStream());
            var itemXml = new XmlDocument();
            itemXml.LoadXml(itemReader.ReadToEnd());

            var itemList = itemXml.SelectNodes("//atom:entry/atom:content/m:properties/d:Title", xmlnspm);

            lblCSOMRESTAppWebItems.Text = "";
            foreach (XmlNode itemTitle in itemList)
            {
                lblCSOMRESTAppWebItems.Text += itemTitle.InnerXml + "<br/>";
            }
        }

        protected void btnCSOMRESTGetHostWebList_Click(object sender, EventArgs e)
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(Context);

            HttpWebRequest itemRequest = (HttpWebRequest)HttpWebRequest.Create(
                spContext.SPHostUrl + "/_api/web/lists/GetByTitle('" + ddCSOMRESTHostWebLists.SelectedValue + "')/items");
            itemRequest.Method = "GET";
            itemRequest.Accept = "application/atom+xml";
            itemRequest.ContentType = "application/atom+xml;type=entry";

            if (chkAppOnlyREST.Checked)
            {
                itemRequest.Headers.Add("Authorization", "Bearer " + spContext.AppOnlyAccessTokenForSPHost);
            }
            else
            {
                itemRequest.Headers.Add("Authorization", "Bearer " + spContext.UserAccessTokenForSPHost);
            }

            HttpWebResponse itemResponse = (HttpWebResponse)itemRequest.GetResponse();
            StreamReader itemReader = new StreamReader(itemResponse.GetResponseStream());
            var itemXml = new XmlDocument();
            itemXml.LoadXml(itemReader.ReadToEnd());

            var itemList = itemXml.SelectNodes("//atom:entry/atom:content/m:properties/d:Title", xmlnspm);

            lblCSOMRESTHostWebItems.Text = "";
            foreach (XmlNode itemTitle in itemList)
            {
                lblCSOMRESTHostWebItems.Text += itemTitle.InnerXml + "<br/>";
            }
        }

        protected void btnCSOMRESTNewHostWebListItem_Click(object sender, EventArgs e)
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(Context);

            //Execute a REST request to get the form digest. All POST requests that change the state of resources on the host 
            //Web require the form digest in the request header. 
            HttpWebRequest contextinfoRequest =
                (HttpWebRequest)HttpWebRequest.Create(spContext.SPHostUrl + "/_api/contextinfo");
            contextinfoRequest.Method = "POST";
            contextinfoRequest.ContentType = "text/xml;charset=utf-8";
            contextinfoRequest.ContentLength = 0;

            if (chkAppOnlyREST.Checked)
            {
                contextinfoRequest.Headers.Add("Authorization", "Bearer " + spContext.AppOnlyAccessTokenForSPHost);
            }
            else
            {
                contextinfoRequest.Headers.Add("Authorization", "Bearer " + spContext.UserAccessTokenForSPHost);
            }

            HttpWebResponse contextinfoResponse = (HttpWebResponse)contextinfoRequest.GetResponse();
            StreamReader contextinfoReader = new StreamReader(contextinfoResponse.GetResponseStream(), System.Text.Encoding.UTF8);
            var formDigestXML = new XmlDocument();
            formDigestXML.LoadXml(contextinfoReader.ReadToEnd());
            var formDigestNode = formDigestXML.SelectSingleNode("//d:FormDigestValue", xmlnspm);
            string formDigest = formDigestNode.InnerXml;

            //Execute a REST request to get the list name and the entity type name for the list. 
            //The entity type name is the required type when you construct a request to add a list item. 
            HttpWebRequest listRequest = (HttpWebRequest)HttpWebRequest.Create(
                spContext.SPHostUrl + "/_api/web/lists/GetByTitle('" + ddCSOMRESTHostWebLists.SelectedValue + "')");
            listRequest.Method = "GET";
            listRequest.Accept = "application/atom+xml";
            listRequest.ContentType = "application/atom+xml;type=entry";

            if (chkAppOnlyREST.Checked)
            {
                listRequest.Headers.Add("Authorization", "Bearer " + spContext.AppOnlyAccessTokenForSPHost);
            }
            else
            {
                listRequest.Headers.Add("Authorization", "Bearer " + spContext.UserAccessTokenForSPHost);
            }

            HttpWebResponse listResponse = (HttpWebResponse)listRequest.GetResponse();
            StreamReader listReader = new StreamReader(listResponse.GetResponseStream());
            var listXml = new XmlDocument();
            listXml.LoadXml(listReader.ReadToEnd());

            var entityTypeNode = listXml.SelectSingleNode("//atom:entry/atom:content/m:properties/d:ListItemEntityTypeFullName", xmlnspm);
            var listNameNode = listXml.SelectSingleNode("//atom:entry/atom:content/m:properties/d:Title", xmlnspm);
            string entityTypeName = entityTypeNode.InnerXml;
            string listName = listNameNode.InnerXml;

            //Execute a REST request to add an item to the list. 
            string itemPostBody = "{'__metadata':{'type':'" + entityTypeName + "'}, 'Title':'" + txtCSOMRESTNewHostWebListItem.Text + "'}}";
            Byte[] itemPostData = System.Text.Encoding.ASCII.GetBytes(itemPostBody);

            HttpWebRequest itemRequest = (HttpWebRequest)HttpWebRequest.Create(
                spContext.SPHostUrl + "/_api/web/lists/GetByTitle('" + ddCSOMRESTHostWebLists.SelectedValue + "')/items");
            itemRequest.Method = "POST";
            itemRequest.ContentLength = itemPostBody.Length;
            itemRequest.ContentType = "application/json;odata=verbose";
            itemRequest.Accept = "application/json;odata=verbose";

            if (chkAppOnlyREST.Checked)
            {
                itemRequest.Headers.Add("Authorization", "Bearer " + spContext.AppOnlyAccessTokenForSPHost);
            }
            else
            {
                itemRequest.Headers.Add("Authorization", "Bearer " + spContext.UserAccessTokenForSPHost);
            }

            itemRequest.Headers.Add("X-RequestDigest", formDigest);
            Stream itemRequestStream = itemRequest.GetRequestStream();

            itemRequestStream.Write(itemPostData, 0, itemPostData.Length);
            itemRequestStream.Close();

            HttpWebResponse itemResponse = (HttpWebResponse)itemRequest.GetResponse();

            txtCSOMRESTNewHostWebListItem.Text = "";
            btnCSOMRESTGetHostWebList_Click(sender, e);
        }

        private void PopulateAppWebListsREST()
        {
            try
            {
                var spContext = SharePointContextProvider.Current.GetSharePointContext(Context);

                HttpWebRequest listRequest = (HttpWebRequest)HttpWebRequest.Create(spContext.SPAppWebUrl + "/_api/web/lists");
                listRequest.Method = "GET";
                listRequest.Accept = "application/atom+xml";
                listRequest.ContentType = "application/atom+xml;type=entry";

                if (chkAppOnlyREST.Checked)
                {
                    listRequest.Headers.Add("Authorization", "Bearer " + spContext.AppOnlyAccessTokenForSPAppWeb);
                }
                else
                {
                    listRequest.Headers.Add("Authorization", "Bearer " + spContext.UserAccessTokenForSPAppWeb);
                }

                HttpWebResponse listResponse = (HttpWebResponse)listRequest.GetResponse();
                StreamReader listReader = new StreamReader(listResponse.GetResponseStream());
                var listXml = new XmlDocument();
                listXml.LoadXml(listReader.ReadToEnd());

                var titleList = listXml.SelectNodes("//atom:entry/atom:content/m:properties/d:Title", xmlnspm);

                ddCSOMRESTAppWebLists.Items.Clear();
                foreach (XmlNode title in titleList)
                {
                    ddCSOMRESTAppWebLists.Items.Add(title.InnerXml);
                }
            }
            catch (Exception ex)
            {
                lblRESTExceptionInfo.Text = ex.Message + "<br/>" + ex.StackTrace;
            }
        }

        private void PopulateHostWebListsREST()
        {
            try
            {
                var spContext = SharePointContextProvider.Current.GetSharePointContext(Context);

                HttpWebRequest listRequest = (HttpWebRequest)HttpWebRequest.Create(spContext.SPHostUrl + "/_api/web/lists");
                listRequest.Method = "GET";
                listRequest.Accept = "application/atom+xml";
                listRequest.ContentType = "application/atom+xml;type=entry";

                if (chkAppOnlyREST.Checked)
                {
                    listRequest.Headers.Add("Authorization", "Bearer " + spContext.AppOnlyAccessTokenForSPHost);
                }
                else
                {
                    listRequest.Headers.Add("Authorization", "Bearer " + spContext.UserAccessTokenForSPHost);
                }

                HttpWebResponse listResponse = (HttpWebResponse)listRequest.GetResponse();
                StreamReader listReader = new StreamReader(listResponse.GetResponseStream());
                var listXml = new XmlDocument();
                listXml.LoadXml(listReader.ReadToEnd());

                var titleList = listXml.SelectNodes("//atom:entry/atom:content/m:properties/d:Title", xmlnspm);

                ddCSOMRESTHostWebLists.Items.Clear();
                foreach (XmlNode title in titleList)
                {
                    ddCSOMRESTHostWebLists.Items.Add(title.InnerXml);
                }
            }
            catch (Exception ex)
            {
                lblRESTExceptionInfo.Text = ex.Message + "<br/>" + ex.StackTrace;
            }
        }
    }
}