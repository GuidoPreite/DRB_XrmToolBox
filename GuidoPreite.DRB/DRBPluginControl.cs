﻿using McTools.Xrm.Connection;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;


namespace GuidoPreite.DRB
{
    public partial class DRBPluginControl : PluginControlBase, IMessageBusHost, IHelpPlugin, IGitHubPlugin, IPayPalPlugin
    {
        #region Help
        public string HelpUrl => "https://github.com/GuidoPreite/DRB";
        #endregion

        #region GitHub
        public string RepositoryName => "DRB";
        public string UserName => "GuidoPreite";
        #endregion

        #region PayPal
        public string DonationDescription => "Donation for Dataverse REST Builder";
        public string EmailAccount => "[username]@gmail.com";
        #endregion

        public event EventHandler<MessageBusEventArgs> OnOutgoingMessage;

        private void DRBPluginControl_Load(object sender, EventArgs e)
        {
        }

        public override void UpdateConnection(IOrganizationService newService, ConnectionDetail detail, string actionName, object parameter)
        {
            base.UpdateConnection(newService, detail, actionName, parameter);
            CheckConnection();
        }

        private void SendFetchXMLToFXB(string fetchXML)
        {

            try
            {
                OnOutgoingMessage(this, new MessageBusEventArgs("FetchXML Builder") { TargetArgument = fetchXML });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"SendFetchXMLToFXB Error. Details: {ex.Message}", "Dataverse REST Builder", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CheckConnection()
        {
            Controls.Clear();
            try
            {
                CrmServiceClient serviceClient = ConnectionDetail.GetCrmServiceClient();
                string url = ConnectionDetail.WebApplicationUrl;
                string token = serviceClient.CurrentAccessToken;
                string version = serviceClient.ConnectedOrgVersion.ToString();

                if (string.IsNullOrWhiteSpace(token))
                {
                    MessageBox.Show($"The current connection is not compatible with the XrmToolBox version of Dataverse REST Builder.{Environment.NewLine}{Environment.NewLine}Inside this instance you can install the Managed Solution, click the «Help» button to visit the repository.", "Dataverse REST Builder", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, 0, "https://github.com/GuidoPreite/DRB/releases");
                }
                else
                {
                    LoadDRBWebView(url, token, version);
                }

            }
            catch (Exception ex)
            {
                Controls.Clear();
                MessageBox.Show($"Error during the connection. Details: {ex.Message}", "Dataverse REST Builder", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void LoadDRBWebView(string url, string token, string version)
        {
            string drbFolder = "GuidoPreite.DRB";
            string drbIndexFile = "drb_index.htm";


            // Check if WebView2 component is installed
            bool isWebView2Installed = true;
            try
            {
                string webView2Version = CoreWebView2Environment.GetAvailableBrowserVersionString();
            }
            catch (Exception ex)
            {
                isWebView2Installed = false;
            }

            if (isWebView2Installed == false)
            {
                MessageBox.Show($"Dataverse REST Builder requires the Microsoft Edge WebView2 runtime in order to run.{Environment.NewLine}{Environment.NewLine}Click the «Help» button to visit the Microsoft download page.{Environment.NewLine}{Environment.NewLine}After the runtime is installed please reopen XrmToolBox and try to launch again Dataverse REST Builder.", "Dataverse REST Builder", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, 0, "https://developer.microsoft.com/en-us/microsoft-edge/webview2/");
            }
            else
            {
                try
                {
                    WebView2 wvMain = new WebView2 { Dock = DockStyle.Fill };
                    Controls.Add(wvMain);
                    await wvMain.EnsureCoreWebView2Async();
                    XTBSettings xtbSettings = new XTBSettings { Token = token, Url = url, Version = version };
                    wvMain.CoreWebView2.AddHostObjectToScript("xtbSettings", xtbSettings);
                    string indexPath = Path.Combine(Paths.PluginsPath, drbFolder, drbIndexFile);
                    wvMain.Source = new Uri(indexPath);

                    wvMain.WebMessageReceived += WvMain_WebMessageReceived;

                    RefreshToken(false);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"LoadDRBWebView Error. Details: {ex.Message}", "Dataverse REST Builder", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void WvMain_WebMessageReceived(object sender, Microsoft.Web.WebView2.Core.CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                string message = e.TryGetWebMessageAsString();
                if (!string.IsNullOrWhiteSpace(message))
                {
                    WebViewPostMessage parsedMessage = JsonConvert.DeserializeObject<WebViewPostMessage>(message);
                    switch (parsedMessage.action)
                    {
                        case "sendtofxb":
                            SendFetchXMLToFXB(parsedMessage.data);
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"WvMain_WebMessageReceived Error. Details: {ex.Message}", "Dataverse REST Builder", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RefreshToken(bool normalRun = true)
        {
            try
            {
                CrmServiceClient serviceClient = ConnectionDetail.GetCrmServiceClient();
                string token = serviceClient.CurrentAccessToken;
                DateTimeOffset? checkExpirationTime = ExtractExpirationTimeFromJWT(token);
                if (checkExpirationTime.HasValue)
                {
                    DateTimeOffset expirationTime = checkExpirationTime.Value;
                    DateTimeOffset utcNow = DateTimeOffset.UtcNow;
                    if (expirationTime > utcNow)
                    {
                        // send the token to the loaded web page
                        foreach (Control control in Controls)
                        {
                            if (control is WebView2)
                            {
                                string jsToExecute = "DRB.Common.RefreshXTBToken('" + token + "');";
                                if (normalRun)
                                {
                                    Invoke((MethodInvoker)delegate { ((WebView2)control).ExecuteScriptAsync(jsToExecute); });
                                }
                            }
                        }
                        TimeSpan difference = expirationTime - utcNow;
                        Task.Delay(difference).ContinueWith(task => RefreshToken());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"RefreshToken Error. Details: {ex.Message}", "Dataverse REST Builder", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private DateTimeOffset? ExtractExpirationTimeFromJWT(string token)
        {
            try
            {
                string base64Payload = token.Split('.')[1];
                while (base64Payload.Length % 4 != 0) { base64Payload += '='; }
                byte[] bytePayload = Convert.FromBase64String(base64Payload);
                string stringPayload = Encoding.UTF8.GetString(bytePayload);
                dynamic jsonPayload = JsonConvert.DeserializeObject(stringPayload);
                long unixExpiration = jsonPayload.exp;
                return DateTimeOffset.FromUnixTimeSeconds(unixExpiration);
            }
            catch
            {
                // something went wrong with the extraction
                return null;
            }
        }

        public void OnIncomingMessage(MessageBusEventArgs message)
        {
            // do nothing
        }
    }

    public class WebViewPostMessage
    {
        public string action { get; set; }
        public string data { get; set; }
    }

    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComVisible(true)]
    public class XTBSettings
    {
        public string Token { get; set; }
        public string Url { get; set; }
        public string Version { get; set; }
    }
}