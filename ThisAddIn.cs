using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Web;
using OutlookAddIN_SendSms.ServiceReference1;
using IkcoSaleUtility;
using System.Diagnostics;
using System.Data.SqlClient;

namespace OutlookAddIN_SendSms
{
    public partial class ThisAddIn
    {
        Outlook.Inspectors inspectors;
        private Outlook.Application outlookApplication;
        private Outlook.NameSpace outlookNamespace;
        private Outlook.MAPIFolder inboxFolder;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            inspectors = this.Application.Inspectors;
            inspectors.NewInspector +=
            new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);



            outlookApplication = this.Application;
            outlookNamespace = outlookApplication.GetNamespace("MAPI");

            // Get the Inbox folder
            inboxFolder = outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

            // Subscribe to the NewMailEx event
            outlookApplication.NewMailEx += Application_NewMailEx;
        }
        void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                if (mailItem.EntryID == null)
                {
                    mailItem.Subject = "";
                    mailItem.Body = "";
                }

            }
        }
        public SendSmsContentRequest recipient { get; set; }

        void Application_NewMailEx(string EntryIDCollection)
        {
            string username1 = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            Outlook.NameSpace outlookNameSpace = this.Application.GetNamespace("MAPI");
            Outlook.MailItem mailItem = (Outlook.MailItem)outlookNameSpace.GetItemFromID(EntryIDCollection);
            string senderAddress = mailItem.SenderEmailAddress;
            Debug.WriteLine("Message");
            // Now you can check the sender's address
            // if (senderAddress == "azuredevops@ikco.ir")
            string mobile="";
            string emailAddress = "";

            if (senderAddress == "azuredevops@ikco.ir")
            {
                // Do something if the email is from azuredevops@ikco.ir
                

                Outlook.Recipients recipients = mailItem.Recipients;

                foreach (Outlook.Recipient recipient in recipients)
                {
                     emailAddress = recipient.Address;
                }
                
                //get mobile phone from user database based on email
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    string query = "SELECT [Mobile]  FROM db";
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                mobile = reader.GetString(0);
                            }
                            reader.Close();
                            connection.Close();
                        }
                    }
                }
                int startIndex = mailItem.Subject.IndexOf('>') + 1;
                int endIndex = mailItem.Subject.IndexOf(':');

                string title = mailItem.Subject.Substring(startIndex, endIndex - startIndex);
                recipient = new SendSmsContentRequest
                {
                    mobile = mobile,
                    context = "با سلام." + title + "در انتظار نظر شما مي باشد.",
                    dateSend = SendDate,
                    timeSend = SendTime,
                    userId = "",
                    objectId = "",
                    expiredDate = ExpireDate,
                    expiredTime = "23:59:59",
                };
               var result =  client.SendSmsContentAsync(recipient);
            }
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
