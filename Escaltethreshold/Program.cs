using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;
using System.Diagnostics;
using Microsoft.Exchange.WebServices.Data;
using System.DirectoryServices.ActiveDirectory;
using System.DirectoryServices.AccountManagement;
using System.Globalization;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using System.Configuration;


namespace Escaltethreshold
{
    class Program
    {

        [STAThread]
        static void Main(string[] args)
        {

            DateTime dt = DateTime.Now;

            Trace.WriteLine("Application started -->" + dt + "", "Threshold Mail Listener");
            
            var p = new Program();

            Trace.WriteLine("Checking for Outlook process --> " + dt +".", "TML");
            p.checkoutlook();

            Trace.WriteLine("Starting threshold Model -->."+ dt +"", "TML");
            p.ThresholdListener();

            
        }

        #region Main threshold Listener
        public void ThresholdListener()
        {


            Microsoft.Office.Interop.Outlook.Application myapp = null;
            Microsoft.Office.Interop.Outlook.MAPIFolder myInbox = null;
            Microsoft.Office.Interop.Outlook.NameSpace mapiNameSpace = null;
            Outlook.Application application = null;


            DateTime thisDate = DateTime.Now.Date;
            CultureInfo culture = new CultureInfo("pt-BR");
            string CurrTime = thisDate.ToString("d", culture);

            MainClass m = new MainClass();



            //Write into system Event logs
            String sSource = "Threshold Mail Listerner";
            String sLog = "Application";
            String sEvent = "TML Logs -->";

            //if (!EventLog.SourceExists(sSource))
            //    EventLog.CreateEventSource(sSource, sLog);



            //EventLog.WriteEntry(sSource, sEvent);

            if (Process.GetProcessesByName("OUTLOOK").Count() <= 0)
            {


                try
                {

                    Microsoft.Office.Interop.Outlook.Application myApp = new Microsoft.Office.Interop.Outlook.Application();
                   

                }
                catch (Exception ex)
                {

                    Microsoft.Office.Interop.Outlook.Application myApp = new Microsoft.Office.Interop.Outlook.Application();
                    //throw;
                }

            }
            else
            {

                try
                {
                    //if it is running , creating a new application instance 
                    myapp = Marshal.GetActiveObject("Outlook.Application") as Microsoft.Office.Interop.Outlook.Application;
                }
                catch (COMException)
                {
                    Type type = Type.GetTypeFromProgID("Outlook.Application");
                    myapp = (Outlook.Application)System.Activator.CreateInstance(type);

                }



            }

            mapiNameSpace = myapp.GetNamespace("MAPI");

            //selecting Inbox folder

            myInbox = mapiNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
            mapiNameSpace.SendAndReceive(false); //performs SendRecieve Operation without showing ProgrssDialog

            if (myInbox.Items.Count > 0) //if checking mailcount greater than 0
            {
                string subject = string.Empty;
                string attachments = string.Empty;
                string body = string.Empty;
                string senderName = string.Empty;
                string senderEmail = string.Empty;
                string recepients = string.Empty;
                DateTime creationdate;

                bool isMailItem = true;
                Microsoft.Office.Interop.Outlook.MailItem MyOutlookItem = null;

                try
                {
                    //if the item is not a mail Item Application will throw COM exception
                    MyOutlookItem = ((Microsoft.Office.Interop.Outlook.MailItem)myInbox.Items[4]);
                }
                catch (System.Exception ex)
                {
                    Console.WriteLine(ex.Message + "\nThere Item  is not a Mail Item", "Outlook Reader");
                    isMailItem = false;
                }
             

                if (isMailItem)
                {


                    for (int i = 1; i <= myInbox.Items.Count; i++)
                    {


                        var item = myInbox.Items[i];

                        subject = item.Subject;
                        body = item.Body;

                        if (subject.Contains("THRESHOLD") || body.Contains("Threshold") || body.Contains("Threshold Reporting - Nigeria"))
                        {

                            creationdate = (item.SentOn);
                            subject = subject.Replace('\'', '\"').ToUpper();

                            //Create Appointments

                            int X = m.createAppointment(subject, body, creationdate);


                            //insert into oracle database
                            string isql = "INSERT INTO THRESHOLD_TASK (TASK_SUBJECT ,TASK_START_DATE,TASK_STATUS,TASK_END_DATE,LAST_UPDATE_DATE ," +
                        "CREATION_DATE ,AST_UPDATE_BY, TASK_PRIORITY) Values ('" + subject + "', '" + creationdate + "', 'In Progress',  '" + (creationdate.AddHours(2)) + "'," +
                            " '" + CurrTime + "','" + CurrTime + "','TML', High' )";

                            int ires = m.insupddelClass(isql);





                        }


                    } //End For loop


                } // End of if ismailItem

            }

        }
        #endregion


        #region starting Outlook
        public void startsoutlook()
        {
            var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "outlook.exe"
                }
            };
            process.Start();
            process.WaitForInputIdle();


        }
        #endregion

        #region checking and starting outlook
        public void checkoutlook()
        {
            if (Process.GetProcessesByName("OUTLOOK").Count() <= 0)
            {

                this.startsoutlook();
            }

        }
        #endregion


    }
}