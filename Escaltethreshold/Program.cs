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
using System.Windows.Forms;
using System.Net;
using Office = Microsoft.Office.Core;





namespace Escaltethreshold
{
    class Program
    {

        [STAThread]
        static void Main(string[] args)
        {
            #region Call to main Execution

            DateTime dt = DateTime.Now;

            Trace.WriteLine("<<<<<<<<<<<<<Application Started>>>>>>>>>>>>>>>>>>>" + dt + "", "TML");

            var p = new Program();

            string ipadr = p.GetIP();
            string strttime = dt.ToString();

            Trace.WriteLine("Checking for Network Connection --> " + dt + ".", "TML");
            p.checknetwork();

            Trace.WriteLine("Checking for Outlook process --> " + dt + ".", "TML");
            p.checkoutlook();

            Trace.WriteLine("Starting threshold Model -->" + dt + "", "TML");
            p.ThresholdListener();

            Trace.WriteLine("Endiing threshold Model -->" + dt + " ", "TML");
            Trace.WriteLine("Memory Consumption " + p.processcalc().ToString() + " \n ");

            string xendtime = dt.ToString();
            p.audittrail(ipadr, strttime, xendtime);

            Trace.WriteLine("<<<<<<<<<<<<<Application Ended>>>>>>>>>>>>>>>>>>> " + dt + "\n", "TML");

            #endregion

        }

        #region Main threshold Listener
        public void ThresholdListener()
        {


            Microsoft.Office.Interop.Outlook.Application myapp = null;
            Microsoft.Office.Interop.Outlook.MAPIFolder myInbox = null;
            Microsoft.Office.Interop.Outlook.NameSpace mapiNameSpace = null;
           


            DateTime thisDate = DateTime.Now.Date;
            CultureInfo culture = new CultureInfo("pt-BR");
            string CurrTime = thisDate.ToString("d", culture);

            MainClass m = new MainClass();



            //Write into system Event logs
            //String sSource = "Threshold Mail Listerner";
            //String sLog = "Application";
            //String sEvent = "TML Logs -->";

            //Check if Outlook process is running
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

            Outlook.Items _items = myInbox.Items;
            _items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(Items_ItemAdd);

            Outlook.Items UnReads = myInbox.Items.Restrict("[Unread]=true");

            /** start unread Mails**/
            if (UnReads.Count > 0 || myInbox.Items.Count > 0) //(myInbox.Items.Count > 0)//(xitem is Outlook.MailItem)//(myInbox.Items.Count > 0) //if checking mailcount greater than 0
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

                /** start Mails items **/
                if (isMailItem)
                {

                    /** start loops **/
                    for (int i = 1; i <= UnReads.Count; i++)    //(int i = 1; i <= myInbox.Items.Count; i++)
                    {
                        // var item = myInbox.Items[i];
                     
                        var item = UnReads[i];
                        subject = item.Subject;
                        body = item.Body;



                        /** Search for Keywords **/
                        if (subject.Contains("THRESHOLD") || body.Contains("Threshold") || body.Contains("Threshold Reporting - Nigeria"))
                        {

                            creationdate = (item.SentOn);
                            subject = subject.Replace('\'', '\"').ToUpper();
                 


                            Outlook.Recipients recips = item.Recipients;
                            foreach (Outlook.Recipient recip in recips)
                            {
                                Outlook.PropertyAccessor pa = recip.PropertyAccessor;

                                recepients = (recip.Name);
                            }

                            //Create Appointments
                            int X = m.createAppointment(subject, body, creationdate);


                            var result = MessageBox.Show(" Hello " + recepients + " \n New Appointment/Calendar with Subject " + subject + "", "",
                                             MessageBoxButtons.OK);




                            //generating the sql query
                            string isql = "INSERT INTO c##isng.THRESHOLD_TASK (TASK_SUBJECT ,TASK_START_DATE,TASK_STATUS,TASK_END_DATE,LAST_UPDATE_DATE ," +
                        "CREATION_DATE ,AST_UPDATE_BY, TASK_PRIORITY,TASK_ASSIGN1) Values ('" + subject + "', '" + creationdate + "', 'In Progress',  '" + (creationdate.AddHours(2)) + "'," +
                            " '" + CurrTime + "','" + CurrTime + "','TML', 'High' , '" + recepients + "')";


                            //insert into oracle database
                            int ires = m.insupddelClass(isql);


                            //send Text Messages
                            string xphone = "+2348029998152";//ConfigurationManager.AppSettings["phonenumber"];
                            string msg = " Hello you have an appointment with subject " + subject + " Please check your calendar";
                            m.sendtextmessage(xphone, msg);




                        } /** End search for key words **/


                    } /** End For loop**/

                    /** >>>>>>>>>>>>>>>>>>>>>>> Check for read Messages >>>>>>>>>>>>>>>>>>> **/


                    /** start loops **/
                    for (int i = 1; i <= myInbox.Items.Count; i++)
                    {
                        var item = myInbox.Items[i];
                        //var item = UnReads[i];
                        subject = item.Subject;
                        body = item.Body;

                        DateTime dt = DateTime.Now;

                        DateTime de = dt.Date;

                        DateTime dte = dt.AddDays(-7);
                        DateTime weekdate = dte.Date;

                        DateTime sentdate = (item.SentOn);
                        DateTime newsentdate = sentdate.Date;




                        for (var day = de; day <= sentdate; day = day.AddDays(-7))       //var day = from.Date; day.Date <= thru.Date; day = day.AddDays(1)
                        {

                            /** Search for Keywords **/
                            if (subject.Contains("THRESHOLD") || body.Contains("Threshold") || body.Contains("Threshold Reporting - Nigeria"))
                            {


                                creationdate = (item.SentOn);
                                subject = subject.Replace('\'', '\"').ToUpper();
                                // recepients = item.Recipients;


                                Outlook.Recipients recips = item.Recipients;
                                foreach (Outlook.Recipient recip in recips)
                                {
                                    Outlook.PropertyAccessor pa = recip.PropertyAccessor;

                                    recepients = (recip.Name);
                                }

                                //Create Appointments
                                int X = m.createAppointment(subject, body, creationdate);


                                var result = MessageBox.Show(" Hello " + recepients + " \n New Appointment/Calendar with Subject " + subject + "", "",
                                                 MessageBoxButtons.OK);
                                //if (result == DialogResult.OK)
                                //{
                                //    return;

                                //}




                                //generating the sql query
                                string isql = "INSERT INTO c##isng.THRESHOLD_TASK (TASK_SUBJECT ,TASK_START_DATE,TASK_STATUS,TASK_END_DATE,LAST_UPDATE_DATE ," +
                            "CREATION_DATE ,AST_UPDATE_BY, TASK_PRIORITY,TASK_ASSIGN1) Values ('" + subject + "', '" + creationdate + "', 'In Progress',  '" + (creationdate.AddHours(2)) + "'," +
                                " '" + CurrTime + "','" + CurrTime + "','TML', 'High' , '" + recepients + "')";


                                //insert into oracle database
                                int ires = m.insupddelClass(isql);


                                //send Text Messages
                                string xphone = "+2348029998152";//ConfigurationManager.AppSettings["phonenumber"];
                                string msg = " Hello you have an appointment with subject " + subject + " Please check your calendar";
                                m.sendtextmessage(xphone, msg);




                            } /** End search for key words **/

                        }

                    } /** End For loop**/

                    /** >>>>>>>>>>>>>>>>>>>>> End Check for read messages >>>>>>>>>>>>>>>>> **/




                } /** End of if ismailItem **/

            } /** End Unread Mails**/

        }
        #endregion

        #region garbage collection
        private void Items_ItemAdd(object Item)
        {
            MessageBox.Show("New Mail");
            throw new NotImplementedException();
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

        #region checking Network Connection
        public void checknetwork()
        {
            if (System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable() != true)
            {

                Trace.WriteLine("There is no network Connection ---> Please Check cable \n", "TML");
                var result = MessageBox.Show("There is a problem --> No Network Connection", "",
                    MessageBoxButtons.OK);
                if (result == DialogResult.OK)
                {
                    return;

                }
                return;
            }

        }
        #endregion

        #region checking New Email
        public void ThisListener_Startup(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Outlook.Application myapp = null;
            Microsoft.Office.Interop.Outlook.MAPIFolder myInbox = null;
            Outlook.Application Application = null;


            Outlook.MAPIFolder inbox = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

            try
            {
                inbox.Items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(this.Items_ItemAdd);

            }
            catch (Exception)
            {

                throw;
            }


        }
        #endregion

        #region NewMail event handler.
        public void outLookApp_NewMailEx(string EntryIDCollection)
        {
            MessageBox.Show("You've got a new mail whose EntryIDCollection is \n" + EntryIDCollection,
                    "NOTE", MessageBoxButtons.OK);
        }
        #endregion

        #region NewMail event handler.
        public object processcalc()
        {
            System.Threading.Thread.MemoryBarrier();

            var initialMemory = System.GC.GetTotalMemory(true);
            // body
            var somethingThatConsumesMemory = Enumerable.Range(0, 100000).ToArray();
            // end
            System.Threading.Thread.MemoryBarrier();
            var finalMemory = System.GC.GetTotalMemory(true);
            var consumption = finalMemory - initialMemory;
            return consumption;
        }
        #endregion

        #region Get Local IPAddress.
        public string GetIP()
        {
            string strHostName = "";
            strHostName = System.Net.Dns.GetHostName();
            IPHostEntry ipEntry = System.Net.Dns.GetHostEntry(strHostName);
            IPAddress[] addr = ipEntry.AddressList;
            return addr[addr.Length - 2].ToString();
        }
        #endregion

        #region Audit Trail.
        public void audittrail(string hostname, string starttime, string endtime)
        {

            Guid g = Guid.NewGuid();
            string isql = "INSERT " +
                 "INTO C##ISNG.APPMONITOR" +
                  "(" +
                  "HOSTNAME ," +
                  " CREATED_BY ," +
                   " APPENDTIME ," +
                   "CREATED_ON ," +
                   "LASTMODIFIED ," +
                   " LAST_UPDATED_BY ," +
                  "  LAST_UPDATED_ON ," +
                   " APPSTARTTIME ," +
                       " SESSIONID" +
                  " )" +
              " VALUES" +
                  " (" +
              " '" + hostname + "'," +
              " 'TML' ," +
               " '" + endtime + "'," +
                " sysdate," +
               " 'TML'," +
                " '" + endtime + "'," +
                    " '" + endtime + "'," +
           " '" + starttime + "'," +
            "  '" + g + "' " +
                 ")";
            MainClass m = new MainClass();
            int y = m.insupddelClass(isql);

            if (y == 1)
            {
                Trace.WriteLine("<<<<<<<<Audit Trail Updated>>>>>>>>>>");

            }
            else
            {

                Trace.WriteLine("<<<<<<<Audit Trail Not Updated>>>>>>>>");
            }

        }
        #endregion
    }
}