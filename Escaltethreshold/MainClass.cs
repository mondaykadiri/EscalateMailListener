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
using System.IO;
using System.Configuration;
using Oracle.DataAccess.Client;
using System.Data;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Configuration;
using Twilio;



namespace Escaltethreshold
{
    class MainClass
    {
        ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
        

        #region find outlook items
        public void FindItems()
        {
            ItemView view = new ItemView(10);
            view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Ascending);
            view.PropertySet = new PropertySet(
                BasePropertySet.IdOnly,
                ItemSchema.Subject,
                ItemSchema.DateTimeReceived);

            FindItemsResults<Item> findResults = service.FindItems(
                WellKnownFolderName.Inbox,
                new SearchFilter.SearchFilterCollection(
                    LogicalOperator.Or,
                    new SearchFilter.ContainsSubstring(ItemSchema.Subject, "Threshold"),
                    new SearchFilter.ContainsSubstring(ItemSchema.Body, "Nigeria")),
                view);


            //return findResults
            //Console.WriteLine("Total number of items found: " + findResults.TotalCount.ToString());

            foreach (Item item in findResults)
            {
                // Do something with the item.
            }
        }
        #endregion 

        #region Email reply method
        public void ReplyToMessage(EmailMessage messageToReplyTo, string reply, string cc)
        {
            messageToReplyTo.Reply(reply, true /* replyAll */);
            // Or
            ResponseMessage responseMessage = messageToReplyTo.CreateReply(true);
            responseMessage.BodyPrefix = reply;
            responseMessage.CcRecipients.Add(cc);
            responseMessage.SendAndSaveCopy();
        }
         #endregion 

         #region Email forwarder
         public void ForwardMessage(EmailMessage messageToForward, string forward, string ccrec)
        {
            messageToForward.Forward(forward);
            // Or
            ResponseMessage responseMessage = messageToForward.CreateForward();
            responseMessage.BodyPrefix = forward;
            responseMessage.CcRecipients.Add(ccrec);
            responseMessage.SendAndSaveCopy();
        }
        #endregion

         #region Email forwarder
         public void sendtextmessage(string xTo, string xmsg)
         {
             if (System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable() == true)
             {
                 // Find your Account Sid and Auth Token at twilio.com/user/account 
                 //string AccountSid = "AC45b00a5504e242b8a486ebf4cad405c9";
                 //string AuthToken = "605ec28a7d811660710961fdc3a9f594";
                 //var twilio = new TwilioRestClient(AccountSid, AuthToken);

                 //var message = twilio.SendMessage("[From]", "[To]", null, null, null);
                 //Console.WriteLine(message.Sid); 

                 string AccountSid = "AC45b00a5504e242b8a486ebf4cad405c9";
                 string AuthToken = "605ec28a7d811660710961fdc3a9f594";

                 var twilio = new TwilioRestClient(AccountSid, AuthToken);
                // var message = twilio.SendMessage("+17314724935", xTo, xmsg);
                 var message = twilio.SendMessage("+17314724935", xTo, xmsg);                //("+17314724935", xTo, xmsg,null ,"", AccountSid); 
                 

                 //if (message.Sid != null)
                 //{
                   
                 //    Trace.WriteLine("The Messsage ID is "+ message.Sid+"");
                 //}
                 //else
                 //{
                 //    Trace.WriteLine( "Message Not Sent");

                 //}
             }
             else
             {
                 MessageBox.Show("There is a Network Issue", "", MessageBoxButtons.OK);

             }

         }
          #endregion


         //private void assigntaskexample(string xsubject)
        //{
        //    outlook.application application = globals.thisaddin.application;
        //    outlook.taskitem task = application.createitem(
        //        outlook.olitemtype.oltaskitem) as outlook.taskitem;
        //    task.subject = xsubject;// "tax preparation";
        //    task.startdate = datetime.parse("4/1/2007 8:00 am");
        //    task.duedate = datetime.parse("4/15/2007 8:00 am");
        //    outlook.recurrencepattern pattern =
        //        task.getrecurrencepattern();
        //    pattern.recurrencetype = outlook.olrecurrencetype.olrecursyearly;
        //    pattern.patternstartdate = datetime.parse("4/1/2007");
        //    pattern.noenddate = true;
        //    task.reminderset = true;
        //    task.remindertime = datetime.parse("4/1/2007 8:00 am");
        //    task.assign();
        //    task.recipients.add("accountant@example.com");
        //    task.recipients.resolveall();
        //    task.send();
        //}

         //private void createtodoitemexample()
         //{
         //    // date operations
         //    datetime today = datetime.parse("10:00 am");
         //    timespan duration = timespan.fromdays(1);
         //    datetime tomorrow = today.add(duration);
         //    outlook.mailitem mail = application.session.
         //        getdefaultfolder(
         //        outlook.oldefaultfolders.olfolderinbox).items.find(
         //        "[messageclass]='ipm.note'") as outlook.mailitem;
         //    mail.markastask(outlook.olmarkinterval.olmarktomorrow);
         //    mail.taskstartdate = today;
         //    mail.reminderset = true;
         //    mail.remindertime = tomorrow;
         //    mail.save();
         //}


        //private void AddAppointment()
        //{
        //    try
        //    {
        //        Outlook.AppointmentItem newAppointment =
        //            (Outlook.AppointmentItem)
        //        Microsoft.Office.Interop.Outlook.Application.CreateItem(Outlook.OlItemType.olAppointmentItem);
        //        newAppointment.Start = DateTime.Now.AddHours(2);
        //        newAppointment.End = DateTime.Now.AddHours(3);
        //        newAppointment.Location = "ConferenceRoom #2345";
        //        newAppointment.Body =
        //            "We will discuss progress on the group project.";
        //        newAppointment.AllDayEvent = false;
        //        newAppointment.Subject = "Group Project";
        //        newAppointment.Recipients.Add("Roger Harui");
        //        Outlook.Recipients sentTo = newAppointment.Recipients;
        //        Outlook.Recipient sentInvite = null;
        //        sentInvite = sentTo.Add("Holly Holt");
        //        sentInvite.Type = (int)Outlook.OlMeetingRecipientType
        //            .olRequired;
        //        sentInvite = sentTo.Add("David Junca ");
        //        sentInvite.Type = (int)Outlook.OlMeetingRecipientType
        //            .olOptional;
        //        sentTo.ResolveAll();
        //        newAppointment.Save();
        //        newAppointment.Display(true);
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("The following error occurred: " + ex.Message);
        //    }
        //}



        #region creating Appointment.
        public int createAppointment(string xsubject, string xbody, DateTime xsentdate)
		{
           

			try
			{
                
               
				Outlook.Application outlookApp = new Outlook.Application(); // creates new outlook app
				Outlook.AppointmentItem oAppointment = (Outlook.AppointmentItem)outlookApp.CreateItem(Outlook.OlItemType.olAppointmentItem); // creates a new appointment

				oAppointment.Subject = xsubject; // set the subject
				oAppointment.Body = xbody; // set the body
				oAppointment.Location = "My Office"; // set the location
				oAppointment.Start = xsentdate; // Set the start date 
				oAppointment.End = xsentdate.AddHours(3); // End date 
				oAppointment.ReminderSet = true; // Set the reminder
				oAppointment.ReminderMinutesBeforeStart = 15; // reminder time
			    oAppointment.Importance = Outlook.OlImportance.olImportanceHigh; // appointment importance
				oAppointment.BusyStatus = Outlook.OlBusyStatus.olBusy;
				oAppointment.Save();

				Outlook.MailItem mailItem = oAppointment.ForwardAsVcal(); 

                // email address to send to 
                mailItem.To = "mondaykadiri@gmail.com"; 
             
                mailItem.Send();

				//service.AutodiscoverUrl("monday.kadiri@ng.is.co.za");
				//Appointment appointment = new Appointment(service);
				//appointment.Subject = xsubject;     // "Meditation";
				//appointment.Body = xbody; // "My weekly relaxation time.";
				//appointment.Start = xsentdate; //new DateTime(2008, 1, 1, 18, 0, 0);
				//appointment.End = appointment.Start.AddHours(2);
				//// Occurs every weeks on Tuesday and Thursday
				////appointment.Recurrence = new Recurrence.WeeklyPattern( new DateTime(2008, 1, 1),2, DayOfWeek.Tuesday,DayOfWeek.Thursday);
				//appointment.Save();

				return 1;

			}
			catch (Exception e)
			{
                Trace.WriteLine(e.ToString());
				return 0;

			}



		}
        #endregion 


        #region creates outlook app instance.
        public Outlook.Application runapplicationoutlook()
        {



            Outlook.Application application = null;

            // Check whether there is an Outlook process running.
            if (Process.GetProcessesByName("OUTLOOK").Count() > 0)
            {

                try
                {
                    // If so, use the GetActiveObject method to obtain the process and cast it to an Application object.
                    application = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
                }
                catch (COMException ce)
                {
                    Type type = Type.GetTypeFromProgID("Outlook.Application");
                    application = (Outlook.Application)System.Activator.CreateInstance(type);
                    throw ce;
                }



            }
            else
            {

                // If not, create a new instance of Outlook and log on to the default profile.
                application = new Outlook.Application();
                Outlook.NameSpace nameSpace = application.GetNamespace("MAPI");
                nameSpace.Logon("", "", Type.Missing, Type.Missing);
                nameSpace = null;
            }

            // Return the Outlook Application object.
            return application;





        }
        #endregion


        #region insert update delete class
        public int insupddelClass(string osql)
        {
            try
            {
                var xconn = Properties.Settings.Default.ConnectionString;    //ConfigurationSettings.AppSettings["conOracle"];
                OracleConnection conn = new OracleConnection((xconn));

                string isql = osql;

                OracleCommand cmd = new OracleCommand(isql, conn);
                conn.Open();
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();
                Trace.WriteLine("Information Saved in Database \n");
                return 1;


                conn.Close();
                conn.Dispose();


            }
            catch (Exception ex)
            {
                //string elog = Convert.ToString(ex);
                //this.writelog(elog);
                Trace.WriteLine("Error Message",ex.ToString()+"\n");
                return 0;

            }
        }
        #endregion


        //private void ReminderExample()
        //{
        //    Outlook.AppointmentItem appt = Application.CreateItem(
        //        Outlook.OlItemType.olAppointmentItem)
        //        as Outlook.AppointmentItem;
        //    appt.Subject = "Wine Tasting";
        //    appt.Location = "Napa CA";
        //    appt.Sensitivity = Outlook.OlSensitivity.olPrivate;
        //    appt.Start = DateTime.Parse("10/21/2006 10:00 AM");
        //    appt.End = DateTime.Parse("10/21/2006 3:00 PM");
        //    appt.ReminderSet = true;
        //    appt.ReminderMinutesBeforeStart = 120;
        //    appt.Save();
        //}


        //public void maillistener()
        //{
        //    Outlook.Application outlookApp = new Outlook.Application(); //
        //    //  Outlook.ApplicationClass outLookApp = new Outlook.ApplicationClass();

        //    // Ring up the new message event.
        //    outlookApp.NewMail += ApplicationEvents_11_NewMailEventHandler(outLookApp_NewMailEx);
        //    Console.WriteLine("Please wait for new messages...");
        //    Console.ReadLine();


        //}

        #region NewMail event handler.
        private static void outLookApp_NewMailEx(string EntryIDCollection)
        {
            MessageBox.Show("You've got a new mail whose EntryIDCollection is \n" + EntryIDCollection,
                    "NOTE", MessageBoxButtons.OK);
        }
        #endregion





    }
}
