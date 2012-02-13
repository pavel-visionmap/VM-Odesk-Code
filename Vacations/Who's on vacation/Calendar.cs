#define Debug
using System;
using System.Collections.Generic;
using System.Text;

namespace Who_s_on_vacation
{
    class Calendar
    {
        bool sendEmail;
        String objEntryID;
        Microsoft.Office.Interop.Outlook.MAPIFolder objCalendar;
        Microsoft.Office.Interop.Outlook.Application objOutlook;
        int _Days;
        List<String> _EMails;
        List<EmailEvent> EmailEvents;
        String LOGPATH;

        
        public Calendar(int Days, List<String> EMails, String path, String ID)
        {
            sendEmail = false;
            objEntryID = ID;
            _Days = Days;
            _EMails = EMails;
            EmailEvents = new List<EmailEvent>();

            // set log path from logPath constant from Main()
            LOGPATH = path;
#if Debug
            try
            {
#endif
               
                objOutlook = new Microsoft.Office.Interop.Outlook.Application();
                Microsoft.Office.Interop.Outlook.NameSpace objNS = objOutlook.GetNamespace("MAPI");


                //objEntryID = "000000005BCDA19FC3AF564C9A7D910BCCF3D24642820000";
                //String objEntryID = "00000000E6BD34CD1C0FA04DBA1773A312FE25690100E15948BC84BC624E822DC6493E0F48BE0010CCEC4D970000";
                
                objCalendar = objOutlook.Session.GetFolderFromID(objEntryID, System.Type.Missing);    
            

               
#if Debug

            }
            catch (Exception ex)
            {
                //ToDo:
                // add handling for exception accessing Outlook
                System.IO.File.AppendAllText(LOGPATH,DateTime.Now.ToString() + ": " + ex.Message + Environment.NewLine);
                return;
            }
#endif

            findEvents();
            


        }

        void findEvents()
        {
            Microsoft.Office.Interop.Outlook.Items objEvents;
            objEvents = null;
            History objHistory;
            try
            {
                objHistory = new History();
            }
            catch (Exception ex)
            {
                System.IO.File.AppendAllText(LOGPATH, DateTime.Now.ToString() + " Error Loading History: " + ex.Message + Environment.NewLine);
                return;

            }
                        
                    

                // retrieve all events from Calendar
                objEvents = objCalendar.Items; // objFolder.Items;

                // retrict events to specified days.  Restrict by <= current date + number of days and > that calculation - 1 day.
                // Since will run every day on a schedualer, do not want to send events already sent in previouse days emails.
                // this way will only look for events after date searched previouse day up to number of days to search.

                DateTime restrictDate = DateTime.Now.AddDays(_Days);
            // Date is trestricted from between current date and date provided in params.
                

               // Original: String Filter = "[Start] <= '" + restrictDate.ToString("MM/dd/yyyy") + "' AND [Start] > '" + DateTime.Now.ToString("MM/dd/yyyy") + "'";
            //ISO8601: String Filter = "[Start] <= '" + String.Format("{0:yyyy/MM/dd h:m tt}", restrictDate) + "' AND [Start] > '" + String.Format("{0:M/d/yy h:m tt}", DateTime.Now) + "'"
               // suggested format for Outlook restrict function:
               String Filter = "[Start] <= '" + String.Format("{0:M/d/yyyy h:m tt}", restrictDate) + "' AND [Start] > '" + String.Format("{0:M/d/yyyy h:m tt}", DateTime.Now) + "'";

                // initial test just filtering for any event <= specified number of days
                // String Filter = "[Start] <= '" + restrictDate.ToString("MM/dd/yyyy hh:mm tt") + "'";

                objEvents = objEvents.Restrict(Filter);

                // iterate restricted event items, instantiate EmailEvent class and add to list
                for (int i = 1; i <= objEvents.Count; i++)
                {
                    
                    

#if Debug
            try
            {
#endif
                    Microsoft.Office.Interop.Outlook.AppointmentItem objEvent;
                    objEvent = (Microsoft.Office.Interop.Outlook.AppointmentItem)objEvents[i];

                    // double check since does not always filtering correctly apparently
                    if (objEvent.Start <= restrictDate)
                    {
                        //*** 02/03/12 check if already sent
                        // tests if history id already exists.
                        if (objHistory.FindByID(objEvent.EntryID) != true)
                        {
                            if (sendEmail == false)
                            {
                                sendEmail = true;
                            }
                            EmailEvent objEmailEvent = new EmailEvent();
                            objEmailEvent.EndDate = objEvent.End;
                            objEmailEvent.StartDate = objEvent.Start;
                            objEmailEvent.Subject = objEvent.Subject;




                            EmailEvents.Add(objEmailEvent);
                            // save event id
                            objHistory.AddEvent(objEvent.EntryID, objEvent.Start.ToString(@"yyyy-MM-dd"), objEvent.End.ToString(@"yyyy-MM-dd"));
                        }
                    }

#if Debug
            }
            catch (Exception ex)
            {
                //ToDo:
                // add exception handling for retrieving events
                // continue loop for now
                System.IO.File.AppendAllText(LOGPATH, DateTime.Now.ToString() + " Error retrieving Event: " + ex.Message + Environment.NewLine);
                continue;

            }
#endif


                }
            // persist data.
                SaveIDs(objHistory);





            if (sendEmail == true)
            {
                sendEmails();
            }


        }
        void SaveIDs(History objTemp)
        {
            System.Xml.Serialization.XmlSerializer objXML = new System.Xml.Serialization.XmlSerializer(typeof(History));
            System.IO.StreamWriter objWriter = new System.IO.StreamWriter("History.xml");
            objXML.Serialize(objWriter, objTemp);
            objWriter.Close();


        }
        void sendEmails()
        {
            // iterate events list for specified date range
            // build email body
            System.Net.Mail.SmtpClient objSmtp;
            System.Net.Mail.MailMessage objMailItem;
            //Microsoft.Office.Interop.Outlook.MailItem objMailItem;
#if Debug
            try
            {
#endif
                //objMailItem = (Microsoft.Office.Interop.Outlook.MailItem)objOutlook.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
                objMailItem = new System.Net.Mail.MailMessage();
               
                objMailItem.Subject = "Vacation Alert";
                
                //objMailItem.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatPlain;

                //populate recipient list from params
                System.Text.StringBuilder objRecipients = new StringBuilder();
                for (int i = 0; i < _EMails.Count; i++)
                {
                    

                    objMailItem.To.Add(_EMails[i]);
                    //objRecipients.Append(_EMails[i]);
                    //if (i < _EMails.Count - 1)
                    //{
                    //    objRecipients.Append(";");
                    //}

                }
                //objMailItem.To = objRecipients.ToString();


                System.Text.StringBuilder objMailText = new StringBuilder();
                for (int i = 0; i < EmailEvents.Count; i++)
                {
                    // To do: use outlook to send emails 


                    objMailText.Append(Environment.NewLine);
                    objMailText.Append(Environment.NewLine);
                    objMailText.Append(EmailEvents[i].ToString());
                    objMailText.Append(Environment.NewLine);
                    objMailText.Append(Environment.NewLine);
                    objMailText.Append("---------------------------------------------");



                }
                String SmtpClient = "smtp.east.cox.net";
                objSmtp = new System.Net.Mail.SmtpClient(SmtpClient);
                //System.Net.ICredentials objCred = new System.Net.ICredentialsByHost();
                String strFrom = "Administrator@gmail.com";
                objMailItem.From = new System.Net.Mail.MailAddress(strFrom);
                objSmtp.Credentials = System.Net.CredentialCache.DefaultNetworkCredentials;
                objMailItem.Body = objMailText.ToString();
#if Debug
            }
            catch (Exception ex)
            {
                System.IO.File.AppendAllText(LOGPATH, DateTime.Now.ToString() + " Error builing Email: " + ex.Message + Environment.NewLine);
                return;

            }
#endif

#if Debug
            try
            {
#endif        
                objSmtp.Send(objMailItem);
                //objMailItem.Send();
#if Debug
            }
            catch (Exception ex)
            {
                System.IO.File.AppendAllText(LOGPATH, DateTime.Now.ToString() + " Error sending Email: " + ex.Message + Environment.NewLine);
            }
#endif

           
                      


        }

        public static bool TryParseEmail(String Email)
        {
            //just tests if there is at least an '@' followed by a '.'. 
            String regExp = @"\w*\@\w*\.";
            System.Text.RegularExpressions.Regex objRegEx = new System.Text.RegularExpressions.Regex(regExp);
            return objRegEx.IsMatch(Email);

        }


        

    }
}
