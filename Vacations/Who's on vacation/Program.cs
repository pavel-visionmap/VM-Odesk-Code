#define Debug
using System;
using System.Collections.Generic;
using System.Text;

namespace Who_s_on_vacation
{
    class Program
    {
        static void Main(string[] args)
        {
           

            //set log path
            String LOGPATH = Environment.GetEnvironmentVariable("appdata") + @"\WhosOnVacation\log.txt";
            //create application folder
            if (System.IO.Directory.Exists(Environment.GetEnvironmentVariable("appdata") + @"\WhosOnVacation") == false)
            {
                System.IO.Directory.CreateDirectory(Environment.GetEnvironmentVariable("appdata") + @"\WhosOnVacation");

            }

            int Days = 0;
            List<String> EMails = new List<String>();
            bool result = false;


            //retrieve/validate command line params
            for (int i = 0; i < args.Length; i++)
            {
                if (i == 0)
                {
                    if (int.TryParse(args[i], out Days) == true)
                    {
                        result = true;

                    }
                    else
                    {
                        result = false;

                    }
                }
                else
                {
                    if (Calendar.TryParseEmail(args[i]) == true)
                    {

                        EMails.Add(args[i]);

                    }


                }

            }

            // Days must be present and there must be at least one email address
            String ID;
            XmlSettings objSettings;
#if Debug
             try
            {
            
#endif
                 
            System.Xml.Serialization.XmlSerializer objXML = new System.Xml.Serialization.XmlSerializer(typeof(XmlSettings));
            System.IO.FileStream objFS = new System.IO.FileStream(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\Settings.xml", System.IO.FileMode.Open);
            objSettings = (XmlSettings)objXML.Deserialize(objFS);
            ID = objSettings.CalendarID;

#if Debug
           

            }
            catch(Exception ex)
            {

                System.IO.File.AppendAllText(LOGPATH, DateTime.Now.ToString() + " InvalidArguements");
                return;


            
                        }
#endif


            if (result == true && EMails.Count > 0)
            {
                //do work
                Calendar objCalendar = new Calendar(Days, EMails, LOGPATH, ID);

            }
            else
            {
                //log if invalid command line parameters passed.
                // check if xml has info.
                
                if (objSettings.Days > 0)
                {
                    Days = objSettings.Days;
                    result = true;


                }
                else
                {

                    result = false;
                }

                List<String> objEmails = objSettings.Emails;
                if (result == true && objEmails.Count > 0)
                {
                    Calendar objCalendar = new Calendar(Days, objEmails, LOGPATH, ID);
                }
                else
                {
                    //add error handling
                    System.IO.File.AppendAllText(LOGPATH, DateTime.Now.ToString() + " InvalidArguements");
                    return;
                }

                System.IO.File.AppendAllText(LOGPATH, DateTime.Now.ToString() + " InvalidArguements");
                return;


            }
            
         
        }
    }
}
