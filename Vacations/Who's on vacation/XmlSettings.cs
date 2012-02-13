using System;
using System.Collections.Generic;
using System.Text;
//using System.Xml.Serialization;

namespace Who_s_on_vacation
{
    

    public class XmlSettings
    {

         String m_CalendarID;
        int m_Days;
        List<String> m_Emails;
        public XmlSettings()
        {
        }
        public String CalendarID
        {
            get
            {
                return m_CalendarID;
            }
            set
            {
                m_CalendarID = value;
            }
        }

        public int Days
        {
            get
            {
                return m_Days;
            }
            set
            {
                m_Days = value;
            }
        }

        public List<String> Emails
        {
            get
            {
                return m_Emails;
            }
            set
            {
                 m_Emails = value;
            }
        }


   

    }
}
