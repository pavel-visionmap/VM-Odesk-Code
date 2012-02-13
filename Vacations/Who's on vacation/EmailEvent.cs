using System;
using System.Collections.Generic;
using System.Text;

namespace Who_s_on_vacation
{
    class EmailEvent
    {
        String _subject;
        DateTime _StartDate;
        DateTime _EndDate;

        
        public EmailEvent()
        {

        }

        public String Subject
        {
            //get
            //{
            //    return _subject;
            //}
            set
            {
                _subject = value;
            }
        }

        public DateTime StartDate
        {
            //get
            //{
            //    return _StartDate;
            //}
            set
            {
                _StartDate = value;
            }
        }

        public DateTime EndDate
        {
            //get
            //{
            //    return _EndDate;
            //}
            set
            {
                _EndDate = value;
            }
        }

        public override string ToString()
        {
            // ToDo: return email body

            System.Text.StringBuilder objBody = new StringBuilder();
            objBody.Append("Event: " + _subject );
            objBody.Append(Environment.NewLine);
            objBody.Append("Start: " + _StartDate.ToString(@"yyyy-MM-dd"));
            objBody.Append(Environment.NewLine);
            objBody.Append("End: " + _EndDate.ToString(@"yyyy-MM-dd"));
            return objBody.ToString();
            //return base.ToString();
        }




    }
}
