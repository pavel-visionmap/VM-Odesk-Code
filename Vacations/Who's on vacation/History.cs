#define Debug

using System;
using System.Collections.Generic;
using System.Text;

namespace Who_s_on_vacation
{
    public class HistoryItem
    {

        String m_FromDate;
        String m_ToDate;
        String m_ID;

        public HistoryItem()
        {


        }

        public String FromDate
        {
            get
            {
                return m_FromDate;
            }
            set
            {
                m_FromDate = value;
            }
        }
        public String ToDate
        {
            get
            {
                return m_ToDate;
            }
            set
            {
                m_ToDate = value;
            }
        }
        public String ID
        {
            get
            {
                return m_ID;
            }
            set
            {
                m_ID = value;
            }
        }
    }

    public class HistoryCollection
    {
        History objHistory;
        
        
        internal HistoryCollection(History objTemp)
        {
            objHistory = objTemp;
            

        }

        public HistoryCollection()
        {
            
        }

        public HistoryItem this[int Index]
        {

            get
            {
                return objHistory.objHistoryList[Index];
            }     
                            

        }

        
    }

    public class History
    {
        public List<HistoryItem> objHistoryList;
        public HistoryCollection HistoryItems;
        public History()
        {
#if Debug
            try
            {
#endif
                objHistoryList = new List<HistoryItem>();

                // instantiate list from xml
                GetList();
                HistoryItems = new HistoryCollection(this);
#if Debug
            }
            catch (Exception ex)
            {
                throw ex;
                

            }
#endif
        }

        public bool FindByID(String ID)
        {
            bool found = false;

            for (int i = 0; i < objHistoryList.Count; i++)
            {
                if (String.Compare(ID.Trim(), objHistoryList[i].ID.Trim(), true) == 0)
                {
                    found = true;
                    break;
                }
            

            }

            return found;


        }

        public void Delete(int delIndex)
        {
            //unused member
            objHistoryList.RemoveAt(delIndex);
        }

        public void AddEvent(String ID, String FromDate, String ToDate)
        {
            HistoryItem objHistory = new HistoryItem();
            objHistory.ID = ID;
            objHistory.FromDate = FromDate;
            objHistory.ToDate = ToDate;
            Add(objHistory);

            
        }

        private void Add(HistoryItem objTempAdd)
        {
            objHistoryList.Add(objTempAdd);
        }
        private void GetList()
        {
            System.Xml.XmlDocument objDoc = new System.Xml.XmlDocument();
            try
            {
                objDoc.Load("History.xml");
                System.Xml.XmlNodeList objHistoryNodes = objDoc.GetElementsByTagName("HistoryItem");
                foreach (System.Xml.XmlElement objNode in objHistoryNodes)
                {
                    HistoryItem objItem = new HistoryItem();
                    System.Xml.XmlNodeList objProperties = objNode.ChildNodes;
                    foreach (System.Xml.XmlElement objElement in objProperties)
                    {
                        if (String.Compare(objElement.Name, "ID", true) == 0)
                        {
                            objItem.ID = objElement.InnerText;

                        }
                        else if (String.Compare(objElement.Name, "FromDate", true) == 0)
                        {
                            objItem.FromDate = objElement.InnerText;
                        }
                        else if (String.Compare(objElement.Name, "ToDate", true) == 0)
                        {
                            objItem.ToDate = objElement.InnerText;


                        }



                    }
                    if (DateTime.Parse(objItem.ToDate) > DateTime.Now)
                    {
                        // will only add to list items whos too date has already expired.
                        objHistoryList.Add(objItem);
                    }


                }
            }
            catch (Exception ex)
            {


            }

        }
        public int Count
        {
            get
            {
                return objHistoryList.Count;
            }
        }

        


    }

    
}
