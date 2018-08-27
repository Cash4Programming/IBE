using System;
using System.Collections;
using InstructorBriefcaseExtractor.BLL;

namespace InstructorBriefcaseExtractor.Model
{
    // This class contains all extracted quarters from the web page
    public class Quarters : CollectionBase
    {
        #region Events

        public event InformationalMessage MessageResults;
        protected virtual void SendMessage(object sender, EventArgs e)
        {
            MessageResults?.Invoke(sender, e);
        }

        #endregion

        // Creates an empty collection.
        public Quarters()
        { }

        public void AddQuartersByDate()
        {
            this.Clear();
            Quarter2MonthConfiguration QuarterConfiguration = new Quarter2MonthConfiguration(null);
            Quarters MyQuarters = QuarterConfiguration.GenerateQuarters();

            foreach (Quarter Q in MyQuarters)
            {
                this.Add(Q.Name, Q.YRQ);
            }

        }

        //public void GenerateQuarters()
        //{

        //    this.Clear();
        //    this.Add("FALL  2018", "B892");
        //    this.Add("WINTER  2019", "B893");
        //    this.Add("SPRING  2019", "B894");
        //    this.Add("SUMMER  2019", "B901");
        //    //this.Add("FALL  2019", "B902");
        //    //this.Add("WINTER  2020", "B903");
        //    //this.Add("SPRING  2020", "B904");
        //    //this.Add("SUMMER  2020", "C011");
        //    //this.Add("FALL  2020", "C012");
        //    //this.Add("WINTER  2021", "C013");
        //    //this.Add("SPRING  2021", "C014");
        //}

        public Quarter Add(String Name, String YRQ)
        {
            Quarter value = new Quarter();
            this.List.Add(value);
            value.Name = Name;
            value.YRQ = YRQ;
            return value;
        }

        public new void Clear()
        {
            this.List.Clear();
        }

        public new int Count()
        {
            return this.List.Count;
        }

        public Quarter this[String QuarterName]
        {
            get
            {
                foreach (Quarter myVar in this.List)
                {
                    if (myVar.Name == QuarterName) return myVar;
                }
                return null;
            }
        }

        public Quarter this[int index]
        {
            get { return (Quarter)this.List[index]; }
        }

        public int IndexOf(String QuarterName)
        {
            foreach (Quarter myVar in this.List)
            {
                if (myVar.Name == QuarterName) return this.List.IndexOf(myVar);
            }
            return -1;
        }

        public void Remove(String YRQ)
        {
            foreach (Quarter myVar in this.List)
            {
                if (myVar.YRQ == YRQ)
                {
                    this.List.Remove(myVar);
                    break;
                }
            }
        }

        public new void RemoveAt(int index)
        {
            this.List.RemoveAt(index);
        }
        
    }    // class Quarters
}
