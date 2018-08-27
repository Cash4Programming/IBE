using System;
using System.Collections;
//using System.Collections.Generic;
using System.Text;

namespace InstructorBriefcaseExtractor.DAL
{
    //Didn't Work
    //public class Colleges : Dictionary<string,College>
    //{
    //    public Colleges()
    //    {
    //        YVCC YVCC = new YVCC();
    //        this.Add(YVCC.Key,YVCC.GetInfo);            
    //    }
        
    //    public College FindCollege(string Abbreviation)
    //    {
    //        College C = null;
    //        this.TryGetValue(Abbreviation, out C);
    //        return C;
    //    }
    //}

    public class myDictionary
    {
        public string Key 
        {
            get { return _College.Abbreviation; }            
        }

        private College _College;
        public College College
        {
            get { return _College; }
            set { _College = value; }
        }

        public string DropDownBoxName
        {
            get { return _College.DropDownBoxName; }
        }

        public string Abbreviation
        {
            get { return _College.Abbreviation; }
        }
    }

    public class Colleges : CollectionBase
    {
        public Colleges()
        {
            YVC YVCC = new YVC();
            myDictionary D = new myDictionary();
            D.College = YVCC.GetMyInfo;
            this.List.Add(D);
        }

        public string FindDropDownBoxName(string Abbreviation)
        {
            foreach (myDictionary D in this.List)
            {
                if (D.College.Abbreviation == Abbreviation)
                {
                    return D.DropDownBoxName;
                }
            }
            return "";
        }

        public College FindCollege(string Abbreviation)
        {
            foreach (myDictionary D in this.List)
            {
                if (D.College.Abbreviation == Abbreviation)
                {
                    return D.College;
                }
            }
            return null;
        }
    }
}
