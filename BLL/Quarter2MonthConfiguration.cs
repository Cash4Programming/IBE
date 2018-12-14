﻿using System;
using InstructorBriefcaseExtractor.Model;
using InstructorBriefcaseExtractor.Utility;

namespace InstructorBriefcaseExtractor.BLL
{
    public class Quarter2MonthConfiguration
    {
        private readonly string xmlRootQuarter2MonthLocation = "QUARTERS";
        private readonly string SummerMonthKey = "SummerMonth";
        private readonly string FallMonthKey = "FallMonth";
        private readonly string WinterMonthKey = "WinterMonth";
        private readonly string SpringMonthKey = "SpringMonth";

        private readonly int SummerMonthDefault = 6;   //June
        private readonly int FallMonthDefault = 8;     //August
        private readonly int WinterMonthDefault = 12;  //December
        private readonly int SpringMonthDefault = 3;   //March

        private readonly UserSettings UserSettings;
        public Quarter2MonthConfiguration(UserSettings UserSettings)
        {
            this.UserSettings = UserSettings;
        }

        public Quarter2MonthSettings Load()
        {
            Quarter2MonthSettings Quarter2Month = new Quarter2MonthSettings();
            XMLhelper XML = new XMLhelper(UserSettings);

            try
            {
                Quarter2Month.SummerMonth = Convert.ToInt32(XML.XMLReadFile(xmlRootQuarter2MonthLocation, SummerMonthKey));
            }
            catch
            {
                Quarter2Month.SummerMonth = SummerMonthDefault;
                XML.XMLWriteFile(xmlRootQuarter2MonthLocation, SummerMonthKey, Quarter2Month.SummerMonth.ToString());
            }

            try
            {
                Quarter2Month.FallMonth = Convert.ToInt32(XML.XMLReadFile(xmlRootQuarter2MonthLocation, FallMonthKey));
            }
            catch
            {
                Quarter2Month.FallMonth = FallMonthDefault;
                XML.XMLWriteFile(xmlRootQuarter2MonthLocation, FallMonthKey, Quarter2Month.FallMonth.ToString());
            }

            try
            {
                Quarter2Month.WinterMonth = Convert.ToInt32(XML.XMLReadFile(xmlRootQuarter2MonthLocation, WinterMonthKey));
            }
            catch 
            {
                Quarter2Month.WinterMonth = WinterMonthDefault;
                XML.XMLWriteFile(xmlRootQuarter2MonthLocation, WinterMonthKey, Quarter2Month.WinterMonth.ToString());
            }

            try
            {
                Quarter2Month.SpringMonth = Convert.ToInt32(XML.XMLReadFile(xmlRootQuarter2MonthLocation, SpringMonthKey));
            }
            catch 
            {
                Quarter2Month.SpringMonth = SpringMonthDefault;
                XML.XMLWriteFile(xmlRootQuarter2MonthLocation, SpringMonthKey, Quarter2Month.SpringMonth.ToString());
            }
            
            return Quarter2Month;

        }

        public void Save(Quarter2MonthSettings Quarter2Month)
        {
            XMLhelper XML = new XMLhelper(UserSettings);

            XML.XMLWriteFile(xmlRootQuarter2MonthLocation, SummerMonthKey, Quarter2Month.SummerMonth.ToString());
            XML.XMLWriteFile(xmlRootQuarter2MonthLocation, FallMonthKey, Quarter2Month.FallMonth.ToString());
            XML.XMLWriteFile(xmlRootQuarter2MonthLocation, WinterMonthKey, Quarter2Month.WinterMonth.ToString());
            XML.XMLWriteFile(xmlRootQuarter2MonthLocation, SpringMonthKey, Quarter2Month.SpringMonth.ToString());
        }

        #region Base38 YRQ <--> Quarter Name Conversion

        ////// !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        //////
        ////// Change these parameters to affect what option is auto selected
        //////
        ////// Recognized Values For: QuarterNameTable
        ////// 1 (Summer)
        ////// 2 (Fall)
        ////// 3 (Winter)
        ////// 4 (Spring)
        //////
        ////// Modified from javascript by Michael Jenck 6/14/2011
        //////
        ////// Updated by Michael Jenck 10/21-24/2012 to fix a typo, Added a struct
        ////// Updated by Michael Jenck 08/04/2018 - Deleted unneed items and created
        ////// a quarter generator
        ////// 
        ////// https://lcc.ctc.edu/internal/waYRQGenerator.htm
        ////// 
        ////// 

        private string Base38List = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        private readonly static string[] QuarterNameTable = new string[] { "Summer", "Fall", "Winter", "Spring" };

        public string EncodeYRQFunction(int StartYear, int Quarter)
        {
            // these conditions prevent a negative index or an index out of range for the Base38List string
            if (StartYear < 1900) { throw new Exception("YRQENCODE must start from the year 1900 or later."); }
            if (StartYear > 2259) { throw new Exception("YRQENCODE must end before the year  2259."); }

            int ones = StartYear % 10;
            int YRQThirdChari;
            if (Quarter < 3)
            {
                YRQThirdChari = ones + 1;
            }
            else
            {
                YRQThirdChari = ones;
                ones -= 1;
            }
            
            if (YRQThirdChari > 9)
            {
                YRQThirdChari = 0;
            }
            int FirstThreedigits = (StartYear - ones) / 10;
            int Base38pos = FirstThreedigits - 190;

            string YRQFirstChar = Base38List[Base38pos].ToString();
            string YRQSecondChar = ones.ToString();
            string YRQThirdChar = YRQThirdChari.ToString();

            return YRQFirstChar + YRQSecondChar + YRQThirdChar + Quarter.ToString();
        }

        private string DecodeYRQFunction(int OutputType, string YRQ)
        {
            if (YRQ.Length != 4) { throw new Exception("YRQ must be 4 characters in length."); }

            int OutputOption = OutputType;
            if (OutputType < 0) { OutputOption = 0; }
            if (OutputType > 6) { OutputOption = 6; }

            int pos = Base38List.IndexOf(YRQ[0]);
            if (pos < 0) { throw new Exception("YRQ must start with a digit or a capital letter."); }

            int baseyear = (190 + pos) * 10;
            int startyear = baseyear + Convert.ToInt32(YRQ[1].ToString());
            int endyear = startyear + 1;
            string QuarterNumber = YRQ[3].ToString();
            int Quarter = Convert.ToInt32(QuarterNumber) - 1;
            if ((Quarter < 0) || (3 < Quarter)) { throw new Exception("Quarter must be a number between 1 and 4."); }
            string QuarterName = QuarterNameTable[Quarter];

            string ReturnVal = "";
            //// OutputType
            //// 0 - return: YYYY - YYYY: Quarter Name
            //// 1 - return: Quarter # - Quarter Name
            //// 2 - return: YYYY-YY
            //// 3 - return: YYYY - YY
            //// 4 - return: Quarter Name YYYY             "FALL 2011"
            //// 5 - return: YYYY
            //// 6 - return: Quarter Name (all Caps)
            switch (OutputOption)
            {
                case 0:
                    ReturnVal = startyear.ToString() + "-" + endyear.ToString() + " " + QuarterName;
                    break;
                case 1:
                    ReturnVal = QuarterNumber + " - " + QuarterName;
                    break;
                case 2:
                    ReturnVal = startyear.ToString() + "-" + endyear.ToString().Substring(0, 2);
                    break;
                case 3:
                    ReturnVal = startyear.ToString() + " - " + endyear.ToString().Substring(0, 2);
                    break;
                case 4:
                    if (Quarter >= 2) { ReturnVal = QuarterName + " " + endyear.ToString(); }
                    else ReturnVal = QuarterName + " " + startyear.ToString();
                    break;
                case 5:
                    ReturnVal = startyear.ToString();
                    break;
                case 6:
                    ReturnVal = QuarterName.ToUpper();
                    break;
                default:
                    break;
            }
            return ReturnVal;
        }

        #endregion

        public Quarters GenerateQuarters()
        {
            Quarters MyQuarters = new Quarters();
            int year = DateTime.Now.Year;
            int month = DateTime.Now.Month;
            int SummerFallYear = 0;
            int WinterSpringYear = 0;

            if ((1 <= month) && (month <= 6))
            {
                WinterSpringYear = year;
                SummerFallYear = year - 1;
            }
            else
            {
                SummerFallYear = year;
                WinterSpringYear = year + 1;
            }

            //https://lcc.ctc.edu/internal/waYRQGenerator.htm

            MyQuarters.Add("Summer " + SummerFallYear.ToString(), EncodeYRQFunction(SummerFallYear, 1));
            MyQuarters.Add("Fall " + SummerFallYear.ToString(), EncodeYRQFunction(SummerFallYear, 2));
            MyQuarters.Add("Winter " + WinterSpringYear.ToString(), EncodeYRQFunction(WinterSpringYear, 3));
            MyQuarters.Add("Spring " + WinterSpringYear.ToString(), EncodeYRQFunction(WinterSpringYear, 4));
            MyQuarters.Add("Summer " + WinterSpringYear.ToString(), EncodeYRQFunction(WinterSpringYear, 1));

            return MyQuarters;
        }

        public string CurrentQuarter(Quarter2MonthSettings Quarter2Month)
        {
            int year = DateTime.Now.Year;
            int month = DateTime.Now.Month;
            

            int SummerFallYear = 0;
            int WinterSpringYear = 0;

            if ((1 <= month) && (month <= 6))
            {
                WinterSpringYear = year;
                SummerFallYear = year - 1;
            }
            else
            {
                SummerFallYear = year;
                WinterSpringYear = year + 1;
            }

            if (Quarter2Month.WinterMonth == 1)
            {
                if ((Quarter2Month.WinterMonth <= month) && (month < Quarter2Month.SpringMonth)) return "Winter " + WinterSpringYear.ToString();
            }
            else
            {
                if (Quarter2Month.WinterMonth == month) return "Winter " + WinterSpringYear.ToString();
            }
            if ((Quarter2Month.SpringMonth <= month) && (month < Quarter2Month.SummerMonth)) return "Spring " + WinterSpringYear.ToString();
            if ((Quarter2Month.SummerMonth <= month) && (month < Quarter2Month.FallMonth)) return "Summer " + WinterSpringYear.ToString();

            return "Fall " + SummerFallYear.ToString();          
        }
    }
}
