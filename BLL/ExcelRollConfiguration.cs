using System;

using InstructorBriefcaseExtractor.Model;
using InstructorBriefcaseExtractor.Utility;

namespace InstructorBriefcaseExtractor.BLL
{
    public class ExcelRollConfiguration
    {

        private string xmlEXCELNodeNameLocation;

        private string ExportKey;
        private readonly string ExportWaitListKey = "ExportWaitList";
        private string ExportLabKey;
        private string ClassIncrementKey;
        private string ClassNameColumnLetterKey;
        private string FirstClassKey;
        private string FirstMondayDayCellKey;
        private string FirstNameColumnLetterKey;
        private string HeaderKey;
        //private string HeightKey;
        private string LastNameColumnLetterKey;
        private string MondayDateCellKey;

        public static readonly int ClassIncrementDefault = 40;
        public static readonly int FirstClassDefault = 3;
        public static readonly string FirstMondayDayCellDefault = "P1";
        public static readonly string FirstNameColumnLetterDefault = "C";
        public static readonly string HeaderDefault = "WEEK #";
        public static readonly float HeightDefault = 20.25F;
        public static readonly string LastNameColumnLetterDefault = "B";
        public static readonly string ClassNameColumnLetterDefault = "B";  // contains the irem number and class name
        public static readonly string MondayDateCellDefault = "D1";

        private readonly UserSettings UserSettings;

        private void SetStrings(UserSettings UserSettings)
        {
            if (UserSettings.Version == "1")
            {
                xmlEXCELNodeNameLocation = "EXCEL";
                ExportKey = "Roll.Export";
                ExportLabKey = "Roll.ExportLab";
                ClassIncrementKey = "Roll.ClassIncrement";
                ClassNameColumnLetterKey = "Roll.ClassNameColumnLetter";
                FirstClassKey = "Roll.FirstClass";
                FirstMondayDayCellKey = "Roll.FirstDayCell";
                FirstNameColumnLetterKey = "Roll.FirstNameColumnLetter";
                HeaderKey = "Roll.Header";
                //HeightKey = "Roll.Height";
                LastNameColumnLetterKey = "Roll.LastNameColumnLetter";
                MondayDateCellKey = "Roll.MondayDateCell";
            }
            else
            {
                xmlEXCELNodeNameLocation = "EXCEL-ROLL";
                ExportKey = "Export";
                ExportLabKey = "ExportLab";
                ClassIncrementKey = "ClassIncrement";
                ClassNameColumnLetterKey = "ClassNameColumnLetter";
                FirstClassKey = "FirstClass";
                FirstMondayDayCellKey = "FirstDayCell";
                FirstNameColumnLetterKey = "FirstNameColumnLetter";
                HeaderKey = "Header";
                //HeightKey = "Height";
                LastNameColumnLetterKey = "LastNameColumnLetter";
                MondayDateCellKey = "MondayDateCell";
            }

        }

        public ExcelRollConfiguration(UserSettings UserSettings)
        {
            SetStrings(UserSettings);
            this.UserSettings = UserSettings;
        }

        public ExcelRollSettings Load()
        {
            ExcelRollSettings ExcellRollSettings = new ExcelRollSettings();
            XMLhelper XML = new XMLhelper(UserSettings);

            try
            {
                ExcellRollSettings.Export = Convert.ToBoolean(XML.XMLReadFile(xmlEXCELNodeNameLocation, ExportKey).ToLower());
            }
            catch
            {
                ExcellRollSettings.Export = false;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, ExportKey, ExcellRollSettings.Export.ToString());
            }

            try
            {
                ExcellRollSettings.ExportWaitlist = Convert.ToBoolean(XML.XMLReadFile(xmlEXCELNodeNameLocation, ExportWaitListKey).ToLower());
            }
            catch
            {
                ExcellRollSettings.ExportWaitlist = false;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, ExportWaitListKey, ExcellRollSettings.ExportWaitlist.ToString());
            }


            try
            {
                ExcellRollSettings.ExportLab = Convert.ToBoolean(XML.XMLReadFile(xmlEXCELNodeNameLocation, ExportLabKey).ToLower());
            }
            catch
            {
                ExcellRollSettings.ExportLab = false;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, ExportLabKey, ExcellRollSettings.ExportLab.ToString());
            }

            try
            {
                ExcellRollSettings.ClassIncrement = Convert.ToInt32(XML.XMLReadFile(xmlEXCELNodeNameLocation, ClassIncrementKey));
            }
            catch
            {
                ExcellRollSettings.FirstClass = ClassIncrementDefault;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, ClassIncrementKey, ExcellRollSettings.ClassIncrement.ToString());
            }

            try
            {
                ExcellRollSettings.FirstClass = Convert.ToInt32(XML.XMLReadFile(xmlEXCELNodeNameLocation, FirstClassKey));
            }
            catch
            {
                ExcellRollSettings.FirstClass = FirstClassDefault;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, FirstClassKey, ExcellRollSettings.FirstClass.ToString());
            }

            ExcellRollSettings.ClassNameColumnLetter = XML.XMLReadFile(xmlEXCELNodeNameLocation, ClassNameColumnLetterKey);
            if (ExcellRollSettings.ClassNameColumnLetter == "")
            {
                ExcellRollSettings.ClassNameColumnLetter = ClassNameColumnLetterDefault;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, ClassNameColumnLetterKey, ExcellRollSettings.ClassNameColumnLetter);
            }

            ExcellRollSettings.FirstMondayDayCell = XML.XMLReadFile(xmlEXCELNodeNameLocation, FirstMondayDayCellKey);
            if (ExcellRollSettings.FirstMondayDayCell == "")
            {
                ExcellRollSettings.FirstMondayDayCell = FirstMondayDayCellDefault;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, FirstMondayDayCellKey, ExcellRollSettings.FirstMondayDayCell);
            }

            ExcellRollSettings.FirstNameColumnLetter = XML.XMLReadFile(xmlEXCELNodeNameLocation, FirstNameColumnLetterKey);
            if (ExcellRollSettings.FirstNameColumnLetter == "")
            {
                ExcellRollSettings.FirstNameColumnLetter = FirstNameColumnLetterDefault;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, FirstNameColumnLetterKey, ExcellRollSettings.FirstNameColumnLetter);
            }

            ExcellRollSettings.Header = XML.XMLReadFile(xmlEXCELNodeNameLocation, HeaderKey);
            if (ExcellRollSettings.Header == "")
            {
                ExcellRollSettings.Header = HeaderDefault;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, HeaderKey, ExcellRollSettings.Header);
            }

            //try
            //{
            //    ExcellRollSettings.Height = Convert.ToSingle(XML.XMLReadFile(xmlEXCELNodeNameLocation, HeightKey));
            //}
            //catch
            //{
            //    ExcellRollSettings.Height = HeightDefault;
            //    XML.XMLWriteFile(xmlEXCELNodeNameLocation, HeightKey, ExcellRollSettings.Height.ToString());
            //}

            ExcellRollSettings.LastNameColumnLetter = XML.XMLReadFile(xmlEXCELNodeNameLocation, LastNameColumnLetterKey);
            if (ExcellRollSettings.LastNameColumnLetter == "")
            {
                ExcellRollSettings.LastNameColumnLetter = LastNameColumnLetterDefault;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, LastNameColumnLetterKey, ExcellRollSettings.LastNameColumnLetter);
            }

            ExcellRollSettings.MondayDateCell = XML.XMLReadFile(xmlEXCELNodeNameLocation, MondayDateCellKey);
            if (ExcellRollSettings.MondayDateCell == "")
            {
                ExcellRollSettings.MondayDateCell = MondayDateCellDefault;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, MondayDateCellKey, ExcellRollSettings.MondayDateCell);
            }

            return ExcellRollSettings;
        }

        public void Save(ExcelRollSettings ExcellRollSettings)
        {
            XMLhelper XML = new XMLhelper(UserSettings);

            XML.XMLWriteFile(xmlEXCELNodeNameLocation, ExportKey, ExcellRollSettings.Export.ToString());
            XML.XMLWriteFile(xmlEXCELNodeNameLocation, ExportWaitListKey, ExcellRollSettings.ExportWaitlist.ToString());
            XML.XMLWriteFile(xmlEXCELNodeNameLocation, ExportLabKey, ExcellRollSettings.ExportLab.ToString());
            XML.XMLWriteFile(xmlEXCELNodeNameLocation, ClassIncrementKey, ExcellRollSettings.ClassIncrement.ToString());
            XML.XMLWriteFile(xmlEXCELNodeNameLocation, ClassNameColumnLetterKey, ExcellRollSettings.ClassNameColumnLetter);
            XML.XMLWriteFile(xmlEXCELNodeNameLocation, FirstClassKey, ExcellRollSettings.FirstClass.ToString());
            XML.XMLWriteFile(xmlEXCELNodeNameLocation, FirstMondayDayCellKey, ExcellRollSettings.FirstMondayDayCell);
            XML.XMLWriteFile(xmlEXCELNodeNameLocation, FirstNameColumnLetterKey, ExcellRollSettings.FirstNameColumnLetter);
            XML.XMLWriteFile(xmlEXCELNodeNameLocation, HeaderKey, ExcellRollSettings.Header);
            //XML.XMLWriteFile(xmlEXCELNodeNameLocation, HeightKey, ExcellRollSettings.Height.ToString());
            XML.XMLWriteFile(xmlEXCELNodeNameLocation, LastNameColumnLetterKey, ExcellRollSettings.LastNameColumnLetter);
            XML.XMLWriteFile(xmlEXCELNodeNameLocation, MondayDateCellKey, ExcellRollSettings.MondayDateCell);


        }
    }
}
