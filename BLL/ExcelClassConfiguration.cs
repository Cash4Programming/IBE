using System;
using System.Globalization;
using System.IO;

using InstructorBriefcaseExtractor.Model;
using InstructorBriefcaseExtractor.Utility;

namespace InstructorBriefcaseExtractor.BLL
{
    public class ExcelClassConfiguration
    {
        private readonly string xmlEXCELNodeNameLocation = "EXCEL";

        private readonly string ExportKey = "Export";
        private readonly string ExportWaitListKey = "ExportWaitList";
        private readonly string ExportMiddleInitialKey = "ExportMI";
        private readonly string ExportSIDKey = "ExportSID";
        private readonly string ExportSIDLast4Key = "ExportSIDLast4";
        private readonly string FirstNameColumnLetterKey = "FirstNameColumnLetter";
        private readonly string FirstStudentKey = "FirstStudent";
        private readonly string ItemCellKey = "ItemCell";
        private readonly string LastNameColumnLetterKey = "LastNameColumnLetter";
        private readonly string SIDColumnLetterKey = "SIDColumnLetter";
        private readonly string SIDLast4ColumnLetterKey = "SIDLast4ColumnLetter";
        private string TemplateNameKey;
        private readonly string TemplateDirectoryKey = "TemplateDirectory";
        private string SaveFileNameKey;
        private readonly string SaveFileDirectoryKey = "SaveFileDirectory";
        private readonly string TemplateCopyNameIncludesQuarterKey = "TemplateCopyNameIncludesQuarter";
        private readonly string OptHead1ColumnLetterKey = "OptHead1ColumnLetter";
        private readonly string ExportoptHead1Key = "ExportoptHead1";
        private readonly string OptHead2ColumnLetterKey = "OptHead2ColumnLetter";
        private readonly string ExportoptHead2Key = "ExportoptHead2";
        private readonly string OptHead3ColumnLetterKey = "OptHead3ColumnLetter";
        private readonly string ExportoptHead3Key = "ExportoptHead3";
        public static readonly string OptHead1ColumnLetterDefault = "F";        
        public static readonly string OptHead2ColumnLetterDefault = "F";
        public static readonly string OptHead3ColumnLetterDefault = "F";
        public static readonly string OptHead1NameDefault = "OptHead1";
        public static readonly string OptHead2NameDefault = "OptHead2";
        public static readonly string OptHead3NameDefault = "OptHead3";
        public static readonly string OptHeadWaitlistNameDefault = "OptHeadWaitList";
        public static readonly string SaveAsFileNameDafault = "Grades.xls";
        public static readonly string TemplateFileNameDafault = "GradeBook_Template_Percent.xls";
        public static readonly string FirstNameColumnLetterDefault = "E";
        public static readonly int FirstStudentDefault = 23;
        public static readonly string ItemCellDefault = "A1";
        public static readonly string LastNameColumnLetterDefault = "D";
        public static readonly string SIDColumnLetterDefault = "C";
        public static readonly string SIDLast4ColumnLetterDefault = "G";

        //private readonly int Version;
        private readonly UserSettings UserSettings;

        private void SetStrings(UserSettings UserSettings)
        {
            if (UserSettings.Version == "1")
            {
                SaveFileNameKey = "TemplateCopyName";
                TemplateNameKey = "TemplateName";
            }
            else
            {
                TemplateNameKey = "TemplateFileName";
                SaveFileNameKey = "SaveFileName";
            }
        }

        public ExcelClassConfiguration(UserSettings UserSettings)
        {
            SetStrings(UserSettings);

            //this.Version = Version;
            this.UserSettings = UserSettings;
        }

        public static string GenerateSaveFileName(ExcelClassSettings ExcelClassSettings, Quarter Quarter)
        {
            return ExcelClassSettings.SaveFileDirectory + Quarter.Name.ToUpper() + "_" + ExcelClassSettings.SaveFileName;
        }

        public ExcelClassSettings Load()
        {
            ExcelClassSettings ExcelClassSettings = new ExcelClassSettings();

            XMLhelper XML = new XMLhelper(UserSettings);

            try
            {
                ExcelClassSettings.Export = Convert.ToBoolean(XML.XMLReadFile(xmlEXCELNodeNameLocation, ExportKey).ToLower());
            }
            catch
            {
                ExcelClassSettings.Export = false;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, ExportKey, ExcelClassSettings.Export.ToString(CultureInfo.CurrentCulture));
            }

            try
            {
                ExcelClassSettings.ExportWaitlist = Convert.ToBoolean(XML.XMLReadFile(xmlEXCELNodeNameLocation, ExportWaitListKey).ToLower());
            }
            catch
            {
                ExcelClassSettings.ExportWaitlist = false;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, ExportWaitListKey, ExcelClassSettings.ExportWaitlist.ToString(CultureInfo.CurrentCulture));
            }

            try
            {
                ExcelClassSettings.ExportMiddleInitial = Convert.ToBoolean(XML.XMLReadFile(xmlEXCELNodeNameLocation, ExportMiddleInitialKey).ToLower());
            }
            catch
            {
                ExcelClassSettings.ExportMiddleInitial = false;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, ExportMiddleInitialKey, ExcelClassSettings.ExportMiddleInitial.ToString(CultureInfo.CurrentCulture));
            }

            try
            {
                ExcelClassSettings.ExportSID = Convert.ToBoolean(XML.XMLReadFile(xmlEXCELNodeNameLocation, ExportSIDKey).ToLower());
            }
            catch
            {
                ExcelClassSettings.ExportSID = true;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, ExportSIDKey, ExcelClassSettings.ExportSID.ToString(CultureInfo.CurrentCulture));
            }

            try
            {
                ExcelClassSettings.ExportSIDLast4 = Convert.ToBoolean(XML.XMLReadFile(xmlEXCELNodeNameLocation, ExportSIDLast4Key).ToLower());
            }
            catch
            {
                ExcelClassSettings.ExportSIDLast4 = false;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, ExportSIDLast4Key, ExcelClassSettings.ExportSIDLast4.ToString(CultureInfo.CurrentCulture));
            }

            try
            {
                ExcelClassSettings.FirstStudent = Convert.ToInt32(XML.XMLReadFile(xmlEXCELNodeNameLocation, FirstStudentKey).ToLower());                
            }
            catch
            {
                ExcelClassSettings.FirstStudent = FirstStudentDefault;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, FirstStudentKey, ExcelClassSettings.FirstStudent.ToString(CultureInfo.CurrentCulture));
            }

            try
            {
                ExcelClassSettings.SaveFileNameIncludesQuarter = Convert.ToBoolean(XML.XMLReadFile(xmlEXCELNodeNameLocation, TemplateCopyNameIncludesQuarterKey).ToLower());
            }
            catch
            {
                ExcelClassSettings.SaveFileNameIncludesQuarter = true;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, TemplateCopyNameIncludesQuarterKey, ExcelClassSettings.SaveFileNameIncludesQuarter.ToString(CultureInfo.CurrentCulture));
            }


            ExcelClassSettings.FirstNameColumnLetter = XML.XMLReadFile(xmlEXCELNodeNameLocation, FirstNameColumnLetterKey);
            if (ExcelClassSettings.FirstNameColumnLetter == "")
            {
                ExcelClassSettings.FirstNameColumnLetter = FirstNameColumnLetterDefault;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, FirstNameColumnLetterKey, ExcelClassSettings.FirstNameColumnLetter);
            }

            ExcelClassSettings.ItemCell = XML.XMLReadFile(xmlEXCELNodeNameLocation, ItemCellKey);
            if (ExcelClassSettings.ItemCell == "")
            {
                ExcelClassSettings.ItemCell = ItemCellDefault;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, ItemCellKey, ExcelClassSettings.ItemCell);
            }

            ExcelClassSettings.LastNameColumnLetter = XML.XMLReadFile(xmlEXCELNodeNameLocation, LastNameColumnLetterKey);
            if (ExcelClassSettings.LastNameColumnLetter == "")
            {
                ExcelClassSettings.LastNameColumnLetter = LastNameColumnLetterDefault;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, LastNameColumnLetterKey, ExcelClassSettings.LastNameColumnLetter);
            }

            #region Option ColumnLetterKey

            ExcelClassSettings.OptHead1ColumnLetter = XML.XMLReadFile(xmlEXCELNodeNameLocation, OptHead1ColumnLetterKey);
            if (ExcelClassSettings.OptHead1ColumnLetter == "")
            {
                ExcelClassSettings.OptHead1ColumnLetter = OptHead1ColumnLetterDefault;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, OptHead1ColumnLetterKey, ExcelClassSettings.OptHead1ColumnLetter);
            }

            ExcelClassSettings.OptHead2ColumnLetter = XML.XMLReadFile(xmlEXCELNodeNameLocation, OptHead2ColumnLetterKey);
            if (ExcelClassSettings.OptHead2ColumnLetter == "")
            {
                ExcelClassSettings.OptHead2ColumnLetter = OptHead2ColumnLetterDefault;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, OptHead2ColumnLetterKey, ExcelClassSettings.OptHead2ColumnLetter);
            }


            ExcelClassSettings.OptHead3ColumnLetter = XML.XMLReadFile(xmlEXCELNodeNameLocation, OptHead3ColumnLetterKey);
            if (ExcelClassSettings.OptHead3ColumnLetter == "")
            {
                ExcelClassSettings.OptHead3ColumnLetter = OptHead3ColumnLetterDefault;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, OptHead3ColumnLetterKey, ExcelClassSettings.OptHead3ColumnLetter);
            }
            #endregion

            #region Option Header Names
            ExcelClassSettings.HeaderNames.Header1 = XML.XMLReadFile(xmlEXCELNodeNameLocation, OptHead1NameDefault);
            if (ExcelClassSettings.HeaderNames.Header1 == "")
            {
                ExcelClassSettings.HeaderNames.Header1 = OptionHeaders.HeaderDefault;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, OptHead1NameDefault, ExcelClassSettings.HeaderNames.Header1);
            }

            ExcelClassSettings.HeaderNames.Header2 = XML.XMLReadFile(xmlEXCELNodeNameLocation, OptHead2NameDefault);
            if (ExcelClassSettings.HeaderNames.Header2 == "")
            {
                ExcelClassSettings.HeaderNames.Header2 = OptionHeaders.HeaderDefault;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, OptHead2NameDefault, ExcelClassSettings.HeaderNames.Header2);
            }

            ExcelClassSettings.HeaderNames.Header3 = XML.XMLReadFile(xmlEXCELNodeNameLocation, OptHead3NameDefault);
            if (ExcelClassSettings.HeaderNames.Header3 == "")
            {
                ExcelClassSettings.HeaderNames.Header3 = OptionHeaders.HeaderDefault;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, OptHead3NameDefault, ExcelClassSettings.HeaderNames.Header3);
            }
            ExcelClassSettings.HeaderNames.WaitListHeader = XML.XMLReadFile(xmlEXCELNodeNameLocation, OptHeadWaitlistNameDefault);
            if (ExcelClassSettings.HeaderNames.WaitListHeader == "")
            {
                ExcelClassSettings.HeaderNames.WaitListHeader = OptionHeaders.HeaderDefault;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, OptHeadWaitlistNameDefault, ExcelClassSettings.HeaderNames.WaitListHeader);
            }
            #endregion

            #region Export Option Values
            try
            {
                ExcelClassSettings.ExportoptHead1 = Convert.ToBoolean(XML.XMLReadFile(xmlEXCELNodeNameLocation, ExportoptHead1Key).ToLower());
            }
            catch
            {
                ExcelClassSettings.ExportoptHead1 = false;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, ExportoptHead1Key, ExcelClassSettings.ExportoptHead1.ToString(CultureInfo.CurrentCulture));
            }

            try
            {
                ExcelClassSettings.ExportoptHead2 = Convert.ToBoolean(XML.XMLReadFile(xmlEXCELNodeNameLocation, ExportoptHead2Key).ToLower());
            }
            catch
            {
                ExcelClassSettings.ExportoptHead2 = false;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, ExportoptHead2Key, ExcelClassSettings.ExportoptHead2.ToString(CultureInfo.CurrentCulture));
            }



            try
            {
                ExcelClassSettings.ExportoptHead3 = Convert.ToBoolean(XML.XMLReadFile(xmlEXCELNodeNameLocation, ExportoptHead3Key).ToLower());
            }
            catch
            {
                ExcelClassSettings.ExportoptHead3 = false;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, ExportoptHead3Key, ExcelClassSettings.ExportoptHead3.ToString(CultureInfo.CurrentCulture));
            }

            #endregion

            ExcelClassSettings.SIDColumnLetter = XML.XMLReadFile(xmlEXCELNodeNameLocation, SIDColumnLetterKey);
            if (ExcelClassSettings.SIDColumnLetter == "")
            {
                ExcelClassSettings.SIDColumnLetter = SIDColumnLetterDefault;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, SIDColumnLetterKey, ExcelClassSettings.SIDColumnLetter);
            }

            ExcelClassSettings.SIDLast4ColumnLetter = XML.XMLReadFile(xmlEXCELNodeNameLocation, SIDLast4ColumnLetterKey);
            if (ExcelClassSettings.SIDLast4ColumnLetter == "")
            {
                ExcelClassSettings.SIDLast4ColumnLetter = SIDLast4ColumnLetterDefault;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, SIDLast4ColumnLetterKey, ExcelClassSettings.SIDLast4ColumnLetter);
            }

            ExcelClassSettings.TemplateDirectory = XML.XMLReadFile(xmlEXCELNodeNameLocation, TemplateDirectoryKey);
            if (ExcelClassSettings.TemplateDirectory == "")
            {
                ExcelClassSettings.TemplateDirectory = UserSettings.MyDocuments;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, TemplateDirectoryKey, ExcelClassSettings.TemplateDirectory);
            }

            ExcelClassSettings.TemplateFileName = XML.XMLReadFile(xmlEXCELNodeNameLocation, TemplateNameKey);
            if (ExcelClassSettings.TemplateFileName == "")
            {
                ExcelClassSettings.TemplateFileName = TemplateFileNameDafault;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, TemplateNameKey, ExcelClassSettings.TemplateFileName);
            }

            ExcelClassSettings.SaveFileName = XML.XMLReadFile(xmlEXCELNodeNameLocation, SaveFileNameKey);
            if (ExcelClassSettings.SaveFileName == "")
            {
                ExcelClassSettings.SaveFileName = SaveAsFileNameDafault;
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, SaveFileNameKey, ExcelClassSettings.SaveFileName);
            }

            ExcelClassSettings.SaveFileDirectory = XML.XMLReadFile(xmlEXCELNodeNameLocation, SaveFileDirectoryKey);
            if (ExcelClassSettings.SaveFileDirectory == "")
            {
                ExcelClassSettings.SaveFileDirectory = ExcelClassSettings.UserDesktop();
                XML.XMLWriteFile(xmlEXCELNodeNameLocation, SaveFileDirectoryKey, ExcelClassSettings.SaveFileDirectory);
            }

            return ExcelClassSettings;
        }

        public void Save(ExcelClassSettings ExcelClassSettings)
        {
            XMLhelper XML = new XMLhelper(UserSettings);

            XML.XMLWriteFile(xmlEXCELNodeNameLocation, ExportKey, ExcelClassSettings.Export.ToString(CultureInfo.CurrentCulture));
            XML.XMLWriteFile(xmlEXCELNodeNameLocation, ExportWaitListKey, ExcelClassSettings.ExportWaitlist.ToString(CultureInfo.CurrentCulture));
            XML.XMLWriteFile(xmlEXCELNodeNameLocation, ExportMiddleInitialKey, ExcelClassSettings.ExportMiddleInitial.ToString(CultureInfo.CurrentCulture));
            XML.XMLWriteFile(xmlEXCELNodeNameLocation, ExportSIDKey, ExcelClassSettings.ExportSID.ToString(CultureInfo.CurrentCulture));
            XML.XMLWriteFile(xmlEXCELNodeNameLocation, ExportSIDLast4Key, ExcelClassSettings.ExportSIDLast4.ToString(CultureInfo.CurrentCulture));
            XML.XMLWriteFile(xmlEXCELNodeNameLocation, FirstStudentKey, ExcelClassSettings.FirstStudent.ToString(CultureInfo.CurrentCulture));
            XML.XMLWriteFile(xmlEXCELNodeNameLocation, TemplateCopyNameIncludesQuarterKey, ExcelClassSettings.SaveFileNameIncludesQuarter.ToString(CultureInfo.CurrentCulture));
            XML.XMLWriteFile(xmlEXCELNodeNameLocation, ExportoptHead1Key, ExcelClassSettings.ExportoptHead1.ToString(CultureInfo.CurrentCulture));
            XML.XMLWriteFile(xmlEXCELNodeNameLocation, FirstNameColumnLetterKey, ExcelClassSettings.FirstNameColumnLetter);
            XML.XMLWriteFile(xmlEXCELNodeNameLocation, ItemCellKey, ExcelClassSettings.ItemCell);
            XML.XMLWriteFile(xmlEXCELNodeNameLocation, LastNameColumnLetterKey, ExcelClassSettings.LastNameColumnLetter);

            XML.XMLWriteFile(xmlEXCELNodeNameLocation, OptHead1ColumnLetterKey, ExcelClassSettings.OptHead1ColumnLetter);
            XML.XMLWriteFile(xmlEXCELNodeNameLocation, OptHead2ColumnLetterKey, ExcelClassSettings.OptHead2ColumnLetter);
            XML.XMLWriteFile(xmlEXCELNodeNameLocation, OptHead3ColumnLetterKey, ExcelClassSettings.OptHead3ColumnLetter);
            XML.XMLWriteFile(xmlEXCELNodeNameLocation, OptHeadWaitlistNameDefault, ExcelClassSettings.HeaderNames.WaitListHeader);

            XML.XMLWriteFile(xmlEXCELNodeNameLocation, OptHead1NameDefault, ExcelClassSettings.HeaderNames.Header1);
            XML.XMLWriteFile(xmlEXCELNodeNameLocation, OptHead2NameDefault, ExcelClassSettings.HeaderNames.Header2);
            XML.XMLWriteFile(xmlEXCELNodeNameLocation, OptHead3NameDefault, ExcelClassSettings.HeaderNames.Header3);

            XML.XMLWriteFile(xmlEXCELNodeNameLocation, SIDColumnLetterKey, ExcelClassSettings.SIDColumnLetter);
            XML.XMLWriteFile(xmlEXCELNodeNameLocation, SIDLast4ColumnLetterKey, ExcelClassSettings.SIDLast4ColumnLetter);
            XML.XMLWriteFile(xmlEXCELNodeNameLocation, TemplateDirectoryKey, ExcelClassSettings.TemplateDirectory);
            XML.XMLWriteFile(xmlEXCELNodeNameLocation, TemplateNameKey, ExcelClassSettings.TemplateFileName);
            XML.XMLWriteFile(xmlEXCELNodeNameLocation, SaveFileNameKey, ExcelClassSettings.SaveFileName);
            
        }
    }
}
