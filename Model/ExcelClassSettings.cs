using System;
namespace InstructorBriefcaseExtractor.Model
{
    
    public class ExcelClassSettings
    {
        private readonly MyDirectory MyTemplateDirectory;
        private readonly MyDirectory MySaveAsDirectory;
        public OptionHeaders HeaderNames;

        public static string UserDesktop()
        {
            MyDirectory TempDirectory = new MyDirectory
            {
                Directory = System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };

            return TempDirectory.Directory;
        }

        public ExcelClassSettings()
        {
            MyTemplateDirectory = new MyDirectory();
            MySaveAsDirectory = new MyDirectory();
            HeaderNames = new OptionHeaders();
        }

        public bool Export { get; set; }
        public bool ExportWaitlist { get; set; }

        public string TemplateFileName { get; set; }
        public string TemplateDirectory
        {
            get
            { return MyTemplateDirectory.Directory; }
            set
            { MyTemplateDirectory.Directory = value; }
        }
        public string TemplatePathandFileName
        {
            get
            {
                return MyTemplateDirectory.Directory + TemplateFileName;
            }
        }
        public string SaveFileName { get; set; }
        public string SaveFileDirectory
        {
            get
            { return MySaveAsDirectory.Directory; }
            set
            { MySaveAsDirectory.Directory = value; }
        }
        public string SaveAsPathandFileName
        {
            get
            {
                return MySaveAsDirectory.Directory + SaveFileName;
            }
        }
        public bool SaveFileNameIncludesQuarter { get; set; }

        public bool ExportSID { get; set; }
        public string SIDColumnLetter { get; set; }
        public string SIDLast4ColumnLetter { get; set; }
        public bool ExportSIDLast4 { get; set; }

        public bool ExportMiddleInitial { get; set; }
        public string LastNameColumnLetter { get; set; }
        public string FirstNameColumnLetter { get; set; }
        
        public string ItemCell { get; set; }
        public int FirstStudent { get; set; }

        public string OptHead1ColumnLetter { get; set; }
        public bool ExportoptHead1 { get; set; }

        public string OptHead2ColumnLetter { get; set; }
        public bool ExportoptHead2 { get; set; }

        public string OptHead3ColumnLetter { get; set; }
        public bool ExportoptHead3 { get; set; }

        public ExcelClassSettings Clone()
        {
            ExcelClassSettings Clone = new ExcelClassSettings
            {
                Export = Export,
                TemplateFileName = TemplateFileName,
                TemplateDirectory = TemplateDirectory,
                SaveFileName = SaveFileName,
                SaveFileDirectory = SaveFileDirectory,
                SaveFileNameIncludesQuarter = SaveFileNameIncludesQuarter,
                ExportSID = ExportSID,
                SIDColumnLetter = SIDColumnLetter,
                SIDLast4ColumnLetter = SIDLast4ColumnLetter,
                ExportSIDLast4 = ExportSIDLast4,
                ExportMiddleInitial = ExportMiddleInitial,
                LastNameColumnLetter = LastNameColumnLetter,
                FirstNameColumnLetter = FirstNameColumnLetter,
                ItemCell = ItemCell,
                FirstStudent = FirstStudent,
                OptHead1ColumnLetter = OptHead1ColumnLetter,
                ExportoptHead1 = ExportoptHead1,
                OptHead2ColumnLetter = OptHead2ColumnLetter,
                ExportoptHead2 = ExportoptHead2,
                OptHead3ColumnLetter = OptHead3ColumnLetter,
                ExportoptHead3 = ExportoptHead3
            };
            return Clone;
        }
    }
}
