using System;

namespace InstructorBriefcaseExtractor.Model
{
    public class ExcelRollSettings
    {
        public bool Export { get; set; }
        public bool ExportWaitlist { get; set; }
        public bool ExportLab { get; set; }
        public string Header { get; set; }
        public int FirstClass { get; set; }
        public int ClassIncrement { get; set; }
        public string ClassNameColumnLetter { get; set; }
        public string LastNameColumnLetter { get; set; }
        public string FirstNameColumnLetter { get; set; }
        public string MondayDateCell { get; set; }
        public string FirstMondayDayCell { get; set; }
        private DateTime _FirstDay;
        public DateTime FirstDay {
            get { return _FirstDay; }
            set {
                _FirstDay = value;
                if (_FirstDay.DayOfWeek == DayOfWeek.Monday) { FirstMondayDay = _FirstDay; }
                if (_FirstDay.DayOfWeek == DayOfWeek.Tuesday) { FirstMondayDay = _FirstDay.AddDays(-1); }
                if (_FirstDay.DayOfWeek == DayOfWeek.Wednesday) { FirstMondayDay = _FirstDay.AddDays(-2); }
                if (_FirstDay.DayOfWeek == DayOfWeek.Thursday) { FirstMondayDay = _FirstDay.AddDays(-3); }
                if (_FirstDay.DayOfWeek == DayOfWeek.Friday) { FirstMondayDay = _FirstDay.AddDays(-4); }
            }

        }
        public DateTime FirstMondayDay { get; private set; }

        public ExcelRollSettings Clone()
        {
            ExcelRollSettings Clone = new ExcelRollSettings();
            Clone.Export = Export;
            Clone.ExportLab = ExportLab;
            Clone.Header = Header;
            Clone.FirstClass = FirstClass;
            Clone.ClassIncrement = ClassIncrement;
            Clone.ClassNameColumnLetter = ClassNameColumnLetter;
            Clone.LastNameColumnLetter = LastNameColumnLetter;
            Clone.FirstNameColumnLetter = FirstNameColumnLetter;
            Clone.MondayDateCell = MondayDateCell;
            Clone.FirstMondayDayCell = FirstMondayDayCell;
            Clone.FirstDay = FirstDay;
            Clone.FirstMondayDay = FirstMondayDay;

            return Clone;
        }
    }
}
