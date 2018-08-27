using System;


namespace InstructorBriefcaseExtractor.Model
{
    // All Model delegates and EventArgs
    public delegate void InformationalMessage(object sender, EventArgs e);
    public class Information : EventArgs
    {
        public Information(String Message)
        {
            this.Message = Message;
        }

        public String Message { get; private set; }
    }

    public delegate void FileExist(object sender, EventArgs e);
    public class FileInformation : EventArgs
    {
        public FileInformation(String DialogBoxTitle, String OldFileandDirectory)
        {
            StrDialogBoxTitle = DialogBoxTitle;
            int LastSlash = OldFileandDirectory.LastIndexOf('\\');

            StrOldDirectory = OldFileandDirectory.Substring(0, LastSlash + 1);
            StrNewDirectory = StrOldDirectory;

            StrOldFileName = OldFileandDirectory.Substring(LastSlash + 1, OldFileandDirectory.Length - LastSlash - 1);
            StrNewFileName = StrOldFileName;

            blnCancelExport = false;
            blnCancelALLExport = false;
            blnOverWriteAll = false;
        }

        public FileInformation(String DialogBoxTitle, String OldDirectory, String OldFileName)
        {
            StrDialogBoxTitle = DialogBoxTitle;

            StrOldDirectory = OldDirectory;
            if (!StrOldDirectory.EndsWith("\\")) { StrOldDirectory += "\\"; };
            StrNewDirectory = StrOldDirectory;

            StrOldFileName = OldFileName;
            StrNewFileName = StrOldFileName;

            blnCancelExport = false;
            blnCancelALLExport = false;
            blnOverWriteAll = false;
        }

        private String StrDialogBoxTitle;
        public String DialogBoxTitle
        {
            get { return StrDialogBoxTitle; }
            set { StrDialogBoxTitle = value; }
        }

        private String StrOldDirectory;
        public String OldDirectory
        {
            get { return StrOldDirectory; }
        }

        private String StrOldFileName;
        public String OldFileName
        {
            get { return StrOldFileName; }
            set { StrOldFileName = value; }
        }

        public String OldFileNameandPath
        {
            get { return StrOldDirectory + StrOldFileName; }
        }

        private String StrNewDirectory;
        public String NewDirectory
        {
            get { return StrNewDirectory; }
            set
            {
                StrNewDirectory = value;
                if (!StrNewDirectory.EndsWith("\\")) { StrNewDirectory += "\\"; };
            }
        }

        private String StrNewFileName;
        public String NewFileName
        {
            get { return StrNewFileName; }
            set { StrNewFileName = value; }
        }

        public String NewFileNameandPath
        {
            get { return StrNewDirectory + StrNewFileName; }
        }

        private bool blnCancelExport;
        public bool CancelExport
        {
            get { return blnCancelExport; }
            set { blnCancelExport = value; }
        }

        private bool blnCancelALLExport;
        public bool CancelALLExport
        {
            get { return blnCancelALLExport; }
            set { blnCancelALLExport = value; }
        }

        private bool blnOverWriteAll;
        public bool OverWriteAll
        {
            get { return blnOverWriteAll; }
            set { blnOverWriteAll = value; }
        }
    }
}
