using System;
using System.Globalization;
using System.Text;
using System.IO;

using InstructorBriefcaseExtractor.Model;
using InstructorBriefcaseExtractor.Utility;

namespace InstructorBriefcaseExtractor.BLL
{
    
    internal class MTG
    {
        #region Events

        public event FileExist FileExiStrequest;
        protected virtual void FileExistsMessage(object sender, EventArgs e)
        {
            FileExiStrequest?.Invoke(sender, e);
        }

        public event InstructorBriefcaseExtractor.Model.InformationalMessage MessageResults;
        protected virtual void SendMessage(object sender, EventArgs e)
        {
            MessageResults?.Invoke(sender, e);
        }

        #endregion

        public MTG()
        {}

        public void Export(MTGSettings MTGSettings, UserSettings UserSettings, Courses myCourses)
        {
            //==========================================================
            //
            // IN: UserInfo  - user specific Options
            //     myCourses - contains the courses to export
            //
            // Returns: Nothing - Events are raised to send status messages
            //
            // DESCRIPTION
            //    This procedure exports to an Outlook Distribution list
            //==========================================================            
            bool DeleteFileError;

            // Verify that MTG is to be exported
            if (!MTGSettings.Export) { return; }
            // Create Directory if it does not exist
            if (!Directory.Exists(MTGSettings.Directory))
               { Directory.CreateDirectory(MTGSettings.Directory); }

            try
            {
                foreach (Course C in myCourses)
                {
                    StringBuilder SB = new StringBuilder(2880); // 72 character per student * 40 students

                    if (C.Export)
                    {

                        // Export to Making the grade output format
                        // Last,First,SID,Sex,GroupCode,Cit System,Counselor Name,Parent Name,
                        // Street Address,City,St,Zip Code,Book Number,Lock,,Home Phone,
                        // Custom Field 01,Custom Field 02,Custom Field 03,Custom Field 04,Custom Field 05,
                        // Custom Field 06,Custom Field 07,Custom Field 08,,,,,,,,,,
                        int StudentCount = C.Students.Length;  // this is a 1 dimensional array
                        for (int i = 0; i < StudentCount; i++)
                        {
                            Student S = C.Students[i];

                            SB.Append(S.LastName + ", ");    // Student Last Name
                            SB.Append(S.FirstName + ",");   // Student First Name
                            SB.Append(S.SID + ",");         // SID
                            SB.Append(",");                 // Sex
                            SB.Append(",");                 // GroupCode
                            SB.Append(",");                 // Grd Sys
                            SB.Append(",");                 // Cit Mark
                            SB.Append(",");                 // Counselor Name
                            SB.Append(",");                 // Parent Name
                            SB.Append(",");                 // Street Address
                            SB.Append(",");                 // City
                            SB.Append(",");                 // State
                            SB.Append(",");                 // Zip Code
                            SB.Append(",");                 // Book Number
                            SB.Append(",");                 // Locker Bin
                            SB.Append(",");                 // Birth Date
                            SB.Append(S.Phone + ",");        // Home or Work Phone
                            SB.Append(",");                 // Line Library Comments
                            SB.Append(",");                 // Custom Field 01
                            SB.Append(",");                 // Custom Field 02
                            SB.Append(",");                 // Custom Field 03
                            SB.Append(",");                 // Custom Field 04
                            SB.Append(",");                 // Custom Field 05
                            SB.Append(",");                 // Custom Field 06
                            SB.Append(",");                 // Custom Field 07
                            SB.Append(",");                 // Custom Field 08
                            SB.Append(",");                 // Custom Field 09
                            SB.Append(",");                 // Custom Field 10
                            SB.Append(",");                 // MAT
                            SB.Append(",");                 // SPL
                            SB.Append(",");                 // LIT
                            SB.Append(",");                 // WRT
                            SB.Append(",");                 // ORL
                            SB.Append(",");                 // GRM
                            SB.Append(",");                 // COM
                            SB.Append(",");                 // PAR
                            SB.Append(",");                 // QFT
                            SB.Append(",");                 // CAT0
                            if (i < StudentCount) { SB.Append("\r\n"); }
                        }

                        if (MTGSettings.ExportWaitlist)
                        {

                            // Export to Making the grade output format
                            // Last,First,SID,Sex,GroupCode,Cit System,Counselor Name,Parent Name,
                            // Street Address,City,St,Zip Code,Book Number,Lock,,Home Phone,
                            // Custom Field 01,Custom Field 02,Custom Field 03,Custom Field 04,Custom Field 05,
                            // Custom Field 06,Custom Field 07,Custom Field 08,,,,,,,,,,
                            int WaitlistCount = C.Waitlist.Length;  // this is a 1 dimensional array
                            for (int i = 0; i < WaitlistCount; i++)
                            {
                                Student S = C.Waitlist[i];

                                SB.Append(S.LastName + ", ");    // Student Last Name
                                SB.Append(S.FirstName + ",");   // Student First Name
                                SB.Append(S.SID + ",");         // SID
                                SB.Append(",");                 // Sex
                                SB.Append(",");                 // GroupCode
                                SB.Append(",");                 // Grd Sys
                                SB.Append(",");                 // Cit Mark
                                SB.Append(",");                 // Counselor Name
                                SB.Append(",");                 // Parent Name
                                SB.Append(",");                 // Street Address
                                SB.Append(",");                 // City
                                SB.Append(",");                 // State
                                SB.Append(",");                 // Zip Code
                                SB.Append(",");                 // Book Number
                                SB.Append(",");                 // Locker Bin
                                SB.Append(",");                 // Birth Date                                
                                SB.Append(",");                 // Home or Work Phone
                                SB.Append(",");                 // Line Library Comments
                                SB.Append(",");                 // Custom Field 01
                                SB.Append(",");                 // Custom Field 02
                                SB.Append(",");                 // Custom Field 03
                                SB.Append(",");                 // Custom Field 04
                                SB.Append(",");                 // Custom Field 05
                                SB.Append(",");                 // Custom Field 06
                                SB.Append(",");                 // Custom Field 07
                                SB.Append(",");                 // Custom Field 08
                                SB.Append(",");                 // Custom Field 09
                                SB.Append(",");                 // Custom Field 10
                                SB.Append(",");                 // MAT
                                SB.Append(",");                 // SPL
                                SB.Append(",");                 // LIT
                                SB.Append(",");                 // WRT
                                SB.Append(",");                 // ORL
                                SB.Append(",");                 // GRM
                                SB.Append(",");                 // COM
                                SB.Append(",");                 // PAR
                                SB.Append(",");                 // QFT
                                SB.Append(",");                 // CAT0
                                if (i < WaitlistCount) { SB.Append("\r\n"); }
                            }
                        }

                        string DiskNames = C.DiskName(MTGSettings.Underscore) + ".txt";
                        FileInformation FE = new FileInformation("The file " + DiskNames + " already exists.", MTGSettings.Directory, DiskNames);
                        DeleteFileError = false;

                        // check to see if file exists
                        SendMessage(this, new Information("Preparing " + DiskNames + "."));
                        if (File.Exists(FE.OldFileNameandPath))
                        {
                            SendMessage(this, new Information("File already exists."));
                            if (UserSettings.OverWriteAll)
                            {
                                try
                                {
                                    SendMessage(this, new Information("Attempting to replace."));
                                    File.Delete(FE.OldFileNameandPath);
                                }
                                catch
                                {
                                    DeleteFileError = true;
                                }
                            }
                            else
                            {
                                FileExistsMessage(this, FE);
                                UserSettings.OverWriteAll = FE.OverWriteAll;

                                if (FE.CancelALLExport)
                                {
                                    SendMessage(this, new Information("Canceling MTG export."));
                                    return;
                                }
                                if ((!FE.CancelExport) && (FE.OldFileNameandPath == FE.NewFileNameandPath))
                                {
                                    try
                                    {
                                        File.Delete(FE.OldFileNameandPath);
                                    }
                                    catch
                                    {
                                        DeleteFileError = true;
                                    }
                                }
                            }
                        }

                        if (FE.CancelExport)
                        {
                            SendMessage(this, new Information("User Canceled export for file."));
                        }
                        else if (DeleteFileError)
                        {
                            SendMessage(this, new Information("Unable to delete file!"));
                        }
                        else
                        {
                            try
                            {
                                SendMessage(this, new Information("Writing file to disk."));
                                File.WriteAllText(FE.NewFileNameandPath, SB.ToString());
                            }
                            catch (Exception Ex)
                            {
                                SendMessage(this, new Information("While trying to write " + FE.OldFileNameandPath + "An error occurred.\r\n" + Ex.Message));
                            }
                        }
                    }
                }
            }
            catch (Exception Ex)
            {
                SendMessage(this, new Information("An unexpected error occurred - " + Ex.Message));
                throw;
            }
        }
    }
}
