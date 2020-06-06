using System;
using System.Globalization;
using System.Text;
using System.IO;

using InstructorBriefcaseExtractor.Model;
using InstructorBriefcaseExtractor.Utility;

namespace InstructorBriefcaseExtractor.BLL
{
    internal class Clicker
    {
        #region Events

        #region Message

        public event InstructorBriefcaseExtractor.Model.InformationalMessage MessageResults;
        protected virtual void SendMessage(object sender, EventArgs e)
        {
            MessageResults?.Invoke(sender, e);
        }

        #endregion

        #region File Exist Request
        public event FileExist FileExiStrequest;
        protected virtual void FileExistsMessage(object sender, EventArgs e)
        {
            FileExiStrequest?.Invoke(sender, e);
        }

        #endregion

        #endregion

        public static GenericSettingsCollection ClickerFormat()
        {
            GenericSettingsCollection ClickerFormatCollection = new GenericSettingsCollection
            {
                { 1, "#, Student Name, UniqueID" }
            };

            return ClickerFormatCollection;
        }

        public Clicker()
        {

        }

        public ClickerConfiguration Options { get; set; }

        public void Export(ClickerSettings ClickerSettings, UserSettings UserSettings, Courses myCourses)
        {
            //==========================================================
            //
            // IN: UserInfo  - user specific Options
            //     myCourses - contains the courses to export
            //
            // Returns: Nothing - Events are raised to send status messages
            //
            // DESCRIPTION
            //    This procedure exports to an Clicker output file
            //==========================================================            
            bool DeleteFileError;

            // Verify that MTG is to be exported
            if (!ClickerSettings.Export) { return; }
            // Create Directory if it does not exist
            if (!Directory.Exists(ClickerSettings.Directory))
            { Directory.CreateDirectory(ClickerSettings.Directory); }

            try
            {
                foreach (Course C in myCourses)
                {
                    StringBuilder SB = new StringBuilder(2880); // 72 character per student * 40 students

                    if (C.Export)
                    {
                        // Export to clicker output format
                        // Index, Last,First,UniqueID (Made up of last name first initial and ID                        
                        int StudentCount = C.Students.Length;  // this is a 1 dimensional array
                        for (int i = 0; i < StudentCount; i++)
                        {
                            if (ClickerSettings.SelectedValue == 1)
                            {
                                SB.Append((i + 1).ToString(CultureInfo.CurrentCulture)).Append(". ");
                                Student S = C.Students[i];
                                SB.Append(S.LastName + ",");    // Student Last Name
                                SB.Append(S.FirstName + ",");   // Student First Name
                                SB.Append(S.LastName).Append(S.FirstName[0]).Append((i + 1).ToString(CultureInfo.CurrentCulture));         // UniqueID
                            }
                            else if (ClickerSettings.SelectedValue == 2)
                            {
                                // need to define the format
                                SB.Append((i + 1).ToString(CultureInfo.CurrentCulture)).Append(". ");
                                Student S = C.Students[i];
                                SB.Append(S.LastName + ",");    // Student Last Name
                                SB.Append(S.FirstName + ",");   // Student First Name
                                SB.Append(S.LastName).Append(S.FirstName[0]).Append((i + 1).ToString(CultureInfo.CurrentCulture));         // UniqueID
                            }
                            if (i < StudentCount) { SB.Append("\r\n"); }
                        }

                        if (ClickerSettings.ExportWaitlist)
                        {
                            // Export to clicker output format
                            // Index, Last,First,UniqueID (Made up of last name first initial and ID                        
                            int WaitlistCount = C.Waitlist.Length;  // this is a 1 dimensional array
                            for (int i = 0; i < WaitlistCount; i++)
                            {
                                int j = i + StudentCount;
                                Student S = C.Waitlist[i];
                                if (ClickerSettings.SelectedValue == 1)
                                {
                                    SB.Append((j + 1).ToString(CultureInfo.CurrentCulture)).Append(". ");
                                    SB.Append(S.LastName + ",");    // Student Last Name
                                    SB.Append(S.FirstName + ",");   // Student First Name
                                    SB.Append(S.LastName).Append(S.FirstName[0]).Append((j + 1).ToString(CultureInfo.CurrentCulture));         // UniqueID
                                }
                                else if (ClickerSettings.SelectedValue == 2)
                                {
                                    // need to define the format
                                    SB.Append((j + 1).ToString(CultureInfo.CurrentCulture)).Append(". ");
                                    SB.Append(S.LastName + ",");    // Student Last Name
                                    SB.Append(S.FirstName + ",");   // Student First Name
                                    SB.Append(S.LastName).Append(S.FirstName[0]).Append((j + 1).ToString(CultureInfo.CurrentCulture));         // UniqueID
                                }
                                if (i < WaitlistCount) { SB.Append("\r\n"); }
                            }

                        }

                        string DiskNames = C.DiskName(ClickerSettings.Underscore) + "_Roster.txt";
                        FileInformation FE = new FileInformation("The file " + DiskNames + " already exists.", ClickerSettings.Directory, DiskNames);
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
                                    SendMessage(this, new Information("Canceling Clicker export."));
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