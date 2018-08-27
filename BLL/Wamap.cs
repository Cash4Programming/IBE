using System;
using System.Text;
using System.IO;

using InstructorBriefcaseExtractor.Model;
using InstructorBriefcaseExtractor.Utility;

namespace InstructorBriefcaseExtractor.BLL
{
    internal class Wamap
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

        private string StripIllegalCharacters(string ToModify)
        {
            return ToModify.Replace("-", "");
        }

        public void Export(WamapSettings WamapSettings, UserSettings UserSettings, Courses myCourses)
        {
            //==========================================================
            //
            // IN: UserInfo  - user specific Options
            //     myCourses - contains the courses to export
            //
            // Returns: Nothing - Events are raised to send status messages
            //
            // DESCRIPTION
            //    This procedure exports to an WebAssign upload files
            //==========================================================
            bool DeleteFileError = false;
            
            // Verify that WebAssign is to be exported
            if (!WamapSettings.Export) { return; }

            // Create Directory if it does not exist
            if (!Directory.Exists(WamapSettings.Directory))
            { Directory.CreateDirectory(WamapSettings.Directory); }

            // waUserNames contain a list of all WAMAP Usernames as they must
            // be unique
            // No header row
            // password,desired username,Lastname Firstname,4,5,6,email address
            try
            {
                foreach (Course C in myCourses)
                {
                    if (C.Export)
                    {
                        // sbCourse - upload to WAMAP
                        // sbStudent - print and cut out for student
                        StringBuilder sbCourse = new StringBuilder(2880); // 72 character per student * 40 students
                        StringBuilder sbStudent = new StringBuilder(2880); // 72 character per student * 40 students

                        SendMessage(this, new Information("Exporting " + C.Name));

                        string StrSEPERATOR_TYPE= ",";
                        string StrUserName = "";

                        int StudentCount = C.Students.Length;  // this is a 1 dimensional array
                        for (int i = 0; i < StudentCount; i++)
                        {
                            Student S = C.Students[i];

                            StrUserName = StripIllegalCharacters(S.SID);
                            
                            // Eliminate illegal character (-) from username and password
                            string StrWamapUserName = StripIllegalCharacters(StrUserName).ToLower();
                            string StrWamapPassword = StripIllegalCharacters(S.SID).ToLower();

                            // password,desired username,Lastname Firstname,section,5,6,email address
                            sbCourse.Append(StrWamapPassword + StrSEPERATOR_TYPE);                // 1-password
                            sbCourse.Append(StrWamapUserName.ToLower() + StrSEPERATOR_TYPE);      // 2-desired username
                            sbCourse.Append(S.LastName + " " + S.FirstName + StrSEPERATOR_TYPE);  // 3-Lastname Firstname
                            sbCourse.Append(C.ItemNumber + StrSEPERATOR_TYPE);                    // 4 - section code - class item number
                            sbCourse.Append(StrSEPERATOR_TYPE);                                   // 5
                            sbCourse.Append(StrSEPERATOR_TYPE);                                   // 6
                            sbCourse.Append(S.Email.ToLower() + "\r\n");                          // 7-email address
                            
                            // create the student text file that contains the student name
                            // their username and the assigned password
                            // Example
                            //------------------------
                            //| BRISCOE, KERI        |
                            //------------------------
                            //| Username |  Password |
                            //
                            //| kbriscoe |  briscoe  |
                            //------------------------

                            int intUserNameHeader = 0;  // space if bigger than the UserNameUser
                            int intPasswordHeader = 0;  // space if bigger than the User's Password 
                            int intUserNameUser = 0;    // space if bigger than the UserNameUser Header
                            int intPasswordUser = 0;    // space if bigger than the User's Password
                            int intSpacer = 0;          // number of - in the seperator lines
                            int intFillnameSpacer = 0;  // spaces for full name

                            string StrUserNameHeader = "Username";
                            string StrPasswordHeader = "Password";
                            string StrFullName = S.LastName + ", " + S.FirstName;

                            // Adjust the header/user spaces
                            if (StrUserNameHeader.Length > StrUserName.Length)
                            {
                                intUserNameHeader = 0;
                                intUserNameUser = StrUserNameHeader.Length - StrWamapUserName.Length;
                            }
                            else
                            {
                                intUserNameHeader = StrWamapUserName.Length - StrUserNameHeader.Length;
                                intUserNameUser = 0;
                            }

                            if (StrPasswordHeader.Length > StrUserName.Length)
                            {
                                intPasswordHeader = 0;
                                intPasswordUser = StrPasswordHeader.Length - StrWamapPassword.Length;
                            }
                            else
                            {
                                intPasswordHeader = StrWamapPassword.Length - StrPasswordHeader.Length;
                                intPasswordUser = 0;
                            }

                            intSpacer = 10 + StrUserNameHeader.Length + intUserNameHeader +
                                             StrPasswordHeader.Length + intPasswordHeader;
                            string StrSpacer = new String('-',intSpacer);

                            if ((intSpacer - 3) > StrFullName.Length)
                            {
                                intFillnameSpacer = (intSpacer - StrFullName.Length) - 4;
                            }
                            else
                            {
                                intFillnameSpacer = 0;
                            }

                            sbStudent.Append(StrSpacer + "\r\n");
                            sbStudent.Append("| " + StrFullName + new String(' ', intFillnameSpacer) + " |" + "\r\n");
                            sbStudent.Append(StrSpacer + "\r\n");
                            sbStudent.Append("| " + StrUserNameHeader + new String(' ', intUserNameHeader)                                                   
                                                  + " | " + StrPasswordHeader + new String(' ', intPasswordHeader) 
                                                  + " |" + "\r\n\r\n");

                            sbStudent.Append("| " + StrWamapUserName.ToLower() + new String(' ', intUserNameUser)                                            
                                                   + " | " + StrWamapPassword.ToLower() + new String(' ', intPasswordUser) 
                                                   + " |" + "\r\n");
                            sbStudent.Append(StrSpacer + "\r\n\r\n\r\n");
                        }

                        String[] DiskNames = { C.DiskName(WamapSettings.Underscore) + ".txt", C.DiskName(WamapSettings.Underscore) + "_Students.txt" };
                        String[] FileContents = { sbCourse.ToString(), sbStudent.ToString() };


                        for (int i = 0; i < 1; i++)  // 1 = only export the rooster list - modified on 2014-04-04
                        {
                            FileInformation FE = new FileInformation("The file " + DiskNames[i] + " already exists.", WamapSettings.Directory, DiskNames[i]);
                            DeleteFileError = false;

                            // check to see if file exists
                            SendMessage(this, new Information("Preparing " + DiskNames[i] + "."));
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
                                        SendMessage(this, new Information("Canceling WebAssign export."));
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
                                    File.WriteAllText(FE.NewFileNameandPath, FileContents[i]);
                                }
                                catch (Exception Ex)
                                {
                                    SendMessage(this, new Information("While trying to write " + FE.OldFileNameandPath + "An error occurred.\r\n" + Ex.Message));
                                }
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
