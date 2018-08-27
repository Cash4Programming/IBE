using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

using InstructorBriefcaseExtractor.Model;
using InstructorBriefcaseExtractor.Utility;

namespace InstructorBriefcaseExtractor.BLL
{
    internal class WebAssign
    {
        #region Events

        public event FileExist FileExiStrequest;
        protected virtual void FileExistsMessage(object sender, EventArgs e)
        {
            FileExiStrequest?.Invoke(sender, e);
        }

        public event InstructorBriefcaseExtractor.Model.InformationalMessage MessageResults;
        protected virtual void SendMessage(object sender, Information Information)
        {
            MessageResults?.Invoke(sender, Information);
        }

        #endregion

        private string StripIllegalCharacters(string ToModify)
        {
            return ToModify.Replace("-", "");
        }

        public void Export(WebAssignSettings WebAssignSettings, UserSettings UserSettings, Courses myCourses)
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
            if (!WebAssignSettings.Export) { return; }

            // Create Directory if it does not exist
            if (!Directory.Exists(WebAssignSettings.Directory))
            { Directory.CreateDirectory(WebAssignSettings.Directory); }

            // waUserNames contain a list of all WebAssign Usernames as they must
            // be unique
            Dictionary<String, int> waUserNames = new Dictionary<String, int>();
            try
            {
                foreach (Course C in myCourses)
                {
                    waUserNames.Clear();
                    if (C.Export)
                    {
                        StringBuilder sbCourse = new StringBuilder(2880); // 72 character per student * 40 students
                        StringBuilder sbStudent = new StringBuilder(2880); // 72 character per student * 40 students

                        SendMessage(this, new Information("Exporting " + C.Name));

                        string StrSEPERATOR_TYPE;
                        string StrUserName = "";
                        string StrFullName = "";
                        string StrUserPassword = "";                        

                        if (WebAssignSettings.SEPARATOR == WebAssignSeperator.Tab)
                        { StrSEPERATOR_TYPE = "\t"; }  // see escape characters in help 
                        else
                        { StrSEPERATOR_TYPE = ","; }

                        int StudentCount = C.Students.Length;  // this is a 1 dimensional array
                        for (int i = 0; i < StudentCount; i++)
                        {
                            Student S = C.Students[i];

                            switch (WebAssignSettings.Username)
                            {
                                case WebAssignUserName.First_Name_Initial_plus_Last_Name:
                                    StrUserName = S.FirstName.Substring(0, 1) + S.LastName;
                                    break;
                                case WebAssignUserName.First_Name_plus_Last_Name:
                                    StrUserName = S.FirstName + S.LastName;
                                    break;
                                case WebAssignUserName.First_Name_Initial_plus_Middle_Initial_plus_Last_Name:
                                    StrUserName = S.FirstName.Substring(0, 1);
                                    if ((S.MiddleName != null) && (S.MiddleName != ""))
                                    {
                                        StrUserName += S.MiddleName.Substring(0, 1);
                                    }
                                    StrUserName += S.LastName;
                                    break;
                                default:
                                    break;
                            }
                            StrUserName = StripIllegalCharacters(StrUserName);

                            StrFullName = S.LastName + ", " + S.FirstName;
                            if (WebAssignSettings.UserPassword == WebAssignPassword.SID)
                            { StrUserPassword = S.SID; }
                            else
                            { StrUserPassword = S.LastName; }

                            if (waUserNames.ContainsKey(StrUserName))
                            {
                                string msg = "Username for " + StrFullName + " already exists.  Replacing '" + StrUserName + "' with ";
                                int NewStudentIndex = 1;
                                do
                                {
                                    // Username has already been used - try and create a new one                           
                                    switch (WebAssignSettings.Username)
                                    {
                                        case WebAssignUserName.First_Name_Initial_plus_Last_Name:
                                            StrUserName = S.FirstName.Substring(0, 1) + NewStudentIndex + "." + S.LastName;
                                            break;
                                        case WebAssignUserName.First_Name_plus_Last_Name:
                                            StrUserName = S.FirstName + NewStudentIndex + "." + S.LastName;
                                            break;
                                        case WebAssignUserName.First_Name_Initial_plus_Middle_Initial_plus_Last_Name:
                                            StrUserName = S.FirstName.Substring(0, 1);
                                            if ((S.MiddleName != null) && (S.MiddleName != ""))
                                            {
                                                StrUserName += S.MiddleName.Substring(0, 1);
                                            }
                                            StrUserName += NewStudentIndex + "." + S.LastName;
                                            break;
                                        default:
                                            break;
                                    }
                                    NewStudentIndex++;
                                } while (waUserNames.ContainsKey(StrUserName));
                                
                                StrUserName = StripIllegalCharacters(StrUserName);
                                if (waUserNames.ContainsKey(StrUserName))
                                {
                                    // use full name
                                    StrUserName = S.FirstName;
                                    if ((S.MiddleName != null) && (S.MiddleName != ""))
                                    {
                                        StrUserName += "." + S.MiddleName;
                                    }
                                    StrUserName += "." + S.LastName;
                                    StrUserName = StripIllegalCharacters(StrUserName);
                                }

                                if (waUserNames.ContainsKey(StrUserName))
                                {
                                    // ERROR
                                    msg += "Unable to calculate a unique user name for " + StrUserName + "'.  Using full name.";
                                }
                                else
                                {
                                    msg += "'" + StrUserName + "'.";
                                    waUserNames.Add(StrUserName, 1);
                                }
                                SendMessage(this, new Information(msg));
                            }
                            else
                            {
                                waUserNames.Add(StrUserName, 1);
                            }

                            // Eliminate illegal character (-) from username and password
                            string StrWebassignUserName = StripIllegalCharacters(StrUserName).ToLower();
                            string StrWebassignPassword = StripIllegalCharacters(StrUserPassword).ToLower();

                            sbCourse.Append( StrWebassignUserName + StrSEPERATOR_TYPE);    // username
                            sbCourse.Append( StrFullName.ToLower() + StrSEPERATOR_TYPE);     // fullname
                            //sbCourse.Append( StrWebassignPassword + StrSEPERATOR_TYPE);  // password
                            sbCourse.Append( S.Email.ToLower() + StrSEPERATOR_TYPE);          // email
                            sbCourse.Append(StripIllegalCharacters(S.SID).ToLower() + "\r\n");   // SID

                            int intInstitutionHeader = 0;
                            int intUserNameHeader = 0;
                            int intPasswordHeader = 0;
                            int intInstitutionUser = 0;
                            int intUserNameUser = 0;
                            int intPasswordUser = 0;
                            int intSpacer = 0;
                                                        
                            string StrInstitutionHeader = "Institution";
                            string StrUserNameHeader = "Username";
                            string StrPasswordHeader = "Password";
                            string StrInstitution = WebAssignSettings.Institution;

                            int intFillnameSpacer = 0;

                            // Adjust the header/user spaces
                            if (StrUserNameHeader.Length > StrUserName.Length)
                            {
                                intUserNameHeader = 0;
                                intUserNameUser = StrUserNameHeader.Length - StrWebassignUserName.Length;
                            }
                            else
                            {
                                intUserNameHeader = StrWebassignUserName.Length - StrUserNameHeader.Length;
                                intUserNameUser = 0;
                            }

                            if (StrInstitutionHeader.Length > StrInstitution.Length)
                            {
                                intInstitutionHeader = 0;
                                intInstitutionUser = StrInstitutionHeader.Length - StrInstitution.Length;
                            }
                            else
                            {
                                intInstitutionHeader = StrInstitution.Length - StrInstitutionHeader.Length;
                                intInstitutionUser = 0;
                            }

                            if (StrPasswordHeader.Length > StrUserPassword.Length)
                            {
                                intPasswordHeader = 0;
                                intPasswordUser = StrPasswordHeader.Length - StrWebassignPassword.Length;
                            }
                            else
                            {
                                intPasswordHeader = StrWebassignPassword.Length - StrPasswordHeader.Length;
                                intPasswordUser = 0;
                            }

                            intSpacer = 10 + StrInstitutionHeader.Length + intInstitutionHeader +
                                             StrUserNameHeader.Length + intUserNameHeader +
                                             StrPasswordHeader.Length + intPasswordHeader;
                            string StrSpacer = new String('-',intSpacer);
                            //                For intLoop = 0 To intSpacer
                            //                    StrSpacer = StrSpacer & "-"
                            //                Next
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
                                                   + " | " + StrInstitutionHeader + new String(' ', intInstitutionHeader)
                                                   + " | " + StrPasswordHeader + new String(' ', intPasswordHeader) 
                                                   + " |" + "\r\n\r\n");

                            sbStudent.Append("| " + StrWebassignUserName.ToLower() + new String(' ', intUserNameUser)
                                                   + " | " + StrInstitution + new String(' ', intInstitutionUser)
                                                   + " | " + StrWebassignPassword.ToLower() + new String(' ', intPasswordUser) 
                                                   + " |" + "\r\n");
                            sbStudent.Append(StrSpacer + "\r\n\r\n\r\n");
                        }

                        String[] DiskNames = { C.DiskName(WebAssignSettings.Underscore) + ".txt", C.DiskName(WebAssignSettings.Underscore) + "_Students.txt" };
                        String[] FileContents = { sbCourse.ToString(), sbStudent.ToString() };

                        
                        for (int i = 0; i < 1; i++)  // 1 = only export the rooster list - modified on 2014-04-04
                        {
                            FileInformation FE = new FileInformation("The file " + DiskNames[i] + " already exists.", WebAssignSettings.Directory, DiskNames[i]);
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
