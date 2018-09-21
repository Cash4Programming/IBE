using System;
using System.IO;
using System.Reflection; // Required for Missing.Value 
using MSOffice = Microsoft.Office.Interop;

using InstructorBriefcaseExtractor.Model;


namespace InstructorBriefcaseExtractor.BLL
{
    internal class Excel
    {
        #region Events

        public event InstructorBriefcaseExtractor.Model.InformationalMessage MessageResults;
        protected virtual void SendMessage(object sender, EventArgs e) => MessageResults?.Invoke(sender, e);

        //protected virtual void SendMessage(object sender, EventArgs e)
        //{
        //    InstructorBriefcaseExtractor.Model.InformationalMessage handler = MessageResults;
        //    if (handler != null)
        //    {
        //        handler(sender, e);
        //    }
        //}

        public event FileExist FileExiStrequest;
        protected virtual void FileExistsMessage(object sender, EventArgs e)
        {
            FileExiStrequest?.Invoke(sender, e);
        }

        //protected virtual void FileExistsMessage(object sender, EventArgs e)
        //{
        //    FileExist handler = FileExiStrequest;
        //    if (handler != null)
        //    {
        //        handler(sender, e);
        //    }
        //}

        #endregion

        #region Properties and Variables
        private static readonly string CLASS_TEMPLATE = "class";    // lowercase
        private static readonly string ROLL_TEMPLATE = "roll";      // lowercase
        private static readonly string LAB_TEMPLATE = "lab";        // lowercase
        private static readonly string CONFIG_TEMPLATE = "config";  // lowercase

        // Get hooks to MS Outlook
        MSOffice.Excel.Application ExcelApp = null;            // To open workbook
        MSOffice.Excel.Workbook ExcelWorkbook = null;          // Workbook
        MSOffice.Excel._Worksheet ExcelClassWorksheet = null;  // Shortcut
        MSOffice.Excel._Worksheet ExcelRollWorksheet = null;   // Shortcut
        MSOffice.Excel._Worksheet ExcelLabWorksheet = null;   // Shortcut
        MSOffice.Excel._Worksheet ExcelConfigWorksheet = null;   // Shortcut
        MSOffice.Excel._Worksheet mySheet = null;              // loop variable
        MSOffice.Excel._Worksheet myCopy = null;              // Copy variable

        UserSettings UserSettings;
        #endregion

        public Excel(UserSettings UserSettings)
        {
            this.UserSettings = UserSettings;
        }

        private void CreateExcelHandle()
        {
            ExcelApp = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = false,        //hide from end user                
                DisplayAlerts = false  // hide "are you sure prompts"
            };
        }

        public void VerifySheetsExists(ExcelClassSettings ExcelClassSettings, Courses myCourses)
        {
            //==========================================================
            //
            // IN: UserInfo  - user specific Options
            //     myCourses - contains the courses to export
            //
            // Returns: Nothing - Events are raised to send status messages
            //
            // DESCRIPTION
            //    This procedure makes sure that a sheet with the class 
            // name exists in the template spreadsheet.
            //==========================================================

            // Does the template Exist?
            if (!File.Exists(ExcelClassSettings.TemplatePathandFileName))
            {
                throw new Exception("The Excel template " + ExcelClassSettings.TemplatePathandFileName + " does not exist");
            }

            // Classes contain a unique list of - Make sure that a single copy of each exits
            ExcelLinks ClassPointer = new ExcelLinks(myCourses.Count());
            for (int i = 0; i < myCourses.Count(); i++)
            {
                Course C = myCourses[i];
                ClassPointer.Add(C.Name, C.ItemNumber, i);
            }

            // Is Excel open?
            if (ExcelApp == null) { CreateExcelHandle(); }

            try
            {
                // open template
                ExcelWorkbook = ExcelApp.Workbooks.Open(
                    ExcelClassSettings.TemplatePathandFileName,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                // COUNT starts at 1!!!!!!!!!
                for (int i = 1; i < (ExcelWorkbook.Sheets.Count + 1); i++)
                {
                    mySheet = (MSOffice.Excel.Worksheet)ExcelWorkbook.Sheets[i];
                    if (mySheet.Name.ToLower() == CLASS_TEMPLATE)
                    {
                        ExcelClassWorksheet = mySheet;
                    }
                    if (mySheet.Name.ToLower() == ROLL_TEMPLATE)
                    {
                        ExcelRollWorksheet = mySheet;
                    }
                    if (mySheet.Name.ToLower() == LAB_TEMPLATE)
                    {
                        ExcelLabWorksheet = mySheet;
                    }
                }

                if (ExcelClassWorksheet == null)
                {
                    throw new Exception("The Excel template does not contain a 'class' template!");
                }

                if (ExcelRollWorksheet == null)
                {
                    throw new Exception("The Excel template does not contain a 'roll' template!");
                }

                if (ExcelLabWorksheet == null)
                {
                    // try and copy the roll sheet as the Lab
                    string CopyName = ROLL_TEMPLATE + " (2)";
                    ExcelRollWorksheet.Copy(Missing.Value, ExcelRollWorksheet);
                    // Get a handle of the sheet
                    ExcelLabWorksheet = (MSOffice.Excel.Worksheet)ExcelWorkbook.Sheets[CopyName];
                    // Change the sheet name
                    ExcelLabWorksheet.Name = "Lab";
                }

                // now loop through all Classes and verify that they exist
                bool FoundSheet = false;

                for (int myClass = 0; myClass < ClassPointer.Count(); myClass++)
                {
                    FoundSheet = false;  // set to not found
                    for (int i = 1; i < (ExcelWorkbook.Sheets.Count + 1); i++)
                    {
                        mySheet = (MSOffice.Excel.Worksheet)ExcelWorkbook.Sheets[i];
                        if (mySheet.Name.ToLower() == ClassPointer[myClass].ClassName.ToLower())
                        {
                            FoundSheet = true;
                            break;
                        }
                    }
                    if (!FoundSheet)
                    {
                        string CopyName = ExcelClassWorksheet.Name + " (2)";
                        ExcelClassWorksheet.Activate();
                        // create a copy of the class sheet
                        ExcelClassWorksheet.Copy(Missing.Value, ExcelClassWorksheet);
                        // Get a handle of the sheet
                        mySheet = (MSOffice.Excel.Worksheet)ExcelWorkbook.Sheets[CopyName];
                        // Change the sheet name
                        mySheet.Name = ClassPointer[myClass].ClassName;
                    }
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                // clean up local varables
                mySheet = null;
                ExcelClassWorksheet = null;
                try
                {
                    ExcelWorkbook.Save();
                    ExcelWorkbook.Close(Missing.Value, Missing.Value, Missing.Value);
                }
                catch
                { }
                ExcelWorkbook = null;
            }
        }

        private int CellLetterToRow(string ColumnName)
        {
            if (ColumnName.Length == 1)
            {
                return (int)ColumnName[0] - (int)'A' + 1;
            }
            else
            {
                // Untested as of 02/16/2008
                int ReturnValue = (int)ColumnName[0] - (int)'A' + 1;
                ReturnValue = ReturnValue * 26 + 26; // Get the correct multiple of 26
                ReturnValue += (int)ColumnName[1] - (int)'A' + 1;

                return ReturnValue;
            }

        }

        private bool CreateConfigurationFromSpreadSheet(ExcelRollSettings ExcelRollSettings, ExcelClassSettings ExcelClassSettings)
        {
            //==========================================================
            //
            // IN: ExcelClassSettings - Settings for the Class Sheets
            //     ExcelRollSettings - Settings for the Roll Sheet
            //     myCourses - contains the courses to export
            //     SelectedQuarter - Contains the YRQ and Quarter Name if needed for
            //
            // Returns: Nothing - Events are raised to send status messages
            //
            // DESCRIPTION
            //    This procedure exports Student information to an Excel
            // spreadsheet.
            //==========================================================
            bool RetVal = false;
            // Does the template Exist?
            if (!File.Exists(ExcelClassSettings.TemplatePathandFileName))
            { throw new Exception("The Excel template " + ExcelClassSettings.TemplatePathandFileName + " does not exist"); }

            // Is Excel open?
            if (ExcelApp == null) { CreateExcelHandle(); }

            // Open the workbook
            try
            {
                // open template
                SendMessage(this, new Information("Opening the workbook copy."));
                ExcelWorkbook = ExcelApp.Workbooks.Open(
                    ExcelClassSettings.TemplatePathandFileName,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value);

                ExcelApp.DisplayAlerts = false; // supress "are you sure" prompts
                SendMessage(this, new Information("Trying to find the config worksheet."));

                // COUNT starts at 1!!!!!!!!!  
                int NumberOfSheets = ExcelWorkbook.Sheets.Count;
                for (int i = NumberOfSheets; i > 0; i--)
                {
                    mySheet = (MSOffice.Excel._Worksheet)ExcelWorkbook.Sheets[i];

                    if (mySheet.Name.ToLower() == CONFIG_TEMPLATE)
                    {
                        ExcelConfigWorksheet = mySheet;
                        break;
                    }
                }
                if (ExcelConfigWorksheet == null)
                {
                    // TODO - make a new workshhet
                    mySheet = (MSOffice.Excel._Worksheet)ExcelWorkbook.Sheets[1];
                    // create a new worksheet
                    ExcelConfigWorksheet = (MSOffice.Excel._Worksheet)ExcelWorkbook.Sheets.Add(mySheet, Missing.Value, Missing.Value, Missing.Value);
                    // Change the sheet name
                    ExcelConfigWorksheet.Name = CONFIG_TEMPLATE;
                }


                // write the data names
                ExcelConfigWorksheet.Cells[6, 1].Value = "FirstNameColumnLetter";
                ExcelConfigWorksheet.Cells[7, 1].Value = "FirstStudent";
                ExcelConfigWorksheet.Cells[8, 1].Value = "ItemCell";
                ExcelConfigWorksheet.Cells[9, 1].Value = "LastNameColumnLetter";
                ExcelConfigWorksheet.Cells[10, 1].Value = "SIDColumnLetter";
                ExcelConfigWorksheet.Cells[11, 1].Value = "SIDLast4ColumnLetter";
                ExcelConfigWorksheet.Cells[12, 1].Value = "OptHead1ColumnLetter";
                ExcelConfigWorksheet.Cells[13, 1].Value = "OptHead2ColumnLetter";
                ExcelConfigWorksheet.Cells[14, 1].Value = "OptHead3ColumnLetter";

                ExcelConfigWorksheet.Cells[23, 1].Value = "ClassIncrement";
                ExcelConfigWorksheet.Cells[24, 1].Value = "FirstClass";
                ExcelConfigWorksheet.Cells[25, 1].Value = "FirstMondayDayCell";
                ExcelConfigWorksheet.Cells[26, 1].Value = "FirstNameColumnLetter";
                ExcelConfigWorksheet.Cells[27, 1].Value = "LastNameColumnLetter";
                ExcelConfigWorksheet.Cells[28, 1].Value = "MondayDateCell";
                ExcelConfigWorksheet.Cells[29, 1].Value = "ClassNameColumnLetter";

                // write the data                
                ExcelConfigWorksheet.Cells[6, 2].Value = ExcelClassSettings.FirstNameColumnLetter;
                ExcelConfigWorksheet.Cells[7, 2].Value = ExcelClassSettings.FirstStudent.ToString();
                ExcelConfigWorksheet.Cells[8, 2].Value = ExcelClassSettings.ItemCell;
                ExcelConfigWorksheet.Cells[9, 2].Value = ExcelClassSettings.LastNameColumnLetter;
                ExcelConfigWorksheet.Cells[10, 2].Value = ExcelClassSettings.SIDColumnLetter;
                ExcelConfigWorksheet.Cells[11, 2].Value = ExcelClassSettings.SIDLast4ColumnLetter;
                ExcelConfigWorksheet.Cells[12, 2].Value = ExcelClassSettings.OptHead1ColumnLetter;
                ExcelConfigWorksheet.Cells[13, 2].Value = ExcelClassSettings.OptHead2ColumnLetter;
                ExcelConfigWorksheet.Cells[14, 2].Value = ExcelClassSettings.OptHead3ColumnLetter;

                ExcelConfigWorksheet.Cells[23, 2].Value = ExcelRollSettings.ClassIncrement.ToString();
                ExcelConfigWorksheet.Cells[24, 2].Value = ExcelRollSettings.FirstClass.ToString();
                ExcelConfigWorksheet.Cells[25, 2].Value = ExcelRollSettings.FirstMondayDayCell;
                ExcelConfigWorksheet.Cells[26, 2].Value = ExcelRollSettings.FirstNameColumnLetter;
                ExcelConfigWorksheet.Cells[27, 2].Value = ExcelRollSettings.LastNameColumnLetter;
                ExcelConfigWorksheet.Cells[28, 2].Value = ExcelRollSettings.MondayDateCell;
                ExcelConfigWorksheet.Cells[29, 2].Value = ExcelRollSettings.ClassNameColumnLetter;
            }
            catch (Exception Ex)
            {
                string M = Ex.Message;

            }
            finally
            {
                // clean up local varables
                mySheet = null;
                myCopy = null;
                ExcelClassWorksheet = null;
                ExcelApp.Visible = true;        //show to end user
                ExcelApp.DisplayAlerts = true;
                try
                {
                    ExcelWorkbook.Save();
                }
                catch
                {
                    System.Windows.Forms.MessageBox.Show("Unable to save workbook.", "Saved Failed!");
                }
                finally
                {
                    ExcelWorkbook = null;
                }
                RetVal = true;
            }
            return RetVal;
        }

        public bool LoadConfigurationFromSpreadSheet(ExcelRollSettings ExcelRollSettings, ExcelClassSettings ExcelClassSettings)
        {
            //==========================================================
            //
            // IN: ExcelClassSettings - Settings for the Class Sheets
            //     ExcelRollSettings - Settings for the Roll Sheet
            //     myCourses - contains the courses to export
            //     SelectedQuarter - Contains the YRQ and Quarter Name if needed for
            //
            // Returns: Nothing - Events are raised to send status messages
            //
            // DESCRIPTION
            //    This procedure exports Student information to an Excel
            // spreadsheet.
            //==========================================================
            // Does the template Exist?
            if (!File.Exists(ExcelClassSettings.TemplatePathandFileName))
            { throw new Exception("The Excel template " + ExcelClassSettings.TemplatePathandFileName + " does not exist"); }

            // Is Excel open?
            if (ExcelApp == null) { CreateExcelHandle(); }

            // Open the workbook
            try
            {
                // open template
                SendMessage(this, new Information("Opening the workbook copy."));
                ExcelWorkbook = ExcelApp.Workbooks.Open(
                    ExcelClassSettings.TemplatePathandFileName,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value);

                ExcelApp.DisplayAlerts = false; // supress "are you sure" prompts
                SendMessage(this, new Information("Trying to find the config worksheet."));

                // COUNT starts at 1!!!!!!!!!  
                int NumberOfSheets = ExcelWorkbook.Sheets.Count;
                for (int i = NumberOfSheets; i > 0; i--)
                {
                    mySheet = (MSOffice.Excel._Worksheet)ExcelWorkbook.Sheets[i];

                    if (mySheet.Name.ToLower() == CONFIG_TEMPLATE)
                    {
                        ExcelConfigWorksheet = mySheet;
                        break;
                    }
                }

                if (ExcelConfigWorksheet == null)
                {
                    return CreateConfigurationFromSpreadSheet(ExcelRollSettings, ExcelClassSettings);
                }

                // read the data                
                ExcelClassSettings.FirstNameColumnLetter = ExcelConfigWorksheet.Cells[6, 2].Text;
                ExcelClassSettings.FirstStudent = Convert.ToInt32(ExcelConfigWorksheet.Cells[7, 2].Text);
                ExcelClassSettings.ItemCell = ExcelConfigWorksheet.Cells[8, 2].Text;
                ExcelClassSettings.LastNameColumnLetter = ExcelConfigWorksheet.Cells[9, 2].Text;
                ExcelClassSettings.SIDColumnLetter = ExcelConfigWorksheet.Cells[10, 2].Text;
                ExcelClassSettings.SIDLast4ColumnLetter = ExcelConfigWorksheet.Cells[11, 2].Text;
                ExcelClassSettings.OptHead1ColumnLetter = ExcelConfigWorksheet.Cells[12, 2].Text;
                ExcelClassSettings.OptHead2ColumnLetter = ExcelConfigWorksheet.Cells[13, 2].Text;
                ExcelClassSettings.OptHead3ColumnLetter = ExcelConfigWorksheet.Cells[14, 2].Text;

                ExcelRollSettings.ClassIncrement = Convert.ToInt32(ExcelConfigWorksheet.Cells[23, 2].Text);
                ExcelRollSettings.FirstClass = Convert.ToInt32(ExcelConfigWorksheet.Cells[24, 2].Text);
                ExcelRollSettings.FirstMondayDayCell = ExcelConfigWorksheet.Cells[25, 2].Text;
                ExcelRollSettings.FirstNameColumnLetter = ExcelConfigWorksheet.Cells[26, 2].Text;
                ExcelRollSettings.LastNameColumnLetter = ExcelConfigWorksheet.Cells[27, 2].Text;
                ExcelRollSettings.MondayDateCell = ExcelConfigWorksheet.Cells[28, 2].Text;
                ExcelRollSettings.ClassNameColumnLetter = ExcelConfigWorksheet.Cells[29, 2].Text;

                SendMessage(this, new Information("Finished reading config infornmation."));
            }
            catch (Exception Ex)
            {
                string M = Ex.Message;
                SendMessage(this, new Information(M));
            }

            // clean up local varables
            mySheet = null;
            myCopy = null;
            ExcelConfigWorksheet = null;
            ExcelWorkbook.Close();
            ExcelWorkbook = null;
            return true;

        }

        public void Export(ExcelClassSettings ExcelClassSettings, ExcelRollSettings ExcelRollSettings, Courses myCourses, Quarter SelectedQuarter)
        {
            //==========================================================
            //
            // IN: ExcelClassSettings - Settings for the Class Sheets
            //     ExcelRollSettings - Settings for the Roll Sheet
            //     myCourses - contains the courses to export
            //     SelectedQuarter - Contains the YRQ and Quarter Name if needed for
            //
            // Returns: Nothing - Events are raised to send status messages
            //
            // DESCRIPTION
            //    This procedure exports Student information to an Excel
            // spreadsheet.
            //==========================================================

            string LocalSaveAsFileName = ExcelClassSettings.SaveAsPathandFileName;

            if (ExcelClassSettings.SaveFileNameIncludesQuarter)
            {
                LocalSaveAsFileName = ExcelClassConfiguration.GenerateSaveFileName(ExcelClassSettings, SelectedQuarter);
            }

            // Does the template Exist?
            if (!File.Exists(ExcelClassSettings.TemplatePathandFileName))
            { throw new Exception("The Excel template " + ExcelClassSettings.TemplatePathandFileName + " does not exist"); }

            // Verify the Course sheets exist (just in case the caller didn't call this
            // first
            VerifySheetsExists(ExcelClassSettings, myCourses);

            #region Create a count of of each type of class

            ExcelLinks ClassesPointer = new ExcelLinks(myCourses.Count());

            int TotalExportedCourse = 0;
            SendMessage(this, new Information("Scanning for Classes to export"));
            for (int i = 0; i < myCourses.Count(); i++)
            {
                Course C = myCourses[i];
                if (C.Export)  // make sure the course is to be exported
                {
                    SendMessage(this, new Information("Found " + C.Name));
                    ClassesPointer.Add(C.Name, C.ItemNumber, i);
                    TotalExportedCourse++;
                }
            }

            #endregion

            #region Does the LocalSaveAsFileName Exist?

            FileInformation FE = new FileInformation("The file " + LocalSaveAsFileName + " already exists.", LocalSaveAsFileName);
            bool DeleteFileError = false;

            // check to see if file exists
            SendMessage(this, new Information("Preparing " + LocalSaveAsFileName + "."));
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
                // Copy the template
                File.Copy(ExcelClassSettings.TemplatePathandFileName, FE.NewFileNameandPath);
            }

            #endregion

            // Is Excel open?
            if (ExcelApp == null) { CreateExcelHandle(); }

            // Open the workbook
            try
            {
                // open template
                SendMessage(this, new Information("Opening the workbook copy."));
                ExcelWorkbook = ExcelApp.Workbooks.Open(
                    FE.NewFileNameandPath,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value);

                ExcelApp.DisplayAlerts = false; // supress "are you sure" prompts

                //ExcelApp.Visible = true;

                SendMessage(this, new Information("Delete unneeded sheets."));
                #region Delete unneeded sheets

                // COUNT starts at 1!!!!!!!!!  
                int NumberOfSheets = ExcelWorkbook.Sheets.Count;
                for (int i = NumberOfSheets; i > 0; i--)
                {
                    mySheet = (MSOffice.Excel._Worksheet)ExcelWorkbook.Sheets[i];

                    if (mySheet.Name.ToLower() == ROLL_TEMPLATE)
                    {
                        if (!ExcelRollSettings.Export)                        
                        {
                            mySheet.Delete();
                        }
                    }
                    else if (mySheet.Name.ToLower() == LAB_TEMPLATE)
                    {
                        if (!ExcelRollSettings.ExportLab)
                        {
                            mySheet.Delete();
                        }
                    }
                    else if (ExcelClassSettings.Export)
                    {
                        if (!ClassesPointer.Contains(mySheet.Name.ToLower()))
                        {
                            mySheet.Delete();
                        }
                    }
                    else
                    {
                        mySheet.Delete();
                    }
                }
                #endregion

                //-------------------------------------------------------------------------------------
                //-------------------------------------------------------------------------------------
                //-------------------------------------------------------------------------------------

                #region Create Lab Sheets

                if (ExcelRollSettings.ExportLab)
                {
                    // Find the correct sheet
                    NumberOfSheets = ExcelWorkbook.Sheets.Count;
                    for (int i = NumberOfSheets; i > 0; i--)
                    {
                        mySheet = (MSOffice.Excel._Worksheet)ExcelWorkbook.Sheets[i];
                        if (mySheet.Name.ToLower() == LAB_TEMPLATE)
                        {
                            ExcelLabWorksheet = mySheet;
                            break;
                        }
                    }

                    

                    #region Date Cells

                    // FirstDay Cell - The First Teaching day of the current quarter Cell Location.
                    int CellRow = 0;
                    int CellColumn = 0;
                    string ColumnLetters = "";

                    if (ExcelRollSettings.FirstMondayDayCell == "")
                    {
                        ExcelRollSettings.FirstMondayDayCell = ExcelRollConfiguration.FirstMondayDayCellDefault;
                    }

                    foreach (char c in ExcelRollSettings.FirstMondayDayCell.ToUpper())
                    {
                        if ((int)c > 64)
                        {
                            // Letter
                            ColumnLetters += c;
                        }
                        else
                        {
                            // Row
                            CellRow = CellRow * 10 + (int)c - 48;
                        }
                    }
                    CellColumn = CellLetterToRow(ColumnLetters);
                    string CellValue = ExcelRollSettings.FirstDay.ToShortDateString();
                    SendMessage(this, new Information("Setting the first monday date."));
                    ExcelLabWorksheet.Cells[CellRow, CellColumn] = CellValue; 
                    
                    //MondayDateCell
                    CellRow = 0;
                    CellColumn = 0;
                    ColumnLetters = "";
                    if (ExcelRollSettings.MondayDateCell == "")
                    {
                        ExcelRollSettings.MondayDateCell = ExcelRollConfiguration.MondayDateCellDefault;
                    }
                    foreach (char c in ExcelRollSettings.MondayDateCell.ToUpper())
                    {
                        if ((int)c > 64)
                        {
                            // Letter
                            ColumnLetters += c;
                        }
                        else
                        {
                            // Row
                            CellRow = CellRow * 10 + (int)c - 48;
                        }
                    }
                    CellColumn = CellLetterToRow(ColumnLetters);
                    // CellValue has not changed
                    ExcelLabWorksheet.Cells[CellRow, CellColumn] = CellValue;

                    #endregion
                    
                    #region Add Students to Lab Sheet

                    if (ExcelRollSettings.FirstClass < 3)
                    {
                        ExcelRollSettings.FirstClass = ExcelRollConfiguration.FirstClassDefault;
                    }
                    if (ExcelRollSettings.ClassIncrement < 3)
                    {
                        ExcelRollSettings.ClassIncrement = ExcelRollConfiguration.ClassIncrementDefault;
                    }

                    int Row = ExcelRollSettings.FirstClass - ExcelRollSettings.ClassIncrement;
                    for (int i = 0; i < myCourses.Count(); i++)
                    {
                        Course C = myCourses[i];
                        if (C.Export)  // make sure the course is to be exported
                        {
                            // Adjust starting row
                            // The first time throug Row = -37 + 40 = 3 = ExcelRollSettings.FirstClass
                            Row += ExcelRollSettings.ClassIncrement;

                            // Enter class Header two lines above class
                            CellRow = Row - 2;
                            if (ExcelRollSettings.ClassNameColumnLetter == "")
                            {
                                ExcelRollSettings.ClassNameColumnLetter = ExcelRollConfiguration.ClassNameColumnLetterDefault;
                            }
                            CellColumn = CellLetterToRow(ExcelRollSettings.ClassNameColumnLetter);
                            CellValue = C.SheetName();  // C.Display;
                            ExcelLabWorksheet.Cells[CellRow, CellColumn] = CellValue; 
                            
                            int CourseCount = C.Students.GetLength(0);
                            if (CourseCount > ExcelRollSettings.ClassIncrement)
                            {
                                CourseCount = ExcelRollSettings.ClassIncrement;
                                SendMessage(this, new Information("Course has more students then space in the roll sheet.  Adjust roll sheet and modify Options and re-extract."));
                            }
                            for (int s = 0; s < CourseCount; s++)
                            {
                                CellRow = s + Row;
                                if (ExcelRollSettings.LastNameColumnLetter == "")
                                {
                                    ExcelRollSettings.LastNameColumnLetter = ExcelRollConfiguration.LastNameColumnLetterDefault;
                                }
                                CellColumn = CellLetterToRow(ExcelRollSettings.LastNameColumnLetter);
                                CellValue = C.Students[s].LastName;
                                ExcelLabWorksheet.Cells[CellRow, CellColumn] = CellValue; 
                                
                                if (ExcelRollSettings.FirstNameColumnLetter == "")
                                {
                                    ExcelRollSettings.FirstNameColumnLetter = ExcelRollConfiguration.FirstNameColumnLetterDefault;
                                }
                                CellColumn = CellLetterToRow(ExcelRollSettings.FirstNameColumnLetter);
                                CellValue = C.Students[s].FirstName;
                                ExcelLabWorksheet.Cells[CellRow, CellColumn] = CellValue;                                 
                            }
                        }
                    }

                    #endregion

                }

                #endregion

                //-------------------------------------------------------------------------------------
                //-------------------------------------------------------------------------------------
                //-------------------------------------------------------------------------------------

                #region Create Roll

                if ((ExcelRollSettings.Export) && (ExcelRollWorksheet != null))
                {
                    // Find the correct sheet
                    NumberOfSheets = ExcelWorkbook.Sheets.Count;
                    for (int i = NumberOfSheets; i > 0; i--)
                    {
                        mySheet = (MSOffice.Excel._Worksheet)ExcelWorkbook.Sheets[i];

                        if (mySheet.Name.ToLower() == ROLL_TEMPLATE)
                        {
                            ExcelRollWorksheet = mySheet;
                            break;
                        }
                    }
                    
                    SendMessage(this, new Information("Creating a Roll sheet."));

                    #region Date Cells

                    // FirstDay Cell - The First Teaching day of the current quarter Cell Location.
                    int CellRow = 0;
                    int CellColumn = 0;
                    string ColumnLetters = "";

                    if (ExcelRollSettings.FirstMondayDayCell == "")
                    {
                        ExcelRollSettings.FirstMondayDayCell = ExcelRollConfiguration.FirstMondayDayCellDefault;
                    }

                    foreach (char c in ExcelRollSettings.FirstMondayDayCell.ToUpper())
                    {
                        if ((int)c > 64)
                        {
                            // Letter
                            ColumnLetters += c;
                        }
                        else
                        {
                            // Row
                            CellRow = CellRow * 10 + (int)c - 48;
                        }
                    }
                    CellColumn = CellLetterToRow(ColumnLetters);
                    string CellValue = ExcelRollSettings.FirstDay.ToShortDateString();
                    ExcelRollWorksheet.Cells[CellRow, CellColumn] = CellValue;


                    //MondayDateCell
                    CellRow = 0;
                    CellColumn = 0;
                    ColumnLetters = "";
                    if (ExcelRollSettings.MondayDateCell == "")
                    {
                        ExcelRollSettings.MondayDateCell = ExcelRollConfiguration.MondayDateCellDefault;
                    }
                    foreach (char c in ExcelRollSettings.MondayDateCell.ToUpper())
                    {
                        if ((int)c > 64)
                        {
                            // Letter
                            ColumnLetters += c;
                        }
                        else
                        {
                            // Row
                            CellRow = CellRow * 10 + (int)c - 48;
                        }
                    }
                    CellColumn = CellLetterToRow(ColumnLetters);
                    // CellValue has not changed
                    ExcelRollWorksheet.Cells[CellRow, CellColumn] = CellValue; 

                    #endregion

                    #region Add Students to Roll

                    if (ExcelRollSettings.FirstClass < 3)
                    {
                        ExcelRollSettings.FirstClass = ExcelRollConfiguration.FirstClassDefault;
                    }
                    if (ExcelRollSettings.ClassIncrement < 3)
                    {
                        ExcelRollSettings.ClassIncrement = ExcelRollConfiguration.ClassIncrementDefault;
                    }

                    int Row = ExcelRollSettings.FirstClass - ExcelRollSettings.ClassIncrement;
                    for (int i = 0; i < myCourses.Count(); i++)
                    {
                        Course C = myCourses[i];
                        if (C.Export)  // make sure the course is to be exported
                        {
                            // Adjust starting row
                            // The first time throug Row = -37 + 40 = 3 = ExcelRollSettings.FirstClass
                            Row += ExcelRollSettings.ClassIncrement;

                            // Enter class Header two lines above class
                            CellRow = Row - 2;
                            if (ExcelRollSettings.ClassNameColumnLetter == "")
                            {
                                ExcelRollSettings.ClassNameColumnLetter = ExcelRollConfiguration.ClassNameColumnLetterDefault;
                            }
                            CellColumn = CellLetterToRow(ExcelRollSettings.ClassNameColumnLetter);
                            CellValue = C.SheetName();  // C.Display;
                            ExcelRollWorksheet.Cells[CellRow, CellColumn] = CellValue; 

                            int CourseCount = C.Students.GetLength(0);
                            if (CourseCount > ExcelRollSettings.ClassIncrement)
                            {
                                CourseCount = ExcelRollSettings.ClassIncrement;
                                SendMessage(this, new Information("Course has more students then space in the roll sheet.  Adjust roll sheet and modify Options and re-extract."));
                            }
                            for (int s = 0; s < CourseCount; s++)
                            {
                                CellRow = s + Row;
                                if (ExcelRollSettings.LastNameColumnLetter == "")
                                {
                                    ExcelRollSettings.LastNameColumnLetter = ExcelRollConfiguration.LastNameColumnLetterDefault;
                                }
                                CellColumn = CellLetterToRow(ExcelRollSettings.LastNameColumnLetter);
                                CellValue = C.Students[s].LastName;
                                ExcelRollWorksheet.Cells[CellRow, CellColumn] = CellValue; 

                                if (ExcelRollSettings.FirstNameColumnLetter == "")
                                {
                                    ExcelRollSettings.FirstNameColumnLetter = ExcelRollConfiguration.FirstNameColumnLetterDefault;
                                }
                                CellColumn = CellLetterToRow(ExcelRollSettings.FirstNameColumnLetter);
                                CellValue = C.Students[s].FirstName;
                                ExcelRollWorksheet.Cells[CellRow, CellColumn] = CellValue; 
                            }
                            // WAITLIST
                            // For the start row 
                            int WaitlistCellRow = CourseCount + Row;


                            if (ExcelRollSettings.ExportWaitlist)
                            {
                                CourseCount = C.Waitlist.GetLength(0);
                                for (int s = 0; s < CourseCount; s++)
                                {
                                    CellRow = s + WaitlistCellRow;
                                    if (ExcelRollSettings.LastNameColumnLetter == "")
                                    {
                                        ExcelRollSettings.LastNameColumnLetter = ExcelRollConfiguration.LastNameColumnLetterDefault;
                                    }
                                    CellColumn = CellLetterToRow(ExcelRollSettings.LastNameColumnLetter);
                                    CellValue = C.Waitlist[s].LastName;
                                    ExcelRollWorksheet.Cells[CellRow, CellColumn] = CellValue;

                                    if (ExcelRollSettings.FirstNameColumnLetter == "")
                                    {
                                        ExcelRollSettings.FirstNameColumnLetter = ExcelRollConfiguration.FirstNameColumnLetterDefault;
                                    }
                                    CellColumn = CellLetterToRow(ExcelRollSettings.FirstNameColumnLetter);
                                    CellValue = C.Waitlist[s].FirstName;
                                    ExcelRollWorksheet.Cells[CellRow, CellColumn] = CellValue;
                                }
                            }
                        }
                    }

                    #endregion

                }

                #endregion

                if (ExcelClassSettings.Export)
                {
                    SendMessage(this, new Information("Adjusting the sheets for the selected courses."));
                    #region Adjusting the sheets for the selected courses

                    // add the correct number of copies for each course
                    NumberOfSheets = ExcelWorkbook.Sheets.Count + 1;
                    for (int WorkbookIndex = 1; WorkbookIndex < NumberOfSheets; WorkbookIndex++)
                    {
                        mySheet = (MSOffice.Excel._Worksheet)ExcelWorkbook.Sheets[WorkbookIndex];
                        if (ClassesPointer.Contains(mySheet.Name.ToLower()))
                        {
                            // do we have enough?
                            //  .ToLower() fix 2008-09-19
                            ExcelLink excelLink = ClassesPointer.GetExcelLink(mySheet.Name.ToLower());

                            // make sure that each loop starts with the template
                            mySheet = (MSOffice.Excel._Worksheet)ExcelWorkbook.Sheets[WorkbookIndex];
                            string OriginalName = mySheet.Name;
                            string CopyName = OriginalName + " (2)";


                            // add more sheets??
                            int MyCount = excelLink.Count();
                            int LastClass = MyCount - 1;
                            for (int index = 0; index < MyCount; index++)
                            {
                                string NewName = mySheet.Name + " (" + excelLink.GetPointers[index].ItemNumber + ")";
                                if (index == LastClass)
                                {
                                    // rename last sheet
                                    mySheet.Activate();
                                    mySheet.Name = NewName;
                                }
                                else
                                {
                                    // create a copy of the class sheet
                                    mySheet.Copy(Missing.Value, mySheet);
                                    // Get a handle of the sheet
                                    myCopy = (MSOffice.Excel.Worksheet)ExcelWorkbook.Sheets[CopyName];
                                    // Change the sheet name
                                    myCopy.Name = NewName;

                                }
                            }
                        }
                    }

                    NumberOfSheets = ExcelWorkbook.Sheets.Count + 1;
                    for (int WorkbookIndex = 1; WorkbookIndex < NumberOfSheets; WorkbookIndex++)
                    {
                        mySheet = (MSOffice.Excel._Worksheet)ExcelWorkbook.Sheets[WorkbookIndex];
                        string Name = mySheet.Name;
                        int start = Name.IndexOf("(") + 1;
                        int end = Name.IndexOf(")");
                        if ((end - start) > 0)
                        {
                            string ItemNumber = Name.Substring(start, end - start);
                            string OriginalName = Name.Substring(0, start - 3);

                            ExcelLink E = ClassesPointer.GetExcelNewNameLink(Name);
                            E.Set(ItemNumber, WorkbookIndex);
                        }
                    }

                    #endregion

                    SendMessage(this, new Information("Adding student names to sheets."));
                    #region Add names to each sheet
                    ExcelApp.Visible = true;

                    for (int j = 1; j < (ExcelWorkbook.Sheets.Count + 1); j++)
                    {
                        mySheet = (MSOffice.Excel._Worksheet)ExcelWorkbook.Sheets[j];
                        int myCourseIndex = ClassesPointer.IndexOf(j);

                        if (myCourseIndex != -1)
                        {
                            // Item Cell
                            int CellRow = 0;
                            int CellColumn = 0;
                            string ColumnLetters = "";

                            if (ExcelClassSettings.ItemCell == "")
                            {
                                ExcelClassSettings.ItemCell = ExcelClassConfiguration.ItemCellDefault;
                            }

                            foreach (char c in ExcelClassSettings.ItemCell.ToUpper())
                            {
                                if ((int)c > 64)
                                {
                                    // Letter
                                    ColumnLetters += c;
                                }
                                else
                                {
                                    // Row
                                    CellRow = CellRow * 10 + (int)c - 48;
                                }
                            }
                            CellColumn = CellLetterToRow(ColumnLetters);
                            string CellValue = "ID = " + myCourses[myCourseIndex].ItemNumber;
                            mySheet.Cells[CellRow, CellColumn] = CellValue;

                            // For the student Length
                            int StudentStart = myCourses[myCourseIndex].Students.GetLowerBound(0);
                            int StudentEnd = myCourses[myCourseIndex].Students.GetUpperBound(0);

                            // first line of students
                            int FirstStudent = ExcelClassSettings.FirstStudent;

                            if (FirstStudent < 1) { FirstStudent = ExcelClassConfiguration.FirstStudentDefault; }

                            for (int i = StudentStart; i <= StudentEnd; i++)
                            {
                                CellRow = FirstStudent + i - StudentStart;
                                // Last Name
                                if (ExcelClassSettings.LastNameColumnLetter == "")
                                {
                                    ExcelClassSettings.LastNameColumnLetter = ExcelClassConfiguration.LastNameColumnLetterDefault;
                                }
                                CellColumn = CellLetterToRow(ExcelClassSettings.LastNameColumnLetter);
                                CellValue = myCourses[myCourseIndex].Students[i].LastName;

                                mySheet.Cells[CellRow, CellColumn] = CellValue;

                                // First Name + MI
                                if (ExcelClassSettings.FirstNameColumnLetter == "")
                                {
                                    ExcelClassSettings.FirstNameColumnLetter = ExcelClassConfiguration.FirstNameColumnLetterDefault;
                                }
                                CellColumn = CellLetterToRow(ExcelClassSettings.FirstNameColumnLetter);
                                CellValue = myCourses[myCourseIndex].Students[i].FirstName;
                                if (ExcelClassSettings.ExportMiddleInitial)
                                {
                                    if (myCourses[myCourseIndex].Students[i].MiddleName.Length > 0)
                                    {
                                        CellValue += " " + myCourses[myCourseIndex].Students[i].MiddleName.Substring(0, 1);
                                    }
                                }
                                mySheet.Cells[CellRow, CellColumn] = CellValue;


                                // Export Option Items
                                if (ExcelClassSettings.ExportoptHead1)
                                {
                                    if (ExcelClassSettings.OptHead1ColumnLetter == "")
                                    {
                                        ExcelClassSettings.OptHead1ColumnLetter = ExcelClassConfiguration.OptHead1ColumnLetterDefault;
                                    }
                                    CellColumn = CellLetterToRow(ExcelClassSettings.OptHead1ColumnLetter);
                                    CellValue = myCourses[myCourseIndex].Students[i].Opt1;
                                    mySheet.Cells[CellRow, CellColumn] = CellValue;
                                }


                                if (ExcelClassSettings.ExportoptHead2)
                                {
                                    if (ExcelClassSettings.OptHead2ColumnLetter == "")
                                    {
                                        ExcelClassSettings.OptHead2ColumnLetter = ExcelClassConfiguration.OptHead2ColumnLetterDefault;
                                    }
                                    CellColumn = CellLetterToRow(ExcelClassSettings.OptHead2ColumnLetter);
                                    CellValue = myCourses[myCourseIndex].Students[i].Opt2;
                                    mySheet.Cells[CellRow, CellColumn] = CellValue;
                                }


                                if (ExcelClassSettings.ExportoptHead3)
                                {
                                    if (ExcelClassSettings.OptHead3ColumnLetter == "")
                                    {
                                        ExcelClassSettings.OptHead3ColumnLetter = ExcelClassConfiguration.OptHead3ColumnLetterDefault;
                                    }
                                    CellColumn = CellLetterToRow(ExcelClassSettings.OptHead3ColumnLetter);
                                    CellValue = myCourses[myCourseIndex].Students[i].Opt3;
                                    mySheet.Cells[CellRow, CellColumn] = CellValue;
                                }
                                
                                if (ExcelClassSettings.ExportSID)
                                {
                                    if (ExcelClassSettings.SIDColumnLetter == "")
                                    {
                                        ExcelClassSettings.SIDColumnLetter = ExcelClassConfiguration.SIDColumnLetterDefault;
                                    }
                                    CellColumn = CellLetterToRow(ExcelClassSettings.SIDColumnLetter);
                                    CellValue = myCourses[myCourseIndex].Students[i].SID;
                                    mySheet.Cells[CellRow, CellColumn] = CellValue;
                                }
                                if (ExcelClassSettings.ExportSIDLast4)
                                {
                                    if (ExcelClassSettings.SIDLast4ColumnLetter == "")
                                    {
                                        ExcelClassSettings.SIDLast4ColumnLetter = ExcelClassConfiguration.SIDLast4ColumnLetterDefault;
                                    }
                                    CellColumn = CellLetterToRow(ExcelClassSettings.SIDLast4ColumnLetter);
                                    CellValue = myCourses[myCourseIndex].Students[i].SID;
                                    CellValue = "'" + CellValue.Substring(CellValue.Length - 4, 4);
                                    mySheet.Cells[CellRow, CellColumn] = CellValue;
                                }
                            }

                            // WAITLIST
                            // For the start row 
                            int WaitlistCellRow = FirstStudent + StudentEnd - StudentStart + 1;

                            // reset index for waitlist
                            StudentStart = myCourses[myCourseIndex].Waitlist.GetLowerBound(0);
                            StudentEnd = myCourses[myCourseIndex].Waitlist.GetUpperBound(0);

                            
                            if (ExcelClassSettings.ExportWaitlist)
                            {
                                for (int i = StudentStart; i <= StudentEnd; i++)
                                {
                                    CellRow = WaitlistCellRow + i - StudentStart;
                                    // Last Name
                                    if (ExcelClassSettings.LastNameColumnLetter == "")
                                    {
                                        ExcelClassSettings.LastNameColumnLetter = ExcelClassConfiguration.LastNameColumnLetterDefault;
                                    }
                                    CellColumn = CellLetterToRow(ExcelClassSettings.LastNameColumnLetter);
                                    CellValue = myCourses[myCourseIndex].Waitlist[i].LastName;

                                    mySheet.Cells[CellRow, CellColumn] = CellValue;

                                    // First Name + MI
                                    if (ExcelClassSettings.FirstNameColumnLetter == "")
                                    {
                                        ExcelClassSettings.FirstNameColumnLetter = ExcelClassConfiguration.FirstNameColumnLetterDefault;
                                    }
                                    CellColumn = CellLetterToRow(ExcelClassSettings.FirstNameColumnLetter);
                                    CellValue = myCourses[myCourseIndex].Waitlist[i].FirstName;
                                    if (ExcelClassSettings.ExportMiddleInitial)
                                    {
                                        if (myCourses[myCourseIndex].Waitlist[i].MiddleName.Length > 0)
                                        {
                                            CellValue += " " + myCourses[myCourseIndex].Waitlist[i].MiddleName.Substring(0, 1);
                                        }
                                    }
                                    mySheet.Cells[CellRow, CellColumn] = CellValue;

                                    if (ExcelClassSettings.ExportSID)
                                    {
                                        if (ExcelClassSettings.SIDColumnLetter == "")
                                        {
                                            ExcelClassSettings.SIDColumnLetter = ExcelClassConfiguration.SIDColumnLetterDefault;
                                        }
                                        CellColumn = CellLetterToRow(ExcelClassSettings.SIDColumnLetter);
                                        CellValue = myCourses[myCourseIndex].Waitlist[i].SID;
                                        mySheet.Cells[CellRow, CellColumn] = CellValue;
                                    }
                                    if (ExcelClassSettings.ExportSIDLast4)
                                    {
                                        if (ExcelClassSettings.SIDLast4ColumnLetter == "")
                                        {
                                            ExcelClassSettings.SIDLast4ColumnLetter = ExcelClassConfiguration.SIDLast4ColumnLetterDefault;
                                        }
                                        CellColumn = CellLetterToRow(ExcelClassSettings.SIDLast4ColumnLetter);
                                        CellValue = myCourses[myCourseIndex].Waitlist[i].SID;
                                        CellValue = "'" + CellValue.Substring(CellValue.Length - 4, 4);
                                        mySheet.Cells[CellRow, CellColumn] = CellValue;
                                    }
                                }
                            }
                        }
                    }

                    #endregion

                    SendMessage(this, new Information("Delete Unwanted Class Sheets."));
                    #region Delete Template Sheets

                    //for (int loop = 0; loop < Classes.Count; loop++)
                    //{
                    //    for (int i = 1; i < (ExcelWorkbook.Sheets.Count + 1); i++)
                    //    {
                    //        mySheet = (MSOffice.Excel._Worksheet)ExcelWorkbook.Sheets[i];
                    //        if (Classes.ContainsKey(mySheet.Name.ToLower()))
                    //        {
                    //            mySheet.Delete();
                    //        }
                    //    }
                    //}

                    #endregion
                }
                SendMessage(this, new Information("Finished exporting to Excel."));
            }
            catch
            {
                throw;
            }
            finally
            {
                // clean up local varables
                mySheet = null;
                myCopy = null;
                ExcelClassWorksheet = null;
                ExcelApp.Visible = true;        //show to end user
                ExcelApp.DisplayAlerts = true;
                try
                {
                    ExcelWorkbook.Save();
                }
                catch
                {
                    System.Windows.Forms.MessageBox.Show("Unable to save workbook.", "Saved Failed!");
                }
                finally
                {
                    ExcelWorkbook = null;
                }

            }
        }
    }
}
