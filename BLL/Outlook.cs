using System;
using MSOffice = Microsoft.Office.Interop;
using InstructorBriefcaseExtractor.Model;

namespace InstructorBriefcaseExtractor.BLL
{
    internal class Outlook
    {
        #region Events

        public event InstructorBriefcaseExtractor.Model.InformationalMessage MessageResults;

        protected virtual void SendMessage(object sender, EventArgs e)
        {
            MessageResults?.Invoke(sender, e);
        }

        #endregion

        public Outlook()
        {
        }

        public bool CreateEmail(string EmailTo, string Subject,string Body)
        {
            // Get hooks to MS Outlook
            MSOffice.Outlook.Application OutlookApp;
            MSOffice.Outlook.MailItem myMailItem;
            bool ReturnValue = true;
            try
            {
                OutlookApp = new MSOffice.Outlook.Application();                
                myMailItem = (MSOffice.Outlook.MailItem)OutlookApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
                myMailItem.Subject = Subject;
                myMailItem.To = EmailTo;
                myMailItem.Body = Body;
                myMailItem.Importance = Microsoft.Office.Interop.Outlook.OlImportance.olImportanceLow;
                myMailItem.Display(true);
            }
            catch 
            {
                OutlookApp = null;
                ReturnValue= false;
            }

            myMailItem = null;
            OutlookApp = null;

            return ReturnValue;
        }

        public void Export(OutlookSettings OutlookSettings, Courses myCourses)
        {
            //==========================================================
            //
            // IN: OutlookSettings  - user specific Options
            //     myCourses        - contains the courses to export
            //
            // Returns: Nothing - Events are raised to send status messages
            //
            // DESCRIPTION
            //    This procedure exports to an Outlook Distribution list
            //==========================================================
            // Verify that outlook is to be exported
            if (!OutlookSettings.Export) { return; } 

            // Get hooks to MS Outlook
            MSOffice.Outlook.Application OutlookApp;            
            MSOffice.Outlook.DistListItem myDistList;
            MSOffice.Outlook.MailItem myMailItem;
            MSOffice.Outlook.Recipients myRecipients;            

            try
            {
                SendMessage(this, new Information("Attempting to Open Outlook"));
                OutlookApp = new MSOffice.Outlook.Application();
            }
            catch (Exception Ex)
            {
                SendMessage(this, new Information("Unable toOpen Outlook - \r\n" + Ex.Message));
                OutlookApp = null;
                return;
            }           

            try
            {
                foreach (Course C in myCourses)
                {
                    if (C.Export)
                    {
                        // Any error are caught in the outermost loop
                        
                        // Create the Distribution List item
                        string DLName = C.DiskName(OutlookSettings.Underscore);
                        SendMessage(this, new Information("Creating Distribution list for - " + C.Name));

                        myDistList = (MSOffice.Outlook.DistListItem)OutlookApp.CreateItem(MSOffice.Outlook.OlItemType.olDistributionListItem);
                        myDistList.DLName = DLName;
                        
                        //SendOutputMessage(new InformationNotice("Creating MailItem.");
                        myMailItem = (MSOffice.Outlook.MailItem)OutlookApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);

                        //SendOutputMessage(new InformationNotice("Creating MailItem.Recipients.");
                        try
                        {
                            myRecipients = myMailItem.Recipients;
                        }
                        catch 
                        {                            
                            throw new Exception("Unable to Create Recipients in the Distribution list.  Please make sure Outlook is open.");
                        }
                        
                        foreach (Student S in C.Students)
                        {
                            string Email = S.ExtractedName + " <" + S.Email + ">";
                            myRecipients.Add(Email);                            
                        }

                        if (OutlookSettings.ExportWaitlist)
                        {
                            foreach (Student S in C.Waitlist)
                            {
                                string Email = S.ExtractedName + " <" + S.Email + ">";
                                myRecipients.Add(Email);
                            }
                        }

                        if (!myRecipients.ResolveAll())
                        {
                            for (int i = 1; i < myRecipients.Count; i++)
                            {
                                if (!myRecipients[i].Resolved)
                                { SendMessage(this, new Information("Could not resolve - " + myRecipients[i].Name)); }
                            }
                        }

                        myDistList.AddMembers(myRecipients); // Add to Distribution list
                        myDistList.Save();                   // myDistList.Saved
                        myDistList.Display(false);           // Show user the new list
                    }
                }
            }
            catch (Exception Ex)
            {
                SendMessage(this, new Information("An unexpected error occurred - " + Ex.Message));
                throw;
            }
            finally
            {
                // Clean up
                myRecipients = null;
                myDistList = null;
                myMailItem = null;
                OutlookApp = null;                
            }

            SendMessage(this, new Information("Export to Outlook Completed."));            
        }

    }
}
