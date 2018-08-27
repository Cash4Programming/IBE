using System;
using System.Collections;

using InstructorBriefcaseExtractor.DAL;     // CISWebRequest
using InstructorBriefcaseExtractor.Utility;

namespace InstructorBriefcaseExtractor.Model
{
    // This class is the collection of all courses
    public class Courses : CollectionBase
    {
        #region Events

        public event InformationalMessage MessageResults;
        protected virtual void SendMessage(object sender, EventArgs e)
        {
            MessageResults?.Invoke(sender, e);
        }

        #endregion

        // Variable names in java(script) code to look for
        private string JavaTicketVariable
        {
            get { return "ticket"; }
        }
        private string JavaRosterVariable
        {
            get { return "new rostrInfo("; }
        }
        private string JavaInstructorVariable
        {
            get { return "instrName"; }
        }

        // Creates an empty collection.
        public Courses()
        {
        }

        internal void AddTestData()
        {
            ExcelClassSettings ExcelClassSettings = new ExcelClassSettings();
            Add("", "MICHAEL JENCK", "MATH 075", "2121", "A782", "Fall 2008", ExcelClassSettings).Export = true;
            Add("", "MICHAEL JENCK", "MATH 075", "2122", "A782", "Fall 2008", ExcelClassSettings).Export = true;
            Add("", "MICHAEL JENCK", "MATH 075", "2123", "A782", "Fall 2008", ExcelClassSettings).Export = true;
        }


        public void AddFromCIS(CISLogin MyLogin,ExcelClassSettings ExcelClassSettings, string ActionURL)
        {
            SendMessage(this, new Information("Retrieving Course URL."));

            // Create the http request string for the Instructor
            string HTTPCoursesRequest = MyLogin.HTTPCoursesRequestString;
            SendMessage(this, new Information(HTTPCoursesRequest));

            // Retrieve the Web page with the courses information
            CISWebRequest myWebRequest = new CISWebRequest();
            string HTTPCourses = myWebRequest.RetrieveHTMLPage(ActionURL, HTTPCoursesRequest, CISWebRequest.HTMLRequestType.POST, 5000);

            if (HTTPCourses.ToLower().IndexOf("service unavailable") > 0)
            {
                SendMessage(this, new Information(HTTPCourses));
                throw new Exception("Service is unavailable.");
            }
            else if (HTTPCourses.ToLower().IndexOf("pin") > 0)
            {
                SendMessage(this, new Information(HTTPCourses));
                throw new Exception("You have entered an incorrect Employee ID or PIN.");
            }
            else
            {
                SendMessage(this, new Information(HTTPCourses));
            }

            // Extract ticket         
            ExtractValue EV = new ExtractValue();
            string ticketvalue = EV.FromJavaScript(HTTPCourses, JavaTicketVariable, "'", "'");
            SendMessage(this, new Information("Found Ticket Value " + ticketvalue));

            // Extract Number of classes
            string InstructorName = EV.FromJavaScript(HTTPCourses, JavaInstructorVariable, "'", "'");
            SendMessage(this, new Information("Found Instructor Name " + InstructorName));

            string RawHTML = "";
            string HTTPRequestString = "";
            String CourseName = "";
            String ItemNumber = "";
            int Start = 0;
            int End = 0;
            int HTTPStart = HTTPCourses.IndexOf(JavaRosterVariable);
            if (HTTPStart >= 0)
            {

                do
                {
                    End = HTTPCourses.IndexOf(")", HTTPStart);
                    String OneCourse = HTTPCourses.Substring(HTTPStart, End - HTTPStart);
                    Start = OneCourse.IndexOf("'") + 1;  // don't include the '
                    End = OneCourse.IndexOf("'", Start + 1);
                    ItemNumber = OneCourse.Substring(Start, End - Start);

                    Start = OneCourse.IndexOf("'", End + 1) + 1;
                    End = OneCourse.IndexOf("'", Start + 1);
                    CourseName = OneCourse.Substring(Start, End - Start).Trim();

                    SendMessage(this, new Information("\r\n\nFound Course " + ItemNumber + " - " + CourseName));

                    // get request String
                    CISCourse myCISCourses = new CISCourse(ItemNumber, ticketvalue, MyLogin.Quarter.YRQ);
                    HTTPRequestString = myCISCourses.HTTPRequestString;

                    // retrieve class data
                    RawHTML = myWebRequest.RetrieveHTMLPage(ActionURL, HTTPRequestString, CISWebRequest.HTMLRequestType.POST, 5000);
                    SendMessage(this, new Information(RawHTML));

                    this.Add(RawHTML, InstructorName, CourseName, ItemNumber, MyLogin.Quarter.YRQ, MyLogin.Quarter.Name, ExcelClassSettings);

                    HTTPStart = HTTPCourses.IndexOf(JavaRosterVariable, HTTPStart + 1);
                } while (HTTPStart > 0);
            }
            else
            {
                // error no courses found
                SendMessage(this, new Information(HTTPCourses));
                throw new Exception("No courses found for " + MyLogin.Quarter.Name + " which I have a YRQ " + MyLogin.Quarter.YRQ);
            }
        }

        private Course Add(String RawHTML, String InstructorName, String CourseName,
                           String ItemNumber, String YRQ, String QuarterName, ExcelClassSettings ExcelClassSettings)
        {
            Course value = new Course(RawHTML, InstructorName, CourseName, ItemNumber, YRQ, QuarterName, ExcelClassSettings);
            this.List.Add(value);
            return value;
        }

        public new void Clear()
        {
            this.List.Clear();
        }

        public new int Count()
        {
            return this.List.Count;
        }

        public void SetAllExport(bool ToExport)
        {
            foreach (Course myVar in this.List)
            {
                myVar.Export = ToExport;
            }
        }

        public Course this[String ItemNumber]
        {
            get
            {
                foreach (Course myVar in this.List)
                {
                    if (myVar.ItemNumber == ItemNumber) return myVar;
                }
                return null;
            }
        }

        public Course this[int index]
        {
            get { return (Course)this.List[index]; }
        }

        public int IndexOf(String ItemNumber)
        {
            foreach (Course myVar in this.List)
            {
                if (myVar.ItemNumber == ItemNumber) return this.List.IndexOf(myVar);
            }
            return -1;
        }

        public void Remove(String ItemNumber)
        {
            foreach (Course myVar in this.List)
            {
                if (myVar.ItemNumber == ItemNumber)
                {
                    this.List.Remove(myVar);
                    break;
                }
            }
        }

        public new void RemoveAt(int index)
        {
            this.List.RemoveAt(index);
        }



    }    // class Courses
}
