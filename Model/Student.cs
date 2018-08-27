
// *********************************************************************
//  Copyright ©2008, 2018 Michael Paul Jenck, All Rights Reserved
//  http://www.jenck.net/
// *********************************************************************
// This program is free software; you can rediStribute it and/or
// modify it under the terms of the GNU General Public License
// as published by the Free Software Foundation; either version 2
// of the License, or (at your Option) any later version.
//
// This program is diStributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program; if not, write to the
// Free Software Foundation, Inc.,
// 59 Temple Place - Suite 330,
// Boston, MA  02111-1307, USA.
// *********************************************************************
// Reference Works: None
// *********************************************************************
// Requires: Nothing
// *********************************************************************
//
// History
//
// Modified 2008-01-12 MPJ - Converted to C# and VS 2005
// Modified 2005-08-18 MPJ - Converted to dot.net
// Created  2004-10-17 MPJ
// *********************************************************************
// DESCRIPTION
//    This class contains an individual student.
// *********************************************************************

namespace InstructorBriefcaseExtractor.Model
{
    // For Yvcc as of 02/12/2009
    // Opt1 = ADD DATE
    // Opt2 = DROP DATE
    // Opt3 = Student Email

    // For Yvcc as of 01/12/2008
    // Opt1 = ADD DATE
    // Opt2 = DROP DATE
    // Opt3 = DAY PHONE
    
    // This class contains the information about a single student
    public class Student
    {
        public Student()
        {
        }
                  
#pragma warning disable IDE1006 // Naming Styles
        public string gr { get; set; }
#pragma warning restore IDE1006 // Naming Styles
        public string Email { get; set; }
        public string EnrolledCredit { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string MiddleName { get; set; }
        public string Opt1 { get; set; }
        public string Opt2 { get; set; }
        public string Opt3 { get; set; }
        public string Opt1Header { get; set; }
        public string Opt2Header { get; set; }
        public string Opt3Header { get; set; }    
        public string SID { get; set; }
        public string SuffuxName { get; set; }

        public string Phone
        {
            get
            {
                if(Opt1Header.ToLower().Contains("phone"))
                { return Opt1; }
                if (Opt2Header.ToLower().Contains("phone"))
                { return Opt2; }
                if (Opt3Header.ToLower().Contains("phone"))
                { return Opt3; }

                return "";
            }
        }

        private string StrExtractedName = "";
        public string ExtractedName
        {
            get { return StrExtractedName; }
            set
            {
                StrExtractedName = value;

                string[] StudentNames = StrExtractedName.Split(' ');

                LastName = StudentNames[0];

                if (StudentNames.Length == 4)
                {
                    SuffuxName = StudentNames[1];
                    FirstName = StudentNames[2];
                    MiddleName = StudentNames[3];
                    Email = FirstName + "." + MiddleName + "." + LastName + "." + SuffuxName;
                }
                else if (StudentNames.Length == 3)
                {
                    FirstName = StudentNames[1];
                    MiddleName = StudentNames[2];
                    Email = FirstName + "." + MiddleName + "." + LastName;
                }
                else if (StudentNames.Length == 2)
                {
                    FirstName = StudentNames[1];
                    Email = FirstName + "." + LastName;
                }
            }
        }

    }
}