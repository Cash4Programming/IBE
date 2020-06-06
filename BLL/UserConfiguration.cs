using System;
using System.Globalization;
using System.Security.Cryptography;

using InstructorBriefcaseExtractor.Model;
using InstructorBriefcaseExtractor.Utility;

namespace InstructorBriefcaseExtractor.BLL
{
    class UserConfiguration
    {
        private readonly string xmlUserNodeNameLocation = "";
        private readonly string EmployeeIDKey = "EmployeeID";
        private readonly string EmployeePINKey = "EmployeePIN";
        private readonly string EmailKey = "Email";
        private readonly string OverWriteAllKey = "OverWrite";
        private readonly string CollegeAbbreviationKey = "CollegeAbbreviation";


        private readonly UserSettings UserSettings;
        
        public UserConfiguration(UserSettings UserSettings)
        {
            this.UserSettings = UserSettings;
            xmlUserNodeNameLocation = UserSettings.UserName;
        }

        public UserSettings Load()
        {
            XMLhelper XML = new XMLhelper(UserSettings);
            Crypto cry = new Crypto(UserSettings);

            UserSettings.CollegeAbbreviation = XML.XMLReadFile(xmlUserNodeNameLocation, CollegeAbbreviationKey);

            string StrEncoded = XML.XMLReadFile(xmlUserNodeNameLocation, EmployeeIDKey);
            UserSettings.EmployeeID = cry.Decrypt(StrEncoded);
            StrEncoded = XML.XMLReadFile(xmlUserNodeNameLocation, EmployeePINKey);
            UserSettings.EmployeePIN = cry.Decrypt(StrEncoded);

            UserSettings.Email = XML.XMLReadFile(xmlUserNodeNameLocation, EmailKey);
    
            try
            {
                UserSettings.OverWriteAll = Convert.ToBoolean(XML.XMLReadFile(xmlUserNodeNameLocation, OverWriteAllKey));
            }
            catch
            {
                UserSettings.OverWriteAll = false;
                XML.XMLWriteFile(xmlUserNodeNameLocation, OverWriteAllKey, UserSettings.OverWriteAll.ToString(CultureInfo.CurrentCulture));                
            }

            return UserSettings;
        }

        public void Save()
        {
            Crypto cry = new Crypto(UserSettings);

            XMLhelper XML = new XMLhelper(UserSettings);

            XML.XMLWriteFile(xmlUserNodeNameLocation, CollegeAbbreviationKey, UserSettings.CollegeAbbreviation);
            if (UserSettings.EmployeeID == "")
            {
                XML.XMLWriteFile(xmlUserNodeNameLocation, EmployeeIDKey, "");
            }
            else
            {
                XML.XMLWriteFile(xmlUserNodeNameLocation, EmployeeIDKey, cry.Encrypt(UserSettings.EmployeeID));
            }

            if (UserSettings.EmployeePIN == "")
            {
                XML.XMLWriteFile(xmlUserNodeNameLocation, EmployeePINKey, "");
            }
            else
            { 
                XML.XMLWriteFile(xmlUserNodeNameLocation, EmployeePINKey, cry.Encrypt(UserSettings.EmployeePIN));
            }

            XML.XMLWriteFile(xmlUserNodeNameLocation, EmailKey, UserSettings.Email);
            XML.XMLWriteFile(xmlUserNodeNameLocation, OverWriteAllKey, UserSettings.OverWriteAll.ToString(CultureInfo.CurrentCulture));
        }
    }
}
