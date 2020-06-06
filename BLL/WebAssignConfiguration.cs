using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using InstructorBriefcaseExtractor.Model;
using InstructorBriefcaseExtractor.Utility;

namespace InstructorBriefcaseExtractor.BLL
{
    public class WebAssignConfiguration
    {
        private readonly string xmlWEBASSIGNNodeNameLocation = "WEBASSIGN";
        private readonly string ExportKey = "Export";
        private readonly string UnderscoreKey = "Underscore";
        private readonly string PasswordKey = "Password";
        private readonly string UsernameKey = "Username";
        private readonly string SEPARATORKey = "SEPERATOR";
        private readonly string DirectoryKey = "Directory";
        private readonly string InstitutionKey = "Institution";

        private readonly UserSettings UserSettings;
        private readonly string DefaultDirectory = "";

        public WebAssignConfiguration(UserSettings UserSettings)
        {
            this.UserSettings = UserSettings;
            DefaultDirectory = UserSettings.MyDocuments + "WebAssign\\";
        }

        public WebAssignSettings Load()
        {
            WebAssignSettings WebAssignSettings = new WebAssignSettings();
            string Return;

            XMLhelper XML = new XMLhelper(UserSettings);

            WebAssignSettings.Directory = XML.XMLReadFile(xmlWEBASSIGNNodeNameLocation, DirectoryKey);
            if (WebAssignSettings.Directory == "")
            {
                WebAssignSettings.Directory = DefaultDirectory;
                XML.XMLWriteFile(xmlWEBASSIGNNodeNameLocation, DirectoryKey, WebAssignSettings.Directory);
            }

            try
            {
                WebAssignSettings.Export = Convert.ToBoolean(XML.XMLReadFile(xmlWEBASSIGNNodeNameLocation, ExportKey).ToLower());
            }
            catch
            {
                WebAssignSettings.Export = false;
                XML.XMLWriteFile(xmlWEBASSIGNNodeNameLocation, ExportKey, WebAssignSettings.Export.ToString(CultureInfo.CurrentCulture));
            }

            Return = XML.XMLReadFile(xmlWEBASSIGNNodeNameLocation, PasswordKey);
            if(Return == "")
            {
                Return = "LAST_NAME";
                XML.XMLWriteFile(xmlWEBASSIGNNodeNameLocation, PasswordKey, Return);
            }
            if (Return == "LAST_NAME")
            {
                WebAssignSettings.UserPassword = WebAssignPassword.LastName;
            }
            else
            {
                WebAssignSettings.UserPassword = WebAssignPassword.SID;
            }

            Return = XML.XMLReadFile(xmlWEBASSIGNNodeNameLocation, SEPARATORKey);
            if (Return == "")
            {
                Return = "TAB";
                XML.XMLWriteFile(xmlWEBASSIGNNodeNameLocation, SEPARATORKey, Return);
            }

            if (Return == "TAB")
            {
                WebAssignSettings.SEPARATOR = WebAssignSeperator.Tab;
            }
            else
            {
                WebAssignSettings.SEPARATOR = WebAssignSeperator.Comma;
            }

            try
            {
                WebAssignSettings.Underscore = Convert.ToBoolean(XML.XMLReadFile(xmlWEBASSIGNNodeNameLocation, UnderscoreKey).ToLower());
            }
            catch
            {
                WebAssignSettings.Underscore = true;
                XML.XMLWriteFile(xmlWEBASSIGNNodeNameLocation, UnderscoreKey, WebAssignSettings.Underscore.ToString(CultureInfo.CurrentCulture));
            }
            WebAssignSettings.Institution = XML.XMLReadFile(xmlWEBASSIGNNodeNameLocation, InstitutionKey);

            Return = XML.XMLReadFile(xmlWEBASSIGNNodeNameLocation, UsernameKey);
            if (Return == "")
            {
                Return = "First_Name_Initial_plus_Last_Name";
                XML.XMLWriteFile(xmlWEBASSIGNNodeNameLocation, UsernameKey, Return);
            }


            if (Return == "First_Name_Initial_plus_Last_Name")
            {
                WebAssignSettings.Username = WebAssignUserName.First_Name_Initial_plus_Last_Name;
            }
            else if (Return == "First_Name_plus_Last_Name")
            {
                WebAssignSettings.Username = WebAssignUserName.First_Name_plus_Last_Name;
            }
            else
            {
                WebAssignSettings.Username = WebAssignUserName.First_Name_Initial_plus_Middle_Initial_plus_Last_Name;
            }

            return WebAssignSettings;
        }

        public void Save(WebAssignSettings WebAssignSettings)
        {
            string Temp;
            XMLhelper XML = new XMLhelper(UserSettings);

            XML.XMLWriteFile(xmlWEBASSIGNNodeNameLocation, DirectoryKey, WebAssignSettings.Directory);
            XML.XMLWriteFile(xmlWEBASSIGNNodeNameLocation, ExportKey, WebAssignSettings.Export.ToString(CultureInfo.CurrentCulture));
            

            if (WebAssignSettings.UserPassword== WebAssignPassword.LastName)
            { Temp = "LAST_NAME"; }
            else
            { Temp = "SID"; }
            XML.XMLWriteFile(xmlWEBASSIGNNodeNameLocation, PasswordKey, Temp);


            if (WebAssignSettings.SEPARATOR == WebAssignSeperator.Tab)
            { Temp = "TAB"; }
            else
            { Temp = "COMMA"; }
            XML.XMLWriteFile(xmlWEBASSIGNNodeNameLocation, SEPARATORKey, Temp);


            if (WebAssignSettings.Username == WebAssignUserName.First_Name_Initial_plus_Last_Name)
            { Temp = "First_Name_Initial_plus_Last_Name"; }
            else if (WebAssignSettings.Username == WebAssignUserName.First_Name_plus_Last_Name)
            { Temp = "First_Name_plus_Last_Name"; }
            else
            { Temp = "First_Name_Initial_plus_Middle_Initial_plus_Last_Name"; }

            XML.XMLWriteFile(xmlWEBASSIGNNodeNameLocation, UsernameKey, Temp);
            XML.XMLWriteFile(xmlWEBASSIGNNodeNameLocation, UnderscoreKey, WebAssignSettings.Underscore.ToString(CultureInfo.CurrentCulture));
            XML.XMLWriteFile(xmlWEBASSIGNNodeNameLocation, InstitutionKey, WebAssignSettings.Institution);
        }
    }
}
