using System;
using InstructorBriefcaseExtractor.Model;
using InstructorBriefcaseExtractor.Utility;

namespace InstructorBriefcaseExtractor.BLL
{
    public class XMLFileVersionConfiguration
    {
        private UserSettings UserSettings;

        public XMLFileVersionConfiguration(UserSettings UserSettings)
        {
            this.UserSettings = UserSettings;
        }

        public string CheckVersion()
        {
            XMLhelper XML = new XMLhelper(UserSettings);

           return XML.XMLReadFile(UserSettings.UserName, XMLhelper.VersionKey);           
        }
    }
}
