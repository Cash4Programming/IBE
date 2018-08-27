using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using InstructorBriefcaseExtractor.Model;
using InstructorBriefcaseExtractor.Utility;

namespace InstructorBriefcaseExtractor.BLL
{
    public class OutlookConfiguration
    {

        private readonly string xmlOUTLOOKNodeNameLocation = "OUTLOOK";
        private readonly string ExportKey = "Export";
        private readonly string UnderscoreKey = "Underscore";

        private readonly UserSettings UserSettings;
        
        public OutlookConfiguration(UserSettings UserSettings)
        {
            this.UserSettings = UserSettings;            
        }


        public OutlookSettings Load()
        {
            OutlookSettings OutlookSettings = new OutlookSettings();
            XMLhelper XML = new XMLhelper(UserSettings);

            
            try
            {
                OutlookSettings.Underscore = Convert.ToBoolean(XML.XMLReadFile(xmlOUTLOOKNodeNameLocation, UnderscoreKey).ToLower());
            }
            catch
            {
                OutlookSettings.Underscore = true;
                XML.XMLWriteFile(xmlOUTLOOKNodeNameLocation, UnderscoreKey, OutlookSettings.Underscore.ToString());
            }

            try
            {
                OutlookSettings.Export = Convert.ToBoolean(XML.XMLReadFile(xmlOUTLOOKNodeNameLocation, ExportKey).ToLower());
            }
            catch
            {
                OutlookSettings.Export = false;
                XML.XMLWriteFile(xmlOUTLOOKNodeNameLocation, ExportKey, OutlookSettings.Export.ToString());
            }

            return OutlookSettings;
        }

        public void Save(OutlookSettings OutlookSettings)
        {
            XMLhelper XML = new XMLhelper(UserSettings);

            XML.XMLWriteFile(xmlOUTLOOKNodeNameLocation, ExportKey, OutlookSettings.Export.ToString());
            XML.XMLWriteFile(xmlOUTLOOKNodeNameLocation, UnderscoreKey, OutlookSettings.Underscore.ToString());
        }
    }

}

