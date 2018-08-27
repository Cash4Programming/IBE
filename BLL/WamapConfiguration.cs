using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using InstructorBriefcaseExtractor.Model;
using InstructorBriefcaseExtractor.Utility;

namespace InstructorBriefcaseExtractor.BLL
{
    class WamapConfiguration
    {
        private readonly string xmlWAMAPNodeNameLocation = "WAMAP";

        private readonly string ExportKey = "Export";
        private readonly string UnderscoreKey = "Underscore";
        private readonly string DirectoryKey = "Directory";


        private readonly UserSettings UserSettings;
        private readonly string DefaultDirectory = "";

        public WamapConfiguration(UserSettings UserSettings)
        {
            this.UserSettings = UserSettings;
            DefaultDirectory = UserSettings.MyDocuments + "WAMAP\\";
        }


        public WamapSettings Load()
        {
            WamapSettings WamapSettings = new WamapSettings();
            XMLhelper XML = new XMLhelper(UserSettings);

            WamapSettings.Directory = XML.XMLReadFile(xmlWAMAPNodeNameLocation, DirectoryKey);
            if (WamapSettings.Directory == "")
            {
                WamapSettings.Directory = DefaultDirectory;
                XML.XMLWriteFile(xmlWAMAPNodeNameLocation, DirectoryKey, WamapSettings.Directory);
            }

            try
            {
                WamapSettings.Underscore = Convert.ToBoolean(XML.XMLReadFile(xmlWAMAPNodeNameLocation, UnderscoreKey).ToLower());
            }
            catch
            {
                WamapSettings.Underscore = true;
                XML.XMLWriteFile(xmlWAMAPNodeNameLocation, UnderscoreKey, WamapSettings.Underscore.ToString());
            }

            try
            {
                WamapSettings.Export = Convert.ToBoolean(XML.XMLReadFile(xmlWAMAPNodeNameLocation, ExportKey).ToLower());
            }
            catch
            {
                WamapSettings.Export = false;
                XML.XMLWriteFile(xmlWAMAPNodeNameLocation, ExportKey, WamapSettings.Export.ToString());
            }

            return WamapSettings;
        }

        public void Save(WamapSettings WamapSettings)
        {
            XMLhelper XML = new XMLhelper(UserSettings);

            XML.XMLWriteFile(xmlWAMAPNodeNameLocation, DirectoryKey, WamapSettings.Directory);
            XML.XMLWriteFile(xmlWAMAPNodeNameLocation, ExportKey, WamapSettings.Export.ToString());
            XML.XMLWriteFile(xmlWAMAPNodeNameLocation, UnderscoreKey, WamapSettings.Underscore.ToString());
        }
    }

}
