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
    class MTGConfiguration
    {
        private readonly string xmlMTGNodeNameLocation = "MTG";
        private readonly string ExportKey = "Export";
        private readonly string UnderscoreKey = "Underscore";
        private readonly string DirectoryKey = "Directory";

        private readonly UserSettings UserSettings = null;
        private readonly string DefaultDirectory = "";

        public MTGConfiguration(UserSettings UserSettings)
        {
            this.UserSettings = UserSettings;
            DefaultDirectory = UserSettings.MyDocuments + "MTG\\";
        }

        public MTGSettings Load()
        {
            MTGSettings MTGSettings = new MTGSettings();
            XMLhelper XML = new XMLhelper(UserSettings);

            MTGSettings.Directory = XML.XMLReadFile(xmlMTGNodeNameLocation, DirectoryKey);
            if (MTGSettings.Directory == "")
            {
                MTGSettings.Directory = DefaultDirectory;
                XML.XMLWriteFile(xmlMTGNodeNameLocation, DirectoryKey, MTGSettings.Directory);
            }
            
            try
            {
                MTGSettings.Export = Convert.ToBoolean(XML.XMLReadFile(xmlMTGNodeNameLocation, ExportKey).ToLower());
            }
            catch
            {
                MTGSettings.Export = false;
                XML.XMLWriteFile(xmlMTGNodeNameLocation, ExportKey, MTGSettings.Export.ToString(CultureInfo.CurrentCulture));
            }

            try
            {
                MTGSettings.Underscore = Convert.ToBoolean(XML.XMLReadFile(xmlMTGNodeNameLocation, UnderscoreKey).ToLower());
            }
            catch
            {
                MTGSettings.Underscore = true;
                XML.XMLWriteFile(xmlMTGNodeNameLocation, UnderscoreKey, MTGSettings.Underscore.ToString(CultureInfo.CurrentCulture));
            }

            return MTGSettings;
        }

        public void Save(MTGSettings MTGSettings)
        {
            XMLhelper XML = new XMLhelper(UserSettings);

            XML.XMLWriteFile(xmlMTGNodeNameLocation, DirectoryKey, MTGSettings.Directory);
            XML.XMLWriteFile(xmlMTGNodeNameLocation, ExportKey, MTGSettings.Export.ToString(CultureInfo.CurrentCulture));
            XML.XMLWriteFile(xmlMTGNodeNameLocation, UnderscoreKey, MTGSettings.Underscore.ToString(CultureInfo.CurrentCulture));
        }

    }
}
