using System;

using InstructorBriefcaseExtractor.Model;
using InstructorBriefcaseExtractor.Utility;

namespace InstructorBriefcaseExtractor.BLL
{
    class ClickerConfiguration
    {
        private readonly string xmlRootCLICKERLocation = "CLICKER";
        private readonly string ExportKey = "Export";
        private readonly string UnderscoreKey = "Underscore";
        private readonly string DirectoryKey = "Directory";
        private readonly string SelectedOutputkey = "SelectedOutput";

        private readonly UserSettings UserSettings;
        private readonly string DefaultDirectory = "";
        private readonly int SelectedOutputDefault = 1;

        public ClickerConfiguration(UserSettings UserSettings)
        {
            this.UserSettings = UserSettings;
            DefaultDirectory = UserSettings.MyDocuments + "Clicker\\";
        }

        public ClickerSettings Load()
        {
            ClickerSettings ClickerSettings = new ClickerSettings();
            XMLhelper XML = new XMLhelper(UserSettings);

            ClickerSettings.Directory = XML.XMLReadFile(xmlRootCLICKERLocation, DirectoryKey);
            if (ClickerSettings.Directory == "")
            {
                ClickerSettings.Directory = DefaultDirectory;
                XML.XMLWriteFile(xmlRootCLICKERLocation, DirectoryKey, ClickerSettings.Directory);
            }

            try
            {
                ClickerSettings.Export = Convert.ToBoolean(XML.XMLReadFile(xmlRootCLICKERLocation, ExportKey).ToLower());
            }
            catch {
                ClickerSettings.Export = false;
                XML.XMLWriteFile(xmlRootCLICKERLocation, ExportKey, ClickerSettings.Export.ToString());
            }

            try
            {
                ClickerSettings.Underscore = Convert.ToBoolean(XML.XMLReadFile(xmlRootCLICKERLocation, UnderscoreKey).ToLower());
            }
            catch {
                ClickerSettings.Underscore = true;
                XML.XMLWriteFile(xmlRootCLICKERLocation, UnderscoreKey, ClickerSettings.Underscore.ToString());
            }

            try
            {
                ClickerSettings.SelectedValue = Convert.ToInt32(XML.XMLReadFile(xmlRootCLICKERLocation, SelectedOutputkey));
            }
            catch
            {
                ClickerSettings.SelectedValue = SelectedOutputDefault;
                XML.XMLWriteFile(xmlRootCLICKERLocation, SelectedOutputkey, ClickerSettings.SelectedValue.ToString());
            }

            
            return ClickerSettings;
        }

        public void Save(ClickerSettings ClickerSettings)
        {

            XMLhelper XML = new XMLhelper(UserSettings);

            XML.XMLWriteFile(xmlRootCLICKERLocation, DirectoryKey, ClickerSettings.Directory);
            XML.XMLWriteFile(xmlRootCLICKERLocation, ExportKey, ClickerSettings.Export.ToString());
            XML.XMLWriteFile(xmlRootCLICKERLocation, UnderscoreKey, ClickerSettings.Underscore.ToString());
            XML.XMLWriteFile(xmlRootCLICKERLocation, SelectedOutputkey, ClickerSettings.SelectedValue.ToString());
        }
    }
}
