using System.IO;
using System.Xml;
using InstructorBriefcaseExtractor.Model;

namespace InstructorBriefcaseExtractor.Utility
{
    public class XMLhelper
    {
        private readonly string Configuration = "configuration";
        private readonly string xmlUserNameLocation = "";
        private readonly string xmlRootUserLocation = "";
        private readonly string PathandFileName = "";
        private readonly string UserName = "";
        
        public static string VersionKey = "Version";
        public static string VersionNumber = "2";

        public XMLhelper(UserSettings UserSettings)
        {
            xmlUserNameLocation = UserSettings.UserName;
            xmlRootUserLocation = "/"+Configuration + "/" + UserSettings.UserName;
            this.PathandFileName = UserSettings.PathandFileName;            
            this.UserName = UserSettings.UserName;

            if (!File.Exists(PathandFileName))
            {
                XMLWriteFile(UserName, VersionKey, VersionNumber);
            }
        }

        public string XMLReadFile(string NodeLocation, string NodeName)
        {
            string XmlRoot = "";
            if (NodeLocation == xmlUserNameLocation)
            {
                XmlRoot = xmlRootUserLocation;
            }
            else
            {
                XmlRoot = xmlRootUserLocation + "/" + NodeLocation;
            }

            try
            {
                XmlDocument m_xmld;
                XmlNodeList m_nodelist;
                string StrTemp = "";

                //# Create the XML Document
                m_xmld = new XmlDocument();

                //# Load the Xml file
                m_xmld.Load(PathandFileName);

                //# Get the list of name nodes
                m_nodelist = m_xmld.SelectNodes(XmlRoot);

                //# Loop through the nodes
                foreach (XmlNode m_node in m_nodelist)
                {
                    if (m_node.Name == NodeLocation)
                    {
                        //Look for the correct values
                        for (int intloop = 0; intloop < m_node.ChildNodes.Count; intloop++)
                        {
                            if ((m_node.ChildNodes.Item(intloop).Attributes != null))
                            {

                                if (m_node.ChildNodes.Item(intloop).Attributes.GetNamedItem("key") != null)
                                {
                                    //Load the key to see if you have the correct one
                                    StrTemp = m_node.ChildNodes.Item(intloop).Attributes.GetNamedItem("key").Value;
                                    if (StrTemp == NodeName)
                                    {
                                        // return the value
                                        return m_node.ChildNodes.Item(intloop).Attributes.GetNamedItem("value").Value;
                                    }
                                }
                            }
                        }
                        
                    }
                }
                // node was not found - create it
                XMLWriteFile(NodeLocation, NodeName, "");
                return "";
            }
            catch
            {
                throw;
            }
        }

        public void XMLWriteFile(string NodeLocation, string NodeName, string NodeValue)
        {
            string XmlRoot = "";
            if (NodeLocation == xmlUserNameLocation)
            {
                XmlRoot = xmlRootUserLocation;
            }
            else
            {
                XmlRoot = xmlRootUserLocation + "/" + NodeLocation;
            }

            XmlDocument m_xmld = new XmlDocument();
            try
            {

                XmlNodeList m_nodelist;
                string StrTemp = "";
                bool BlnFoundAttribute = false;

                // Create the Attribute
                XmlAttribute aValue = m_xmld.CreateAttribute("value");
                aValue.Value = NodeValue;

                // Does the file Exist?
                if (!File.Exists(PathandFileName)) {
                    XmlDocument xmlCreate = new XmlDocument();
                    // Create the Root
                    XmlElement Root = xmlCreate.CreateElement(Configuration);
                    XmlNode AppNode  = xmlCreate.AppendChild(Root);
                    xmlCreate.AppendChild(Root);

                    // Create the username
                    XmlElement  MyName = xmlCreate.CreateElement(NodeLocation);
                    AppNode.AppendChild(MyName);

                    // Save to disk
                    xmlCreate.Save(PathandFileName);
                }
                
                //# Load the Xml file
                m_xmld.Load(PathandFileName);
                
                //# Get the list of name nodes
                m_nodelist = m_xmld.SelectNodes(XmlRoot);

                if(m_nodelist.Count == 0)
                {
                    // need to create the Element
                    XmlElement ElementName = m_xmld.CreateElement(NodeLocation);

                    XmlNodeList DocumentNode = m_xmld.SelectNodes(xmlRootUserLocation);
                    foreach (XmlNode MyName in DocumentNode)
                    {
                        // There should be only on - I am using a loop to make sure the 
                        // If something has been added - this should write the data in the correct section
                        if (MyName.Name == UserName)
                        {
                            MyName.AppendChild(ElementName);
                            m_xmld.Save(PathandFileName);
                            break;
                        }
                    }


                    // now select the Created Element
                    m_nodelist = m_xmld.SelectNodes(XmlRoot);
                }


                //# Loop through the nodes
                foreach (XmlNode m_node in m_nodelist)
                {
                    if (m_node.Name == NodeLocation)
                    {
                        //Look for the correct values
                        for (int intloop = 0; intloop < m_node.ChildNodes.Count; intloop++)
                        {
                            if ((m_node.ChildNodes.Item(intloop).Attributes != null))
                            {
                                //Load the key to see if you have the correct one
                                if (m_node.ChildNodes.Item(intloop).Attributes.GetNamedItem("key") != null)
                                {
                                    StrTemp = m_node.ChildNodes.Item(intloop).Attributes.GetNamedItem("key").Value;
                                    if (StrTemp == NodeName)
                                    {
                                        // return the value
                                        m_node.ChildNodes.Item(intloop).Attributes.SetNamedItem(aValue);
                                        BlnFoundAttribute = true;
                                        break;  // intloop exit
                                    }
                                }
                            }
                        }
                        if (!BlnFoundAttribute)
                        {
                            // create entry
                            //Reference
                            // http://groups.google.com/group/microsoft.public.xml/browse_frm/thread/d722490293c78e6/99e2eacfe9cdbd4a?lnk=st+q=vb.net+XmlNode+attribute+add+rnum=5+hl=en#99e2eacfe9cdbd4a
                            //
                            XmlNode nn = m_xmld.CreateElement("add");
                            XmlAttribute aKey = m_xmld.CreateAttribute("key");

                            // Set values
                            aKey.Value = NodeName;

                            // Append attributes to new node
                            nn.Attributes.Append(aKey);
                            nn.Attributes.Append(aValue);

                            // Add node to current node
                            m_node.AppendChild(nn);
                        }
                    }

                }

                // Write the result to disc
                m_xmld.Save(PathandFileName);
            }
            catch
            {
                throw;
            }
            finally
            {
                if (m_xmld != null) { m_xmld = null; }
            }
        }


    }
}
