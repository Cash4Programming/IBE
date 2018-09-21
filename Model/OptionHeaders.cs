using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InstructorBriefcaseExtractor.Model
{
    public class OptionHeaders
    {
        public readonly static string HeaderDefault = "Unknown";
        public OptionHeaders()
        {
            Header1 = HeaderDefault;
            Header2 = HeaderDefault;
            Header3 = HeaderDefault;
            WaitListHeader = HeaderDefault;
        }
        public string Header1 { get; set; }
        public string Header2 { get; set; }
        public string Header3 { get; set; }
        public string WaitListHeader { get; set; }
    }
}
