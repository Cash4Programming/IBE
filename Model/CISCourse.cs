

namespace InstructorBriefcaseExtractor.Model
{
    // This class generates the request string for a specific course
    // for a list of students in that course
    public class CISCourse
    {
        private readonly string StrRequest;
        private readonly string StrYRQVariable;
        private readonly string StrTicketVariable;
        private readonly string StrItemVariable;
        private readonly string StrItemNumber;
        private readonly string StrTicketValue;
        private readonly string StrYRQ;

        public CISCourse(string ItemNumber, string TicketValue, string YRQ)
        {
            StrYRQ = YRQ;
            StrItemNumber = ItemNumber;
            StrTicketValue = TicketValue;
            StrRequest = "request=roster";
            StrTicketVariable = "&tkt=";
            StrItemVariable = "&item=";
            StrYRQVariable = "&yrq=";
        }

        //"&request=roster" +
        //"&tkt=" + StrTicketValue +
        //"&yrq=" + StrYRQ +
        //"&item=" + StrItemNumber;

        public string HTTPRequestString
        {
            get
            {
                // Http request + Ticket + YRQ + Item
                return StrRequest + StrTicketVariable + StrTicketValue
                                  + StrYRQVariable + StrYRQ
                                  + StrItemVariable + StrItemNumber;
            }
        }                
    }
}
