

namespace InstructorBriefcaseExtractor.Model
{
    // This class generates the request string for the login page to retreive a 
    // list of courses taught
    public class CISLogin
    {
        private readonly string ID_fieldName = "empid";
        private readonly string PIN_fieldName = "pin";
        private readonly string Request_fieldName = "request";
        private readonly string YRQ_fieldName = "yrq";
        private readonly string RequestType_fieldName = "ipalogon";

        private readonly string StrRequest = "";
        private readonly string StrYRQVariable = "";
        private readonly string StrUserName = "";
        private readonly string StrPasswordFieldName = "";        

        public string RequestURL { get; set; } 
        public string PostURL { get; set; }


        public CISLogin(UserSettings UserSettings, Quarter SelectedQuarter)
        {
            this.Quarter = SelectedQuarter;
            this.EmployeeID = UserSettings.EmployeeID;
            this.EmployeePIN = UserSettings.EmployeePIN;

            StrRequest = Request_fieldName + "=" + RequestType_fieldName;
            StrUserName = "&" + ID_fieldName + "=";
            StrPasswordFieldName = "&" + PIN_fieldName + "=";
            StrYRQVariable = "&" + YRQ_fieldName + "=";
            
        }

        public string EmployeeID { get; set; }
        public string EmployeePIN { get; set; }
        public Quarter Quarter { get; set; }
        
        //"&request=ipalogon" +
        //"&empid=" + StrEmployeeID +
        //"&pin=" + StrEmployeePIN +
        //"&yrq=" + StrYRQ +

        public string HTTPCoursesRequestString
        {
            get 
            {
                // Http request + Username + Password + YRQ
                return StrRequest + StrUserName + this.EmployeeID
                                  + StrPasswordFieldName + this.EmployeePIN
                                  + StrYRQVariable + Quarter.YRQ; 
            }
        }


    }
}
