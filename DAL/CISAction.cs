using System;

using InstructorBriefcaseExtractor.Model;
using InstructorBriefcaseExtractor.Utility;

namespace InstructorBriefcaseExtractor.DAL
{
    class CISAction
    {

        #region Events

        public event InformationalMessage MessageResults;
        protected virtual void SendMessage(object sender, EventArgs e)
        {
            MessageResults?.Invoke(sender, e);
        }

        #endregion

        private readonly string StrAction;

        public CISAction()
        {
            StrAction = "Action=";
        }


        // This extraction the Action URL form the login webpage.
        // it is used for further requests.
        // The URL is from the College object
        public string PostURL(string URL)
        {
            SendMessage(this, new Information("Webpage Location = " + URL));

            CISWebRequest myWebRequest = new CISWebRequest();
            string result = myWebRequest.RetrieveHTMLPage(URL);  // get the page 

            // Extract From result the POST Action URL "https://www.ctc.edu/cgi-bin/rq160/"
            // and return it to the caller.

            ExtractValue EV = new ExtractValue();
            return EV.FromJavaScript(result, StrAction, "\"", "\"");            
        }
    }
}
