using System;
using System.Collections.Generic;
using System.Text;

using System.Net;         // Needed for RetrieveHTMLPage  - WebRequest and CredentialCache
using System.IO;          // Needed for RetrieveHTMLPage  - Stream and StreamReader

 //*********************************************************************
 // Copyright ©2008 Michael Paul Jenck, All Rights Reserved
 // http://www.jenck.net/
 //*********************************************************************
 //This program is free software; you can rediStribute it and/or
 //modify it under the terms of the GNU General Public License
 //as published by the Free Software Foundation; either version 3
 //of the License, or (at your Option) any later version.

 //This program is diStributed in the hope that it will be useful,
 //but WITHOUT ANY WARRANTY; without even the implied warranty of
 //MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 //GNU General Public License for more details.

 //You should have received a copy of the GNU General Public License
 //along with this program; if not, write to the
 //Free Software Foundation, Inc.,
 //59 Temple Place - Suite 330,
 //Boston, MA  02111-1307, USA.
 // *********************************************************************
 // Reference Works: NONE
 // *********************************************************************
 // Requires: Nothing
 // *********************************************************************
 // 
 // History
 // 
 // Created  2008-01-20 MPJ
 // Update   2018-01-26 MPJ - WebClient Openread method is failing
 // *********************************************************************
 // DESCRIPTION
 //    This class does a web request for the CIS Website
 // *********************************************************************

namespace InstructorBriefcaseExtractor.DAL
{
    // This is the class that request the web page from the remote server
    public class CISWebRequest
    {
        private CookieContainer myFormCookieJar;
        public CookieContainer CookieJar
        {
            get { return myFormCookieJar; }
            set { myFormCookieJar = value; }
        }

        public CISWebRequest()
        {
            myFormCookieJar = new CookieContainer();
        }

        private void BuildWebRequestStream(HttpWebRequest WebRequest, string PostData)
        {
            //==========================================================
            //
            // IN: WebRequest - the webrequest object
            //     PostData   - the data to be encoded 
            //
            // Returns: the webrequest with the encoded data
            //
            // DESCRIPTION
            //    This method builds the request Stream for WebRequest
            //
            //  Reference:
            //  http://groups-beta.google.com/group/microsoft.public.dotnet.languages.vb/browse_frm/thread/c22d840e7e90de3b/675e484e7594bb8a?lnk=st&q=WebClient+vb.net&rnum=3&hl=en#675e484e7594bb8a
            //==========================================================            
            Byte[] bytes = ASCIIEncoding.ASCII.GetBytes(PostData);
            WebRequest.ContentLength = bytes.Length;

            Stream StreamOut = WebRequest.GetRequestStream();
            StreamOut.Write(bytes, 0, bytes.Length);
            StreamOut.Close();
        }

        public string RetrieveHTMLPage(string URL)
        {
            //==========================================================
            //
            // IN: URL - the website address ( must be correctly formed!)
            //
            // Returns: the returned Stream - converted to a string 
            //
            // DESCRIPTION
            //    This function gets the requested URL as returns it in 
            //  the form of a String
            //
            //  Reference:
            //  http://groups-beta.google.com/group/microsoft.public.dotnet.languages.vb/browse_frm/thread/c22d840e7e90de3b/675e484e7594bb8a?lnk=st&q=WebClient+vb.net&rnum=3&hl=en#675e484e7594bb8a
            //==========================================================
            WebClient client = new WebClient();
            
            // Added 2018-01-27 
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            string HTMLPage = "";
            try
            {
                Stream Stream = client.OpenRead(URL);
                StreamReader reader = new StreamReader(Stream);
                HTMLPage = reader.ReadToEnd();
                reader.Close();
            }
            catch
            {
                HTMLPage = "";
                throw;
            }

            return HTMLPage;
        }

        public enum HTMLRequestType {POST, GET};

        public string RetrieveHTMLPage(string URL,
                                       string PostData,
                                       HTMLRequestType RequestType,
                                       int myTimeOut)
        {
            //==========================================================
            //
            // IN: URL         - the web address to test
            //     PostData    - the data to post to the URL
            //     RequestType - either POST or GET
            //     myTimeOut   - The number of milliseconds to wait before the request times out.
            //
            // Returns: the HTTPWebResponce
            //
            // DESCRIPTION
            //    This method post the data and then returns the responce
            //
            // REFERENCE
            //  http://groups.google.com/group/microsoft.public.dotnet.general/browse_frm/thread/e3c38240e4956e08/da19a7168a9b141d?lnk=st&q=HttpWebRequest.++vb.net+https&rnum=5&hl=en#da19a7168a9b141d
            //
            // If you go through a PROXY see this article
            // http://groups.google.com/group/microsoft.public.dotnet.framework.webservices/browse_frm/thread/c3a0c439943cac22/4da0ff45b5b3c095?lnk=st&q=parse+HTML+page+parameters+vb.net&rnum=2&hl=en#4da0ff45b5b3c095
            //
            //==========================================================           
            HttpWebRequest myWebRequest = (HttpWebRequest)WebRequest.Create(URL);            
            string RetrieveHTMLPage = "";

            if (myTimeOut < 1000)
            {
                myTimeOut = 1000; // must be at least 1 second
            }

            try
            {
                myWebRequest = (HttpWebRequest)WebRequest.Create(URL);
                myWebRequest.CookieContainer = myFormCookieJar;
                myWebRequest.Credentials = CredentialCache.DefaultCredentials; // MONO TODO
                myWebRequest.UserAgent = "BriefcaseExtractor_Ver2";
                myWebRequest.KeepAlive = true;
                myWebRequest.Headers.Set("Pragma", "no-cache");
                myWebRequest.Timeout = myTimeOut;

                if (RequestType == HTMLRequestType.POST)
                {
                    // write Stream to the web request
                    myWebRequest.Method = "POST";
                    myWebRequest.ContentType = "application/x-www-form-urlencoded";
                    myWebRequest.ContentLength = PostData.Length;
                    BuildWebRequestStream(myWebRequest, PostData);
                }
                else
                {
                    myWebRequest.Method = "GET";
                }

                // Retrieve responce
                HttpWebResponse myWebResponse = (HttpWebResponse)myWebRequest.GetResponse();                
                StreamReader sr = new StreamReader(myWebResponse.GetResponseStream());
                RetrieveHTMLPage = sr.ReadToEnd().ToString();
                sr.Close();
                myWebResponse.Close();
            }
            catch
            {
                myWebRequest = null;
                throw;  // rethrow
            }

            return RetrieveHTMLPage;
        }

    }
}
