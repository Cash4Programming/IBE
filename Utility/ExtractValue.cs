using System;
using System.Collections.Generic;
using System.Text;

namespace InstructorBriefcaseExtractor.Utility
{
    // This class extracts a value for the javascrip from the returned web page HTML code
    public class ExtractValue
    {
        public String FromJavaScript(String HTTP, String StartLocation, String StartChar, String EndChar)
        {
            try
            {
                // Extract ticket
                //
                // 8/3/2018 Added .ToLower() to avoid case errors
                //
                int Start = HTTP.ToLower().IndexOf(StartLocation.ToLower());
                Start = HTTP.ToLower().IndexOf(StartChar.ToLower(), Start + 1) + StartChar.Length;   // don't include the StartChar
                int End = HTTP.ToLower().IndexOf(EndChar.ToLower(), Start + 1);

                return HTTP.Substring(Start, End - Start).Trim();
            }
            catch 
            {
                return "";
            }            
        }

    }
}
