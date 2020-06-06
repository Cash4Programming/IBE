using System;
using System.Security.Cryptography;

using InstructorBriefcaseExtractor.Model;

namespace InstructorBriefcaseExtractor.Utility
{
    // This is used to hide the username and pin when it is saved to disk
    public class Crypto
    {
        private readonly UserSettings MyUserSettings;
        public Crypto(UserSettings UserSettings)
        {
            MyUserSettings = UserSettings;
        }

        public string Encrypt(string Str_in)
        {            
            if (MyUserSettings.Version == "1")
            {
                return EncryptTripleDES(Str_in, "");
            }
            else
            {
                return Protect(Str_in);
            }
        }

        public string Decrypt(string Str_in)
        {
            if (MyUserSettings.Version == "1")
            {
                return DecryptTripleDES(Str_in, "");
            }
            else
            {
                return Unprotect(Str_in); 
            }
        }

        #region Version 2 - Using ProtectedData with a scope of CurrentUser

        private static byte[] GetBytes(string str)
        {
            byte[] bytes = new byte[str.Length * sizeof(char)];
            System.Buffer.BlockCopy(str.ToCharArray(), 0, bytes, 0, bytes.Length);
            return bytes;
        }

        private static string GetString(byte[] bytes)
        {
            char[] chars = new char[bytes.Length / sizeof(char)];
            System.Buffer.BlockCopy(bytes, 0, chars, 0, bytes.Length);
            return new string(chars);
        }

        private static string Protect(string Str_in)
        {

            if(Str_in=="") { return ""; }
            try
            {
                byte[] secret = GetBytes(Str_in);

                //System.Windows.Forms.MessageBox.Show("Protect - Before ProtectedData.Protect", "IBE-Info");
                // Encrypt the data using DataProtectionScope.CurrentUser. The result can be decrypted
                //  only by the same current user.
                byte[] returnval = ProtectedData.Protect(secret, null, DataProtectionScope.LocalMachine);
                
                // convert to base64 string
                return Convert.ToBase64String(returnval);
            }
            catch (Exception Ex)
            {
                string Error = Ex.Message + "\r\n" + Ex.InnerException + "\r\n" + Ex.Data + "\r\n\r\n" + Ex.Source;            
                System.Windows.Forms.MessageBox.Show(Error, "IBE-Error");
                throw;
            }
        }

        public static string Unprotect(string Encrypted_Str_in)
        {
            if (Encrypted_Str_in == "") { return ""; }
            try
            {
                byte[] secret = Convert.FromBase64String(Encrypted_Str_in);
                //Decrypt the data using DataProtectionScope.CurrentUser.
                byte[] returnval = ProtectedData.Unprotect(secret, null, DataProtectionScope.LocalMachine);                
                return GetString(returnval);
            }
            catch (CryptographicException)
            {
                return "";
            }
            catch (Exception Ex)
            {
                string Error = Ex.Message + "\r\n" + Ex.InnerException + "\r\n" + Ex.Data + "\r\n" + Ex.Source;
                System.Windows.Forms.MessageBox.Show(Error, "IBE-Error");
                throw;
            }
        }

        #endregion

        // ---------------------------------------------------------------------------------
        // ---------------------------------------------------------------------------------
        // ---------------------------------------------------------------------------------
        // ---------------------------------------------------------------------------------

        #region Version 1 - Retired
        // This is really bad - but for now it works.
        // this needs to be broken down so these key would be much harder to find.
        // 01/13/2008 MPJ
        // Generated from https://www.grc.com/passwords.htm
        private readonly static String _GenericKey1 = "TW3vQuKmB6rZflRmevZFJvbEOvNF877IiCUq142B7YTLwcEKXNUSEOlPkBN4qe2";

        private String GenericKey()
        {
            System.Security.Principal.WindowsIdentity UserName = System.Security.Principal.WindowsIdentity.GetCurrent();
            int l = UserName.Name.Length;
            int GKL = _GenericKey1.Length;
            return _GenericKey1.Substring(0, l) + UserName.Name + _GenericKey1.Substring(l, GKL - l);
        }

        private String VerifyKey(String sKey)
        {
            if (sKey == "")
            {
                return GenericKey();
            }
            return sKey;
        }

        private string EncryptTripleDES(String sIn, String sKey)
        {
            TripleDESCryptoServiceProvider DES = new TripleDESCryptoServiceProvider();
            MD5CryptoServiceProvider hashMD5 = new MD5CryptoServiceProvider();

            if (sIn == "") { return ""; }
            sKey = VerifyKey(sKey);
            DES.Key = hashMD5.ComputeHash(System.Text.ASCIIEncoding.ASCII.GetBytes(sKey));
            DES.Mode = CipherMode.ECB;
            ICryptoTransform DESEncrypt = DES.CreateEncryptor();
            Byte[] Buffer = System.Text.ASCIIEncoding.ASCII.GetBytes(sIn);

            return Convert.ToBase64String(DESEncrypt.TransformFinalBlock(Buffer, 0, Buffer.Length));
        }

        private string DecryptTripleDES(String sOut, String sKey)
        {
            TripleDESCryptoServiceProvider DES = new TripleDESCryptoServiceProvider();
            MD5CryptoServiceProvider hashMD5 = new MD5CryptoServiceProvider();


            if (sOut == "") { return ""; }
            sKey = VerifyKey(sKey);
            DES.Key = hashMD5.ComputeHash(System.Text.ASCIIEncoding.ASCII.GetBytes(sKey));
            DES.Mode = CipherMode.ECB;
            ICryptoTransform DESDecrypt = DES.CreateDecryptor();
            Byte[] Buffer = Convert.FromBase64String(sOut);
            return System.Text.ASCIIEncoding.ASCII.GetString(DESDecrypt.TransformFinalBlock(Buffer, 0, Buffer.Length));
        }
        #endregion
    }
}