using System;
using System.IO;
using System.Text.RegularExpressions;

namespace EWS.Includes
{
    public static class Validation
    {
        public static bool CredentialsError(HandleAppLog hf, String path, String accessSpecifier, String watchWord)
        {
            try
            {
                if ((accessSpecifier != null && watchWord != null))
                {
                    if (accessSpecifier.Length > 0 && watchWord.Length > 0)
                    {
                        if (System.Configuration.ConfigurationManager.AppSettings["accessSpecifier"].Trim() == accessSpecifier && System.Configuration.ConfigurationManager.AppSettings["watchWord"].Trim() == watchWord)
                        {
                            return false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (ex.ToString().Length > 0)
                {
                    path += "Exception\\";
                    hf.FileSystemLog(path, "Validatation", "Exception", ex.ToString());
                    return true;
                }
            }
            return true;
        }
        public static bool EmailError(string mail, bool encrypt)
        {
            if (string.IsNullOrEmpty(mail))
                return true;
            if (encrypt)
                mail = CoverAES.DecryptStringAES(mail);
            mail = mail.Trim().ToUpper();
            string pattern = @"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$";
            if (!Regex.IsMatch(mail, pattern))
                return true;
            return false;
        }
        public static bool FormatError(string str)
        {
            // Check if the string is not null or empty
            if (string.IsNullOrEmpty(str))
            {
                return true;
            }
            // Check if the string meets a minimum length 3 requirement
            if (str.Length < 3)
            {
                return true;
            }
            return false;
        }
        public static bool AttachmentError(GetItem item)
        {
            if (string.IsNullOrEmpty(item.FileName))
                return true;
            if (item.FileContent == null)
                return true;
            if (item.FileContent.Length < 1)
                return true;
            return false;
        }
    }
}