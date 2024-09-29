using System;
//using System.Collections.Generic;
//using System.Linq;
using System.Security.Cryptography;
using System.Text;
//using System.Web;

namespace EWS.Includes
{
    public class CoverRSAsme
    {
        public CoverRSAsme()
        { }
        public static string EncryptData(string data, string publicKey)
        {
            data = Compression.Zip(data);
            using (RSACryptoServiceProvider rsa = new RSACryptoServiceProvider())
            {
                rsa.FromXmlString(publicKey);
                byte[] dataBytes = Encoding.UTF8.GetBytes(data.ToString());
                byte[] encryptedData = rsa.Encrypt(dataBytes, false);
                return Convert.ToBase64String(encryptedData);
            }
        }
        public static string DecryptData(string data, string privateKey)
        {
            byte[] encryptedData = Convert.FromBase64String(data);
            using (RSACryptoServiceProvider rsa = new RSACryptoServiceProvider())
            {
                rsa.FromXmlString(privateKey);
                byte[] decryptedData = rsa.Decrypt(encryptedData, false);
                string decryptedString = Encoding.UTF8.GetString(decryptedData);
                decryptedString = Compression.Unzip(decryptedString);
                return decryptedString;
            }
        }
        ~CoverRSAsme()
        {
            GC.Collect();
        }
    }
}