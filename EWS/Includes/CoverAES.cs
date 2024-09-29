using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using static EWS.Includes.Enumeration;
namespace EWS.Includes
{
    public class CoverAES
    {
        public static string KeyString
        {
            get; set;
        }
        public static string IVString
        {
            get; set;
        }
        private static byte[] Key
        {
            get; set;
        }
        private static byte[] IV
        {
            get; set;
        }
        public CoverAES()
        { }
        public static string EncryptStringAES(string plainText, bool bitByte = false)
        {
            Key = Encoding.UTF8.GetBytes(KeyString);
            IV = Encoding.UTF8.GetBytes(IVString);
            if (plainText == null || plainText.Length <= 0)
                return string.Empty;
            if (Key == null || Key.Length <= 0)
                throw new ArgumentNullException("Key");
            if (IV == null || IV.Length <= 0)
                throw new ArgumentNullException("IV");
            byte[] encrypted;
            using (Aes aesAlg = Aes.Create())
            {
                aesAlg.Key = Key;
                aesAlg.IV = IV;                
                ICryptoTransform encryptor = aesAlg.CreateEncryptor(aesAlg.Key, aesAlg.IV);
                using (MemoryStream msEncrypt = new MemoryStream())
                {
                    using (CryptoStream csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
                    {
                        using (StreamWriter swEncrypt = new StreamWriter(csEncrypt))
                        {
                            swEncrypt.Write(plainText);
                        }
                        encrypted = msEncrypt.ToArray();
                    }
                }
            }
            if (bitByte)
                return System.Text.Encoding.Default.GetString(encrypted);
            else
                return Convert.ToBase64String(encrypted);
        }
        public static string DecryptStringAES(string hiddenText)
        {
            try
            {
                byte[] str = Convert.FromBase64String(hiddenText);
                if (str == null || str.Length <= 0)
                    throw new ArgumentNullException("hiddenText");
                Key = Encoding.UTF8.GetBytes(KeyString);
                IV = Encoding.UTF8.GetBytes(IVString);
                if (Key == null || Key.Length <= 0) throw new ArgumentNullException("Key");
                if (IV == null || IV.Length <= 0) throw new ArgumentNullException("IV");
                string plaintext = null; using (Aes aesAlg = Aes.Create())
                {
                    aesAlg.Key = Key;
                    aesAlg.IV = IV;                    
                    ICryptoTransform decryptor = aesAlg.CreateDecryptor(aesAlg.Key, aesAlg.IV);
                    using (MemoryStream msDecrypt = new MemoryStream(str))
                    {
                        using (CryptoStream csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read))
                        {
                            using (StreamReader srDecrypt = new StreamReader(csDecrypt))
                            {
                                plaintext = srDecrypt.ReadToEnd();
                            }
                        }
                    }
                }
                return plaintext;
            }
            catch (Exception ex)
            {
                if (ex.ToString().Length > 0)
                    return string.Empty;
            }
            return string.Empty;
        }
        public static byte[] EncryptByteArrayAES(byte[] plainBytes)
        {
            Key = Encoding.UTF8.GetBytes(KeyString);
            IV = Encoding.UTF8.GetBytes(IVString);
            if (plainBytes == null || plainBytes.Length <= 0) return null;
            if (Key == null || Key.Length <= 0) throw new ArgumentNullException("Key");
            if (IV == null || IV.Length <= 0) throw new ArgumentNullException("IV");
            using (Aes aesAlg = Aes.Create())
            {
                aesAlg.Key = Key;
                aesAlg.IV = IV;
                ICryptoTransform encryptor = aesAlg.CreateEncryptor(aesAlg.Key, aesAlg.IV);
                using (MemoryStream msEncrypt = new MemoryStream())
                {
                    using (CryptoStream csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
                    {
                        csEncrypt.Write(plainBytes, 0, plainBytes.Length);
                        csEncrypt.FlushFinalBlock();
                    }
                    return msEncrypt.ToArray();
                }
            }
        }
        public static byte[] DecryptByteArrayAES(byte[] hiddenBytes)
        {
            if (hiddenBytes == null || hiddenBytes.Length <= 0)
                throw new ArgumentNullException("EncryptedBytes");
            Key = Encoding.UTF8.GetBytes(KeyString);
            IV = Encoding.UTF8.GetBytes(IVString);
            if (Key == null || Key.Length <= 0) throw new ArgumentNullException("Key");
            if (IV == null || IV.Length <= 0) throw new ArgumentNullException("IV");
            using (Aes aesAlg = Aes.Create())
            {
                aesAlg.Key = Key;
                aesAlg.IV = IV;
                ICryptoTransform decryptor = aesAlg.CreateDecryptor(aesAlg.Key, aesAlg.IV);
                using (MemoryStream msDecrypt = new MemoryStream(hiddenBytes))
                {
                    using (CryptoStream csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read))
                    {
                        using (MemoryStream msPlain = new MemoryStream())
                        {
                            csDecrypt.CopyTo(msPlain);
                            return msPlain.ToArray();
                        }
                    }
                }
            }
        }
        ~CoverAES()
        {
            GC.Collect();
        }
    }
}