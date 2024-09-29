using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.IO;
using System.Linq;
using System.Web;
using System.Text;

namespace EWS.Includes
{
    public class Compression
    {

        public static string Zip(string str)
        {
            var bytes = Encoding.UTF8.GetBytes(str);
            using (var msi = new MemoryStream(bytes))
            using (var mso = new MemoryStream())
            {
                using (var gs = new GZipStream(mso, CompressionMode.Compress))
                {
                    CopyTo(msi, gs);
                }
                return Encoding.UTF8.GetString(mso.ToArray());
            }
        }
        public static string Unzip(string str)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(str);
            using (var msi = new MemoryStream(bytes))
            using (var mso = new MemoryStream())
            {
                using (var gs = new GZipStream(msi, CompressionMode.Decompress))
                {
                    CopyTo(gs, mso);
                }
                return Encoding.UTF8.GetString(mso.ToArray());
            }
        }
        public static void CopyTo(Stream src, Stream dest)
        {
            byte[] bytes = new byte[4096];

            int cnt;

            while ((cnt = src.Read(bytes, 0, bytes.Length)) != 0)
            {
                dest.Write(bytes, 0, cnt);
            }
        }
        public static bool SaveByteArrayToFile(byte[] fileContent, string filePath)
        {
            if (fileContent.Length > 0 && !string.IsNullOrEmpty(filePath))
            {
                string fileName = filePath + $"{DateTime.Now.ToString("yyyyMMddHHmmssfff")}{new Random().Next()}.xlsx";
                File.WriteAllBytes(fileName, fileContent);
            }
            else
                return false;
            return true;
        }

    }
}