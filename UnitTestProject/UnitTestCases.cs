using System;
using System.IO;
using System.ServiceModel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using static EWS.Includes.CoverAES;
using EWS;
using System.Web;
using EWS.Includes;

namespace UnitTestProject
{
    [TestClass]
    public class UnitTestCases
    {
        [TestMethod]
        public void GetItemTestMethod()
        {
            // Blank Item
            GetItem item = new GetItem();
            Console.WriteLine(GetItem.QueryItem(ref item, "role", "pass", "ID|TO|FROM|SUBJECT|HEADERS|BODY|DTR"));
            Console.WriteLine(item.MakePipeDelimitedString());
        }
        [TestMethod]
        public void GetItemTestEncryptedMethod()
        {
            KeyString = "ed1b64d08c9da16a80d3971340c7d698";
            IVString = "d68c69039117ba10";
            // Blank Item
            GetItem item = new GetItem();
            Console.WriteLine(GetItem.QueryItem(ref item, "role", "pass", "ID|TO|FROM|HEADERS|SUBJECT|BODY|DTR"));
            item.To = DecryptStringAES(item.To);
            item.From = DecryptStringAES(item.From);
            item.Headers = DecryptStringAES(item.Headers);
            item.Subject = DecryptStringAES(item.Subject);
            item.Body = DecryptStringAES(item.Body);
            item.DTR = DecryptStringAES(item.DTR);
            Console.WriteLine(item.MakePipeDelimitedString());
        }
        [TestMethod]
        public void CreateItemTestMethod()
        {
            KeyString = "ed1b64d08c9da16a80d3971340c7d698";
            IVString = "d68c69039117ba10";
            // Blank Item
            GetItem item = new GetItem();
            // Read Attachment File
            byte[] fileAttachment = File.ReadAllBytes("..\\..\\Include\\IRP.xlsx");
            string fileName = $"{DateTime.Now.ToString("yyyyMMddHHmmssfff")}{new Random().Next()}.xlsx";
            //File.WriteAllBytes(fileName, fileAttachment);
            Console.WriteLine($"TestCase Blank Item:{GetItem.SendItem(ref item, string.Empty, string.Empty)}");
            // Empty Initialized Item
            item.Id = string.Empty;
            item.To = string.Empty;
            item.From = string.Empty;
            item.Headers = string.Empty;
            item.Subject = string.Empty;
            item.Body = string.Empty;
            item.DTR = string.Empty;
            item.FileName = string.Empty;
            item.FileContent = null;
            Console.WriteLine($"TestCase Empty Initialized Item:{GetItem.SendItem(ref item, "role", "pass")}");
            // Only Email Flag Zero
            item.Flag = 0;
            item.Id = string.Empty;
            item.To = "shahid.here.business@gmail.com";
            item.From = string.Empty;
            item.Headers = string.Empty;
            item.Subject = "Only Email Flag Zero";
            item.Body = "Hi, It's good to hear from you soon!";
            item.DTR = string.Empty;
            item.FileName = string.Empty;
            item.FileContent = fileAttachment;
            Console.WriteLine($"TestCase Only Email Flag Zero:{GetItem.SendItem(ref item, "role", "pass")}");
            // Only Email Flag One
            item.Flag = 1;
            item.Id = string.Empty;
            item.To = "shahid.here.business@gmail.com";
            item.From = string.Empty;
            item.Headers = string.Empty;
            item.Subject = "Only Email Flag One";
            item.Body = "Hi, It's good to hear from you soon!";
            item.DTR = string.Empty;
            item.FileName = string.Empty;
            item.FileContent = fileAttachment;
            Console.WriteLine($"TestCase Only Email Flag One:{GetItem.SendItem(ref item, "role", "pass")}");
            // Attachment Email Without Flag            
            item.Id = string.Empty;
            item.To = "shahid.here.business@gmail.com";
            item.From = string.Empty;
            item.Headers = string.Empty;
            item.Subject = "Attachment Email Without Flag";
            item.Body = "Hi, It's good to hear from you soon!";
            item.DTR = string.Empty;
            item.FileName = string.Empty;
            item.FileContent = fileAttachment;
            Console.WriteLine($"TestCase Attachment Email Without Flag :{GetItem.SendItem(ref item, "role", "pass")}");
            // Attachment Email Flag Two with Empty FileName
            item.Flag = 2;
            item.Id = string.Empty;
            item.To = "shahid.here.business@gmail.com";
            item.From = string.Empty;
            item.Headers = string.Empty;
            item.Subject = "Attachment Email Flag Two with Empty FileName";
            item.Body = "Hi, It's good to hear from you soon!";
            item.DTR = string.Empty;
            item.FileName = string.Empty;
            item.FileContent = fileAttachment;
            Console.WriteLine($"TestCase Attachment Empty FileName Flag Two:{GetItem.SendItem(ref item, "role", "pass")}");
            // Attachment Email Flag Two Success
            item.Flag = 2;
            item.Id = string.Empty;
            item.To = "shahid.here.business@gmail.com";
            item.From = string.Empty;
            item.Headers = string.Empty;
            item.Subject = "Attachment Email Flag Two Success";
            item.Body = "Hi, It's good to hear from you soon!";
            item.DTR = string.Empty;
            item.FileName = "IRPs.xlsx";
            item.FileContent = fileAttachment;
            Console.WriteLine($"TestCase Attachment Email Flag Two Success:{GetItem.SendItem(ref item, "role", "pass")}");
        }
        [TestMethod]
        public void CreateItemEncryptedMethod()
        {
            KeyString = "ed1b64d08c9da16a80d3971340c7d698";
            IVString = "d68c69039117ba10";
            // Blank Item
            GetItem item = new GetItem();
            // Read Attachment File
            byte[] fileAttachment = File.ReadAllBytes("..\\..\\Include\\IRP.xlsx");
            string fileName = $"{DateTime.Now.ToString("yyyyMMddHHmmssfff")}{new Random().Next()}.xlsx";
            //File.WriteAllBytes(fileName, fileAttachment);
            // Encrypted Email Without Flag Success
            item.Id = string.Empty;
            item.To = EncryptStringAES("shahid.here.business@gmail.com");
            item.From = string.Empty;
            item.Headers = string.Empty;
            item.Subject = EncryptStringAES("Encrypted Email Without Flag Success");
            item.Body = EncryptStringAES("Hi, It's good to hear from you soon!");
            item.DTR = string.Empty;
            item.FileName = EncryptStringAES("IRPs.xlsx");
            item.FileContent = EncryptByteArrayAES(fileAttachment);
            Console.WriteLine($"TestCase Encrypted Email Without Flag Success:{GetItem.SendItem(ref item, "role", "pass")}");
            // Encrypted Attachment Email Flag Two Success
            item.Flag = 2;
            item.Id = string.Empty;
            item.To = EncryptStringAES("shahid.here.business@gmail.com");
            item.From = string.Empty;
            item.Headers = string.Empty;
            item.Subject = EncryptStringAES("Encrypted Attachment Email Flag Two Success");
            item.Body = EncryptStringAES("Hi, It's good to hear from you soon!");
            item.DTR = string.Empty;
            item.FileName = EncryptStringAES("IRPs.xlsx");
            item.FileContent = EncryptByteArrayAES(fileAttachment);
            Console.WriteLine($"TestCase Encrypted Attachment Email Flag Two Success:{GetItem.SendItem(ref item, "role", "pass")}");
        }
    }
}