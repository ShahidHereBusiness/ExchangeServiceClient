using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;
using EWS.Includes;
using Microsoft.Exchange.WebServices.Data;
using static EWS.Includes.CoverAES;
using static EWS.Includes.Enumeration;
using static EWS.Includes.Validation;
using static EWS.Includes.Compression;

namespace EWS.Includes
{
    public class GetItem
    {
        // Initializations
        private static ExchangeService services = null;
        private static String uri = null;
        private static String EWSSpecifier = null;
        private static String EWSWords = null;
        private static EmailMessage email = null;
        Item item = null;
        //private static HandleAppLog hf = null;
        private static string path = string.Empty;
        private static string remoteAddress = string.Empty;
        private static bool WebHost = true;
        private static Boolean flagRSA = false;
        /// <summary>
        /// Constructor Function
        /// </summary>
        public GetItem() { }
        public int Flag
        {
            get; set;
        }
        public string Id
        {
            get; set;
        }
        public string From
        {
            get; set;
        }
        public string To
        {
            get; set;
        }
        public string Headers
        {
            get; set;
        }
        public string Subject
        {
            get; set;
        }
        public string Body
        {
            get; set;
        }
        public string DTR
        {
            get; set;
        }
        public string FileName
        {
            get; set;
        }
        public byte[] FileContent
        {
            get; set;
        }
        public static ResponseEnum SendItem(ref GetItem item, string accessSpecifier, string watchWord)
        {
            using (HandleAppLog hf = new HandleAppLog(remoteAddress))
            {
                try
                {
                    // Service Initialization
                    if (!Instantiate())
                        return ResponseEnum.IntializationError;
                    // Not White Listed               
                    if (WebHost)
                        return ResponseEnum.NotAllowedURI;

                    #region Configure Log Encryption                    
                    string msg = EncryptStringAES(item.MakePipeDelimitedString());
                    #endregion
                    if (FormatError(path))
                        return ResponseEnum.ConfigurationError;
                    else if (FormatError(KeyString) || FormatError(IVString))
                        return ResponseEnum.AESConfigurationError;
                    else if (CredentialsError(hf, path, accessSpecifier, watchWord))
                        return ResponseEnum.InvalidCredentials;
                    else if (FormatError(uri) || FormatError(EWSSpecifier) || FormatError(EWSWords))
                        return ResponseEnum.ExternalCredentialsError;
                    else if (EmailError(item.To, flagRSA) || FormatError(item.Subject) || FormatError(item.Body))
                        return ResponseEnum.FormatError;
                    else if ((item.Flag == (int)FlagEnum.IsAttachment) && AttachmentError(item))
                        return ResponseEnum.FormatError;
                    else
                    {
                        #region Validate Encrypted Log Request
                        if (hf.FileSystemLog(path, "MakeSend", "Receipt", msg) > 0)
                            return ResponseEnum.FileSystemLogFailure;
                        #endregion

                        // Decrypt Before Sending Email
                        if (flagRSA)
                        {
                            item.To = (string.IsNullOrEmpty(item.To)) ? string.Empty : DecryptStringAES(item.To);
                            item.From = (string.IsNullOrEmpty(item.From)) ? string.Empty : DecryptStringAES(item.From);
                            item.Headers = (string.IsNullOrEmpty(item.Headers)) ? string.Empty : DecryptStringAES(item.Headers);
                            item.Subject = (string.IsNullOrEmpty(item.Subject)) ? string.Empty : DecryptStringAES(item.Subject);
                            item.Body = (string.IsNullOrEmpty(item.Body)) ? string.Empty : DecryptStringAES(item.Body);
                            item.DTR = (string.IsNullOrEmpty(item.DTR)) ? string.Empty : DecryptStringAES(item.DTR);
                            item.FileName = (AttachmentError(item)) ? string.Empty : DecryptStringAES(item.FileName);
                            item.FileContent = (AttachmentError(item)) ? null : DecryptByteArrayAES(item.FileContent);
                        }
                        // Write XLSX Test Case Decrypted
                        //if (item.FileContent != null)
                        //    SaveByteArrayToFile(item.FileContent, path);
                        // Email from Mailbox
                        if (!item.SendMessage())
                            return ResponseEnum.UnexpectedFailure;
                        #region Encrypted Log Response
                        msg = EncryptStringAES(item.MakePipeDelimitedString());
                        hf.FileSystemLog(path, "MakeSend", "Response", msg);
                        #endregion
                    }
                }
                catch (Exception ex)
                {
                    #region Exception Log                
                    hf.FileSystemLog(path, "MakeSend", "Exception", EncryptStringAES(item.MakePipeDelimitedString()));
                    #endregion
                    if (ex.ToString().Length > 0)
                        return ResponseEnum.UnhandledServing;
                }
                finally
                {
                    item.Id = string.Empty;
                    item.To = string.Empty;
                    item.From = string.Empty;
                    item.Headers = string.Empty;
                    item.Subject = string.Empty;
                    item.Body = string.Empty;
                    item.DTR = string.Empty;
                    item.FileName = string.Empty;
                    item.FileContent = null;
                }
            }
            return ResponseEnum.Success;
        }
        public static ResponseEnum QueryItem(ref GetItem item, string accessSpecifier, string watchWord, string paramsList)
        {
            using (HandleAppLog hf = new HandleAppLog(remoteAddress))
            {
                try
                {
                    // Service Initialization
                    if (!Instantiate())
                        return ResponseEnum.IntializationError;
                    // Not White Listed               
                    if (WebHost)
                        return ResponseEnum.NotAllowedURI;

                    #region Configure Log Encryption                
                    string msg = EncryptStringAES(item.MakePipeDelimitedString());
                    #endregion

                    if (FormatError(path))
                        return ResponseEnum.ConfigurationError;
                    else if (FormatError(KeyString) || FormatError(IVString))
                        return ResponseEnum.AESConfigurationError;
                    else if (CredentialsError(hf, path, accessSpecifier, watchWord))
                        return ResponseEnum.InvalidCredentials;
                    else if (FormatError(uri) || FormatError(EWSSpecifier) || FormatError(EWSWords))
                        return ResponseEnum.ExternalCredentialsError;
                    else
                    {
                        #region Validate Encrypted Log Request
                        if (hf.FileSystemLog(path, "HitSend", "Receipt", msg) > 0)
                            return ResponseEnum.FileSystemLogFailure;
                        #endregion
                        // Get Unread Email from Mailbox
                        if (!item.GetEnqueue(paramsList))
                            return ResponseEnum.UnexpectedFailure;
                        // Set Email as Read required
                        if (item.Flag == (int)FlagEnum.MarkRead)
                            if (!item.SetDequeue())
                                return ResponseEnum.UnexpectedFailure;
                        #region Encrypted Log Response
                        msg = EncryptStringAES(item.MakePipeDelimitedString());
                        hf.FileSystemLog(path, "HitSend", "Response", msg);
                        #endregion
                        // Encrypt Before Sending GetItem
                        if (flagRSA)
                        {
                            item.To = (string.IsNullOrEmpty(item.To)) ? string.Empty : EncryptStringAES(item.To);
                            item.From = (string.IsNullOrEmpty(item.From)) ? string.Empty : EncryptStringAES(item.From);
                            item.Headers = (string.IsNullOrEmpty(item.Headers)) ? string.Empty : EncryptStringAES(item.Headers);
                            item.Subject = (string.IsNullOrEmpty(item.Subject)) ? string.Empty : EncryptStringAES(item.Subject);
                            item.Body = (string.IsNullOrEmpty(item.Body)) ? string.Empty : EncryptStringAES(item.Body);
                            item.DTR = (string.IsNullOrEmpty(item.DTR)) ? string.Empty : EncryptStringAES(item.DTR);
                        }
                    }
                }
                catch (Exception ex)
                {
                    #region Exception Log                
                    hf.FileSystemLog(path, "HitSend", "Exception", EncryptStringAES(item.MakePipeDelimitedString()));
                    #endregion
                    if (ex.ToString().Length > 0)
                        return ResponseEnum.UnhandledServing;
                }
            }
            return ResponseEnum.Success;
        }
        /// <summary>
        /// Get Pipe delimited all Attribute from Object
        /// </summary>
        /// <returns></returns>
        public string MakePipeDelimitedString()
        {
            var properties = GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance);
            var values = new object[properties.Length];

            for (int i = 0; i < properties.Length; i++)
            {
                values[i] = properties[i].GetValue(this);
            }
            string str = string.Join("|", values);
            return str.Substring(0, str.Length - 1);
        }
        /// <summary>
        /// Get Next Email Queued for Traversing
        /// </summary>
        /// <returns>0/1</returns>
        public Boolean GetEnqueue(string respParam)
        {
            using (HandleAppLog hf = new HandleAppLog(remoteAddress))
            {
                string[] paramList = respParam.ToLower().Trim().Split('|');
                try
                {
                    // Create a search filter for unread emails
                    SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));
                    // Define the Inbox folder
                    Folder inbox = Folder.Bind(services, WellKnownFolderName.Inbox);
                    // Set the number of items to retrieve
                    int numItems = 1;
                    // Retrieve the emails
                    FindItemsResults<Item> findResults = services.FindItems(WellKnownFolderName.Inbox, searchFilter, new ItemView(numItems));
                    // Get Singleton Email Instance Received
                    if (findResults.Items.Count < 1)
                        return true;
                    this.item = findResults.Items[0];
                    this.Id = item.Id.ToString();
                    email = EmailMessage.Bind(services, item.Id, new PropertySet(ItemSchema.Attachments));
                    email.Load();
                    // Output email details
                    if (Array.IndexOf(paramList, "to") > -1)
                    {
                        this.To = string.Empty;// Receipient Emails
                        foreach (EmailAddress toAdd in email.ToRecipients)
                        {
                            this.To += toAdd.Address + ";";
                        }
                    }
                    if (Array.IndexOf(paramList, "from") > -1)
                        this.From = email.From.Address;//Sender Email                
                    if (Array.IndexOf(paramList, "header") > -1)
                    {
                        int count = email.InternetMessageHeaders.Count;
                        this.Headers = string.Empty;// Email Headers
                        for (int i = 0; i < count; i++)
                        {
                            this.Headers += email.InternetMessageHeaders[i].Name + "|" + email.InternetMessageHeaders[i].Value + System.Environment.NewLine;
                        }
                    }
                    if (Array.IndexOf(paramList, "subject") > -1)
                        this.Subject = email.Subject;//Email Subject
                    if (Array.IndexOf(paramList, "body") > -1)
                        this.Body = email.Body.Text;// Email Body Text
                    if (Array.IndexOf(paramList, "dtr") > -1)
                        this.DTR = email.DateTimeReceived.ToString();//Email Receive Date & Time
                }
                catch (Exception ex)
                {
                    #region Exception Log                
                    hf.FileSystemLog(path, "GetEnqueue", "Exception", EncryptStringAES(ex.ToString()));
                    #endregion
                    return false;
                }
            }
            return true;
        }
        /// <summary>
        /// Respond to Mark Completed
        /// </summary>
        /// <returns>True/False</returns>
        public Boolean SetDequeue()
        {
            using (HandleAppLog hf = new HandleAppLog(remoteAddress))
            {
                try
                {
                    if (item != null)
                    {
                        // Bind to the email message
                        email = EmailMessage.Bind(services, item.Id);
                        // Mark the email as read
                        email.IsRead = true;
                        email.Update(ConflictResolutionMode.AutoResolve);
                    }
                }
                catch (Exception ex)
                {
                    #region Exception Log                
                    hf.FileSystemLog(path, "SetDequeue", "Exception", EncryptStringAES(ex.ToString()));
                    #endregion
                    return false;
                }
            }
            return true;
        }
        /// <summary>
        /// Send Formatted Email
        /// </summary>
        /// <returns>True/False</returns>
        public Boolean SendMessage()
        {
            using (HandleAppLog hf = new HandleAppLog(remoteAddress))
            {
                try
                {
                    if (To.Trim().Length > 0 && Subject.Trim().Length > 0 && Body.Trim().Length > 0)
                    {
                        EmailMessage email = new EmailMessage(services);
                        email.From = EWSSpecifier.Trim();
                        email.ToRecipients.Add(To.Replace(";", ""));// Validate Addressee with hot fix semi colon
                        email.Subject = Subject;// Validate AIM
                        email.Body = new MessageBody(Body);// Validate Language
                        if (Flag == (int)Enumeration.FlagEnum.IsAttachment)
                            email.Attachments.AddFileAttachment(FileName, FileContent);  //Named file Attachment                                     
                        email.SendAndSaveCopy(WellKnownFolderName.SentItems);// Send and Save in Sent Item
                    }
                }
                catch (Exception ex)
                {
                    #region Exception Log                
                    hf.FileSystemLog(path, "SendMessage", "Exception", EncryptStringAES(ex.ToString()));
                    #endregion
                    return false;
                }
            }
            return true;
        }
        /// <summary>
        /// URI Validation
        /// </summary>
        /// <param name="redirectionUrl">Exchange URL</param>
        /// <returns></returns>
        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            Uri redirectionUri = new Uri(redirectionUrl);
            // The redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                return true;
            }
            return false;
        }
        /// <summary>
        /// Destructor Function
        /// </summary>
        public static bool Instantiate()
        {
            try
            {
                #region Validate White Listed                
                // Localized
                if (HttpContext.Current == null)
                    WebHost = false;
                else// Web Hosting
                {
                    remoteAddress = HttpContext.Current.Request.ServerVariables["REMOTE_ADDR"].ToString().Trim();
                    string[] allowedAddress = ConfigurationManager.AppSettings["WhiteListed"].ToString().Split('|');
                    if (allowedAddress.Contains(remoteAddress))
                        WebHost = false;
                }
                #endregion                                
                // Credentials
                uri = ConfigurationManager.AppSettings["uri"];
                EWSSpecifier = ConfigurationManager.AppSettings["EWSSpecifier"];
                EWSWords = ConfigurationManager.AppSettings["EWSWords"];
                // RSA Config
                flagRSA = Boolean.Parse(ConfigurationManager.AppSettings["encryptItem"].ToString().Trim());
                // AES 
                KeyString = ConfigurationManager.AppSettings["AESKey"];
                IVString = ConfigurationManager.AppSettings["AESiv"];
                // Log Path
                path = ConfigurationManager.AppSettings["path"] + DateTime.Now.Year + "\\" + DateTime.Now.Month + "\\" + DateTime.Now.Day + "\\";
                // Exchange Service Configuration
                services = new ExchangeService(ExchangeVersion.Exchange2013_SP1);//Exchange2010_SP2 Version 1.0 Exchange2007_SP1
                services.Credentials = new WebCredentials(EWSSpecifier, EWSWords);
                services.Url = new Uri(uri);
                //services.AutodiscoverUrl(EWSSpecifier, RedirectionUrlValidationCallback);
                return true;
            }
            catch (Exception ex)
            {
                if (ex.ToString().Length > 0)
                    return false;
            }
            return false;
        }
        ~GetItem()
        {
            GC.Collect();
        }
    }
}