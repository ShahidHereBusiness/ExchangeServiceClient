using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Services;
using EWS.Includes;

namespace EWS
{
    /// <summary>
    /// Summary description for Exchage Web Service
    /// </summary>
    [WebService(Name = "Email Parsing", Namespace = "https://domain/SOA/", Description = "System Email Parsing & ICT")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    // [System.Web.Script.Services.ScriptService]
    public class XchangeWebService : System.Web.Services.WebService
    {
        /// <summary>
        /// /// Parse Email and Relay
        /// </summary>
        /// <param name="accessSpecifier"></param>
        /// <param name="watchWord"></param>
        /// <param name="paramsList"></param>
        /// <param name="item">Email Object</param>
        /// <returns></returns>
        [WebMethod(Description = "Obtains Email Attributes", MessageName = "GetItem")]
        public int HitSend(string accessSpecifier, string watchWord, string paramsList, ref GetItem item)
        {
            return (int)GetItem.QueryItem(ref item, accessSpecifier, watchWord, paramsList);
        }
        /// <summary>
        /// Send Email Item from Exchange Group
        /// </summary>
        /// <param name="accessSpecifier"></param>
        /// <param name="watchWord"></param>
        /// <param name="item"></param>
        /// <returns></returns>
        [WebMethod(Description = "Sends Formatted Email", MessageName = "CreateItem")]
        public int MakeSend(string accessSpecifier, string watchWord, ref GetItem item)
        {
            return (int)GetItem.SendItem(ref item, accessSpecifier, watchWord);
        }
    }
}