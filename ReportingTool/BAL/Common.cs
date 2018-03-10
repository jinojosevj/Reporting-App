#region NameSpace
using System;
using System.IO;
using System.Web;
using System.Text;
using System.Data;
using System.Configuration;
using System.Security.Cryptography;
using System.Net.Mail;
#endregion NameSpace


namespace ReportingTool.BAL
{
    public class Common
    {

        #region Base64Encode
        /// <summary>
        /// Base64Encode
        /// </summary>
        /// <param name="plainText"></param>
        /// <returns></returns>
        public static string Base64Encode(string plainText)
        {
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(plainText);
            return System.Convert.ToBase64String(plainTextBytes);
        }
        #endregion Base64Encode


        #region Base64Decode
        // For Decoding

        public static string Base64Decode(string base64EncodedData)
        {
            object misValue = System.Reflection.Missing.Value;

            var base64EncodedBytes = System.Convert.FromBase64String(base64EncodedData);
            return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
        }
        #endregion Base64Decode

    }
}