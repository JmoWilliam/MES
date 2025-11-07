using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Text;
using MESIII;
using System.Security.Cryptography;
using System.Globalization;
using Newtonsoft.Json.Converters;
using NiceLabel.SDK;

public class MESConfig
{
    public static ISDClient clientNewMes = new ISDClient(0);
    public static ISDClient clientOldMes = new ISDClient(1);
    public static ISDClient clientErp = new ISDClient(2);
    public static CmnSrvLib cmn = new CmnSrvLib();
    public static readonly IsoDateTimeConverter JsonConvertSetting = new IsoDateTimeConverter()
    {
        DateTimeFormat = "yyyy-MM-dd HH:mm:ss"
    };
    
    public static string CKEDIT_IMAGE_FILE { private set; get; }
    public static string USER_DOMAIN { set; get; }
    public static string USER_PACKAGE { set; get; }
    public static string USER_PROGRAM { set; get; }
    public static int COMPANY_ID { private set; get; }
    public static int USER_ID { private set; get; }
    public static string DEFAULT_USER { private set; get; }
    public static string USER_IMAGE { private set; get; }
    public static int ROLE_ID { private set; get; }
    public static string CKEDIT_IMAGE { private set; get; }
    public static string HOST { private set; get; }
    public static string PORT { private set; get; }
    public static string APPLY_FILE { private set; get; }
    public static string EXCEL_FILE { private set; get; }

    #region //GetEncoding001
    public static string GetEncoding001(string PARAMETER)
    {
        try
        {
            PARAMETER = cmn.MyEncode001(PARAMETER);

        }
        catch (Exception e1)
        {
            return e1.Message;
        }

        return PARAMETER;

    }
    #endregion

    #region //Base64StringToString
    public static string Base64StringToString(string base64, string sKey = "00000000", string sIV = "00000000")
    {
        if (base64 != "")
        {
            try
            {
                base64 = Add(base64);
                base64 = base64.Substring(0, base64.Length - 1);
                string[] sInput = base64.Split("-".ToCharArray());
                byte[] data = new byte[sInput.Length];
                for (int i = 0; i < sInput.Length; i++)
                {
                    data[i] = byte.Parse(sInput[i], NumberStyles.HexNumber);
                }
                DESCryptoServiceProvider DES = new DESCryptoServiceProvider();
                DES.Key = ASCIIEncoding.ASCII.GetBytes(sKey);
                DES.IV = ASCIIEncoding.ASCII.GetBytes(sIV);
                ICryptoTransform desencrypt = DES.CreateDecryptor();
                byte[] result = desencrypt.TransformFinalBlock(data, 0, data.Length);

                return Encoding.UTF8.GetString(result);
            }
            catch { }

            return "解密出错！";
        }
        else
        {
            return "";
        }
    }
    
    private static string Add(string str)
    {
        string res = "";

        for (int i = 0; i < str.Length - 1; i++)
        {
            int j;
            Math.DivRem(i, 2, out j);
            if (j == 0)
                res += str.Substring(i, 2) + "-";
        }
        return res;
    }
    #endregion

    #region //LogData     
    public static void LogData(string Message)
    {
        string LogDirectory = "LogData\\";
        try
        {
            DateTime dt = DateTime.Now;
            string logPath = AppDomain.CurrentDomain.BaseDirectory + LogDirectory;
            if (!Directory.Exists(logPath)) { Directory.CreateDirectory(logPath); }

            string logFile = logPath + string.Format("{0:yyyyMMdd}", dt) + ".log";
            if (!System.IO.File.Exists(logFile))
            {
                using (StreamWriter sw = System.IO.File.CreateText(logFile))
                {
                }
            }
            // 加入 LOG 資料
            string UserTime = string.Format("{0:HH:mm:ss}", dt);
            using (StreamWriter sw = System.IO.File.AppendText(logFile))
            {
                sw.WriteLine(UserTime + ":" + Message);
            }
        }
        catch (Exception ee)
        {
            LogData(ee.Message);
        }
    }
    #endregion
}
