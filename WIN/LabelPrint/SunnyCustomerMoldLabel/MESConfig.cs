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

public class MESConfig
{
    public static ISDClient client;
    public static ISDWeb web;
    public static ISDWeb web01;
    public static ISDWeb webERP = new ISDWeb(1);
    public static ISDWeb webIOT = new ISDWeb(2);
    public static CmnSrvLib cmn = new CmnSrvLib();

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

    public static void GetParameter(int SERVER_ID)
    {
        CmnSrvLib cmn = new CmnSrvLib(SERVER_ID);
        client = new ISDClient(SERVER_ID);
        web = new ISDWeb(SERVER_ID);
        web01 = new ISDWeb(1);
        string sParam = "";
        DataSet oDS = new DataSet();
        COMPANY_ID = -1;
        DEFAULT_USER = "";
        USER_IMAGE = "";
        ROLE_ID = -1;
        CKEDIT_IMAGE = "";
        HOST = "";
        PORT = "";

        try
        {
            sParam = "<root><PARAMETER_INFO>";
            sParam += "<PARAM_NO></PARAM_NO>";
            sParam += "</PARAMETER_INFO></root>";
            oDS = MESConfig.client.ctEnumerateData("SYSSO.QryParameterInfo001", sParam);
            if (cmn.CheckEOF(oDS))
            {
                foreach (DataRow dr in oDS.Tables[0].Rows)
                {
                    string PARAM_NO = cmn.cap_string(dr["PARAM_NO"]);

                    switch (PARAM_NO)
                    {
                        case "COMPANY_ID": COMPANY_ID = cmn.cap_int(dr["PARAM_VALUE"]); break;
                        case "DEFAULT_USER": DEFAULT_USER = cmn.cap_string(dr["PARAM_VALUE"]); break;
                        case "USER_IMAGE": USER_IMAGE = cmn.cap_string(dr["PARAM_VALUE"]); break;
                        case "ROLE_ID": ROLE_ID = Convert.ToInt32(dr["PARAM_VALUE"]); break;
                        case "CKEDIT_IMAGE": CKEDIT_IMAGE = cmn.cap_string(dr["PARAM_VALUE"]); break;
                        case "HOST": HOST = cmn.cap_string(dr["PARAM_VALUE"]); break;
                        case "PORT": PORT = cmn.cap_string(dr["PARAM_VALUE"]); break;
                        case "APPLY_FILE": APPLY_FILE = cmn.cap_string(dr["PARAM_VALUE"]); break;
                        case "EXCEL_FILE": EXCEL_FILE = cmn.cap_string(dr["PARAM_VALUE"]); break;
                    }
                }
            }
        }
        catch (Exception e1)
        {
            LogData(e1.Message);
        }
    }

    #region //GetRoleName001
    public static string GetRoleName001(int ROLE_ID)
    {
        string ROLE_NAME = "";
        try
        {
            string sParam = "<root><ROLE_INFO>";
            sParam += "<ROLE_ID>" + ROLE_ID + "</ROLE_ID>";
            sParam += "<ROLE_NAME_CH></ROLE_NAME_CH>";
            sParam += "<STATUS_NO></STATUS_NO>";
            sParam += "</ROLE_INFO>";
            sParam += "<PAGE_INFO>";
            sParam += "<PAGE_INDEX>-1</PAGE_INDEX>";
            sParam += "<PAGE_SIZE>-1</PAGE_SIZE>";
            sParam += "<ORDER_BY></ORDER_BY>";
            sParam += "</PAGE_INFO>";
            sParam += "</root>";
            DataSet oDS = MESConfig.web.ctEnumerateData("SYSSO.QryRoleInfo001", sParam);
            if (cmn.CheckEOF(oDS))
            {
                ROLE_NAME = oDS.Tables[0].Rows[0]["ROLE_NAME_CH"].ToString();
            }
        }
        catch (Exception e1)
        {
            return e1.Message;
        }

        return ROLE_NAME;

    }
    #endregion

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
