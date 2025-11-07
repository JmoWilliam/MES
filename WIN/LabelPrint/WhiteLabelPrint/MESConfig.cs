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
    #region //LogData     
    public static void LogData(string Message)
    {
        string LogDirectory = "LogData\\";
        try
        {
            DateTime dt = DateTime.Now;
            string logPath = AppDomain.CurrentDomain.BaseDirectory + LogDirectory;
            if (!Directory.Exists(logPath)) { Directory.CreateDirectory(logPath); }

            string logFile = logPath + string.Format("{0:yyyyMMdd}", dt) + ".txt";
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
