using Helpers;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;

namespace ERPSynchronize
{
    class BomSubstitutionSynchronize
    {
        readonly string company;
        readonly string secretKey;
        //public static string domainUrl = "https://192.168.20.112:26668/";
        public static string domainUrl = "https://bm.zy-tech.com.tw/";
        public static Logger logger = LogManager.GetCurrentClassLogger();

        public BomSubstitutionSynchronize(string company, string secretKey)
        {
            this.company = company;
            this.secretKey = secretKey;
        }

        public void Init()
        {
            try
            {
                string targetDate = DateTime.Now.AddDays(-2).ToString("yyyyMMdd HH:mm:ss"),
                    targetUrl = string.Format("{0}{1}", domainUrl, "api/ERP/BomSubstitutionSynchronize");

                var postData = new List<KeyValuePair<string, string>>
                {
                    new KeyValuePair<string, string>("Company", company),
                    new KeyValuePair<string, string>("SecretKey", secretKey),
                    new KeyValuePair<string, string>("UpdateDate", targetDate)
                };

                string response = BaseHelper.PostWebRequest(targetUrl, postData);

                if (response.TryParseJson(out JObject tempJObject))
                {
                    JObject resultJson = JObject.Parse(response);

                    Console.WriteLine("狀態：" + resultJson["status"].ToString());
                    Console.WriteLine("回傳訊息：" + resultJson["msg"].ToString());

                    if (resultJson["status"].ToString() != "success") logger.Error(resultJson["msg"].ToString());
                }
                else
                {
                    logger.Error(response);
                }
            }
            catch (Exception e)
            {
                logger.Error(e.Message);
            }
        }
    }
}