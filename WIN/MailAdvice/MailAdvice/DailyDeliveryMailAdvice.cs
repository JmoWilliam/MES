using Helpers;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;

namespace MailAdvice
{
    class DailyDeliveryMailAdvice
    {
        readonly string company;
        readonly string secretKey;
        //public static string domainUrl = "https://192.168.20.112:26668/";
        public static string domainUrl = "https://bm.zy-tech.com.tw/";
        public static Logger logger = LogManager.GetCurrentClassLogger();

        public DailyDeliveryMailAdvice(string company, string secretKey)
        {
            this.company = company;
            this.secretKey = secretKey;
        }

        public void Init()
        {
            try
            {
                string targetUrl = string.Format("{0}{1}", domainUrl, "api/ERP/DailyDelivery");

                var postData = new List<KeyValuePair<string, string>>
                {
                    new KeyValuePair<string, string>("Company", company),
                    new KeyValuePair<string, string>("SecretKey", secretKey)
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