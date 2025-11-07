using Helpers;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ERPSynchronize
{
    class MoProgressDetailNotInventoryMAMO
    {
        readonly string company;
        readonly string secretKey;
        readonly string channelId;
        //public static string domainUrl = "http://192.168.134.33:16668/";
        public static string domainUrl = "https://bm.zy-tech.com.tw/";
        public static Logger logger = LogManager.GetCurrentClassLogger();

        public MoProgressDetailNotInventoryMAMO(string company, string secretKey, string channelId)
        {
            this.company = company;
            this.secretKey = secretKey;
            this.channelId = channelId;
        }

        public void Init()
        {
            try
            {
                string targetUrl = string.Format("{0}{1}", domainUrl, "api/MesReport/MoProgressDetailNotInventoryMamoEIP");

                var postData = new List<KeyValuePair<string, string>>
                {
                    new KeyValuePair<string, string>("Company", company),
                    new KeyValuePair<string, string>("SecretKey", secretKey),
                    new KeyValuePair<string, string>("ChannelId", channelId)
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
