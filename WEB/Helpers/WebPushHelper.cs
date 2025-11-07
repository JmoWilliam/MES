using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WebPush;

namespace Helpers
{
    public class WebPushHelper
    {
        readonly string company;
        readonly string secretKey;
        readonly string subject;
        readonly string publicKey;
        readonly string privateKey;
        //public static string domainUrl = "https://192.168.20.112:26668/";
        public static string domainUrl = "https://bm.zy-tech.com.tw/";

        public WebPushHelper()
        {
            this.company = @"JMO";
            //this.secretKey = @"57670A2810BE5E3AD3D6FB91661B14382D775CF11BC0F0FEB6CF523A0DF4335D";
            this.secretKey = @"FA65D160076DA566CBFC342852580C27DB49415B94CF07CA47B4B9BB194AE8B3";
            this.subject = @"mailto:jmozytech@gmail.com";
            this.publicKey = @"BKM0tOp6v26V36ll4FmozbA9xWvIg_qpqjOU47t1dXx70gli8rr-o7EcmB98ex65NKA12xT9Olrh6DEPfLBaLM4";
            this.privateKey = @"x7zFVlBy8vbM2Otvt-SznUJWE7Mw8UdmQpqkCgKG7V0";
        }

        public void SendPush(List<PushNotificationUser> pushNotificationUsers, string title, string message
            , List<PushNotificationActions> pushNotificationActions = null, List<PushNotificationActionUrls> pushNotificationActionUrls = null)
        {
            if (pushNotificationUsers.Count > 0)
            {
                foreach (var user in pushNotificationUsers)
                {
                    PushSubscription subscription = new PushSubscription(user.ApiEndpoint, user.Keysp256dh, user.Keysauth);
                    var vapidDetails = new VapidDetails(subject, publicKey, privateKey);

                    var payload = new PushNotification
                    {
                        title = title,
                        body = message,
                        actions = pushNotificationActions,
                        urls = pushNotificationActionUrls
                    };

                    var webPushClient = new WebPushClient();
                    try
                    {
                        webPushClient.SendNotification(subscription, JsonConvert.SerializeObject(payload), vapidDetails);
                    }
                    catch (WebPushException)
                    {
                        string targetUrl = string.Format("{0}{1}", domainUrl, "api/BAS/DeleteSubscription");

                        var postData = new List<KeyValuePair<string, string>>
                        {
                            new KeyValuePair<string, string>("Company", company),
                            new KeyValuePair<string, string>("SecretKey", secretKey),
                            new KeyValuePair<string, string>("PushSubscriptionId", user.PushSubscriptionId.ToString()),
                            new KeyValuePair<string, string>("ApiEndpoint", user.ApiEndpoint)
                        };

                        string response = BaseHelper.PostWebRequest(targetUrl, postData);

                        if (response.TryParseJson(out JObject tempJObject))
                        {
                            JObject resultJson = JObject.Parse(response);

                            if (resultJson["status"].ToString() != "success") throw new SystemException(resultJson["msg"].ToString());
                        }
                        else
                        {
                            throw new SystemException(response);
                        }
                    }
                }
            }
        }
    }
}
