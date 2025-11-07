using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Helpers
{
    public class WeixinHelper
    {
        private readonly string token = "jmo";
        private readonly string appID = "wx0882fad396ff10a8";
        private readonly string encodingAESKey = "tg3ONPdFUcbCo9jmJ7oC7TaSKQXIGwSfCPPFr6BUwpU";
        private readonly string appSecret = "e1a4df85ce50e9419b5b469493fee842";

        public string accessToken = "";

        public WeixinHelper()
        {
            accessToken = GetAccessToken();
        }

        #region //GetAccessToken 取得令牌
        /// <summary>
        /// 取得令牌
        /// </summary>
        /// <returns></returns>
        private string GetAccessToken()
        {
            string responseText = BaseHelper.PostWebRequest("http://192.168.20.112:10002/api/Weixin/GetAccessToken");

            if (JsonConvert.DeserializeObject<dynamic>(responseText).status == "success")
            {
                var response = JsonConvert.DeserializeObject<dynamic>(responseText);

                return response.data[0].Plaintext.ToString();
            }
            else
            {
                var data = new
                {
                    grant_type = "client_credential",
                    appid = appID,
                    secret = appSecret
                };

                responseText = BaseHelper.PostWebRequest("https://api.weixin.qq.com/cgi-bin/stable_token", JsonConvert.SerializeObject(data), Encoding.UTF8);
                var response = JsonConvert.DeserializeObject<dynamic>(responseText);

                if (response.access_token != null)
                {
                    var updateData = new
                    {
                        Plaintext = response.access_token.ToString(),
                        ExpiresIn = Convert.ToInt32(response.expires_in)
                    };

                    responseText = BaseHelper.PostWebRequest("http://192.168.20.112:10002/api/Weixin/UpdateAccessToken", JsonConvert.SerializeObject(updateData), Encoding.UTF8);
                    var responseUpdate = JsonConvert.DeserializeObject<dynamic>(responseText);

                    if (responseUpdate.status.ToString() == "success")
                    {
                        return response.access_token.ToString();
                    }
                    else
                    {
                        throw new SystemException(responseUpdate.msg.ToString());
                    }
                }
                else
                {
                    throw new SystemException("Token請求失敗");
                }
            }
        }
        #endregion

        #region //CheckSignature 檢查簽章
        /// <summary>
        /// 
        /// </summary>
        /// <param name="signature"></param>
        /// <param name="timestamp"></param>
        /// <param name="nonce"></param>
        /// <returns></returns>
        public bool CheckSignature(string signature, string timestamp, string nonce)
        {
            List<string> temp = new List<string>
                {
                    token,
                    timestamp,
                    nonce
                };
            temp = temp.Select(x => x).OrderBy(x => x).ToList();

            string tempStr = string.Join("", temp);
            string checkSignature = BaseHelper.Sha1Encrypt(tempStr).ToLower();

            return checkSignature == signature;
        }
        #endregion

        #region //GetServerIP 取得微信API主機IP
        public object GetServerIP()
        {
           string responseText = BaseHelper.PostWebRequest(string.Format("https://api.weixin.qq.com/cgi-bin/get_api_domain_ip?access_token={0}", accessToken));
            var response = JsonConvert.DeserializeObject<dynamic>(responseText);

            return response;
        }
        #endregion
    }
}
