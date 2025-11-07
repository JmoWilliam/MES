using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Http;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Data.SqlClient;
using Dapper;
using NLog;

namespace Helpers
{
    public class LinePushNotifiHelper
    {

        //public static LinePushNotifiHelper linePushNotifiHelper = new LinePushNotifiHelper();
        private string apiLineUrl = "https://notify-bot.line.me";
        private string apiMainUrl = "http://192.168.134.52:16668/";
        public JObject jsonResponse = new JObject();
        public string dataRequest = "";
        public static DynamicParameters dynamicParameters = new DynamicParameters();
        public static Logger logger = LogManager.GetCurrentClassLogger();
        public string ClientID = "qJ1mhzvc0cRF1NKKBdLd0G";
        public string SecretKey = "QyAIPN5arM0jD7iKzP1uNYwuzSLniR9fWYeWiNkyhmw";
        public string callbackURL  = "http://192.168.134.52:16668/Event/LineNotifyTest2";
       // private string LineLoginURL = "https://notify-bot.line.me/oauth/authorize?response_type=code&client_id=" + linePushNotifiHelper.ClientID + "&redirect_uri=" + linePushNotifiHelper.callbackURL + "&scope=notify&state=Gpai";




        #region // 跳轉LINE登入畫面
        public string LineLogin()
        {
            try
            {
                string apiUrl = Path.Combine(apiMainUrl, "api/LinePush/LineLogin");

                #region //Request
                using (HttpClient client = new HttpClient())
                {
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                    {
                        MultipartFormDataContent multipart = new MultipartFormDataContent();
                        //multipart.Add(new StringContent(userId), "userId");
                        multipart.Add(new StringContent(ClientID), "client_id");
                        multipart.Add(new StringContent(callbackURL), "redirect_uri");
                        httpRequestMessage.Content = multipart;

                        using (HttpResponseMessage httpResponseMessage = client.SendAsync(httpRequestMessage).GetAwaiter().GetResult())
                        {
                            if (httpResponseMessage.IsSuccessStatusCode)
                            {
                                #region //取得回傳訊息
                                var response = httpResponseMessage.Content.ReadAsStringAsync();
                                jsonResponse = JObject.Parse(response.Result);
                                #endregion
                            }
                            else
                            {
                                throw new SystemException(httpResponseMessage.StatusCode.ToString());
                            }
                        }
                    }
                }
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region // GetLineNotifyToken 取得Token
        public string GetLineNotifyToken(string code)
        {
            try
            {
                string apiUrl = Path.Combine(apiLineUrl, "/oauth/token");

                #region //Request
                using (HttpClient client = new HttpClient())
                {
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                    {
                        MultipartFormDataContent multipart = new MultipartFormDataContent();
                        //multipart.Add(new StringContent(userId), "userId");
                        multipart.Add(new StringContent("authorization_code"), "grant_type");
                        multipart.Add(new StringContent(code), "code");
                        multipart.Add(new StringContent(callbackURL), "redirect_uri");
                        multipart.Add(new StringContent(ClientID), "client_id");
                        multipart.Add(new StringContent(SecretKey), "client_secret");
                        httpRequestMessage.Content = multipart;

                        using (HttpResponseMessage httpResponseMessage = client.SendAsync(httpRequestMessage).GetAwaiter().GetResult())
                        {
                            if (httpResponseMessage.IsSuccessStatusCode)
                            {
                                #region //取得回傳訊息
                                var response = httpResponseMessage.Content.ReadAsStringAsync();
                                jsonResponse = JObject.Parse(response.Result);
                                #endregion
                            }
                            else
                            {
                                throw new SystemException(httpResponseMessage.StatusCode.ToString());
                            }
                        }
                    }
                }
                #endregion
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion




    }
}
