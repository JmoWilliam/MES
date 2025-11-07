using Dapper;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web;
using System.Configuration;
using System.Net.Http.Headers;

namespace Helpers
{
    public class BpmHelper
    {
        #region //GetBpmToken
        public static string GetBpmToken(string BpmServerPath, string BpmAccount, string BpmPassword)
        {
            string postUrl = BpmServerPath + "/WebServiceStart/services/getToken";
            using (HttpClient httpClient = new HttpClient())
            {
                using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, postUrl))
                {
                    httpRequestMessage.Content = new MultipartFormDataContent
                    {
                        { new StringContent(BpmAccount), "appId" },
                        { new StringContent(BpmPassword), "appKey" }
                    };

                    using (HttpResponseMessage httpResponseMessage = httpClient.SendAsync(httpRequestMessage).GetAwaiter().GetResult())
                    {
                        if (httpResponseMessage.IsSuccessStatusCode)
                        {
                            #region //Response
                            JObject jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                data = httpResponseMessage.Content.ReadAsStringAsync().Result.ToString()
                            });
                            #endregion
                            return jsonResponse.ToString();
                        }
                        else
                        {
                            #region //Response
                            JObject jsonResponse = JObject.FromObject(new
                            {
                                status = "error",
                                data = httpResponseMessage.Content.ReadAsStringAsync().Result.ToString()
                            });
                            #endregion
                            return jsonResponse.ToString();
                        }
                    }
                }
            }
        }
        #endregion

        #region //取得BpmUser資訊
        public static string GetBpmUser(SqlConnection sqlConnection, int UserId)
        {
            DynamicParameters dynamicParameters = new DynamicParameters();
            JObject jsonResponse = new JObject();
            string sql = @"SELECT a.UserId, a.BpmUserId, a.BpmUserNo, a.BpmUserName
                            , a.BpmRoleId, a.BpmRoleName, a.BpmDepId, a.BpmDepNo, a.BpmDepName
                            FROM BPM.BpmUser a
                            INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                            WHERE b.UserId = @UserId";
            dynamicParameters.Add("UserId", UserId);

            var systemTokenResult = sqlConnection.Query(sql, dynamicParameters);

            if (systemTokenResult.Count() > 0)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    data = systemTokenResult
                });
                #endregion
            }
            else
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    data = "此工號查無Bpm使用者資訊!"
                });
                #endregion
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //PostFormToBpm
        public static string PostFormToBpm(string token, string proId, string memId, string rolId, string startMethod, object artInsAppData, string BpmServerPath)
        {
            try
            {
                string apiUrl = BpmServerPath + "/WebServiceStart/services/CreateProcess";

                #region //Request
                JObject jsonData = JObject.FromObject(new
                {
                    proId,
                    memId,
                    rolId,
                    startMethod,
                    authKey = token,
                    artInsAppData
                });
                #endregion

                UTF8Encoding dataEncode = new UTF8Encoding();

                using (HttpClient httpClient = new HttpClient())
                {
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                    {
                        httpRequestMessage.Content = new StringContent(JsonConvert.SerializeObject(jsonData), dataEncode, "application/json");

                        using (HttpResponseMessage httpResponseMessage = httpClient.SendAsync(httpRequestMessage).GetAwaiter().GetResult())
                        {
                            if (httpResponseMessage.IsSuccessStatusCode)
                            {
                                return httpResponseMessage.Content.ReadAsStringAsync().Result.ToString();
                            }
                            else
                            {
                                return "伺服器連線異常";
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        #endregion
    }
}
