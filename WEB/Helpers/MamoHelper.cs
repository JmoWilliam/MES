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
using System.Web;
using System.Configuration;

namespace Helpers
{
    public class MamoHelper
    {
        private string apiMainUrl = "";
        //private string apiMainUrl = "https://bm.zy-tech.com.tw/";
        //private string apiMainUrl = "http://192.168.134.33:16668/";
        //private string apiMainUrl = "http://192.168.134.53:16668/";
        public JObject jsonResponse = new JObject();
        public string dataRequest = "";
        public static DynamicParameters dynamicParameters = new DynamicParameters();
        public static Logger logger = LogManager.GetCurrentClassLogger();
        private string SecretKey = "";
        public string MainConnectionStrings = "";
        public string sql = "";

        public MamoHelper()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            GetServerUrl();
        }

        #region //GetServerUrl
        private void GetServerUrl()
        {
            HttpRequest request = HttpContext.Current.Request;
            Uri serverUrl = request.Url;
            string scheme = serverUrl.Scheme;
            string ipAddress = serverUrl.Host;
            int port = serverUrl.Port;
            string portServer = "";
            if (port > 0)
            {
                portServer = ":" + port.ToString();
            }
            apiMainUrl = scheme + "://" + ipAddress + portServer + "/";
        }
        #endregion

        #region //團隊
        #region //CreateTeams 建立團隊 
        public string CreateTeams(string Company, int UserId, string TeamName, string Remark)
        {
            try
            {
                string apiUrl = Path.Combine(apiMainUrl, "api/MAMO/CreateTeams");

                #region //取得金鑰
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.KeyText 
                            FROM BAS.ApiKey a 
                            INNER JOIN BAS.Company b ON a.CompanyId = b.CompanyId
                            WHERE b.CompanyNo = @CompanyNo
                            AND a.Purpose = 'MAMO API'";
                    dynamicParameters.Add("CompanyNo", Company);

                    var ApiKeyResult = sqlConnection.Query(sql, dynamicParameters);
                    if (ApiKeyResult.Count() <= 0) throw new SystemException("金鑰錯誤!!");

                    foreach (var item in ApiKeyResult)
                    {
                        SecretKey = item.KeyText;
                    }
                }
                #endregion

                #region //Request
                using (HttpClient client = new HttpClient())
                {
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                    {
                        MultipartFormDataContent multipart = new MultipartFormDataContent();
                        multipart.Add(new StringContent(Company), "Company");
                        multipart.Add(new StringContent(UserId.ToString()), "UserId");
                        multipart.Add(new StringContent(SecretKey), "SecretKey");
                        multipart.Add(new StringContent(TeamName), "TeamName");
                        multipart.Add(new StringContent(Remark), "Remark");
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

        #region //DeleteTeams 刪除團隊
        public string DeleteTeams(string Company, int UserId, int TeamId)
        {
            try
            {
                string apiUrl = Path.Combine(apiMainUrl, "api/MAMO/DeleteTeams");

                #region //取得金鑰
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.KeyText 
                            FROM BAS.ApiKey a 
                            INNER JOIN BAS.Company b ON a.CompanyId = b.CompanyId
                            WHERE b.CompanyNo = @CompanyNo
                            AND a.Purpose = 'MAMO API'";
                    dynamicParameters.Add("CompanyNo", Company);

                    var ApiKeyResult = sqlConnection.Query(sql, dynamicParameters);
                    if (ApiKeyResult.Count() <= 0) throw new SystemException("金鑰錯誤!!");

                    foreach (var item in ApiKeyResult)
                    {
                        SecretKey = item.KeyText;
                    }
                }
                #endregion

                #region //Request
                using (HttpClient client = new HttpClient())
                {
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                    {
                        MultipartFormDataContent multipart = new MultipartFormDataContent();
                        multipart.Add(new StringContent(Company), "Company");
                        multipart.Add(new StringContent(UserId.ToString()), "UserId");
                        multipart.Add(new StringContent(SecretKey), "SecretKey");
                        multipart.Add(new StringContent(TeamId.ToString()), "TeamId");
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

        #region //RestoreTeams 重啟團隊
        public string RestoreTeams(string Company, int UserId, int TeamId)
        {
            try
            {
                string apiUrl = Path.Combine(apiMainUrl, "api/MAMO/RestoreTeams");

                #region //取得金鑰
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.KeyText 
                            FROM BAS.ApiKey a 
                            INNER JOIN BAS.Company b ON a.CompanyId = b.CompanyId
                            WHERE b.CompanyNo = @CompanyNo
                            AND a.Purpose = 'MAMO API'";
                    dynamicParameters.Add("CompanyNo", Company);

                    var ApiKeyResult = sqlConnection.Query(sql, dynamicParameters);
                    if (ApiKeyResult.Count() <= 0) throw new SystemException("金鑰錯誤!!");

                    foreach (var item in ApiKeyResult)
                    {
                        SecretKey = item.KeyText;
                    }
                }
                #endregion

                #region //Request
                using (HttpClient client = new HttpClient())
                {
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                    {
                        MultipartFormDataContent multipart = new MultipartFormDataContent();
                        multipart.Add(new StringContent(Company), "Company");
                        multipart.Add(new StringContent(UserId.ToString()), "UserId");
                        multipart.Add(new StringContent(SecretKey), "SecretKey");
                        multipart.Add(new StringContent(TeamId.ToString()), "TeamId");
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

        #region //AddTeamMembers 新增團隊成員
        public string AddTeamMembers(string Company, int UserId, int TeamId, List<string> UserNo)
        {
            try
            {
                string apiUrl = Path.Combine(apiMainUrl, "api/MAMO/AddTeamMembers");

                #region //取得金鑰
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.KeyText 
                            FROM BAS.ApiKey a 
                            INNER JOIN BAS.Company b ON a.CompanyId = b.CompanyId
                            WHERE b.CompanyNo = @CompanyNo
                            AND a.Purpose = 'MAMO API'";
                    dynamicParameters.Add("CompanyNo", Company);

                    var ApiKeyResult = sqlConnection.Query(sql, dynamicParameters);
                    if (ApiKeyResult.Count() <= 0) throw new SystemException("金鑰錯誤!!");

                    foreach (var item in ApiKeyResult)
                    {
                        SecretKey = item.KeyText;
                    }
                }
                #endregion

                #region //Request
                using (HttpClient client = new HttpClient())
                {
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                    {
                        MultipartFormDataContent multipart = new MultipartFormDataContent();
                        multipart.Add(new StringContent(Company), "Company");
                        multipart.Add(new StringContent(UserId.ToString()), "UserId");
                        multipart.Add(new StringContent(SecretKey), "SecretKey");
                        multipart.Add(new StringContent(TeamId.ToString()), "TeamId");

                        for (int i = 0; i < UserNo.Count(); i++)
                        {
                            multipart.Add(new StringContent(UserNo[i]), "UserNo");
                        }

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

        #region //DeleteTeamMembers 刪除團隊成員
        public string DeleteTeamMembers(string Company, int UserId, int TeamId, List<string> UserNo)
        {
            try
            {
                string apiUrl = Path.Combine(apiMainUrl, "api/MAMO/DeleteTeamMembers");

                #region //取得金鑰
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.KeyText 
                            FROM BAS.ApiKey a 
                            INNER JOIN BAS.Company b ON a.CompanyId = b.CompanyId
                            WHERE b.CompanyNo = @CompanyNo
                            AND a.Purpose = 'MAMO API'";
                    dynamicParameters.Add("CompanyNo", Company);

                    var ApiKeyResult = sqlConnection.Query(sql, dynamicParameters);
                    if (ApiKeyResult.Count() <= 0) throw new SystemException("金鑰錯誤!!");

                    foreach (var item in ApiKeyResult)
                    {
                        SecretKey = item.KeyText;
                    }
                }
                #endregion

                #region //Request
                using (HttpClient client = new HttpClient())
                {
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                    {
                        MultipartFormDataContent multipart = new MultipartFormDataContent();
                        multipart.Add(new StringContent(Company), "Company");
                        multipart.Add(new StringContent(UserId.ToString()), "UserId");
                        multipart.Add(new StringContent(SecretKey), "SecretKey");
                        multipart.Add(new StringContent(TeamId.ToString()), "TeamId");

                        for (int i = 0; i < UserNo.Count(); i++)
                        {
                            multipart.Add(new StringContent(UserNo[i]), "UserNo");
                        }

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
        #endregion

        #region //頻道
        #region //CreateChannels 建立頻道 
        public string CreateChannels(string Company, int UserId, int TeamId, string ChannelName, string Remark)
        {
            try
            {
                string apiUrl = Path.Combine(apiMainUrl, "api/MAMO/CreateChannels");

                #region //取得金鑰
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.KeyText 
                            FROM BAS.ApiKey a 
                            INNER JOIN BAS.Company b ON a.CompanyId = b.CompanyId
                            WHERE b.CompanyNo = @CompanyNo
                            AND a.Purpose = 'MAMO API'";
                    dynamicParameters.Add("CompanyNo", Company);

                    var ApiKeyResult = sqlConnection.Query(sql, dynamicParameters);
                    if (ApiKeyResult.Count() <= 0) throw new SystemException("金鑰錯誤!!");

                    foreach (var item in ApiKeyResult)
                    {
                        SecretKey = item.KeyText;
                    }
                }
                #endregion

                #region //Request
                using (HttpClient client = new HttpClient())
                {
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                    {
                        MultipartFormDataContent multipart = new MultipartFormDataContent();
                        multipart.Add(new StringContent(Company), "Company");
                        multipart.Add(new StringContent(UserId.ToString()), "UserId");
                        multipart.Add(new StringContent(SecretKey), "SecretKey");
                        multipart.Add(new StringContent(TeamId.ToString()), "TeamId");
                        multipart.Add(new StringContent(ChannelName), "ChannelName");
                        multipart.Add(new StringContent(Remark), "Remark");
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

        #region //DeleteChannels 刪除頻道 
        public string DeleteChannels(string Company, int UserId, int ChannelId)
        {
            try
            {
                string apiUrl = Path.Combine(apiMainUrl, "api/MAMO/DeleteChannels");

                #region //取得金鑰
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.KeyText 
                            FROM BAS.ApiKey a 
                            INNER JOIN BAS.Company b ON a.CompanyId = b.CompanyId
                            WHERE b.CompanyNo = @CompanyNo
                            AND a.Purpose = 'MAMO API'";
                    dynamicParameters.Add("CompanyNo", Company);

                    var ApiKeyResult = sqlConnection.Query(sql, dynamicParameters);
                    if (ApiKeyResult.Count() <= 0) throw new SystemException("金鑰錯誤!!");

                    foreach (var item in ApiKeyResult)
                    {
                        SecretKey = item.KeyText;
                    }
                }
                #endregion

                #region //Request
                using (HttpClient client = new HttpClient())
                {
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                    {
                        MultipartFormDataContent multipart = new MultipartFormDataContent();
                        multipart.Add(new StringContent(Company), "Company");
                        multipart.Add(new StringContent(UserId.ToString()), "UserId");
                        multipart.Add(new StringContent(SecretKey), "SecretKey");
                        multipart.Add(new StringContent(ChannelId.ToString()), "ChannelId");
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

        #region //RestoreChannels 重啟頻道 
        public string RestoreChannels(string Company, int UserId, int ChannelId)
        {
            try
            {
                string apiUrl = Path.Combine(apiMainUrl, "api/MAMO/RestoreChannels");

                #region //取得金鑰
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.KeyText 
                            FROM BAS.ApiKey a 
                            INNER JOIN BAS.Company b ON a.CompanyId = b.CompanyId
                            WHERE b.CompanyNo = @CompanyNo
                            AND a.Purpose = 'MAMO API'";
                    dynamicParameters.Add("CompanyNo", Company);

                    var ApiKeyResult = sqlConnection.Query(sql, dynamicParameters);
                    if (ApiKeyResult.Count() <= 0) throw new SystemException("金鑰錯誤!!");

                    foreach (var item in ApiKeyResult)
                    {
                        SecretKey = item.KeyText;
                    }
                }
                #endregion

                #region //Request
                using (HttpClient client = new HttpClient())
                {
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                    {
                        MultipartFormDataContent multipart = new MultipartFormDataContent();
                        multipart.Add(new StringContent(Company), "Company");
                        multipart.Add(new StringContent(UserId.ToString()), "UserId");
                        multipart.Add(new StringContent(SecretKey), "SecretKey");
                        multipart.Add(new StringContent(ChannelId.ToString()), "ChannelId");
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

        #region //AddChannelMembers 新增頻道成員 
        public string AddChannelMembers(string Company, int UserId, int ChannelId, List<string> UserNo)
        {
            try
            {
                string apiUrl = Path.Combine(apiMainUrl, "api/MAMO/AddChannelMembers");

                #region //取得金鑰
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.KeyText 
                            FROM BAS.ApiKey a 
                            INNER JOIN BAS.Company b ON a.CompanyId = b.CompanyId
                            WHERE b.CompanyNo = @CompanyNo
                            AND a.Purpose = 'MAMO API'";
                    dynamicParameters.Add("CompanyNo", Company);

                    var ApiKeyResult = sqlConnection.Query(sql, dynamicParameters);
                    if (ApiKeyResult.Count() <= 0) throw new SystemException("金鑰錯誤!!");

                    foreach (var item in ApiKeyResult)
                    {
                        SecretKey = item.KeyText;
                    }
                }
                #endregion

                #region //Request
                using (HttpClient client = new HttpClient())
                {
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                    {
                        MultipartFormDataContent multipart = new MultipartFormDataContent();
                        multipart.Add(new StringContent(Company), "Company");
                        multipart.Add(new StringContent(UserId.ToString()), "UserId");
                        multipart.Add(new StringContent(SecretKey), "SecretKey");
                        multipart.Add(new StringContent(ChannelId.ToString()), "ChannelId");

                        for (int i = 0; i < UserNo.Count(); i++)
                        {
                            multipart.Add(new StringContent(UserNo[i]), "UserNo");
                        }

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

        #region //DeleteChannelMembers 刪除頻道成員 
        public string DeleteChannelMembers(string Company, int UserId, int ChannelId, List<string> UserNo)
        {
            try
            {
                string apiUrl = Path.Combine(apiMainUrl, "api/MAMO/DeleteChannelMembers");

                #region //取得金鑰
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.KeyText 
                            FROM BAS.ApiKey a 
                            INNER JOIN BAS.Company b ON a.CompanyId = b.CompanyId
                            WHERE b.CompanyNo = @CompanyNo
                            AND a.Purpose = 'MAMO API'";
                    dynamicParameters.Add("CompanyNo", Company);

                    var ApiKeyResult = sqlConnection.Query(sql, dynamicParameters);
                    if (ApiKeyResult.Count() <= 0) throw new SystemException("金鑰錯誤!!");

                    foreach (var item in ApiKeyResult)
                    {
                        SecretKey = item.KeyText;
                    }
                }
                #endregion

                #region //Request
                using (HttpClient client = new HttpClient())
                {
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                    {
                        MultipartFormDataContent multipart = new MultipartFormDataContent();
                        multipart.Add(new StringContent(Company), "Company");
                        multipart.Add(new StringContent(UserId.ToString()), "UserId");
                        multipart.Add(new StringContent(SecretKey), "SecretKey");
                        multipart.Add(new StringContent(ChannelId.ToString()), "ChannelId");

                        for (int i = 0; i < UserNo.Count(); i++)
                        {
                            multipart.Add(new StringContent(UserNo[i]), "UserNo");
                        }

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

        #region //SendMessage 寄送訊息(單人/頻道)
        public string SendMessage(string Company, int UserId, string PushType, string SendId, string Content, List<string> Tags, List<int> Files)
        {
            try
            {
                string apiUrl = Path.Combine(apiMainUrl, "api/MAMO/SendMessage");

                #region //取得金鑰
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.KeyText 
                            FROM BAS.ApiKey a 
                            INNER JOIN BAS.Company b ON a.CompanyId = b.CompanyId
                            WHERE b.CompanyNo = @CompanyNo
                            AND a.Purpose = 'MAMO API'";
                    dynamicParameters.Add("CompanyNo", Company);

                    var ApiKeyResult = sqlConnection.Query(sql, dynamicParameters);
                    if (ApiKeyResult.Count() <= 0) throw new SystemException("金鑰錯誤!!");

                    foreach (var item in ApiKeyResult)
                    {
                        SecretKey = item.KeyText;
                    }
                }
                #endregion

                #region //Request
                using (HttpClient client = new HttpClient())
                {
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                    {
                        MultipartFormDataContent multipart = new MultipartFormDataContent();
                        multipart.Add(new StringContent(Company), "Company");
                        multipart.Add(new StringContent(UserId.ToString()), "UserId");
                        multipart.Add(new StringContent(SecretKey), "SecretKey");
                        multipart.Add(new StringContent(PushType), "PushType");
                        multipart.Add(new StringContent(SendId), "SendId");
                        multipart.Add(new StringContent(Content), "Content");

                        for (int i = 0; i < Tags.Count(); i++)
                        {
                            multipart.Add(new StringContent(Tags[i]), "Tags");
                        }

                        for (int i = 0; i < Files.Count(); i++)
                        {
                            multipart.Add(new StringContent(Files[i].ToString()), "Files");
                        }

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
        #endregion
    }
}
