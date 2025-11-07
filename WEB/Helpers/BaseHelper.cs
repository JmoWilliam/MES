using Dapper;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web;

namespace Helpers
{
    public class BaseHelper
    {
        public static string company = @"JMO";
        public static string secretKey =  @"FA65D160076DA566CBFC342852580C27DB49415B94CF07CA47B4B9BB194AE8B3";
        //public static string domainUrl = "https://192.168.20.112:26668/";
        public static string domainUrl = "https://bm.zy-tech.com.tw/";

        /// <summary>
        /// 處理Data Access回傳判斷
        /// </summary>
        /// <param name="data">輸入json字串</param>
        /// <returns>回傳json物件</returns>
        public static JObject DAResponse(string data)
        {
            JObject jsonResponse = new JObject();

            if (JObject.Parse(data)["status"].ToString() == "error")
            {
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = JObject.Parse(data)["msg"].ToString()
                });
            }
            else if (JObject.Parse(data)["status"].ToString() == "errorForDA")
            {
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = JObject.Parse(data)["msg"].ToString()
                });
            }
            else
            {
                if (JObject.Parse(data)["data"] != null)
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "OK",
                        result = JObject.Parse(data)["data"]
                    });
                }
                else
                {
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = JObject.Parse(data)["msg"].ToString()
                    });
                }
            }

            return jsonResponse;
        }

        /// <summary>
        /// AES解密
        /// </summary>
        /// <param name="phrase"></param>
        /// <param name="key"></param>
        /// <param name="iv"></param>
        /// <returns></returns>
        public static string AESDecrypt(string phrase, string key, string iv)
        {
            var encryptBytes = Convert.FromBase64String(phrase);
            var aes = Aes.Create();
            aes.Mode = CipherMode.CBC;
            aes.Padding = PaddingMode.PKCS7;
            aes.Key = Encoding.UTF8.GetBytes(key);
            aes.IV = Encoding.UTF8.GetBytes(iv);
            var transform = aes.CreateDecryptor();

            try
            {
                return Encoding.UTF8.GetString(transform.TransformFinalBlock(encryptBytes, 0, encryptBytes.Length));
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// AES加密
        /// </summary>
        /// <param name="phrase"></param>
        /// <param name="key"></param>
        /// <param name="iv"></param>
        /// <returns></returns>
        public static string AESEncrypt(string phrase, string key, string iv)
        {
            var sourceBytes = Encoding.UTF8.GetBytes(phrase);
            var aes = Aes.Create();
            aes.Mode = CipherMode.CBC;
            aes.Padding = PaddingMode.PKCS7;
            aes.Key = Encoding.UTF8.GetBytes(key);
            aes.IV = Encoding.UTF8.GetBytes(iv);
            var transform = aes.CreateEncryptor();

            return Convert.ToBase64String(transform.TransformFinalBlock(sourceBytes, 0, sourceBytes.Length));
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="inputArray"></param>
        /// <returns></returns>
        public static string byteArrayToString(byte[] inputArray)
        {
            StringBuilder output = new StringBuilder("");
            for (int i = 0; i < inputArray.Length; i++)
            {
                output.Append(inputArray[i].ToString("X2"));
            }

            return output.ToString();
        }

        /// <summary>
        /// Sha1加密
        /// </summary>
        /// <param name="phrase"></param>
        /// <returns></returns>
        public static string Sha1Encrypt(string phrase)
        {
            UTF8Encoding encoder = new UTF8Encoding();
            SHA1Managed sha1hasher = new SHA1Managed();
            byte[] hashedDataBytes = sha1hasher.ComputeHash(encoder.GetBytes(phrase));

            return byteArrayToString(hashedDataBytes);
        }

        /// <summary>
        /// Sha256加密
        /// </summary>
        /// <param name="phrase"></param>
        /// <returns></returns>
        public static string Sha256Encrypt(string phrase)
        {
            UTF8Encoding encoder = new UTF8Encoding();
            SHA256Managed sha256hasher = new SHA256Managed();
            byte[] hashedDataBytes = sha256hasher.ComputeHash(encoder.GetBytes(phrase));

            return byteArrayToString(hashedDataBytes);
        }

        /// <summary>
        /// 獲取使用者IP
        /// </summary>
        /// <returns></returns>
        public static string ClientIP()
        {
            if (!string.IsNullOrEmpty(HttpContext.Current.Request.ServerVariables["HTTP_X_FORWARDED_FOR"]))
            {
                return HttpContext.Current.Request.ServerVariables["HTTP_X_FORWARDED_FOR"];
            }
            else
            {
                return HttpContext.Current.Request.ServerVariables["REMOTE_ADDR"];
            }
        }

        /// <summary>
        /// 獲取使用者電腦名稱
        /// </summary>
        /// <returns></returns>
        public static string ClientComputer()
        {
            try
            {
                var hostEntry = Dns.GetHostEntry(ClientIP());

                return hostEntry.HostName;
            }
            catch (Exception)
            {
                return "Not support device.";
            }
        }

        /// <summary>
        /// 獲取使用者連線狀態
        /// </summary>
        /// <returns></returns>
        public static string ClientLinkType()
        {
            IPAddress clientIP = IPAddress.Parse(ClientIP());
            bool isInnerIp = false;

            IPHelper iPHelper = new IPHelper(IPAddress.Parse("10.0.0.0"), IPAddress.Parse("10.255.255.255"));
            isInnerIp = iPHelper.IsInRange(clientIP);
            if (!isInnerIp)
            {
                iPHelper = new IPHelper(IPAddress.Parse("172.16.0.0"), IPAddress.Parse("172.31.255.255"));
                isInnerIp = iPHelper.IsInRange(clientIP);
            }
            if (!isInnerIp)
            {
                iPHelper = new IPHelper(IPAddress.Parse("192.168.0.0"), IPAddress.Parse("192.168.255.255"));
                isInnerIp = iPHelper.IsInRange(clientIP);
            }

            return isInnerIp ? "廠內連線" : "外部連線";
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sql"></param>
        /// <param name="dynamicParameters"></param>
        /// <param name="parameterName"></param>
        /// <param name="conditionSql"></param>
        /// <param name="conditionValue"></param>
        /// <param name="standardValue"></param>
        public static void SqlParameter<T>(ref string sql, ref DynamicParameters dynamicParameters, string parameterName, string conditionSql, T conditionValue, T standardValue = default(T))
        {
            int tempInt = 0;
            double tempDouble = 0.0;
            DateTime tempDateTime = default(DateTime);

            if (typeof(T) == typeof(int))
            {
                if (int.TryParse(conditionValue.ToString(), out tempInt))
                {
                    int standard = int.TryParse(standardValue.ToString(), out tempInt) ? Convert.ToInt32(standardValue) : 0;

                    if (Convert.ToInt32(conditionValue) > standard)
                    {
                        sql += conditionSql;
                        dynamicParameters.Add(parameterName, Convert.ToInt32(conditionValue));
                    }
                }
            }
            else if (typeof(T) == typeof(double))
            {
                if (double.TryParse(conditionValue.ToString(), out tempDouble))
                {
                    double standard = double.TryParse(standardValue.ToString(), out tempDouble) ? Convert.ToDouble(standardValue) : 0;

                    if (Convert.ToDouble(conditionValue) > standard)
                    {
                        sql += conditionSql;
                        dynamicParameters.Add(parameterName, Convert.ToDouble(conditionValue));
                    }
                }
            }
            else if (typeof(T) == typeof(string))
            {
                if (standardValue != null)
                {
                    if (DateTime.TryParse(standardValue.ToString() ?? "", out tempDateTime))
                    {
                        sql += conditionSql;
                        dynamicParameters.Add(parameterName, Convert.ToDateTime(conditionValue).ToString("yyyy-MM-dd HH:mm:ss"));
                    }
                }
                else
                {
                    if (conditionValue.ToString().Length > 0)
                    {
                        sql += conditionSql;
                        dynamicParameters.Add(parameterName, conditionValue);
                    }
                }
            }
            else if (typeof(T) == typeof(int[]))
            {
                sql += conditionSql;
                dynamicParameters.Add(parameterName, conditionValue);
            }
            else if (typeof(T) == typeof(double[]))
            {
                sql += conditionSql;
                dynamicParameters.Add(parameterName, conditionValue);
            }
            else if (typeof(T) == typeof(string[]))
            {
                sql += conditionSql;
                dynamicParameters.Add(parameterName, conditionValue);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sqlConnection"></param>
        /// <param name="parameters"></param>
        /// <param name="primaryKey"></param>
        /// <param name="queryColumns"></param>
        /// <param name="queryTable"></param>
        /// <param name="orderBy"></param>
        /// <param name="pageIndex"></param>
        /// <param name="pageSize"></param>
        /// <param name="declarePart"></param>
        /// <returns></returns>
        public static IEnumerable<dynamic> SqlQuery(SqlConnection sqlConnection, DynamicParameters parameters, SqlQuery sqlQuery)
        {
            string sqlTemplate =
              @"{0}
                WITH K
                AS (
                    SELECT {11} {1} {2}
                    {4}
                    WHERE 1=1
                    {6}
                    {7}
                    {8}
                )
                SELECT K.*, [Count].TotalCount
                {10}
                {5}
                INNER JOIN K ON {3}
                CROSS APPLY (
                    SELECT COUNT(1) TotalCount
                    {4}
                    WHERE 1=1
                    {6}
                ) [Count]
                {9}";

            string sqlOrderBy = "", sqlPagination = "";
            if (sqlQuery.orderBy.Length > 0)
            {
                sqlOrderBy = string.Format(@" ORDER BY {0} ", sqlQuery.orderBy);

                if (sqlQuery.pageSize > 0)
                {
                    sqlPagination = sqlOrderBy;
                    sqlPagination += @"OFFSET @PageSize * (@PageIndex - 1) ROWS
                                    FETCH NEXT @PageSize ROWS ONLY";
                    parameters.Add("PageIndex", Math.Max(sqlQuery.pageIndex, 0));
                    parameters.Add("PageSize", sqlQuery.pageSize);
                }
            }

            string sqlJoin = "";
            string[] keyArray = sqlQuery.mainKey.Split(',');

            string sqlAlias = "";
            string[] aliasKeyArray = sqlQuery.aliasKey.Split(',');

            if (aliasKeyArray.Length > 1)
            {
                for (int i = 0; i < aliasKeyArray.Length; i++)
                {
                    if (i != 0) sqlAlias += ",";
                    sqlAlias += keyArray[i] + " " + aliasKeyArray[i];
                }

                sqlQuery.mainKey = sqlAlias;
            }

            for (int i = 0; i < keyArray.Length; i++)
            {
                if (i != 0) sqlJoin += " AND ";

                if (aliasKeyArray.Length > 1)
                {
                    sqlJoin += keyArray[i] + " = K." + aliasKeyArray[i];
                }
                else
                {
                    sqlJoin += keyArray[i] + " = K." + keyArray[i].Split('.')[1];
                }
            }

            object[] args = new object[] {
                sqlQuery.declarePart, //0
                sqlQuery.mainKey, //1，必填
                sqlQuery.auxKey, //2，必填
                sqlJoin, //3
                sqlQuery.auxTables.Length > 0 ? sqlQuery.auxTables : sqlQuery.mainTables, //4，必填
                sqlQuery.mainTables, //5，必填
                sqlQuery.conditions, //6，必填
                sqlQuery.groupBy, //7
                sqlPagination, //8
                sqlOrderBy, //9
                sqlQuery.columns, //10，必填
                sqlQuery.distinct ? "DISTINCT" : "" //11
            };

            sqlTemplate = string.Format(sqlTemplate, args);

            return sqlConnection.Query(sqlTemplate, parameters);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sqlConnection"></param>
        /// <param name="parameters"></param>
        /// <param name="primaryKey"></param>
        /// <param name="queryColumns"></param>
        /// <param name="queryTable"></param>
        /// <param name="orderBy"></param>
        /// <param name="pageIndex"></param>
        /// <param name="pageSize"></param>
        /// <param name="declarePart"></param>
        /// <returns></returns>
        public static async Task<IEnumerable<dynamic>> SqlQueryAsync(SqlConnection sqlConnection, DynamicParameters parameters, SqlQuery sqlQuery)
        {
            string sqlTemplate =
              @"{0}
                WITH K
                AS (
                    SELECT {11} {1} {2}
                    {4}
                    WHERE 1=1
                    {6}
                    {7}
                    {8}
                )
                SELECT K.*, [Count].TotalCount
                {10}
                {5}
                INNER JOIN K ON {3}
                CROSS APPLY (
                    SELECT COUNT(1) TotalCount
                    {4}
                    WHERE 1=1
                    {6}
                ) [Count]
                {9}";

            string sqlOrderBy = "", sqlPagination = "";
            if (sqlQuery.orderBy.Length > 0)
            {
                sqlOrderBy = string.Format(@" ORDER BY {0} ", sqlQuery.orderBy);

                if (sqlQuery.pageSize > 0)
                {
                    sqlPagination = sqlOrderBy;
                    sqlPagination += @"OFFSET @PageSize * (@PageIndex - 1) ROWS
                                    FETCH NEXT @PageSize ROWS ONLY";
                    parameters.Add("PageIndex", Math.Max(sqlQuery.pageIndex, 0));
                    parameters.Add("PageSize", sqlQuery.pageSize);
                }
            }

            string sqlJoin = "";
            string[] keyArray = sqlQuery.mainKey.Split(',');

            string sqlAlias = "";
            string[] aliasKeyArray = sqlQuery.aliasKey.Split(',');

            if (aliasKeyArray.Length > 1)
            {
                for (int i = 0; i < aliasKeyArray.Length; i++)
                {
                    if (i != 0) sqlAlias += ",";
                    sqlAlias += keyArray[i] + " " + aliasKeyArray[i];
                }

                sqlQuery.mainKey = sqlAlias;
            }

            for (int i = 0; i < keyArray.Length; i++)
            {
                if (i != 0) sqlJoin += " AND ";

                if (aliasKeyArray.Length > 1)
                {
                    sqlJoin += keyArray[i] + " = K." + aliasKeyArray[i];
                }
                else
                {
                    sqlJoin += keyArray[i] + " = K." + keyArray[i].Split('.')[1];
                }
            }

            object[] args = new object[] {
                sqlQuery.declarePart, //0
                sqlQuery.mainKey, //1，必填
                sqlQuery.auxKey, //2，必填
                sqlJoin, //3
                sqlQuery.auxTables.Length > 0 ? sqlQuery.auxTables : sqlQuery.mainTables, //4，必填
                sqlQuery.mainTables, //5，必填
                sqlQuery.conditions, //6，必填
                sqlQuery.groupBy, //7
                sqlPagination, //8
                sqlOrderBy, //9
                sqlQuery.columns, //10，必填
                sqlQuery.distinct ? "DISTINCT" : "" //11
            };

            sqlTemplate = string.Format(sqlTemplate, args);

            return await sqlConnection.QueryAsync(sqlTemplate, parameters);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sqlConnection"></param>
        /// <param name="parameters"></param>
        /// <param name="sqlQuery"></param>
        /// <returns></returns>
        public static List<T> SqlQuery<T>(SqlConnection sqlConnection, DynamicParameters parameters, SqlQuery sqlQuery)
        {
            string sqlTemplate =
              @"{0}
                WITH K
                AS (
                    SELECT {11} {1} {2}
                    {4}
                    WHERE 1=1
                    {6}
                    {7}
                    {8}
                )
                SELECT {11} K.*, [Count].TotalCount
                {10}
                {5}
                INNER JOIN K ON {3}
                CROSS APPLY (
                    SELECT COUNT(1) TotalCount
                    {4}
                    WHERE 1=1
                    {6}
                ) [Count]
                {9}";

            string sqlOrderBy = "", sqlPagination = "";
            if (sqlQuery.orderBy.Length > 0)
            {
                sqlOrderBy = string.Format(@" ORDER BY {0} ", sqlQuery.orderBy);

                if (sqlQuery.pageSize > 0)
                {
                    sqlPagination = sqlOrderBy;
                    sqlPagination += @"OFFSET @PageSize * (@PageIndex - 1) ROWS
                                    FETCH NEXT @PageSize ROWS ONLY";
                    parameters.Add("PageIndex", Math.Max(sqlQuery.pageIndex, 0));
                    parameters.Add("PageSize", sqlQuery.pageSize);
                }
            }

            string sqlJoin = "";
            string[] keyArray = sqlQuery.mainKey.Split(',');

            string sqlAlias = "";
            string[] aliasKeyArray = sqlQuery.aliasKey.Split(',');

            if (aliasKeyArray.Length > 1)
            {
                for (int i = 0; i < aliasKeyArray.Length; i++)
                {
                    if (i != 0) sqlAlias += ",";
                    sqlAlias += keyArray[i] + " " + aliasKeyArray[i];
                }

                sqlQuery.mainKey = sqlAlias;
            }

            for (int i = 0; i < keyArray.Length; i++)
            {
                if (i != 0) sqlJoin += " AND ";

                if (aliasKeyArray.Length > 1)
                {
                    sqlJoin += keyArray[i] + " = K." + aliasKeyArray[i];
                }
                else
                {
                    sqlJoin += keyArray[i] + " = K." + keyArray[i].Split('.')[1];
                }
            }

            object[] args = new object[] {
                sqlQuery.declarePart, //0
                sqlQuery.mainKey, //1，必填
                sqlQuery.auxKey, //2，必填
                sqlJoin, //3
                sqlQuery.auxTables.Length > 0 ? sqlQuery.auxTables : sqlQuery.mainTables, //4，必填
                sqlQuery.mainTables, //5，必填
                sqlQuery.conditions, //6，必填
                sqlQuery.groupBy, //7
                sqlPagination, //8
                sqlOrderBy, //9
                sqlQuery.columns, //10，必填
                sqlQuery.distinct ? "DISTINCT" : "" //11
            };

            sqlTemplate = string.Format(sqlTemplate, args);

            return sqlConnection.Query<T>(sqlTemplate, parameters).ToList();
        }

        public static string PostWebRequest(string postUrl)
        {
            using (HttpClient httpClient = new HttpClient())
            {
                using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, postUrl))
                {
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

        public static string GetWebRequest(string postUrl)
        {
            using (HttpClient httpClient = new HttpClient())
            {
                using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Get, postUrl))
                {
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

        /// <summary>
        /// 提供後端將特定資料POST到特定URL，，並取得回傳值。(form格式)
        /// </summary>
        /// <param name="postUrl">Post網址</param>
        /// <param name="postData">Post資料</param>
        /// <returns>返回訊息</returns>
        public static string PostWebRequest(string postUrl, List<KeyValuePair<string, string>> postData)
        {
            using (HttpClient httpClient = new HttpClient())
            {
                httpClient.Timeout = TimeSpan.FromMinutes(60);

                using (HttpResponseMessage httpResponseMessage = httpClient.PostAsync(postUrl, new FormUrlEncodedContent(postData)).GetAwaiter().GetResult())
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

        /// <summary>
        /// 提供後端將特定資料POST到特定URL，，並取得回傳值。(json格式)
        /// </summary>
        /// <param name="postUrl">Post網址</param>
        /// <param name="postData">Post資料</param>
        /// <param name="dataEncode">資料編碼</param>
        /// <returns>返回訊息</returns>
        public static string PostWebRequest(string postUrl, string postData, Encoding dataEncode)
        {
            using (HttpClient httpClient = new HttpClient())
            {
                using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, postUrl))
                {
                    httpRequestMessage.Content = new StringContent(postData, dataEncode, "application/json");

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

        /// <summary>
        /// 提供後端將特定檔案POST到特定URL，並取得回傳值。
        /// </summary>
        /// <param name="postUrl">Post網址</param>
        /// <param name="postFileName">檔案名稱</param>
        /// <param name="postFileContent">檔案內容</param>
        /// <param name="postFileExtension">檔案副檔名</param>
        /// <param name="fileKeyName">檔案欄位名稱</param>
        /// <param name="nameValueCollection"></param>
        /// <returns></returns>
        public static string PostWebRequestAsync(string postUrl, string postFileName, byte[] postFileContent, string postFileExtension, string fileKeyName, NameValueCollection nameValueCollection)
        {
            try
            {
                string boundary = "---------------------------" + DateTime.Now.Ticks.ToString("x");
                byte[] boundarybytes = Encoding.ASCII.GetBytes("\r\n--" + boundary + "\r\n");

                HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(postUrl);
                httpWebRequest.ContentType = "multipart/form-data; boundary=" + boundary;
                httpWebRequest.Method = "POST";
                httpWebRequest.KeepAlive = true;
                httpWebRequest.Credentials = CredentialCache.DefaultCredentials;

                using (Stream stream = httpWebRequest.GetRequestStream())
                {
                    string formdataTemplate = "Content-Disposition: form-data; name=\"{0}\"\r\n\r\n{1}";
                    foreach (string key in nameValueCollection.Keys)
                    {
                        stream.Write(boundarybytes, 0, boundarybytes.Length);
                        string formitem = string.Format(formdataTemplate, key, nameValueCollection[key]);
                        byte[] formitembytes = Encoding.UTF8.GetBytes(formitem);
                        stream.Write(formitembytes, 0, formitembytes.Length);
                    }
                    stream.Write(boundarybytes, 0, boundarybytes.Length);

                    string headerTemplate = "Content-Disposition: form-data; name=\"{0}\"; filename=\"{1}\"\r\nContent-Type: {2}\r\n\r\n";
                    string header = string.Format(headerTemplate, fileKeyName, postFileName, FileHelper.GetMime("." + postFileExtension));
                    byte[] headerbytes = Encoding.UTF8.GetBytes(header);
                    stream.Write(headerbytes, 0, headerbytes.Length);
                    stream.Write(postFileContent, 0, postFileContent.Length);

                    byte[] trailer = Encoding.ASCII.GetBytes("\r\n--" + boundary + "--\r\n");
                    stream.Write(trailer, 0, trailer.Length);
                }

                using (WebResponse webResponse = httpWebRequest.GetResponse())
                {
                    using (Stream stream = webResponse.GetResponseStream())
                    {
                        using (StreamReader streamReader = new StreamReader(stream))
                        {
                            return streamReader.ReadToEnd();
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                return exception.Message;
            }
        }

        /// <summary>
        /// 整合版數字轉字母
        /// </summary>
        /// <param name="Number"></param>
        /// <param name="Index"></param>
        /// <returns></returns>
        public static string MergeNumberToChar(int Number, int Index)
        {
            //string finalCahr = "";
            //string units = "";
            //string tens = "";
            //string hundreds = "";

            //int quotient = Number / 26;
            //int remainder = (Number - (quotient * 26));
            //int square = quotient / 26;
            //int tensRemainder = (quotient - (square * 26));

            //if (remainder == 0) remainder = 26;
            //if (remainder > 0)
            //{
            //    units = NumberToChar(remainder);
            //    finalCahr = units;
            //}
            //if (tensRemainder > 0)
            //{
            //    tens = NumberToChar(tensRemainder);
            //    finalCahr = tens + units;
            //}
            //if (square > 0)
            //{
            //    hundreds = NumberToChar(square);
            //    finalCahr = hundreds + tens + units;
            //}

            //finalCahr += Index.ToString();

            //return finalCahr;

            string s = string.Empty;
            while (Number > 0)
            {
                int m = Number % 26;
                if (m == 0) m = 26;
                s = (char)(m + 64) + s;
                Number = (Number - m) / 26;
            }

            return s + Index;
        }

        /// <summary>
        /// 數字轉字母
        /// </summary>
        /// <param name="Number"></param>
        /// <returns></returns>
        public static string NumberToChar(int Number)
        {
            if (Number >= 1)
            {
                ASCIIEncoding asciiEncoding = new ASCIIEncoding();
                int num = Number + 64;
                byte[] btNumber = new byte[] { (byte)num };
                return asciiEncoding.GetString(btNumber);
            }
            return "數字不在轉換範圍內";
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="mailConfig"></param>
        public static void MailSend(MailConfig mailConfig)
        {
            MimeMessage mimeMessage = new MimeMessage
            {
                Subject = mailConfig.Subject
            };

            #region //設定寄件人
            if (mailConfig.From.Length <= 0) throw new SystemException("【寄件人】資料不能為空!");
            if (mailConfig.From.Split(':').Length > 1)
            {
                if (Regex.IsMatch(mailConfig.From.Split(':')[1], RegexHelper.Email, RegexOptions.IgnoreCase))
                {
                    mimeMessage.From.Add(new MailboxAddress(mailConfig.From.Split(':')[0], mailConfig.From.Split(':')[1]));
                }
                else
                {
                    throw new SystemException("【寄件人】格式錯誤!");
                }
            }
            else
            {
                if (Regex.IsMatch(mailConfig.From, RegexHelper.Email, RegexOptions.IgnoreCase))
                {
                    mimeMessage.From.Add(new MailboxAddress(mailConfig.From, mailConfig.From));
                }
                else
                {
                    throw new SystemException("【寄件人】格式錯誤!");
                }
            }
            #endregion

            #region //設定收件人
            string[] mailToList = mailConfig.MailTo.Split(';');
            if (mailToList.Length <= 0) throw new SystemException("【收件人】資料不能為空!");
            foreach (var mail in mailToList)
            {
                if (mail.Split(':').Length > 1)
                {
                    if (Regex.IsMatch(mail.Split(':')[1], RegexHelper.Email, RegexOptions.IgnoreCase))
                    {
                        mimeMessage.To.Add(new MailboxAddress(mail.Split(':')[0], mail.Split(':')[1]));
                    }
                }
                else
                {
                    if (Regex.IsMatch(mail, RegexHelper.Email, RegexOptions.IgnoreCase))
                    {
                        mimeMessage.To.Add(new MailboxAddress(mail, mail));
                    }
                }
            }
            #endregion

            #region //設定副本
            if (mailConfig.MailCc != null) {
                string[] mailCcList = mailConfig.MailCc.Split(';');
                if (mailCcList.Length > 0)
                {
                    foreach (var mail in mailCcList)
                    {
                        if (mail.Split(':').Length > 1)
                        {
                            if (Regex.IsMatch(mail.Split(':')[1], RegexHelper.Email, RegexOptions.IgnoreCase))
                            {
                                mimeMessage.Cc.Add(new MailboxAddress(mail.Split(':')[0], mail.Split(':')[1]));
                            }
                        }
                        else
                        {
                            if (Regex.IsMatch(mail, RegexHelper.Email, RegexOptions.IgnoreCase))
                            {
                                mimeMessage.Cc.Add(new MailboxAddress(mail, mail));
                            }
                        }
                    }
                }
            }
            #endregion

            #region //設定密件副本
            if (mailConfig.MailBcc != null)
            {
                string[] mailBccList = mailConfig.MailBcc.Split(';');
                if (mailBccList.Length > 0)
                {
                    foreach (var mail in mailBccList)
                    {
                        if (mail.Split(':').Length > 1)
                        {
                            if (Regex.IsMatch(mail.Split(':')[1], RegexHelper.Email, RegexOptions.IgnoreCase))
                            {
                                mimeMessage.Bcc.Add(new MailboxAddress(mail.Split(':')[0], mail.Split(':')[1]));
                            }
                        }
                        else
                        {
                            if (Regex.IsMatch(mail, RegexHelper.Email, RegexOptions.IgnoreCase))
                            {
                                mimeMessage.Bcc.Add(new MailboxAddress(mail, mail));
                            }
                        }
                    }
                }
            }
            #endregion

            #region //設定信件內容
            BodyBuilder bodyBuilder = new BodyBuilder
            {
                HtmlBody = HttpUtility.UrlDecode(mailConfig.HtmlBody),
                TextBody = mailConfig.TextBody
            };
            #endregion

            #region //檢查是否有附件
            if (mailConfig.FileInfo != null)
            {
                Regex NumandEG = new Regex("[^A-Za-z0-9]");
                int i = 0;
                foreach (var item in mailConfig.FileInfo)
                {
                    string NewFileName = "";

                    #region //處理附建檔名

                    #region //若附檔為量測數據上傳檔案情況
                    if (mailConfig.QcFileFlag != null)
                    {
                        string[] FileNameList = item.FileName.Split('-');
                        foreach (var fileName in FileNameList)
                        {
                            if (!NumandEG.IsMatch(fileName))
                            {
                                NewFileName += fileName + "-";
                            }
                            else
                            {
                                if (fileName == "回火前") NewFileName += "Before-";
                                else if (fileName == "回火後") NewFileName += "After-";
                            }
                        }
                    }
                    #endregion

                    if (NewFileName.Length <= 0)
                    {
                        NewFileName = item.FileName;
                    }
                    else
                    {
                        NewFileName = NewFileName.Substring(0, NewFileName.Length - 1);
                    }
                    NewFileName = NewFileName + item.FileExtension;
                    #endregion
                    
                    bodyBuilder.Attachments.Add(NewFileName, item.FileContent);

                    i++;
                }
            }
            #endregion

            mimeMessage.Body = bodyBuilder.ToMessageBody();

            #region //發送信件
            using (SmtpClient smtpClient = new SmtpClient())
            {
                smtpClient.Connect(mailConfig.Host, mailConfig.Port, SecureSelect(mailConfig.SendMode));
                smtpClient.Authenticate(new NetworkCredential(mailConfig.Account, mailConfig.Password));
                smtpClient.Send(mimeMessage);
                smtpClient.Disconnect(true);
            }
            #endregion
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="option"></param>
        /// <returns></returns>
        public static SecureSocketOptions SecureSelect(int option)
        {
            var Mapping = new Dictionary<int, SecureSocketOptions>()
            {
                { 0, SecureSocketOptions.None},
                { 1, SecureSocketOptions.Auto},
                { 2, SecureSocketOptions.SslOnConnect},
                { 3, SecureSocketOptions.StartTls},
                { 4, SecureSocketOptions.StartTlsWhenAvailable}
            };

            return Mapping.TryGetValue(option, out SecureSocketOptions Result) ? Result : SecureSocketOptions.Auto;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="digits"></param>
        /// <returns></returns>
        public static string RandomCode(int digits)
        {
            string Code = "", temp = "";

            for (int i = 1; i <= (digits / 2 + 1); i++)
            {
                Code = Code + StrRight("00" + (DateTime.Now.Second + (39 * Rnd() + 1)), 2);
            }

            for (int i = 1; i <= digits; i++)
            {
                HttpContext.Current.Session["Temp_" + i] = Convert.ToInt32(Asc(StrMid(Code, i, 1)));
            }

            int buffer = 0;
            for (int i = 1; i <= digits; i++)
            {
                buffer = i;
                string step1 = "", step2 = "";

                step1 = HttpContext.Current.Session["Temp_" + buffer].ToString();
                if (buffer == digits) buffer = 0;
                step2 = HttpContext.Current.Session["Temp_" + (buffer + 1)].ToString();

                HttpContext.Current.Session["Code_" + i] = Convert.ToInt32(step1 + step2) * ((1000 * Rnd()) + 1) % 36;
            }

            for (int i = 1; i <= digits; i++)
            {
                if (Convert.ToInt32(HttpContext.Current.Session["Code_" + i]) > 9)
                {
                    HttpContext.Current.Session["Code_" + i] = Chr(Convert.ToInt32(HttpContext.Current.Session["Code_" + i]) + 55);
                }
                temp = temp + HttpContext.Current.Session["Code_" + i];
            }

            return temp;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="character"></param>
        /// <returns></returns>
        private static int Asc(string character)
        {
            if (character.Length == 1)
            {
                ASCIIEncoding asciiEncoding = new ASCIIEncoding();
                int intAsciiCode = asciiEncoding.GetBytes(character)[0];
                return (intAsciiCode);
            }
            else
            {
                throw new ApplicationException("Character is not valid.");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="asciiCode"></param>
        /// <returns></returns>
        private static string Chr(int asciiCode)
        {
            if (asciiCode >= 0 && asciiCode <= 255)
            {
                ASCIIEncoding asciiEncoding = new ASCIIEncoding();
                byte[] byteArray = new byte[] { (byte)asciiCode };
                string strCharacter = asciiEncoding.GetString(byteArray);
                return (strCharacter);
            }
            else
            {
                throw new ApplicationException("ASCII Code is not valid.");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static int Rnd()
        {
            Random r = new Random(Guid.NewGuid().GetHashCode());
            int rnumber = r.Next(1, 10);

            return rnumber;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="OriginalString"></param>
        /// <param name="Length"></param>
        /// <returns></returns>
        public static string StrLeft(string OriginalString, int Length)
        {
            if (OriginalString.Length > Length)
            {
                return OriginalString.Substring(0, Length);
            }
            else
            {
                return OriginalString;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="OriginalString"></param>
        /// <param name="Start"></param>
        /// <param name="Length"></param>
        /// <returns></returns>
        public static string StrMid(string OriginalString, int Start, int Length)
        {
            return OriginalString.Substring(Start, Length);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="OriginalString"></param>
        /// <param name="Length"></param>
        /// <returns></returns>
        public static string StrRight(string OriginalString, int Length)
        {
            if (OriginalString.Length > Length)
            {
                return OriginalString.Substring(OriginalString.Length - Length);
            }
            else
            {
                return OriginalString;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="COMPANY"></param>
        /// <param name="TYPE"></param>
        /// <param name="DOC"></param>
        /// <param name="DOC_NO"></param>
        /// <param name="TRANS"></param>
        /// <param name="USER"></param>
        /// <param name="DATE"></param>
        /// <returns></returns>
        public static string TransferErpAPI(string COMPANY, string TYPE, string DOC, string DOC_NO, string TRANS, string USER, string DATE)
        {
            string postUrl = "http://192.168.20.50/WEBAPI_8200441589/DOTRANS";
            JObject postDataJson = new JObject();
            postDataJson = JObject.FromObject(new
            {
                COMPANY,
                TYPE,
                DOC,
                DOC_NO,
                TRANS,
                USER,
                DATE
            });
            string postData = postDataJson.ToString();
            UTF8Encoding uTF8 = new UTF8Encoding();
            var result = PostWebRequest(postUrl, postData, uTF8);
            return result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="UserNo"></param>
        /// <param name="sqlConnection"></param>
        /// <param name="ErpFunctionNo"></param>
        /// <param name="checkType"></param>
        /// <returns></returns>
        public static string CheckErpAuthority(string UserNo, SqlConnection sqlConnection, string ErpFunctionNo, string checkType)
        {
            #region //查詢USR_GROUP
            DynamicParameters dynamicParameters = new DynamicParameters();
            string sql = @"SELECT LTRIM(RTRIM(a.MF004)) USR_GROUP
                           FROM ADMMF a
                           WHERE MF001 = @MF001";
            dynamicParameters.Add("MF001", UserNo);

            var result = sqlConnection.Query(sql, dynamicParameters);

            if (result.Count() <= 0) throw new SystemException("查無ERP帳號權限，無法進行拋轉及核單!");

            string USR_GROUP = "";
            foreach (var item in result)
            {
                USR_GROUP = item.USR_GROUP;
            }
            #endregion

            return USR_GROUP;

            #region //確認是否為超級使用者，若是則無需檢查其他權限
            dynamicParameters = new DynamicParameters();
            sql = @"SELECT TOP 1 1
                    FROM ADMMF a
                    WHERE a.MF001 = @MF001
                    AND a.MF005 = 'Y'";
            dynamicParameters.Add("MF001", UserNo);

            var admmfResult = sqlConnection.Query(sql, dynamicParameters);

            if (admmfResult.Count() <= 0)
            {
                #region //檢核帳號權限
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT LTRIM(RTRIM(a.MG006)) MG006
                        , LTRIM(RTRIM(b.MF005)) MF005
                        FROM ADMMG a
                        INNER JOIN ADMMF b ON a.MG001 = b.MF001
                        WHERE a.MG001 = @MG001
                        AND a.MG002 = @MG002";
                dynamicParameters.Add("MG001", UserNo);
                dynamicParameters.Add("MG002", ErpFunctionNo);

                var admmgResult = sqlConnection.Query(sql, dynamicParameters);

                if (admmgResult.Count() <= 0) throw new SystemException("此工號【" + UserNo + "】尚未開啟ERP【" + ErpFunctionNo + "】權限，請聯絡資訊人員處理!");

                foreach (var item in admmgResult)
                {
                    if (item.MF005 == "Y") break;

                    string MG006 = item.MG006;
                    char[] permissions = MG006.ToCharArray();
                    //if (checkType == "CONFIRM" || checkType == "C-CONFIRM")
                    //{
                    //    if (permissions[3] == 'N') throw new SystemException("此工號【" + UserNo + "】尚未開啟ERP【核單】權限，請聯絡資訊人員處理!");
                    //}
                    //else if (checkType == "RE-CONFIRM")
                    //{
                    //    if (permissions[4] == 'N') throw new SystemException("此工號【" + UserNo + "】尚未開啟ERP【反確認】權限，請聯絡資訊人員處理!");
                    //}
                    //else if (checkType == "CREATE") 
                    //{
                    //    if (permissions[8] == 'N') throw new SystemException("此工號【" + UserNo + "】尚未開啟ERP【建單】權限，請聯絡資訊人員處理!");
                    //}
                    //else if (checkType == "UPDATE") 
                    //{
                    //    if (permissions[1] == 'N') throw new SystemException("此工號【" + UserNo + "】尚未開啟ERP【修改】權限，請聯絡資訊人員處理!");
                    //}
                }
                #endregion
            }
            #endregion

            //return USR_GROUP;
        }

        /// <summary>
        /// 發送系統通知
        /// </summary>
        /// <param name="notification">通知參數</param>
        public static void SendNotification(Notification notification)
        {
            string response = "";
            JObject tempJObject = new JObject();

            if (notification.NotificationModes.Count > 0)
            {
                foreach (var mode in notification.NotificationModes)
                {
                    switch (mode)
                    {
                        case NotificationMode.Mail:
                            #region //取得使用者Email資料，並發送
                            response = PostWebRequest(
                                string.Format("{0}{1}", domainUrl, "api/BAS/GetNotificationMail"),
                                new List<KeyValuePair<string, string>>
                                {
                                    new KeyValuePair<string, string>("Company", company),
                                    new KeyValuePair<string, string>("SecretKey", secretKey),
                                    new KeyValuePair<string, string>("UserNo", notification.UserNo),
                                });

                            if (response.TryParseJson(out tempJObject))
                            {
                                JObject resultJson = JObject.Parse(response);

                                if (resultJson["status"].ToString() != "success") throw new SystemException(resultJson["msg"].ToString());

                                for (int i = 0; i < resultJson["result"].Count(); i++)
                                {
                                    string mailContent = HttpUtility.UrlDecode(resultJson["result"][i]["MailContent"].ToString());
                                    mailContent = mailContent.Replace("[CustomContent]", notification.LogContent);

                                    MailConfig mailConfig = new MailConfig
                                    {
                                        Host = resultJson["result"][i]["Host"].ToString(),
                                        Port = Convert.ToInt32(resultJson["result"][i]["Port"]),
                                        SendMode = Convert.ToInt32(resultJson["result"][i]["SendMode"]),
                                        From = resultJson["result"][i]["MailFrom"].ToString(),
                                        Subject = notification.LogTitle,
                                        Account = resultJson["result"][i]["Account"].ToString(),
                                        Password = resultJson["result"][i]["Password"].ToString(),
                                        MailTo = resultJson["result"][i]["DisplayName"].ToString() + ":" + resultJson["result"][i]["Email"].ToString(),
                                        MailCc = "",
                                        MailBcc = "",
                                        HtmlBody = mailContent,
                                        TextBody = "-"
                                    };

                                    if (resultJson["result"][i]["Email"].ToString().Length > 0) MailSend(mailConfig);
                                }
                            }
                            else
                            {
                                throw new SystemException(response);
                            }
                            #endregion
                            break;
                        case NotificationMode.Push:
                            #region //取得使用者推播訂閱資料，並發送
                            response = PostWebRequest(
                                string.Format("{0}{1}", domainUrl, "api/BAS/GetSubscription"), 
                                new List<KeyValuePair<string, string>>
                                {
                                    new KeyValuePair<string, string>("Company", company),
                                    new KeyValuePair<string, string>("SecretKey", secretKey),
                                    new KeyValuePair<string, string>("UserNo", notification.UserNo),
                                });

                            if (response.TryParseJson(out tempJObject))
                            {
                                JObject resultJson = JObject.Parse(response);

                                if (resultJson["status"].ToString() != "success") throw new SystemException(resultJson["msg"].ToString());

                                for (int i = 0; i < resultJson["result"].Count(); i++)
                                {
                                    JObject pushInfo = JObject.Parse(resultJson["result"][i]["PushInfo"].ToString());

                                    List<PushNotificationUser> notificationUsers = JsonConvert.DeserializeObject<List<PushNotificationUser>>(pushInfo["data"].ToString());

                                    WebPushHelper webPushHelper = new WebPushHelper();
                                    webPushHelper.SendPush(notificationUsers, notification.LogTitle, notification.LogContent);
                                }
                            }
                            else
                            {
                                throw new SystemException(response);
                            }
                            #endregion
                            break;
                        case NotificationMode.WorkWeixin:
                            break;
                    }
                }

                #region //系統通知紀錄
                response = PostWebRequest(
                    string.Format("{0}{1}", domainUrl, "api/BAS/AddNotification"),
                    new List<KeyValuePair<string, string>>
                    {
                        new KeyValuePair<string, string>("Company", company),
                        new KeyValuePair<string, string>("SecretKey", secretKey),
                        new KeyValuePair<string, string>("Notification", JsonConvert.SerializeObject(notification)),
                    });

                if (response.TryParseJson(out tempJObject))
                {
                    JObject resultJson = JObject.Parse(response);

                    if (resultJson["status"].ToString() != "success") throw new SystemException(resultJson["msg"].ToString());
                }
                else
                {
                    throw new SystemException(response);
                }
                #endregion
            }
            else
            {
                throw new SystemException("無法讀取發送模式!");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="UserId">使用者ID</param>
        /// <param name="CompanyId">公司別</param>
        /// <param name="Status">權限狀態</param>
        /// <param name="FunctionCode">功能名稱</param>
        /// <param name="DetailCode">詳細功能名稱</param>
        /// <returns></returns>
        public static string CheckUserAuthority(int UserId, int CompanyId, string Status, string FunctionCode, string DetailCode, SqlConnection sqlConnection)
        {
            string checkResult = "Y";
            DynamicParameters dynamicParameters = new DynamicParameters();
            string sql = @"SELECT TOP 1 d.UserNo, d.UserName, d.DepartmentNo
                                FROM BAS.FunctionDetail a
                                INNER JOIN BAS.[Function] b ON a.FunctionId = b.FunctionId
                                OUTER APPLY (
                                    SELECT ISNULL((
                                        SELECT TOP 1 1
                                        FROM BAS.RoleFunctionDetail ca
                                        WHERE ca.DetailId = a.DetailId
                                        AND ca.RoleId IN (
                                            SELECT caa.RoleId
                                            FROM BAS.UserRole caa
                                            INNER JOIN BAS.[Role] cab ON caa.RoleId = cab.RoleId
                                            WHERE caa.UserId = @UserId
                                            AND cab.CompanyId = @CompanyId
                                        )
                                    ), 0) Authority
                                ) c
                                OUTER APPLY (
                                    SELECT da.UserNo, da.UserName, db.DepartmentNo
                                    FROM BAS.[User] da
                                    INNER JOIN BAS.Department db ON da.DepartmentId = db.DepartmentId
                                    WHERE da.UserId = @UserId
                                ) d
                                WHERE a.[Status] = @Status
                                AND b.[Status] = @Status
                                AND b.FunctionCode = @FunctionCode
                                AND a.DetailCode = @DetailCode
                                AND c.Authority > 0";
            dynamicParameters.Add("UserId", UserId);
            dynamicParameters.Add("CompanyId", CompanyId);
            dynamicParameters.Add("Status", Status);
            dynamicParameters.Add("FunctionCode", FunctionCode);
            dynamicParameters.Add("DetailCode", DetailCode);

            var resultUser = sqlConnection.Query(sql, dynamicParameters);

            if (resultUser.Count() <= 0)
            {
                checkResult = "N";
            }

            return checkResult;
        }
    }
}

public static class Extend
{
    /// <summary>
    /// 將JSON的字串表示轉換成它的對等的JSON。傳回指示作業是否成功的值。
    /// </summary>
    /// <typeparam name="T">資料型態</typeparam>
    /// <param name="this">判別字串</param>
    /// <param name="result">判別目標</param>
    /// <returns>是否符合該型態</returns>
    public static bool TryParseJson<T>(this string @this, out T result)
    {
        bool success = true;
        var settings = new JsonSerializerSettings
        {
            Error = (sender, args) => { success = false; args.ErrorContext.Handled = true; },
            MissingMemberHandling = MissingMemberHandling.Error
        };

        result = JsonConvert.DeserializeObject<T>(@this, settings);

        return success;
    }

    /// <summary>
    /// 回傳 Enum 的 Description 屬性，如果沒有 Description 屬性就回傳列舉成員名稱
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    public static string GetDescription(this Enum value)
    {
        var field = value.GetType().GetField(value.ToString());
        var attribute = field?.GetCustomAttributes(typeof(DescriptionAttribute), false).FirstOrDefault() as DescriptionAttribute;
        return attribute == null ? value.ToString() : attribute.Description;
    }
}
