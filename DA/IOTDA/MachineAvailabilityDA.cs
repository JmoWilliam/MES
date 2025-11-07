using Dapper;
using Helpers;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace IOTDA
{
    public class MachineAvailabilityDA
    {
        public static string MainConnectionStrings = "";
        public static string IotConnectionStrings = "";


        public static int CurrentCompany = -1;
        public static int CurrentUser = -1;
        public static int CreateBy = -1;
        public static int LastModifiedBy = -1;
        public static DateTime CreateDate = default(DateTime);
        public static DateTime LastModifiedDate = default(DateTime);

        public static string sql = "";
        public static JObject jsonResponse = new JObject();
        public static Logger logger = LogManager.GetCurrentClassLogger();
        public static DynamicParameters dynamicParameters = new DynamicParameters();
        public static SqlQuery sqlQuery = new SqlQuery();

        public MachineAvailabilityDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            IotConnectionStrings = ConfigurationManager.AppSettings["MainIotDb"];

            CreateDate = DateTime.Now;
            LastModifiedDate = DateTime.Now;

            dynamicParameters = new DynamicParameters();
            sqlQuery = new SqlQuery();
        }

        #region //成型機歷史紀錄稼動率(趨勢折線圖)
        public string GetFactoryDayHistory(int ID)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.Machine_Name,b.Utilization_Rate,b.RecordTime
FROM Name_Machine a
INNER JOIN History_Day_Utilization_Rate_Injection_machine_White b ON a.MachineSN=b.MachineSN
WHERE 1=1
AND a.ID = @ID
ORDER BY b.RecordTime DESC
";
                    dynamicParameters.Add("ID", ID);
                    var result = sqlConnection.Query(sql, dynamicParameters);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });
                    #endregion
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion


        #region //成型機台即時狀態列表
        public string GetFactoryRealTimeList()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.ID,a.MachineSN,a.Machine_Name,b.UtilizationRate,e.Status_Name,c.MoldName,'' AS T 
FROM Name_Machine a
INNER JOIN Current_Utilization_Rate_Injection_machine_White b ON a.MachineSN = b.MachineSN
INNER JOIN Current_Parameter_Injection_machine_White c ON a.MachineSN = c.MachineSN
INNER JOIN Current_RunTime_Injection_machine_White d ON a.MachineSN = d.MachineSN 
INNER JOIN StatusCode_Injection_machine_White e ON d.RunTime_Status = e.Status_Code
";

                    var result = sqlConnection.Query(sql, dynamicParameters);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });
                    #endregion
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //機台機時(長條圖)
        public string GetFactoryTimeHistory(int SN, string StartTime, string EndTime)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {


                    sql = @"SELECT a.MachineSN,c.Machine_Name,a.RunTime_Status,b.Status_Name,a.StartTime,a.EndTime,a.Seconds
                FROM ALL_ChangeStatus_Record a
                INNER JOIN StatusCode_Injection_machine_White b ON a.RunTime_Status=b.Status_Code 
                INNER JOIN Name_Machine c ON a.MachineSN=c.MachineSN
                WHERE 1=1 
                AND a.MachineSN = @SN
               AND a.StartTime >= @StartTime
               AND a.EndTime <= @EndTime
                UNION ALL 
               SELECT a.MachineSN, c.Machine_Name, a.RunTime_Status, b.Status_Name,MAX(a.StartTime)StartTime, a.EndTime, a.Seconds 
               FROM ALL_ChangeStatus_Record a 
               INNER JOIN StatusCode_Injection_machine_White b ON a.RunTime_Status=b.Status_Code 
                INNER JOIN Name_Machine c ON a.MachineSN=c.MachineSN 
               WHERE 1=1 
                AND a.MachineSN = 1
                AND a.ID = (SELECT MAX(a1.ID)
                FROM ALL_ChangeStatus_Record a1
               WHERE a1.StartTime <= @StartTime
                AND a.MachineSN = a1.MachineSN)
                GROUP BY a.MachineSN,c.Machine_Name,a.RunTime_Status,b.Status_Name,a.StartTime,a.EndTime,a.Seconds
               ORDER BY a.StartTime DESC
";
                    dynamicParameters.Add("SN", SN);
                    dynamicParameters.Add("StartTime", StartTime);
                    dynamicParameters.Add("EndTime", EndTime);
                    var result = sqlConnection.Query(sql, dynamicParameters);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });
                    #endregion
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region//機台生產產品
        #endregion

        #region//產品生產狀況
        #endregion

        #region//產品不良統計
        #endregion

    }
}
