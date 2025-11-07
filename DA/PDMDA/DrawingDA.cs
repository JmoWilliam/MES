using Dapper;
using Helpers;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Transactions;
using System.Web;

namespace PDMDA
{
    public class DrawingDA
    {
        public string MainConnectionStrings = "";
        public string ErpConnectionStrings = "";
        public string HrmConnectionStrings = "";

        public int CurrentCompany = -1;
        public int CurrentUser = -1;
        public int CreateBy = -1;
        public int LastModifiedBy = -1;
        public DateTime CreateDate = default(DateTime);
        public DateTime LastModifiedDate = default(DateTime);

        public string sql = "";
        public JObject jsonResponse = new JObject();
        public Logger logger = LogManager.GetCurrentClassLogger();
        public DynamicParameters dynamicParameters = new DynamicParameters();
        public SqlQuery sqlQuery = new SqlQuery();

        public DrawingDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            ErpConnectionStrings = ConfigurationManager.AppSettings["ErpDb"];
            HrmConnectionStrings = ConfigurationManager.AppSettings["HrmDb"];

            GetUserInfo();
            CreateDate = DateTime.Now;
            LastModifiedDate = DateTime.Now;

            dynamicParameters = new DynamicParameters();
            sqlQuery = new SqlQuery();
        }

        #region //GetUserInfo 取得使用者資訊
        private void GetUserInfo()
        {
            try
            {
                CurrentCompany = Convert.ToInt32(HttpContext.Current.Session["UserCompany"]);
                CurrentUser = Convert.ToInt32(HttpContext.Current.Session["UserId"]);
                CreateBy = Convert.ToInt32(HttpContext.Current.Session["UserId"]);
                LastModifiedBy = Convert.ToInt32(HttpContext.Current.Session["UserId"]);

                if (HttpContext.Current.Session["CompanySwitch"] != null)
                {
                    if (HttpContext.Current.Session["CompanySwitch"].ToString() == "manual")
                    {
                        CurrentCompany = Convert.ToInt32(HttpContext.Current.Session["CompanyId"]);
                    }
                }
            }
            catch (Exception e)
            {
                logger.Error(e.Message);
            }
        }
        #endregion

        #region //Get
        #region //GetCustomerCad -- 取得客戶設計圖資料 -- Ann 2022-06-09
        public string GetCustomerCad(int CadId, string CustomerMtlItemNo, string CustomerDwgNo
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.CadId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CadId, a.CustomerMtlItemNo, a.CustomerDwgNo";
                    sqlQuery.mainTables =
                        @"FROM PDM.CustomerCad a";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CadId", @" AND a.CadId = @CadId", CadId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerMtlItemNo", @" AND a.CustomerMtlItemNo LIKE '%' + @CustomerMtlItemNo + '%'", CustomerMtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerDwgNo", @" AND a.CustomerDwgNo LIKE '%' + @CustomerDwgNo + '%'", CustomerDwgNo);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.CreateDate DESC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

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

        #region //GetCustomerCadControl -- 取得客戶設計圖版本控制資料 -- Ann 2022-06-09
        public string GetCustomerCadControl(int ControlId, int CadId, string Edition
            , string StartDate, string EndDate, string CustomerMtlItemNo, string ReleasedStatus
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.ControlId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ControlId, a.CadId, a.CadFile, a.Edition, a.[Version], a.ReleasedStatus, a.Cause
                        , FORMAT(a.CreateDate, 'yyyy-MM-dd hh:mm:ss') ChangeDate, a.CadFileAbsolutePath
                        , b.CustomerMtlItemNo, ISNULL(b.CustomerDwgNo, '') CustomerDwgNo
                        , c.UserNo, c.UserName
                        , d.[FileName], d.FileExtension
                        , (
                          SELECT aa.FileId
                          , ab.[FileName], ab.FileExtension
                          FROM PDM.CustomerCadControlFile aa
                          INNER JOIN BAS.[File] ab ON aa.FileId = ab.FileId
                          WHERE aa.ControlId = a.ControlId
                          FOR JSON PATH, ROOT('data')
                        ) OtherFile";
                    sqlQuery.mainTables =
                        @"FROM PDM.CustomerCadControl a
                        INNER JOIN PDM.CustomerCad b ON a.CadId = b.CadId
                        INNER JOIN BAS.[User] c ON a.CreateBy = c.UserId
                        LEFT JOIN BAS.[File] d ON a.CadFile = d.FileId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ControlId", @" AND a.ControlId  = @ControlId", ControlId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CadId", @" AND a.CadId  = @CadId", CadId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Edition", @" AND a.Edition LIKE '%' + @Edition + '%'", Edition);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.CreateDate >= @StartDate", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.CreateDate <= @EndDate", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerMtlItemNo", @" AND b.CustomerMtlItemNo LIKE '%' + @CustomerMtlItemNo + '%'", CustomerMtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ReleasedStatus", @" AND a.ReleasedStatus = @ReleasedStatus", ReleasedStatus);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.CreateDate DESC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

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

        #region //GetCustomerCadEdition -- 取得客戶設計圖版次資料 -- Ann 2022-06-23
        public string GetCustomerCadEdition(int CadId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT DISTINCT a.Edition
                            FROM PDM.CustomerCadControl a
                            WHERE a.CadId = @CadId
                            ORDER BY a.Edition";
                    dynamicParameters.Add("CadId", CadId);

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

        #region //GetRdDesign -- 取得研發設計圖資料 -- Ann 2022-06-27
        public string GetRdDesign(int DesignId, string CustomerMtlItemNo,string MtlItemNo, int MtlItemId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.DesignId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.DesignId, a.CustomerMtlItemNo, a.CustomerCadControlId, a.MtlItemId
                        , ISNULL(b.CadFile, -1) CadFile, ISNULL(b.Cause, '') Cause, ISNULL(b.Edition, '') Edition, ISNULL(b.Version, '') Version, b.CadFileAbsolutePath
                        , ISNULL(c.CustomerDwgNo, '') CustomerDwgNo
                        , ISNULL(d.[FileName], '') FileName, ISNULL(d.FileExtension, '') FileExtension
                        , ISNULL(e.MtlItemNo, '') MtlItemNo, ISNULL(e.MtlItemName, '') MtlItemName, ISNULL(e.MtlItemSpec, '') MtlItemSpec";
                    sqlQuery.mainTables =
                        @"FROM PDM.RdDesign a
                        LEFT JOIN PDM.CustomerCadControl b ON a.CustomerCadControlId = b.ControlId
                        LEFT JOIN PDM.CustomerCad c ON b.CadId = c.CadId
                        LEFT JOIN BAS.[File] d ON b.CadFile = d.FileId
                        LEFT JOIN PDM.MtlItem e ON a.MtlItemId = e.MtlItemId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    dynamicParameters.Add("MtlItemId", MtlItemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DesignId", @" AND a.DesignId = @DesignId", DesignId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemId", @" AND a.MtlItemId = @MtlItemId", MtlItemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerMtlItemNo", @" AND a.CustomerMtlItemNo LIKE '%' + @CustomerMtlItemNo + '%'", CustomerMtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND e.MtlItemNo LIKE '%' + @MtlItemNo + '%'", MtlItemNo);

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.CreateDate DESC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

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

        #region //GetAiRdDesign -- 取得 AI 研發設計圖資料 -- MarkChen 2023-08-23
        // SELECT * FROM PDM.MtlItem
        // MtlItemId	CompanyId	ItemTypeId	MtlItemNo	MtlItemName
        //  1	        2	        NULL	    601A101001  主軸(YCM銑床機用)
        public string GetAiRdDesign(int DesignId, string CustomerMtlItemNo, string MtlItemNo, string MtlItemName, int MtlItemId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.DesignId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.DesignId, a.CustomerMtlItemNo, a.CustomerCadControlId, a.MtlItemId
                        , ISNULL(b.CadFile, -1) CadFile, ISNULL(b.Cause, '') Cause, ISNULL(b.Edition, '') Edition, ISNULL(b.Version, '') Version, b.CadFileAbsolutePath
                        , ISNULL(c.CustomerDwgNo, '') CustomerDwgNo
                        , ISNULL(d.[FileName], '') FileName, ISNULL(d.FileExtension, '') FileExtension
                        , ISNULL(e.MtlItemNo, '') MtlItemNo, ISNULL(e.MtlItemName, '') MtlItemName, ISNULL(e.MtlItemSpec, '') MtlItemSpec";
                    sqlQuery.mainTables =
                        @"FROM PDM.RdDesign a
                        LEFT JOIN PDM.CustomerCadControl b ON a.CustomerCadControlId = b.ControlId
                        LEFT JOIN PDM.CustomerCad c ON b.CadId = c.CadId
                        LEFT JOIN BAS.[File] d ON b.CadFile = d.FileId
                        LEFT JOIN PDM.MtlItem e ON a.MtlItemId = e.MtlItemId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    dynamicParameters.Add("MtlItemId", MtlItemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DesignId", @" AND a.DesignId = @DesignId", DesignId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemId", @" AND a.MtlItemId = @MtlItemId", MtlItemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerMtlItemNo", @" AND a.CustomerMtlItemNo LIKE '%' + @CustomerMtlItemNo + '%'", CustomerMtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND e.MtlItemNo LIKE '%' + @MtlItemNo + '%'", MtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemName", @" AND e.MtlItemName LIKE '%' + @MtlItemName + '%'", MtlItemName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.CreateDate DESC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

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

        #region //GetRdDesignControl -- 取得研發設計圖版本控制資料 -- Ann 2022-06-28
        public string GetRdDesignControl(int ControlId, int DesignId, string Edition, string StartDate, string EndDate, int MtlItemId, string ReleasedStatus
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                //NOTE by MarkChen, 2023-08-23, 因應 AI途程, 需要加取 c.MtlItemId
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.ControlId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ControlId, a.Edition, a.[Version], a.Cause, c.MtlItemId
                        , FORMAT(a.DesignDate, 'yyyy-MM-dd') DesignDate, a.ReleasedStatus
                        , FORMAT(a.CreateDate, 'yyyy-MM-dd hh:mm:ss') CreateDate
                        , a.Cad3DFileAbsolutePath, a.Cad2DFileAbsolutePath, a.Pdf2DFileAbsolutePath, a.JmoFileAbsolutePath
                        , (
                          SELECT xa.FileId, (xa.[FileName] + xa.FileExtension) FileInfo
                          FROM BAS.[File] xa
                          WHERE xa.FileId = a.Cad3DFile
                          FOR JSON PATH, ROOT('data')
                        ) Cad3DFile
                        , (
                          SELECT xb.FileId, (xb.[FileName] + xb.FileExtension) FileInfo
                          FROM BAS.[File] xb
                          WHERE xb.FileId = a.Cad2DFile
                          FOR JSON PATH, ROOT('data')
                        ) Cad2DFile
                        , (
                          SELECT xc.FileId, (xc.[FileName] + xc.FileExtension) FileInfo
                          FROM BAS.[File] xc
                          WHERE xc.FileId = a.Pdf2DFile
                          FOR JSON PATH, ROOT('data')
                        ) Pdf2DFile
                        , (
                          SELECT xd.FileId, (xd.[FileName] + xd.FileExtension) FileInfo
                          FROM BAS.[File] xd
                          WHERE xd.FileId = a.JmoFile
                          FOR JSON PATH, ROOT('data')
                        ) JmoFile
                        , b.UserNo, b.UserName
                        , c.DesignId, c.MtlItemId";
                    sqlQuery.mainTables =
                        @"FROM PDM.RdDesignControl a
                        INNER JOIN BAS.[User] b ON a.CreateBy = b.UserId
                        INNER JOIN PDM.RdDesign c ON a.DesignId = c.DesignId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND c.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    dynamicParameters.Add("DesignId", DesignId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ControlId", @" AND a.ControlId = @ControlId", ControlId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DesignId", @" AND a.DesignId = @DesignId", DesignId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Edition", @" AND a.Edition LIKE '%' + @Edition + '%'", Edition);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.CreateDate >= @StartDate", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.CreateDate <= @EndDate", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemId", @" AND c.MtlItemId = @MtlItemId", MtlItemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ReleasedStatus", @" AND a.ReleasedStatus = @ReleasedStatus", ReleasedStatus);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.Edition, a.Version";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

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

        #region //GetRdDesignControlActiveOnly -- 取得研發設計圖版本控制資料, 只限 發行狀態=是, PDM.RdDesignControl ReleasedStatus Y  -- MarkChen 2023-08-24 
        public string GetRdDesignControlActiveOnly(int CompanyId, int ControlId, int DesignId, string Edition, string StartDate, string EndDate, int MtlItemId, string ReleasedStatus
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                //NOTE by MarkChen, 2023-08-23, 因應 AI途程, 需要加取 c.MtlItemId
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.ControlId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ControlId, a.Edition, a.[Version], a.Cause, c.MtlItemId
                        , FORMAT(a.DesignDate, 'yyyy-MM-dd') DesignDate, a.ReleasedStatus
                        , FORMAT(a.CreateDate, 'yyyy-MM-dd hh:mm:ss') CreateDate
                        , a.Cad3DFileAbsolutePath, a.Cad2DFileAbsolutePath, a.Pdf2DFileAbsolutePath, a.JmoFileAbsolutePath
                        , (
                          SELECT xa.FileId, (xa.[FileName] + xa.FileExtension) FileInfo
                          FROM BAS.[File] xa
                          WHERE xa.FileId = a.Cad3DFile
                          FOR JSON PATH, ROOT('data')
                        ) Cad3DFile
                        , (
                          SELECT xb.FileId, (xb.[FileName] + xb.FileExtension) FileInfo
                          FROM BAS.[File] xb
                          WHERE xb.FileId = a.Cad2DFile
                          FOR JSON PATH, ROOT('data')
                        ) Cad2DFile
                        , (
                          SELECT xc.FileId, (xc.[FileName] + xc.FileExtension) FileInfo
                          FROM BAS.[File] xc
                          WHERE xc.FileId = a.Pdf2DFile
                          FOR JSON PATH, ROOT('data')
                        ) Pdf2DFile
                        , (
                          SELECT xd.FileId, (xd.[FileName] + xd.FileExtension) FileInfo
                          FROM BAS.[File] xd
                          WHERE xd.FileId = a.JmoFile
                          FOR JSON PATH, ROOT('data')
                        ) JmoFile
                        , b.UserNo, b.UserName
                        , c.DesignId, c.MtlItemId,c.CustomerMtlItemNo";
                    sqlQuery.mainTables =
                        @"FROM PDM.RdDesignControl a
                        INNER JOIN BAS.[User] b ON a.CreateBy = b.UserId
                        INNER JOIN PDM.RdDesign c ON a.DesignId = c.DesignId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND c.CompanyId = @CompanyId";

                    //限制 發行狀態 是
                    queryCondition += @" AND a.ReleasedStatus = 'Y' ";
                    if (CompanyId != -1)
                    {
                        dynamicParameters.Add("CompanyId", CompanyId);
                    }
                    else {
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                    }
                    dynamicParameters.Add("DesignId", DesignId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ControlId", @" AND a.ControlId = @ControlId", ControlId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DesignId", @" AND a.DesignId = @DesignId", DesignId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Edition", @" AND a.Edition LIKE '%' + @Edition + '%'", Edition);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.CreateDate >= @StartDate", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.CreateDate <= @EndDate", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemId", @" AND c.MtlItemId = @MtlItemId", MtlItemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ReleasedStatus", @" AND a.ReleasedStatus = @ReleasedStatus", ReleasedStatus);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.Edition, a.Version";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

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

        #region //GetRdDesignEdition -- 取得研發設計圖版次資料 -- Ann 2022-06-28
        public string GetRdDesignEdition(int DesignId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT DISTINCT a.Edition
                            FROM PDM.RdDesignControl a
                            WHERE a.DesignId = @DesignId
                            ORDER BY a.Edition";
                    dynamicParameters.Add("DesignId", DesignId);

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

        #region //GetRdDesignAttribute -- 取得研發設計圖加工屬性 -- Ann 2022-07-19
        public string GetRdDesignAttribute(int ControlId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.AttributeId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.AttributeId, a.ControlId, a.AttributeItemId, a.AttributeValue, a.UpperLimit, a.LowerLimit
                        , b.ItemNo, b.ItemDesc, b.ModeId
                        , c.DesignId, c.Edition
                        , d.ModeNo, d.ModeName";
                    sqlQuery.mainTables =
                        @"FROM PDM.RdDesignAttribute a
                        INNER JOIN PDM.RdAttributeItem b ON a.AttributeItemId = b.AttributeItemId
                        INNER JOIN PDM.RdDesignControl c ON a.ControlId = c.ControlId
                        INNER JOIN MES.ProdMode d ON b.ModeId = d.ModeId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ControlId", @" AND a.ControlId = @ControlId", ControlId);
                    sqlQuery.conditions = queryCondition;
                    string OrderBy = "";
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "c.Edition";

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

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

        #region //GetRdAttributeItem -- 取得加工屬性資料 -- Ann 2022-07-21
        public string GetRdAttributeItem()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.AttributeItemId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.AttributeItemId, a.ModeId, a.ItemNo, a.ItemDesc
                        , b.ModeName";
                    sqlQuery.mainTables =
                        @"FROM PDM.RdAttributeItem a
                        INNER JOIN MES.ProdMode b ON a.ModeId = b.ModeId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    sqlQuery.conditions = queryCondition;
                    string OrderBy = "";
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.AttributeItemId DESC";

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

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

        #region //GetFileInfo -- 取得檔案相關資訊 -- Ann 2023-06-19
        public string GetFileInfo(string FilePath, string MESFolderPath, string DownloadFlag)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    if (DownloadFlag == "RD")
                    {
                        #region //先確認此下載路徑是否為合法路徑
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT FolderPath
                            FROM PDM.RdWhitelist
                            WHERE 1=1";
                        var RdWhitelistResult = sqlConnection.Query(sql, dynamicParameters);

                        bool checkWhitelist = false;
                        foreach (var item in RdWhitelistResult)
                        {
                            if (FilePath.IndexOf(item.FolderPath) != -1)
                            {
                                checkWhitelist = true;
                                break;
                            }
                        }

                        if (checkWhitelist == false)
                        {
                            throw new SystemException("此路徑非合法路徑，無法下載!!");
                        }
                        #endregion
                    }
                    else
                    {
                        if (FilePath.IndexOf(MESFolderPath) == -1) throw new SystemException("此路徑非合法路徑，無法下載!!");

                        #region //處理URL特殊符號
                        if (FilePath.IndexOf("+") != -1) FilePath = FilePath.Replace("+", "%2B");
                        if (FilePath.IndexOf("/") != -1) FilePath = FilePath.Replace("/", "%2F");
                        if (FilePath.IndexOf("?") != -1) FilePath = FilePath.Replace("?", "%3F");
                        if (FilePath.IndexOf("#") != -1) FilePath = FilePath.Replace("#", "%23");
                        if (FilePath.IndexOf("&") != -1) FilePath = FilePath.Replace("&", "%26");
                        if (FilePath.IndexOf("=") != -1) FilePath = FilePath.Replace("=", "%3D");
                        #endregion
                    }

                    List<CadFileInfo> fileInfos = new List<CadFileInfo>();
                    if (File.Exists(FilePath))
                    {
                        CadFileInfo cadFileInfo = new CadFileInfo()
                        {
                            FileName = Path.GetFileNameWithoutExtension(FilePath),
                            FileExtension = Path.GetExtension(FilePath),
                            FileByte = File.ReadAllBytes(FilePath),
                        };

                        fileInfos.Add(cadFileInfo);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = fileInfos
                        });
                        #endregion
                    }
                    else
                    {
                        throw new SystemException("檔案路徑錯誤!!");
                    }
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

        #region //GetRdWhitelist -- 取得RD白名單資料夾 -- Ann 20203-08-07
        public string GetRdWhitelist(int ListId, int DepartmentId, int UserId, string FolderPath)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得使用者部門
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DepartmentId
                            FROM BAS.[User]
                            WHERE UserId = @UserId";
                    dynamicParameters.Add("UserId", UserId);

                    var UserResult = sqlConnection.Query(sql, dynamicParameters);

                    if (UserResult.Count() <= 0) throw new SystemException("使用者資訊錯誤!!");

                    foreach (var item in UserResult)
                    {
                        DepartmentId = item.DepartmentId;
                    }
                    #endregion

                    dynamicParameters = new DynamicParameters();

                    sql = @"
                           WITH cte AS (
                               SELECT 
                                   a.ListId,
                                   a.DepartmentId,
                                   a.FolderPath,
                                   ROW_NUMBER() OVER (PARTITION BY a.FolderPath ORDER BY a.ListId) AS rn
                               FROM PDM.RdWhitelist a
                               WHERE 1 = 1
                           ";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ListId", @" AND a.ListId = @ListId", ListId);
                    //BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DepartmentId", @" AND a.DepartmentId = @DepartmentId", DepartmentId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "FolderPath", @" AND a.FolderPath LIKE '%' + @FolderPath + '%'", FolderPath);

                    sql += @"
                           )
                           SELECT ListId, DepartmentId, FolderPath
                           FROM cte
                           WHERE rn = 1";

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

        #region //GetCheckRdWhitelist -- 確認路徑是否為合法路徑 -- Ann 20203-08-08
        public string GetCheckRdWhitelist(int ListId, string FolderPath, int UserId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得使用者部門
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DepartmentId
                            FROM BAS.[User]
                            WHERE UserId = @UserId";
                    dynamicParameters.Add("UserId", UserId);

                    var UserResult = sqlConnection.Query(sql, dynamicParameters);

                    if (UserResult.Count() <= 0) throw new SystemException("使用者資訊錯誤!!");

                    int DepartmentId = -1;
                    foreach (var item in UserResult)
                    {
                        DepartmentId = item.DepartmentId;
                    }
                    #endregion

                    #region //確認白名單資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT FolderPath
                            FROM PDM.RdWhitelist
                            WHERE ListId = @ListId";
                    dynamicParameters.Add("ListId", ListId);
                    //dynamicParameters.Add("DepartmentId", DepartmentId);

                    var RdWhitelistResult = sqlConnection.Query(sql, dynamicParameters);

                    if (RdWhitelistResult.Count() <= 0) throw new SystemException("白名單資料錯誤!!");

                    string WhiteFolderPath = "";
                    foreach (var item in RdWhitelistResult)
                    {
                        WhiteFolderPath = item.FolderPath;
                    }
                    #endregion

                    if (FolderPath.IndexOf(WhiteFolderPath) == -1) throw new SystemException("此路徑非合法路徑!!");

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = RdWhitelistResult
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

        #region //GetRdDesignVersion -- 取得研發設計圖系統版次資料 -- WuTc 2024-09-07
        public string GetRdDesignVersion(int MtlItemId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT c.Edition, c.[Version], c.ControlId, c.Edition + ' / ' + c.[Version] ver
                                FROM PDM.MtlItem a
	                            INNER JOIN PDM.RdDesign b ON a.MtlItemId = b.MtlItemId
	                            INNER JOIN PDM.RdDesignControl c ON b.DesignId = c.DesignId
                                WHERE a.MtlItemId = @MtlItemId
                            ORDER BY a.[Version]";
                    dynamicParameters.Add("MtlItemId", MtlItemId);

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

        #region//GetGetAutoReadCadDesign -- 取得自動讀CAD圖規格 --WuTc 2024.09.07
        public string GetAutoReadCadDesign(int MtlItemId, int CadControlId, int QcItemId, string BallMark, string Design, string Usl, string Lsl, string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.ArcdId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.MtlItemId, a.VersionNo, a.DesignValue, a.UpperTolerance, a.LowerTolerance
                            , ISNULL(a.QcItemId, ''), ISNULL(e.QcItemNo, '未綁定') QcItemNo , ISNULL(a.QcItemName, '') QcItemName, ISNULL(a.QcItemDesc, '') QcItemDesc
                            , ISNULL(a.BallMark, '') BallMark, ISNULL(a.Remark, '') Remark";
                    sqlQuery.mainTables =
                        @"FROM PDM.AutoReadCadDesign a
	                        INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
	                        INNER JOIN PDM.RdDesign c ON b.MtlItemId = c.MtlItemId
	                        INNER JOIN PDM.RdDesignControl d ON c.DesignId = d.DesignId AND a.VersionNo = d.[Version]
                            LEFT JOIN QMS.QcItem e ON a.QcItemId = e.QcItemId";
                    sqlQuery.auxTables = "";
                    string queryCondition = " AND b.CompanyId = @CompanyId AND a.MtlItemId = @MtlItemId AND d.ControlId = @ControlId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    dynamicParameters.Add("MtlItemId", MtlItemId);
                    dynamicParameters.Add("ControlId", CadControlId);
                    //BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemId", @" AND a.MtlItemId = @MtlItemId", MtlItemId);
                    //BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcItemId", @" OR a.QcItemId = @QcItemId", QcItemId);
                    //BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ControlId", @" AND d.ControlId = @ControlId", CadControlId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "BallMark", @" AND a.BallMark = + @BallMark", BallMark);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Design", @" AND a.DesignValue LIKE '%' + @Design + '%'", Design);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Usl", @" AND a.UpperTolerance LIKE '%' + @Usl + '%'", Usl);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Lsl", @" AND a.LowerTolerance LIKE '%' + @Lsl + '%'", Lsl);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "e.QcItemNo, a.ArcdId DESC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;
                    sqlQuery.distinct = false;
                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

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

        #region//GetAutoReadCadDesignIsDid -- 取得自動讀CAD圖規格 --WuTc 2024.09.07
        public string GetAutoReadCadDesignIsDid(int CompanyId, int MtlItemId, string VersionNo)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT TOP 1 *
	                        FROM PDM.AutoReadCadDesign 
	                        WHERE MtlItemId = @MtlItemId
                            AND VersionNo = @VersionNo";
                    
                    dynamicParameters.Add("MtlItemId", MtlItemId);
                    dynamicParameters.Add("VersionNo", VersionNo);
                    
                    var result = sqlConnection.Query(sql, dynamicParameters);

                    if (result.Count() > 0) throw new SystemException("此品號綁定的圖號版本已有解析規格！");

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
        #endregion

        #region //Add
        #region //AddCustomerCad-- 新增客戶設計圖 -- Ann 2022.06.10
        public string AddCustomerCad(string CustomerMtlItemNo, string CustomerDwgNo)
        {
            try
            {
                if (CustomerMtlItemNo.Length <= 0) throw new SystemException("【部番(客戶品號)】不能為空!");
                if (CustomerMtlItemNo.Length > 50) throw new SystemException("【部番(客戶品號)】長度錯誤!");
                if (CustomerDwgNo.Length <= 0) throw new SystemException("【客戶圖號】不能為空!");
                if (CustomerDwgNo.Length > 50) throw new SystemException("【客戶圖號】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        CustomerMtlItemNo = CustomerMtlItemNo.Trim();

                        #region //判斷 部番+客戶圖號 是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.CustomerCad
                                WHERE CompanyId = @CompanyId
                                AND CustomerMtlItemNo = @CustomerMtlItemNo
                                AND CustomerDwgNo = @CustomerDwgNo";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("CustomerMtlItemNo", CustomerMtlItemNo);
                        dynamicParameters.Add("CustomerDwgNo", CustomerDwgNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【部番(客戶品號) + 客戶圖號】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PDM.CustomerCad (CompanyId, CustomerMtlItemNo, CustomerDwgNo
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.CadId
                                VALUES (@CompanyId, @CustomerMtlItemNo, @CustomerDwgNo
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                CustomerMtlItemNo,
                                CustomerDwgNo,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
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

        #region //AddCustomerCadControl-- 新增客戶設計圖版本 -- Ann 2022.06.10
        public string AddCustomerCadControl(int CadId, string Edition, string EditionType, string Cause
            , string ReleasedStatus, int CadFile, string OtherFile)
        {
            try
            {
                if (Edition.Length <= 0) throw new SystemException("【圖面版次】不能為空!");
                if (Edition.Length > 32) throw new SystemException("【圖面版次】長度錯誤!");
                if (!Regex.IsMatch(EditionType, "^(E|U)$", RegexOptions.IgnoreCase)) throw new SystemException("【變更類型】錯誤!");
                if (Cause.Length <= 0) throw new SystemException("【變更原因】不能為空!");
                if (Cause.Length > 100) throw new SystemException("【變更原因】長度錯誤!");
                if (!Regex.IsMatch(ReleasedStatus, "^(Y|N)$", RegexOptions.IgnoreCase)) throw new SystemException("【發行版本】錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷客戶設計圖資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.CustomerCad
                                WHERE CadId = @CadId";
                        dynamicParameters.Add("CadId", CadId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("客戶設計圖資料錯誤!");
                        #endregion

                        #region //判斷變更類型，並確認版次
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.CustomerCadControl
                                WHERE CadId = @CadId
                                AND Edition = @Edition";
                        dynamicParameters.Add("CadId", CadId);
                        dynamicParameters.Add("Edition", Edition);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);

                        switch (EditionType)
                        {
                            case "E":                                
                                if (result2.Count() > 0) throw new SystemException("【圖面版次】已使用，請重新填寫!");
                                break;
                            case "U":
                                if (result2.Count() <= 0) throw new SystemException("【圖面版次】不正確，請重新填寫!");
                                break;
                        }

                        #region //尋找目前系統版次
                        string CurrentVersion = "";

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ISNULL(MAX(Version), 0) CurrentVersion
                                FROM PDM.CustomerCadControl
                                WHERE CadId = @CadId
                                AND Edition = @Edition";
                        dynamicParameters.Add("CadId", CadId);
                        dynamicParameters.Add("Edition", Edition);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        foreach (var item in result3)
                        {
                            CurrentVersion = (Convert.ToInt32(item.CurrentVersion) + 1).ToString("00");
                        }
                        #endregion
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PDM.CustomerCadControl (CadId, CadFile, Edition, Version, Cause, ReleasedStatus
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ControlId
                                VALUES (@CadId, @CadFile, @Edition, @Version, @Cause, @ReleasedStatus
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CadId,
                                CadFile = CadFile > 0 ? (int?)CadFile : null,
                                Edition,
                                Version = CurrentVersion,
                                Cause,
                                ReleasedStatus,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();

                        #region //上傳其他檔案
                        if (OtherFile.Length > 0)
                        {
                            int ControlId = 0;
                            foreach (var item in insertResult)
                            {
                                ControlId = Convert.ToInt32(item.ControlId);
                            }

                            string[] otherFiles = OtherFile.Split(',');
                            foreach (var file in otherFiles)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO PDM.CustomerCadControlFile (ControlId, FileId
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@ControlId, @FileId
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        ControlId,
                                        FileId = file,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                var tempResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();
                            }
                        }
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
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

        #region //AddCustomerCadControl02-- 新增客戶設計圖版本02(圖檔上傳方式更改) -- Ann 2023-06-19
        public string AddCustomerCadControl02(int CadId, string Edition, string EditionType, string Cause
            , string ReleasedStatus, string CadFilePath, string OtherFile, string ServerPath)
        {
            try
            {
                if (Edition.Length <= 0) throw new SystemException("【圖面版次】不能為空!");
                if (Edition.Length > 32) throw new SystemException("【圖面版次】長度錯誤!");
                if (!Regex.IsMatch(EditionType, "^(E|U)$", RegexOptions.IgnoreCase)) throw new SystemException("【變更類型】錯誤!");
                if (Cause.Length <= 0) throw new SystemException("【變更原因】不能為空!");
                if (Cause.Length > 100) throw new SystemException("【變更原因】長度錯誤!");
                if (CadFilePath.Length <= 0) throw new SystemException("【客戶設計圖】不能為空!!");
                if (!Regex.IsMatch(ReleasedStatus, "^(Y|N)$", RegexOptions.IgnoreCase)) throw new SystemException("【發行版本】錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CompanyNo FROM BAS.Company
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var CompanyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (CompanyResult.Count() <= 0) throw new SystemException("公司別資料錯誤!!");

                        string CompanyNo = "";
                        foreach (var item in CompanyResult)
                        {
                            CompanyNo = item.CompanyNo;
                        }
                        #endregion

                        #region //判斷客戶設計圖資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CustomerMtlItemNo, CustomerDwgNo
                                FROM PDM.CustomerCad
                                WHERE CadId = @CadId";
                        dynamicParameters.Add("CadId", CadId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("客戶設計圖資料錯誤!");

                        string CustomerMtlItemNo = "";
                        string CustomerDwgNo = "";
                        foreach (var item in result)
                        {
                            CustomerMtlItemNo = item.CustomerMtlItemNo;
                            CustomerDwgNo = item.CustomerDwgNo;
                        }
                        #endregion

                        #region //判斷變更類型，並確認版次
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.CustomerCadControl
                                WHERE CadId = @CadId
                                AND Edition = @Edition";
                        dynamicParameters.Add("CadId", CadId);
                        dynamicParameters.Add("Edition", Edition);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);

                        switch (EditionType)
                        {
                            case "E":
                                if (result2.Count() > 0) throw new SystemException("【圖面版次】已使用，請重新填寫!");
                                break;
                            case "U":
                                if (result2.Count() <= 0) throw new SystemException("【圖面版次】不正確，請重新填寫!");
                                break;
                        }

                        #region //尋找目前系統版次
                        string CurrentVersion = "";

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ISNULL(MAX(Version), 0) CurrentVersion
                                FROM PDM.CustomerCadControl
                                WHERE CadId = @CadId
                                AND Edition = @Edition";
                        dynamicParameters.Add("CadId", CadId);
                        dynamicParameters.Add("Edition", Edition);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        foreach (var item in result3)
                        {
                            CurrentVersion = (Convert.ToInt32(item.CurrentVersion) + 1).ToString("00");
                        }
                        #endregion
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PDM.CustomerCadControl (CadId, CadFile, CadFileAbsolutePath, Edition, Version, Cause, ReleasedStatus
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ControlId
                                VALUES (@CadId, @CadFile, @CadFileAbsolutePath, @Edition, @Version, @Cause, @ReleasedStatus
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CadId,
                                CadFile = (int?)null,
                                CadFileAbsolutePath = "",
                                Edition,
                                Version = CurrentVersion,
                                Cause,
                                ReleasedStatus,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();

                        int ControlId = -1;
                        foreach (var item in insertResult)
                        {
                            ControlId = item.ControlId;
                        }

                        #region //整理上傳路徑
                        string FolderPath = Path.Combine(ServerPath, CompanyNo, "Customer", CustomerMtlItemNo, "(" + CustomerDwgNo + ")");
                        #endregion

                        #region //將客戶設計圖檔案複製至共夾
                        string ASCIIPath = "";
                        string FilePath = "";

                        #region //反處理URL特殊符號
                        ASCIIPath = CadFilePath;
                        if (CadFilePath.IndexOf("%2B") != -1) ASCIIPath = ASCIIPath.Replace("%2B", "+");
                        if (CadFilePath.IndexOf("%2F") != -1) ASCIIPath = ASCIIPath.Replace("%2F", "/");
                        if (CadFilePath.IndexOf("%3F") != -1) ASCIIPath = ASCIIPath.Replace("%3F", "?");
                        if (CadFilePath.IndexOf("%23") != -1) ASCIIPath = ASCIIPath.Replace("%23", "#");
                        if (CadFilePath.IndexOf("%26") != -1) ASCIIPath = ASCIIPath.Replace("%26", "&");
                        if (CadFilePath.IndexOf("%3D") != -1) ASCIIPath = ASCIIPath.Replace("%3D", "=");
                        #endregion

                        if (File.Exists(ASCIIPath))
                        {
                            string FileName = Path.GetFileNameWithoutExtension(CadFilePath);
                            string FileExtension = Path.GetExtension(CadFilePath);
                            FilePath = Path.Combine(FolderPath, FileName + "-" + ControlId + FileExtension);
                            if (!Directory.Exists(FolderPath)) { Directory.CreateDirectory(FolderPath); }
                            byte[] cadFileByte = File.ReadAllBytes(ASCIIPath);
                            File.WriteAllBytes(FilePath, cadFileByte);
                        }
                        else
                        {
                            throw new SystemException("客戶設計圖檔案路徑錯誤!!");
                        }
                        #endregion

                        #region //將客戶設計圖路徑存回Control Table
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.CustomerCadControl SET
                                CadFileAbsolutePath = @CadFileAbsolutePath,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ControlId = @ControlId";
                        var parametersObject = new
                        {
                            //CadFileAbsolutePath = HttpUtility.UrlEncode(FilePath),
                            CadFileAbsolutePath = FilePath,
                            LastModifiedDate,
                            LastModifiedBy,
                            ControlId
                        };
                        dynamicParameters.AddDynamicParams(parametersObject);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //上傳其他檔案
                        if (OtherFile.Length > 0)
                        {
                            string[] otherFiles = OtherFile.Split(',');
                            foreach (var file in otherFiles)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO PDM.CustomerCadControlFile (ControlId, FileId
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@ControlId, @FileId
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        ControlId,
                                        FileId = file,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                var tempResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();
                            }
                        }
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
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

        #region //AddRdDesign-- 新增研發設計圖 -- Ann 2022.06.27
        public string AddRdDesign(string CustomerMtlItemNo, int CustomerCadControlId, int MtlItemId)
        {
            try
            {
                if (CustomerMtlItemNo.Length <= 0) throw new SystemException("【部番(客戶品號)】不能為空!");
                if (CustomerMtlItemNo.Length > 50) throw new SystemException("【部番(客戶品號)】長度錯誤!");
                CustomerMtlItemNo = CustomerMtlItemNo.Trim();
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別DB
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var companyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion

                        string mtlItemNo = "";
                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //判斷【部番(客戶品號)】是否重複
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM PDM.RdDesign
                                    WHERE CompanyId = @CompanyId
                                    AND CustomerMtlItemNo = @CustomerMtlItemNo";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("CustomerMtlItemNo", CustomerMtlItemNo);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            //if (result.Count() > 0) throw new SystemException("【部番(客戶品號)】重複，請重新輸入!");
                            #endregion

                            if (MtlItemId > 0)
                            {
                                #region //取得品號資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MtlItemNo, a.TransferStatus
	                                    , ISNULL(EffectiveDate, '') as EffectiveDate
	                                    , ISNULL(ExpirationDate, '') as ExpirationDate
                                        FROM PDM.MtlItem a
                                        WHERE a.MtlItemId = @MtlItemId";
                                dynamicParameters.Add("MtlItemId", MtlItemId);
                                var MtlItemResult = sqlConnection.Query(sql, dynamicParameters);

                                #region //確認品號資料是否正確
                                if (MtlItemResult.Count() <= 0) throw new SystemException("品號資料錯誤!!");
                                #endregion

                                foreach (var item in MtlItemResult)
                                {
                                    mtlItemNo = item.MtlItemNo;

                                    #region //判斷品號是否拋轉ERP
                                    if (item.TransferStatus != "Y") throw new SystemException("此品號尚未拋轉ERP!!");
                                    #endregion

                                    #region //判斷品號是否已經失效
                                    var effectiveDate = item.EffectiveDate.ToString("yyyy-MM-dd");
                                    if (effectiveDate != "1900-01-01" && CreateDate.CompareTo(item.EffectiveDate) < 0)
                                        throw new SystemException(string.Format("【品號】</br>{0}</br>{1}生效</br>目前無法使用!", mtlItemNo, effectiveDate));

                                    var expirationDate = item.ExpirationDate.ToString("yyyy-MM-dd");
                                    if (expirationDate != "1900-01-01" && CreateDate.CompareTo(item.ExpirationDate) > 0)
                                        throw new SystemException(string.Format("【品號】</br>{0}</br>{1}到期</br>已無法使用!", mtlItemNo, expirationDate));
                                    #endregion
                                }
                                #endregion

                                #region //判斷ERP品號生效日與失效日
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT LTRIM(RTRIM(MB030)) MB030, LTRIM(RTRIM(MB031)) MB031
                                        FROM INVMB
                                        WHERE MB001 = @MB001";
                                dynamicParameters.Add("MB001", mtlItemNo);

                                var INVMBResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (INVMBResult.Count() <= 0) throw new SystemException("此品號不存在於ERP中!!");

                                foreach (var item in INVMBResult)
                                {
                                    if (item.MB030 != "" && item.MB030 != null)
                                    {
                                        #region //判斷日期需大於或等於生效日
                                        string EffectiveDate = item.MB030;
                                        string effYear = EffectiveDate.Substring(0, 4);
                                        string effMonth = EffectiveDate.Substring(4, 2);
                                        string effDay = EffectiveDate.Substring(6, 2);
                                        DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                        int effresult = DateTime.Compare(CreateDate, effFullDate);
                                        if (effresult < 0) throw new SystemException("此品號尚未生效，無法使用!!");
                                        #endregion
                                    }

                                    if (item.MB031 != "" && item.MB031 != null)
                                    {
                                        #region //判斷日期需小於或等於失效日
                                        string ExpirationDate = item.MB031;
                                        string effYear = ExpirationDate.Substring(0, 4);
                                        string effMonth = ExpirationDate.Substring(4, 2);
                                        string effDay = ExpirationDate.Substring(6, 2);
                                        DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                        int effresult = DateTime.Compare(CreateDate, effFullDate);
                                        if (effresult > 0) throw new SystemException("此品號已失效，無法使用!!");
                                        #endregion
                                    }
                                }
                                #endregion
                            }

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO PDM.RdDesign (CompanyId, CustomerMtlItemNo, CustomerCadControlId, MtlItemId
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.DesignId
                                    VALUES (@CompanyId, @CustomerMtlItemNo, @CustomerCadControlId, @MtlItemId
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    CompanyId = CurrentCompany,
                                    CustomerMtlItemNo,
                                    CustomerCadControlId = CustomerCadControlId > 0 ? (int?)CustomerCadControlId : null,
                                    MtlItemId = MtlItemId > 0 ? (int?)MtlItemId : null,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            int rowsAffected = insertResult.Count();

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                msg = "(" + rowsAffected + " rows affected)"
                            });
                            #endregion
                        }
                    }

                    transactionScope.Complete();
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

        #region //AddRdDesignControl-- 新增研發設計圖版本 -- Ann 2022.06.28
        public string AddRdDesignControl(int DesignId, string Edition, string EditionType, string Cause
            , string DesignDate, string ReleasedStatus, int Cad3DFile, int Cad2DFile, int Pdf2DFile, int JmoFile)
        {
            try
            {
                if (Edition.Length <= 0) throw new SystemException("【圖面版次】不能為空!");
                if (Edition.Length > 32) throw new SystemException("【圖面版次】長度錯誤!");
                if (!Regex.IsMatch(EditionType, "^(E|U)$", RegexOptions.IgnoreCase)) throw new SystemException("【變更類型】錯誤!");
                if (Cause.Length <= 0) throw new SystemException("【變更原因】不能為空!");
                if (Cause.Length > 100) throw new SystemException("【變更原因】長度錯誤!");
                if (DesignDate.Length <=0) throw new SystemException("【出圖日期】不能為空!");
                if (!Regex.IsMatch(ReleasedStatus, "^(Y|N)$", RegexOptions.IgnoreCase)) throw new SystemException("【發行版本】錯誤!");
                if (Cad2DFile <= 0) throw new SystemException("【2D研發設計圖】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷客戶設計圖資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.RdDesign
                                WHERE DesignId = @DesignId";
                        dynamicParameters.Add("DesignId", DesignId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("研發設計圖資料錯誤!");
                        #endregion

                        #region //判斷變更類型，並確認版次
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.RdDesignControl
                                WHERE DesignId = @DesignId
                                AND Edition = @Edition";
                        dynamicParameters.Add("DesignId", DesignId);
                        dynamicParameters.Add("Edition", Edition);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);

                        switch (EditionType)
                        {
                            case "E":
                                if (result2.Count() > 0) throw new SystemException("【圖面版次】已使用，請重新填寫!");
                                break;
                            case "U":
                                if (result2.Count() <= 0) throw new SystemException("【圖面版次】不正確，請重新填寫!");
                                break;
                        }

                        #region //尋找目前系統版次
                        string CurrentVersion = "";

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ISNULL(MAX(Version), 0) CurrentVersion
                                FROM PDM.RdDesignControl
                                WHERE DesignId = @DesignId
                                AND Edition = @Edition";
                        dynamicParameters.Add("DesignId", DesignId);
                        dynamicParameters.Add("Edition", Edition);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        foreach (var item in result3)
                        {
                            CurrentVersion = (Convert.ToInt32(item.CurrentVersion) + 1).ToString("00");
                        }
                        #endregion
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PDM.RdDesignControl (DesignId, Edition, Version, Cause
                                , DesignDate, ReleasedStatus
                                , Cad3DFile, Cad2DFile, Pdf2DFile, JmoFile
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ControlId
                                VALUES (@DesignId, @Edition, @Version, @Cause, @DesignDate, @ReleasedStatus
                                , @Cad3DFile, @Cad2DFile, @Pdf2DFile, @JmoFile
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DesignId,
                                Edition,
                                Version = CurrentVersion,
                                Cause,
                                DesignDate,
                                ReleasedStatus,
                                Cad3DFile = Cad3DFile > 0 ? (int?)Cad3DFile : null,
                                Cad2DFile = Cad2DFile > 0 ? (int?)Cad2DFile : null,
                                Pdf2DFile = Pdf2DFile > 0 ? (int?)Pdf2DFile : null,
                                JmoFile = JmoFile > 0 ? (int?)JmoFile : null,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
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

        #region //AddRdDesignControl02-- 新增研發設計圖版本02(圖檔上傳方式更改) -- Ann 2023-06-26
        public string AddRdDesignControl02(int DesignId, string Edition, string EditionType, string Cause, string DesignDate, string ReleasedStatus
            , string Cad3DFileAbsolutePath, string Cad2DFileAbsolutePath, string Pdf2DFileAbsolutePath, string JmoFileAbsolutePath, string ServerPath)
        {
            try
            {
                if (Edition.Length <= 0) throw new SystemException("【圖面版次】不能為空!");
                if (Edition.Length > 32) throw new SystemException("【圖面版次】長度錯誤!");
                if (!Regex.IsMatch(EditionType, "^(E|U)$", RegexOptions.IgnoreCase)) throw new SystemException("【變更類型】錯誤!");
                if (Cause.Length <= 0) throw new SystemException("【變更原因】不能為空!");
                if (Cause.Length > 100) throw new SystemException("【變更原因】長度錯誤!");
                if (DesignDate.Length <= 0) throw new SystemException("【出圖日期】不能為空!");
                if (!Regex.IsMatch(ReleasedStatus, "^(Y|N)$", RegexOptions.IgnoreCase)) throw new SystemException("【發行版本】錯誤!");
                if (Cad2DFileAbsolutePath.Length <= 0) throw new SystemException("【2D研發設計圖】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CompanyNo FROM BAS.Company
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var CompanyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (CompanyResult.Count() <= 0) throw new SystemException("公司別資料錯誤!!");

                        string CompanyNo = "";
                        foreach (var item in CompanyResult)
                        {
                            CompanyNo = item.CompanyNo;
                        }
                        #endregion

                        #region //判斷研發設計圖資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ISNULL(b.MtlItemNo, '') MtlItemNo
                                FROM PDM.RdDesign a
                                LEFT JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                WHERE a.DesignId = @DesignId";
                        dynamicParameters.Add("DesignId", DesignId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("研發設計圖資料錯誤!");

                        string MtlItemNo = "";
                        foreach (var item in result)
                        {
                            MtlItemNo = item.MtlItemNo;
                        }
                        #endregion

                        #region //判斷變更類型，並確認版次
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.RdDesignControl
                                WHERE DesignId = @DesignId
                                AND Edition = @Edition";
                        dynamicParameters.Add("DesignId", DesignId);
                        dynamicParameters.Add("Edition", Edition);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);

                        switch (EditionType)
                        {
                            case "E":
                                if (result2.Count() > 0) throw new SystemException("【圖面版次】已使用，請重新填寫!");
                                break;
                            case "U":
                                if (result2.Count() <= 0) throw new SystemException("【圖面版次】不正確，請重新填寫!");
                                break;
                        }

                        #region //尋找目前系統版次
                        string CurrentVersion = "";

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ISNULL(MAX(Version), 0) CurrentVersion
                                FROM PDM.RdDesignControl
                                WHERE DesignId = @DesignId
                                AND Edition = @Edition";
                        dynamicParameters.Add("DesignId", DesignId);
                        dynamicParameters.Add("Edition", Edition);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        foreach (var item in result3)
                        {
                            CurrentVersion = (Convert.ToInt32(item.CurrentVersion) + 1).ToString("00");
                        }
                        #endregion
                        #endregion

                        #region //INSERT PDM.RdDesignControl
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PDM.RdDesignControl (DesignId, Edition, Version, Cause
                                , DesignDate, ReleasedStatus
                                , Cad3DFileAbsolutePath, Cad2DFileAbsolutePath, Pdf2DFileAbsolutePath, JmoFileAbsolutePath
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ControlId
                                VALUES (@DesignId, @Edition, @Version, @Cause, @DesignDate, @ReleasedStatus
                                , @Cad3DFileAbsolutePath, @Cad2DFileAbsolutePath, @Pdf2DFileAbsolutePath, @JmoFileAbsolutePath
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DesignId,
                                Edition,
                                Version = CurrentVersion,
                                Cause,
                                DesignDate,
                                ReleasedStatus,
                                Cad3DFileAbsolutePath = "",
                                Cad2DFileAbsolutePath = "",
                                Pdf2DFileAbsolutePath = "",
                                JmoFileAbsolutePath = "",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();

                        int ControlId = -1;
                        foreach (var item in insertResult)
                        {
                            ControlId = item.ControlId;
                        }
                        #endregion

                        #region //整理上傳路徑
                        if (MtlItemNo == "") MtlItemNo = "暫無品號-" + ControlId.ToString();
                        string FolderPath = Path.Combine(ServerPath, CompanyNo, "RdDesign", MtlItemNo);
                        #endregion

                        string ASCIIPath = "";
                        #region //將3D研發設計圖檔案複製至共夾
                        string Cad3DFilePath = "";
                        if (Cad3DFileAbsolutePath.Length > 0)
                        {
                            ASCIIPath = Cad3DFileAbsolutePath;
                            #region //反處理URL特殊符號
                            if (Cad3DFileAbsolutePath.IndexOf("%2B") != -1) ASCIIPath = ASCIIPath.Replace("%2B", "+");
                            if (Cad3DFileAbsolutePath.IndexOf("%2F") != -1) ASCIIPath = ASCIIPath.Replace("%2F", "/");
                            if (Cad3DFileAbsolutePath.IndexOf("%3F") != -1) ASCIIPath = ASCIIPath.Replace("%3F", "?");
                            if (Cad3DFileAbsolutePath.IndexOf("%23") != -1) ASCIIPath = ASCIIPath.Replace("%23", "#");
                            if (Cad3DFileAbsolutePath.IndexOf("%26") != -1) ASCIIPath = ASCIIPath.Replace("%26", "&");
                            if (Cad3DFileAbsolutePath.IndexOf("%3D") != -1) ASCIIPath = ASCIIPath.Replace("%3D", "=");
                            #endregion

                            if (File.Exists(ASCIIPath))
                            {
                                string FileName = Path.GetFileNameWithoutExtension(Cad3DFileAbsolutePath);
                                string FileExtension = Path.GetExtension(Cad3DFileAbsolutePath);
                                Cad3DFilePath = Path.Combine(FolderPath, FileName + "-" + ControlId + FileExtension);
                                if (!Directory.Exists(FolderPath)) { Directory.CreateDirectory(FolderPath); }
                                byte[] cadFileByte = File.ReadAllBytes(ASCIIPath);
                                File.WriteAllBytes(Cad3DFilePath, cadFileByte);
                            }
                            else
                            {
                                throw new SystemException("3D研發設計圖路徑有誤!!");
                            }
                        }
                        #endregion

                        #region //將2D研發設計圖檔案複製至共夾
                        string Cad2DFilePath = "";
                        if (Cad2DFileAbsolutePath.Length > 0)
                        {
                            ASCIIPath = Cad2DFileAbsolutePath;
                            #region //反處理URL特殊符號
                            if (Cad2DFileAbsolutePath.IndexOf("%2B") != -1) ASCIIPath = ASCIIPath.Replace("%2B", "+");
                            if (Cad2DFileAbsolutePath.IndexOf("%2F") != -1) ASCIIPath = ASCIIPath.Replace("%2F", "/");
                            if (Cad2DFileAbsolutePath.IndexOf("%3F") != -1) ASCIIPath = ASCIIPath.Replace("%3F", "?");
                            if (Cad2DFileAbsolutePath.IndexOf("%23") != -1) ASCIIPath = ASCIIPath.Replace("%23", "#");
                            if (Cad2DFileAbsolutePath.IndexOf("%26") != -1) ASCIIPath = ASCIIPath.Replace("%26", "&");
                            if (Cad2DFileAbsolutePath.IndexOf("%3D") != -1) ASCIIPath = ASCIIPath.Replace("%3D", "=");
                            #endregion

                            if (File.Exists(ASCIIPath))
                            {
                                string FileName = Path.GetFileNameWithoutExtension(Cad2DFileAbsolutePath);
                                string FileExtension = Path.GetExtension(Cad2DFileAbsolutePath);
                                Cad2DFilePath = Path.Combine(FolderPath, FileName + "-" + ControlId + FileExtension);
                                if (!Directory.Exists(FolderPath)) { Directory.CreateDirectory(FolderPath); }
                                byte[] cadFileByte = File.ReadAllBytes(ASCIIPath);
                                File.WriteAllBytes(Cad2DFilePath, cadFileByte);
                            }
                            else
                            {
                                throw new SystemException("2D研發設計圖檔案路徑錯誤!!");
                            }
                        }
                        else
                        {
                            throw new SystemException("2D研發設計圖路徑不能為空!!");
                        }
                        #endregion

                        #region //將Pdf2D檔案複製至共夾
                        string Pdf2DFilePath = "";
                        if (Pdf2DFileAbsolutePath.Length > 0)
                        {
                            ASCIIPath = Pdf2DFileAbsolutePath;
                            #region //反處理URL特殊符號
                            if (Pdf2DFileAbsolutePath.IndexOf("%2B") != -1) ASCIIPath = ASCIIPath.Replace("%2B", "+");
                            if (Pdf2DFileAbsolutePath.IndexOf("%2F") != -1) ASCIIPath = ASCIIPath.Replace("%2F", "/");
                            if (Pdf2DFileAbsolutePath.IndexOf("%3F") != -1) ASCIIPath = ASCIIPath.Replace("%3F", "?");
                            if (Pdf2DFileAbsolutePath.IndexOf("%23") != -1) ASCIIPath = ASCIIPath.Replace("%23", "#");
                            if (Pdf2DFileAbsolutePath.IndexOf("%26") != -1) ASCIIPath = ASCIIPath.Replace("%26", "&");
                            if (Pdf2DFileAbsolutePath.IndexOf("%3D") != -1) ASCIIPath = ASCIIPath.Replace("%3D", "=");
                            #endregion

                            if (File.Exists(ASCIIPath))
                            {
                                string FileName = Path.GetFileNameWithoutExtension(Pdf2DFileAbsolutePath);
                                string FileExtension = Path.GetExtension(Pdf2DFileAbsolutePath);
                                Pdf2DFilePath = Path.Combine(FolderPath, FileName + "-" + ControlId + FileExtension);
                                if (!Directory.Exists(FolderPath)) { Directory.CreateDirectory(FolderPath); }
                                byte[] cadFileByte = File.ReadAllBytes(ASCIIPath);
                                File.WriteAllBytes(Pdf2DFilePath, cadFileByte);
                            }
                            else
                            {
                                throw new SystemException("Pdf2D檔案路徑錯誤!!");
                            }
                        }
                        #endregion

                        #region //將Jmo檔案複製至共夾
                        string JmoFilePath = "";
                        if (JmoFileAbsolutePath.Length > 0)
                        {
                            ASCIIPath = JmoFileAbsolutePath;
                            #region //反處理URL特殊符號
                            if (JmoFileAbsolutePath.IndexOf("%2B") != -1) ASCIIPath = ASCIIPath.Replace("%2B", "+");
                            if (JmoFileAbsolutePath.IndexOf("%2F") != -1) ASCIIPath = ASCIIPath.Replace("%2F", "/");
                            if (JmoFileAbsolutePath.IndexOf("%3F") != -1) ASCIIPath = ASCIIPath.Replace("%3F", "?");
                            if (JmoFileAbsolutePath.IndexOf("%23") != -1) ASCIIPath = ASCIIPath.Replace("%23", "#");
                            if (JmoFileAbsolutePath.IndexOf("%26") != -1) ASCIIPath = ASCIIPath.Replace("%26", "&");
                            if (JmoFileAbsolutePath.IndexOf("%3D") != -1) ASCIIPath = ASCIIPath.Replace("%3D", "=");
                            #endregion

                            if (File.Exists(ASCIIPath))
                            {
                                string FileName = Path.GetFileNameWithoutExtension(JmoFileAbsolutePath);
                                string FileExtension = Path.GetExtension(JmoFileAbsolutePath);
                                JmoFilePath = Path.Combine(FolderPath, FileName + "-" + ControlId + FileExtension);
                                if (!Directory.Exists(FolderPath)) { Directory.CreateDirectory(FolderPath); }
                                byte[] cadFileByte = File.ReadAllBytes(ASCIIPath);
                                File.WriteAllBytes(JmoFilePath, cadFileByte);
                            }
                            else
                            {
                                throw new SystemException("JMO檔案路徑錯誤!!");
                            }
                        }
                        #endregion

                        #region //將檔案路徑存回Control Table
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.RdDesignControl SET
                                Cad3DFileAbsolutePath = @Cad3DFileAbsolutePath,
                                Cad2DFileAbsolutePath = @Cad2DFileAbsolutePath,
                                Pdf2DFileAbsolutePath = @Pdf2DFileAbsolutePath,
                                JmoFileAbsolutePath = @JmoFileAbsolutePath,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ControlId = @ControlId";
                        var parametersObject = new
                        {
                            Cad3DFileAbsolutePath = Cad3DFilePath != "" ? Cad3DFilePath : null,
                            Cad2DFileAbsolutePath = Cad2DFilePath != "" ? Cad2DFilePath : null,
                            Pdf2DFileAbsolutePath = Pdf2DFilePath != "" ? Pdf2DFilePath : null,
                            JmoFileAbsolutePath = JmoFilePath != "" ? JmoFilePath : null,
                            LastModifiedDate,
                            LastModifiedBy,
                            ControlId
                        };
                        dynamicParameters.AddDynamicParams(parametersObject);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
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

        #region//AddAutoReadCadDesign --新增自動讀CAD圖規格資訊
        public string AddAutoReadCadDesign(int MtlItemId, int ControlId, string Edition, int QcItemId, string QcItemNo, string QcItemName, string QcItemDesc, string DesignValue, string UpperTolerance, string LowerTolerance
            , string BallMark, string Unit, string Remark)
        {
            try
            {
                if (MtlItemId== -1) throw new SystemException("查無【品號】!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷 品號 是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.MtlItem
                                WHERE CompanyId = @CompanyId
                                AND MtlItemId = @MtlItemId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("MtlItemId", MtlItemId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() != 1) throw new SystemException("查無【品號】");
                        #endregion

                        #region //取得量測項目資訊
                        if (QcItemId != -1 || QcItemNo != "")
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.QcItemId, a.QcItemNo, a.QcItemName, a.QcItemDesc
                                FROM QMS.QcItem a
                                INNER JOIN QMS.QcClass b ON a.QcClassId = b.QcClassId
                                INNER JOIN QMS.QcGroup c ON b.QcGroupId = c.QcGroupId
                                WHERE c.CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcItemNo", @" AND a.QcItemNo = @QcItemNo", QcItemNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcItemId", @" AND a.QcItemId = @QcItemId", QcItemId);
                            var resultQcItem = sqlConnection.Query(sql, dynamicParameters);
                            if (resultQcItem.Count() > 0)
                            {
                                foreach (var item in resultQcItem)
                                {
                                    QcItemId = item.QcItemId;
                                    QcItemNo = item.QcItemNo;
                                    QcItemName = item.QcItemName;
                                    QcItemDesc = item.QcItemDesc;
                                }
                            }
                        }                        
                        #endregion

                        #region //判斷 規格 是否存在，若品號+版本+公司+量測項目存在，則不新增
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
	                            FROM PDM.AutoReadCadDesign a
	                            INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
	                            INNER JOIN PDM.RdDesign c ON b.MtlItemId = c.MtlItemId
	                            INNER JOIN PDM.RdDesignControl d ON c.DesignId = d.DesignId
                                WHERE b.CompanyId = @CompanyId
                                AND a.MtlItemId = @MtlItemId
                                AND d.ControlId = @MtlItemId
                                AND a.QcItemId = @QcItemId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("MtlItemId", MtlItemId);
                        dynamicParameters.Add("ControlId", ControlId);
                        dynamicParameters.Add("QcItemId", QcItemId);
                        var resultJduge = sqlConnection.Query(sql, dynamicParameters);
                        if (resultJduge.Count() == 1) throw new SystemException("此【量測項目：" + QcItemNo + "】已存在規格，請選擇另一個【量測項目】");
                        #endregion

                        #region //查詢系統版次
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MtlItemId, c.Edition, c.[Version]
                                FROM PDM.MtlItem a
	                            INNER JOIN PDM.RdDesign b ON a.MtlItemId = b.MtlItemId
	                            INNER JOIN PDM.RdDesignControl c ON b.DesignId = c.DesignId
                                WHERE 1=1";

                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemId", @" AND a.MtlItemId = @MtlItemId", MtlItemId);
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ControlId", @" AND c.ControlId = @ControlId", ControlId);
                        if (Edition != "") BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Edition", @" AND c.Edition = @Edition", Edition);
                        var resultVersion = sqlConnection.Query(sql, dynamicParameters);
                        string VersionNo = "";
                        if (resultVersion.Count() > 0)
                        {
                            foreach(var item in resultVersion)
                            {
                                VersionNo = item.Version;
                            }
                        }
                        else
                        {
                            throw new SystemException("查無該圖面【系統版次" + VersionNo + "】");
                        }
                        #endregion

                        #region //INSERT Design
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PDM.AutoReadCadDesign (MtlItemId, VersionNo, QcItemId, QcItemName
                                , QcItemDesc, DesignValue, UpperTolerance, LowerTolerance
                                , BallMark, Unit, Remark
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ArcdId
                                VALUES (@MtlItemId, @VersionNo, @QcItemId, @QcItemName
                                , @QcItemDesc, @DesignValue, @UpperTolerance, @LowerTolerance
                                , @BallMark, @Unit, @Remark
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MtlItemId,
                                VersionNo,
                                QcItemId,
                                QcItemName,
                                QcItemDesc,
                                DesignValue,
                                UpperTolerance,
                                LowerTolerance,
                                BallMark,
                                Unit,
                                Remark,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
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

        #endregion

        #region //Update
        #region //UpdateCustomerCad -- 更新客戶設計圖資料 -- Ann 2022.06.09
        public string UpdateCustomerCad(int CadId, string CustomerMtlItemNo, string CustomerDwgNo)
        {
            try
            {
                if (CustomerMtlItemNo.Length <= 0) throw new SystemException("【部番(客戶品號)】不能為空!");
                if (CustomerMtlItemNo.Length > 50) throw new SystemException("【部番(客戶品號)】長度錯誤!");
                if (CustomerDwgNo.Length <= 0) throw new SystemException("【客戶圖號】不能為空!");
                if (CustomerDwgNo.Length > 50) throw new SystemException("【客戶圖號】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷客戶設計圖資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.CustomerCad
                                WHERE CadId = @CadId";
                        dynamicParameters.Add("CadId", CadId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("客戶設計圖資料錯誤!");
                        #endregion

                        #region //判斷 部番+客戶圖號 是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.CustomerCad
                                WHERE CompanyId = @CompanyId
                                AND CustomerMtlItemNo = @CustomerMtlItemNo
                                AND CustomerDwgNo = @CustomerDwgNo
                                AND CadId != @CadId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("CustomerMtlItemNo", CustomerMtlItemNo);
                        dynamicParameters.Add("CustomerDwgNo", CustomerDwgNo);
                        dynamicParameters.Add("CadId", CadId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【部番(客戶品號) + 客戶圖號】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.CustomerCad SET
                                CustomerMtlItemNo = @CustomerMtlItemNo,
                                CustomerDwgNo = @CustomerDwgNo,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE CadId = @CadId";
                        var parametersObject = new
                        {
                            CustomerMtlItemNo,
                            CustomerDwgNo,
                            LastModifiedDate,
                            LastModifiedBy,
                            CadId
                        };
                        dynamicParameters.AddDynamicParams(parametersObject);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
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

        #region //UpdateRdDesign -- 更新研發設計圖資料 -- Ann 2022.06.27
        public string UpdateRdDesign(int DesignId, string CustomerMtlItemNo, int CustomerCadControlId, int MtlItemId)
        {
            try
            {
                if (DesignId <= 0) throw new SystemException("【客戶設計圖編號】不能為空!");
                if (CustomerMtlItemNo.Length <= 0) throw new SystemException("【部番(客戶品號)】不能為空!");
                if (CustomerMtlItemNo.Length > 50) throw new SystemException("【部番(客戶品號)】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別DB
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var companyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //判斷客戶設計圖資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM PDM.RdDesign
                                    WHERE DesignId = @DesignId";
                            dynamicParameters.Add("DesignId", DesignId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("研發設計圖資料錯誤!");
                            #endregion

                            #region //判斷【部番(客戶品號)】是否重複
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM PDM.RdDesign
                                    WHERE CompanyId = @CompanyId
                                    AND CustomerMtlItemNo = @CustomerMtlItemNo
                                    AND DesignId != @DesignId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("CustomerMtlItemNo", CustomerMtlItemNo);
                            dynamicParameters.Add("DesignId", DesignId);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            //if (result2.Count() > 0) throw new SystemException("【部番(客戶品號)】重複，請重新輸入!");
                            #endregion

                            #region //判斷客戶設計圖版本資料是否正確
                            if (CustomerCadControlId > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM PDM.CustomerCadControl
                                        WHERE ControlId = @ControlId";
                                dynamicParameters.Add("ControlId", CustomerCadControlId);

                                var result3 = sqlConnection.Query(sql, dynamicParameters);
                                if (result3.Count() <= 0) throw new SystemException("客戶設計圖版本資料錯誤!");
                            }
                            #endregion

                            #region //判斷品號資料是否正確
                            if (MtlItemId > 0)
                            {
                                #region //確認品號資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MtlItemNo, a.TransferStatus
                                        FROM PDM.MtlItem a
                                        WHERE a.MtlItemId = @MtlItemId";
                                dynamicParameters.Add("MtlItemId", MtlItemId);

                                var MtlItemResult = sqlConnection.Query(sql, dynamicParameters);

                                if (MtlItemResult.Count() <= 0) throw new SystemException("品號資料錯誤!!");

                                string MtlItemNo = "";
                                foreach (var item in MtlItemResult)
                                {
                                    if (item.TransferStatus != "Y") throw new SystemException("此品號尚未拋轉ERP!!");
                                    MtlItemNo = item.MtlItemNo;
                                }
                                #endregion

                                #region //判斷ERP品號生效日與失效日
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT LTRIM(RTRIM(MB030)) MB030, LTRIM(RTRIM(MB031)) MB031
                                        FROM INVMB
                                        WHERE MB001 = @MB001";
                                dynamicParameters.Add("MB001", MtlItemNo);

                                var INVMBResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (INVMBResult.Count() <= 0) throw new SystemException("此品號不存在於ERP中!!");

                                foreach (var item in INVMBResult)
                                {
                                    if (item.MB030 != "" && item.MB030 != null)
                                    {
                                        #region //判斷日期需大於或等於生效日
                                        string EffectiveDate = item.MB030;
                                        string effYear = EffectiveDate.Substring(0, 4);
                                        string effMonth = EffectiveDate.Substring(4, 2);
                                        string effDay = EffectiveDate.Substring(6, 2);
                                        DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                        int effresult = DateTime.Compare(CreateDate, effFullDate);
                                        if (effresult < 0) throw new SystemException("此品號尚未生效，無法使用!!");
                                        #endregion
                                    }

                                    if (item.MB031 != "" && item.MB031 != null)
                                    {
                                        #region //判斷日期需小於或等於失效日
                                        string ExpirationDate = item.MB031;
                                        string effYear = ExpirationDate.Substring(0, 4);
                                        string effMonth = ExpirationDate.Substring(4, 2);
                                        string effDay = ExpirationDate.Substring(6, 2);
                                        DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                        int effresult = DateTime.Compare(CreateDate, effFullDate);
                                        if (effresult > 0) throw new SystemException("此品號已失效，無法使用!!");
                                        #endregion
                                    }
                                }
                                #endregion
                            }
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PDM.RdDesign SET
                                    CustomerMtlItemNo = @CustomerMtlItemNo,
                                    CustomerCadControlId = @CustomerCadControlId,
                                    MtlItemId = @MtlItemId,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE DesignId = @DesignId";
                            var parametersObject = new
                            {
                                CustomerMtlItemNo,
                                CustomerCadControlId = CustomerCadControlId > 0 ? CustomerCadControlId.ToString() : null,
                                MtlItemId = MtlItemId > 0 ? MtlItemId.ToString() : null,
                                LastModifiedDate,
                                LastModifiedBy,
                                DesignId
                            };
                            dynamicParameters.AddDynamicParams(parametersObject);

                            int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                msg = "(" + rowsAffected + " rows affected)"
                            });
                            #endregion
                        }
                    }

                    transactionScope.Complete();
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

        #region //UpdateCustomerCadControlReleasedStatus -- 更新客戶設計圖發行狀態 -- Ann 2022.06.28
        public string UpdateCustomerCadControlReleasedStatus(int ControlId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷客戶設計圖版本資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ReleasedStatus
                                FROM PDM.CustomerCadControl
                                WHERE ControlId = @ControlId";
                        dynamicParameters.Add("ControlId", ControlId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("客戶設計圖版本資料錯誤!");
                        #endregion

                        #region //調整為相反狀態
                        string ReleasedStatus = "";
                        foreach (var item in result)
                        {
                            ReleasedStatus = item.ReleasedStatus;
                        }

                        switch (ReleasedStatus)
                        {
                            case "Y":
                                ReleasedStatus = "N";
                                break;
                            case "N":
                                ReleasedStatus = "Y";
                                break;
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.CustomerCadControl SET
                                ReleasedStatus = @ReleasedStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ControlId = @ControlId";
                        var parametersObject = new
                        {
                            ReleasedStatus,
                            LastModifiedDate,
                            LastModifiedBy,
                            ControlId
                        };
                        dynamicParameters.AddDynamicParams(parametersObject);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
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

        #region //UpdateRdDesignControlReleasedStatus -- 更新研發設計圖發行狀態 -- Ann 2022.06.28
        public string UpdateRdDesignControlReleasedStatus(int ControlId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷研發設計圖版本資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ReleasedStatus
                                FROM PDM.RdDesignControl
                                WHERE ControlId = @ControlId";
                        dynamicParameters.Add("ControlId", ControlId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("研發設計圖版本資料錯誤!");
                        #endregion

                        #region //調整為相反狀態
                        string ReleasedStatus = "";
                        foreach (var item in result)
                        {
                            ReleasedStatus = item.ReleasedStatus;
                        }

                        switch (ReleasedStatus)
                        {
                            case "Y":
                                ReleasedStatus = "N";
                                break;
                            case "N":
                                ReleasedStatus = "Y";
                                break;
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.RdDesignControl SET
                                ReleasedStatus = @ReleasedStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ControlId = @ControlId";
                        var parametersObject = new
                        {
                            ReleasedStatus,
                            LastModifiedDate,
                            LastModifiedBy,
                            ControlId
                        };
                        dynamicParameters.AddDynamicParams(parametersObject);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
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
        #endregion

        #region //Delete
        #region //DeleteCustomerCad -- 刪除客戶設計圖資料 -- Ann 2022.06.10
        public string DeleteCustomerCad(int CadId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷客戶設計圖資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.CustomerCad
                                WHERE CadId = @CadId";
                        dynamicParameters.Add("CadId", CadId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("客戶設計圖資料錯誤!");
                        #endregion

                        #region //判斷客戶設計圖是否有版本紀錄
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.CustomerCadControl
                                WHERE CadId = @CadId";
                        dynamicParameters.Add("CadId", CadId);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() > 0) throw new SystemException("客戶設計圖已有版本紀錄，無法刪除!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PDM.CustomerCad
                                WHERE CadId = @CadId";
                        dynamicParameters.Add("CadId", CadId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
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

        #region //DeleteRdDesign -- 刪除研發設計圖資料 -- Ann 2022.06.28
        public string DeleteRdDesign(int DesignId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷研發設計圖資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.RdDesign
                                WHERE DesignId = @DesignId";
                        dynamicParameters.Add("DesignId", DesignId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("研發設計圖資料錯誤!");
                        #endregion

                        #region //判斷研發設計圖是否有版本紀錄
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.RdDesignControl
                                WHERE DesignId = @DesignId";
                        dynamicParameters.Add("DesignId", DesignId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("研發設計圖已有版本紀錄，無法刪除!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PDM.RdDesign
                                WHERE DesignId = @DesignId";
                        dynamicParameters.Add("DesignId", DesignId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
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
        #endregion

        #region //Export Excel
        #region //GetDrawingInfo -- 取得出圖資料 -- Ann 2024-11-04
        public string GetDrawingInfo(string Year, string Month)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //客戶設計圖
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT COUNT(1) Count
                            FROM PDM.CustomerCadControl a 
                            INNER JOIN PDM.CustomerCad b ON a.CadId = b.CadId
                            WHERE Month(a.CreateDate) = @Month
                            AND YEAR(a.CreateDate) = @Year
                            AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("Month", Month);
                    dynamicParameters.Add("Year", Year);
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    var CustomerCadControlResult = sqlConnection.Query(sql, dynamicParameters);

                    int customerCadCount = 0;
                    foreach (var item in CustomerCadControlResult)
                    {
                        customerCadCount = item.Count;
                    }
                    #endregion

                    #region //研發3D、2D、JMO
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT COUNT(Cad3DFile) + COUNT(Cad3DFileAbsolutePath) Cad3DFileCount
                            , COUNT(Cad2DFile) + COUNT(Cad2DFileAbsolutePath) Cad2DFileCount
                            , COUNT(JmoFile) + COUNT(JmoFileAbsolutePath) JmoFileCount
                            FROM PDM.RdDesignControl a 
                            INNER JOIN PDM.RdDesign b ON a.DesignId = b.DesignId
                            WHERE Month(a.CreateDate) = @Month
                            AND YEAR(a.CreateDate) = @Year
                            AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("Month", Month);
                    dynamicParameters.Add("Year", Year);
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    var RdDesignControlResult = sqlConnection.Query(sql, dynamicParameters);

                    int cad3DFileCount = 0;
                    int cad2DFileCount = 0;
                    int jmoFileCount = 0;
                    foreach (var item in RdDesignControlResult)
                    {
                        cad3DFileCount = item.Cad3DFileCount;
                        cad2DFileCount = item.Cad2DFileCount;
                        jmoFileCount = item.Cad3DFileCount;
                    }
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        customerCadCount,
                        cad3DFileCount,
                        cad2DFileCount,
                        jmoFileCount,
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
        #endregion

        #region //API
        #region //AutoDownloadCustomerCad -- 自動下載客戶設計圖二進制檔案存為實體檔案 -- Ann 2023-06-12
        public string AutoDownloadCustomerCad(string ServerPath, int CompanyId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope(TransactionScopeOption.Required, new TimeSpan(0,2,0)))
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CompanyNo FROM BAS.Company
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        var CompanyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (CompanyResult.Count() <= 0) throw new SystemException("公司別資料錯誤!!");

                        string CompanyNo = "";
                        foreach (var item in CompanyResult)
                        {
                            CompanyNo = item.CompanyNo;
                        }
                        #endregion

                        #region //PDM.CustomerCad CadFile
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ControlId
                                , b.FileId, b.FileContent, b.[FileName], b.FileExtension
                                , c.CustomerMtlItemNo, c.CustomerDwgNo
                                FROM PDM.CustomerCadControl a 
                                INNER JOIN BAS.[File] b ON a.CadFile = b.FileId
                                INNER JOIN PDM.CustomerCad c ON a.CadId = c.CadId
                                WHERE c.CompanyId = @CompanyId
                                AND a.CadFileAbsolutePath IS NULL";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        var CadResult = sqlConnection.Query(sql, dynamicParameters);

                        var invalidFileNameChars = Path.GetInvalidFileNameChars();
                        List<int> ErrorControlIdList = new List<int>();
                        foreach (var item in CadResult)
                        {
                            if (item.FileName.IndexOfAny(invalidFileNameChars) != -1
                                || item.CustomerMtlItemNo.IndexOfAny(invalidFileNameChars) != -1
                                || item.CustomerDwgNo.IndexOfAny(invalidFileNameChars) != -1)
                            {
                                ErrorControlIdList.Add(item.ControlId);
                                continue;
                            }

                            #region //下載至實體路徑
                            string CadPath = @"\\" + CompanyNo + "\\Customer\\" + item.CustomerMtlItemNo + "(" + item.CustomerDwgNo + ")\\";
                            CadPath = Path.Combine(ServerPath + CadPath);
                            if (!Directory.Exists(CadPath)) { Directory.CreateDirectory(CadPath); }
                            string ServerPath2 = Path.Combine(CadPath, item.FileName + "-" + item.ControlId + item.FileExtension);
                            byte[] fileContent = (byte[])item.FileContent;
                            File.WriteAllBytes(ServerPath2, fileContent); // Requires System.IO
                            #endregion

                            #region //UPDATE PDM.CustomerCadControl CadFileAbsolutePath
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PDM.CustomerCadControl SET
                                    CadFileAbsolutePath = @CadFileAbsolutePath,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE ControlId = @ControlId";
                            var parametersObject = new
                            {
                                CadFileAbsolutePath = ServerPath2,
                                LastModifiedDate,
                                LastModifiedBy,
                                item.ControlId
                            };
                            dynamicParameters.AddDynamicParams(parametersObject);

                            int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //其他檔案
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT d.FileId, d.FileContent, d.[FileName], d.FileExtension
                                    FROM PDM.CustomerCadControlFile a 
                                    INNER JOIN PDM.CustomerCadControl b ON a.ControlId = b.ControlId
                                    INNER JOIN PDM.CustomerCad c ON b.CadId = c.CadId
                                    INNER JOIN BAS.[File] d ON a.FileId = d.FileId
                                    WHERE a.ControlId = @ControlId";
                            dynamicParameters.Add("ControlId", item.ControlId);

                            var CustomerCadControlFileResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item2 in CustomerCadControlFileResult)
                            {
                                if (item2.FileName.IndexOfAny(invalidFileNameChars) != -1)
                                {
                                    ErrorControlIdList.Add(item.ControlId);
                                    continue;
                                }

                                #region //下載至實體路徑
                                string CadPath2 = @"\\" + CompanyNo + "\\Customer\\" + item.CustomerMtlItemNo + "(" + item.CustomerDwgNo + ")\\";
                                CadPath2 = Path.Combine(ServerPath + CadPath2);
                                if (!Directory.Exists(CadPath)) { Directory.CreateDirectory(CadPath); }
                                string ServerPath3 = Path.Combine(CadPath2, item2.FileName + "-" + item2.ControlId + item2.FileExtension);
                                byte[] fileContent2 = (byte[])item2.FileContent;
                                File.WriteAllBytes(ServerPath3, fileContent2); // Requires System.IO
                                #endregion

                                #region //UPDATE PDM.CustomerCadControlFile CadOtherFileAbsolutePath
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE PDM.CustomerCadControlFile SET
                                        CadOtherFileAbsolutePath = @CadOtherFileAbsolutePath,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE ControlId = @ControlId
                                        AND FileId = @FileId";
                                var parametersObject2 = new
                                {
                                    CadOtherFileAbsolutePath = ServerPath3,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    item.ControlId,
                                    item.FileId
                                };
                                dynamicParameters.AddDynamicParams(parametersObject2);

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                            #endregion
                        }
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = ErrorControlIdList.ToString()
                        });
                        #endregion
                    }

                    transactionScope.Complete();
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

        #region //AutoDownloadRdDesignCad -- 自動下載RD研發設計圖二進制檔案存為實體檔案 -- Ann 2023-06-13
        public string AutoDownloadRdDesignCad(string ServerPath, int CompanyId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CompanyNo FROM BAS.Company
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        var CompanyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (CompanyResult.Count() <= 0) throw new SystemException("公司別資料錯誤!!");

                        string CompanyNo = "";
                        foreach (var item in CompanyResult)
                        {
                            CompanyNo = item.CompanyNo;
                        }
                        #endregion

                        #region //PDM.CustomerCad CadFile
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.ControlId
                                , ISNULL(b.FileId, -1) Cad3DFileId, b.[FileName] Cad3DFileName, b.FileExtension Cad3DFileExtension, b.FileContent Cad3DFileContent
                                , ISNULL(c.FileId, -1) Cad2DFileId, c.[FileName] Cad2DFileName, c.FileExtension Cad2DFileExtension, c.FileContent Cad2DFileContent
                                , ISNULL(d.FileId, -1) Pdf2DFileId, d.[FileName] Pdf2DFileName, d.FileExtension Pdf2DFileExtension, d.FileContent Pdf2DFileContent
                                , ISNULL(e.FileId, -1) JmoFileId, e.[FileName] JmoFileName, e.FileExtension JmoFileExtension, e.FileContent JmoFileContent
                                , ISNULL(g.MtlItemNo, '無綁定品號:' + CONVERT(VARCHAR(100), a.ControlId)) MtlItemNo
                                FROM PDM.RdDesignControl a 
                                LEFT JOIN BAS.[File] b ON a.Cad3DFile = b.FileId
                                LEFT JOIN BAS.[File] c ON a.Cad2DFile = c.FileId
                                LEFT JOIN BAS.[File] d ON a.Pdf2DFile = d.FileId
                                LEFT JOIN BAS.[File] e ON a.JmoFile = e.FileId
                                INNER JOIN PDM.RdDesign f ON a.DesignId = f.DesignId
                                LEFT JOIN PDM.MtlItem g ON f.MtlItemId = g.MtlItemId
                                WHERE a.Cad3DFileAbsolutePath IS NULL
                                AND f.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        var CadResult = sqlConnection.Query(sql, dynamicParameters);

                        var invalidFileNameChars = Path.GetInvalidFileNameChars();
                        List<int> ErrorControlIdList = new List<int>();
                        foreach (var item in CadResult)
                        {
                            #region //建立此品號資料夾
                            string CadPath = @"\\" + CompanyNo + "\\RdDesign\\" + item.MtlItemNo + "\\";
                            string CadPath2 = @"\\" + CompanyNo + "\\RdDesign\\" + item.MtlItemNo + "\\";
                            CadPath = Path.Combine(ServerPath + CadPath);
                            if (!Directory.Exists(CadPath)) { Directory.CreateDirectory(CadPath); }
                            #endregion

                            #region //下載3D研發設計圖
                            string Cad3DFilePath = "";
                            string Cad3DFileAbsolutePath = "";
                            if (item.Cad3DFileId > 0)
                            {
                                if (item.Cad3DFileName.IndexOfAny(invalidFileNameChars) != -1)
                                {
                                    ErrorControlIdList.Add(item.ControlId);
                                    continue;
                                }
                                Cad3DFileAbsolutePath = CadPath2 + item.Cad3DFileName + "-" + item.Cad3DFileId + item.Cad3DFileExtension;
                                Cad3DFilePath = Path.Combine(CadPath, item.Cad3DFileName + "-" + item.Cad3DFileId + item.Cad3DFileExtension);
                                byte[] fileContent = (byte[])item.Cad3DFileContent;
                                File.WriteAllBytes(Cad3DFilePath, fileContent); // Requires System.IO
                            }
                            #endregion

                            #region //下載2D研發設計圖
                            string Cad2DFilePath = "";
                            string Cad2DFileAbsolutePath = "";
                            if (item.Cad2DFileId > 0)
                            {
                                if (item.Cad2DFileName.IndexOfAny(invalidFileNameChars) != -1)
                                {
                                    ErrorControlIdList.Add(item.ControlId);
                                    continue;
                                }
                                Cad2DFileAbsolutePath = CadPath2 + item.Cad2DFileName + "-" + item.Cad2DFileId + item.Cad2DFileExtension;
                                Cad2DFilePath = Path.Combine(CadPath, item.Cad2DFileName + "-" + item.Cad2DFileId + item.Cad2DFileExtension);
                                byte[] fileContent = (byte[])item.Cad2DFileContent;
                                File.WriteAllBytes(Cad2DFilePath, fileContent); // Requires System.IO
                            }
                            #endregion

                            #region //下載Pdf2D研發設計圖
                            string Pdf2DFilePath = "";
                            string Pdf2DFileAbsolutePath = "";
                            if (item.Pdf2DFileId > 0)
                            {
                                if (item.Pdf2DFileName.IndexOfAny(invalidFileNameChars) != -1)
                                {
                                    ErrorControlIdList.Add(item.ControlId);
                                    continue;
                                }
                                Pdf2DFileAbsolutePath = CadPath2 + item.Pdf2DFileName + "-" + item.Pdf2DFileId + item.Pdf2DFileExtension;
                                Pdf2DFilePath = Path.Combine(CadPath, item.Pdf2DFileName + "-" + item.Pdf2DFileId + item.Pdf2DFileExtension);
                                byte[] fileContent = (byte[])item.Pdf2DFileContent;
                                File.WriteAllBytes(Pdf2DFilePath, fileContent); // Requires System.IO
                            }
                            #endregion

                            #region //下載JMO研發設計圖
                            string JmoFilePath = "";
                            string JmoFileAbsolutePath = "";
                            if (item.JmoFileId > 0)
                            {
                                if (item.JmoFileName.IndexOfAny(invalidFileNameChars) != -1)
                                {
                                    ErrorControlIdList.Add(item.ControlId);
                                    continue;
                                }
                                JmoFileAbsolutePath = CadPath2 + item.JmoFileName + "-" + item.JmoFileId + item.JmoFileExtension;
                                JmoFilePath = Path.Combine(CadPath, item.JmoFileName + "-" + item.JmoFileId + item.JmoFileExtension);
                                byte[] fileContent = (byte[])item.JmoFileContent;
                                File.WriteAllBytes(JmoFilePath, fileContent); // Requires System.IO
                            }
                            #endregion

                            #region //UPDATE PDM.CustomerCadControl CadFileAbsolutePath
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PDM.RdDesignControl SET
                                    Cad3DFileAbsolutePath = @Cad3DFileAbsolutePath,
                                    Cad2DFileAbsolutePath = @Cad2DFileAbsolutePath,
                                    Pdf2DFileAbsolutePath = @Pdf2DFileAbsolutePath,
                                    JmoFileAbsolutePath = @JmoFileAbsolutePath,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE ControlId = @ControlId";
                            var parametersObject = new
                            {
                                Cad3DFileAbsolutePath = Cad3DFileAbsolutePath != "" ? Cad3DFileAbsolutePath : null,
                                Cad2DFileAbsolutePath = Cad2DFileAbsolutePath != "" ? Cad2DFileAbsolutePath : null,
                                Pdf2DFileAbsolutePath = Pdf2DFileAbsolutePath != "" ? Pdf2DFileAbsolutePath : null,
                                JmoFileAbsolutePath = JmoFileAbsolutePath != "" ? JmoFileAbsolutePath : null,
                                LastModifiedDate,
                                LastModifiedBy,
                                item.ControlId
                            };
                            dynamicParameters.AddDynamicParams(parametersObject);

                            int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = ErrorControlIdList.ToString()
                        });
                        #endregion
                    }

                    transactionScope.Complete();
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

        #region//GetDrawingData 製令 設計圖路徑
        public string GetDrawingData(int MoId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT f.Cad2DFileAbsolutePath,f.Pdf2DFileAbsolutePath
                                   ,f.Cad3DFileAbsolutePath,JmoFileAbsolutePath
                            FROM MES.ManufactureOrder a
                                INNER JOIN MES.WipOrder b ON a.WoId=b.WoId
                                INNER JOIN PDM.MtlItem c ON b. MtlItemId=c.MtlItemId
                                INNER JOIN MES.MoRouting d ON a.MoId=d.MoId
                                INNER JOIN MES.RoutingItem e on d.RoutingItemId=e.RoutingItemId
                                INNER JOIN PDM.RdDesignControl f ON e.ControlId=f.ControlId
                            WHERE a.MoId=@MoId ";
                    dynamicParameters.Add("MoId", MoId);
                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("此製令查無圖面資訊");
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

        #endregion

        #region Model
        #region //CadFileInfo
        public class CadFileInfo
        {
            public string FileName { get; set; }
            public string FileExtension { get; set; }
            public byte[] FileByte { get; set; }
        }
        #endregion
        #endregion
    }
}
