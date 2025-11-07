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
using System.Transactions;
using System.Web;
using System.Threading;
using Newtonsoft.Json;

namespace PDMDA
{
    public class ResearchAndDevelopmentDA
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

        public ResearchAndDevelopmentDA()
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

        #region//Get
        #region //GetUserAuthority -- 取得專利使用者權限-- Ted 2024-01-10
        public string GetUserAuthority(string UserNo)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {

                    DynamicParameters dynamicParameters = new DynamicParameters();
                    int Num = -1;
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.DetailCode,x.AuthorityBase FROM BAS.FunctionDetail a
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
                                        INNER JOIN BAS.[User] cac ON caa.UserId = cac.UserId
                                        WHERE cac.UserNo = @UserNo
                                        AND cab.CompanyId = @CompanyId
                                    )
                                ), 0) Authority
                            ) c
                            OUTER APPLY (
                                SELECT da.UserNo, da.UserName, db.DepartmentNo
                                FROM BAS.[User] da
                                INNER JOIN BAS.Department db ON da.DepartmentId = db.DepartmentId
                                WHERE da.UserNo = @UserNo
                            ) d
                            OUTER APPLY (
                                SELECT STUFF((SELECT ','+ Convert(nvarchar(MAX),(x.DetailCode)) 
                                FROM BAS.FunctionDetail x
                                INNER JOIN BAS.[Function] x1 ON x.FunctionId = x1.FunctionId
                                WHERE x1.FunctionCode = 'DesignPatentManagement'
                                FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 1, '') AS AuthorityBase
                            ) x
                            WHERE a.[Status] = 'A'
                            AND b.[Status] = 'A'
                            AND b.FunctionCode = 'DesignPatentManagement'
                            AND c.Authority > 0";
                    dynamicParameters.Add("UserNo", UserNo);
                    dynamicParameters.Add("CompanyId", CurrentCompany);
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

        #region //GetDesignPatentExcel -- 取得專利Exce資料 -- Shintokuro 2023.12.27
        public string GetDesignPatentExcel(
             int DpId, string ApplicationNo, string ApplicationDate, string Applicant,
             string PublicationNo, string PublicationDate, string AnnouncementNo, string AnnouncementDate,
             string ImplementationExample, string LensQuantity, string AperturePosition, string HalfFieldofView, string ImageHeight,
             string L1RefractivePower, string L1aSurfaceType, string L1bSurfaceType, string L2RefractivePower, string L2aSurfaceType, string L2bSurfaceType,
             string L3RefractivePower, string L3aSurfaceType, string L3bSurfaceType, string L4RefractivePower, string L4aSurfaceType, string L4bSurfaceType,
             string L5RefractivePower, string L5aSurfaceType, string L5bSurfaceType, string L6RefractivePower, string L6aSurfaceType, string L6bSurfaceType,
             string L7RefractivePower, string L7aSurfaceType, string L7bSurfaceType, string L8RefractivePower, string L8aSurfaceType, string L8bSurfaceType,
             string L9RefractivePower, string L9aSurfaceType, string L9bSurfaceType, string LXRefractivePower, string LXaSurfaceType, string LXbSurfaceType
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.DpId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @"  ,a.ApplicationNo, a.Applicant, a.AnnouncementNo, a.PublicationNo, 
                            a.ImplementationExample, a.LensQuantity, a.AperturePosition, a.HalfFieldofView, a.ImageHeight, 
                            FORMAT(a.ApplicationDate, 'yyyy/MM/dd') ApplicationDate,
                            FORMAT(a.PublicationDate, 'yyyy/MM/dd') PublicationDate,
                            FORMAT(a.AnnouncementDate, 'yyyy/MM/dd') AnnouncementDate,
                            a.L1bSurfaceType, a.L1RefractivePower, a.L1aSurfaceType, a.L2RefractivePower, a.L2aSurfaceType, a.L2bSurfaceType, 
                            a.L3RefractivePower, a.L3aSurfaceType, a.L3bSurfaceType, a.L4RefractivePower, a.L4aSurfaceType, a.L4bSurfaceType, 
                            a.L5RefractivePower, a.L5aSurfaceType, a.L5bSurfaceType, a.L6RefractivePower, a.L6aSurfaceType, a.L6bSurfaceType, 
                            a.L7RefractivePower, a.L7aSurfaceType, a.L7bSurfaceType, a.L8RefractivePower, a.L8aSurfaceType, a.L8bSurfaceType, 
                            a.L9RefractivePower, a.L9aSurfaceType, a.L9bSurfaceType, a.LXRefractivePower, a.LXaSurfaceType, a.LXbSurfaceType, 
                            a.CreateDate, a.LastModifiedDate, a.CreateBy, a.LastModifiedBy
                          ";
                    sqlQuery.mainTables =
                        @"FROM RFI.DesignPatent a
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DpId", @" AND a.DpId = @DpId", DpId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ApplicationNo", @" AND a.ApplicationNo LIKE '%' + @ApplicationNo + '%'", ApplicationNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ApplicationDate", @" AND a.ApplicationDate = @ApplicationDate", ApplicationDate);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Applicant", @" AND a.Applicant LIKE '%' + @Applicant + '%'", Applicant);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PublicationNo", @" AND a.PublicationNo LIKE '%' + @PublicationNo + '%'", PublicationNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PublicationDate", @" AND a.PublicationDate = @PublicationDate", PublicationDate);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AnnouncementNo", @" AND a.AnnouncementNo LIKE '%' + @AnnouncementNo + '%'", AnnouncementNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AnnouncementDate", @" AND a.AnnouncementDate = @AnnouncementDate", AnnouncementDate);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ImplementationExample", @" AND a.ImplementationExample = @ImplementationExample", ImplementationExample);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "LensQuantity", @" AND a.LensQuantity = @LensQuantity", LensQuantity);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AperturePosition", @" AND a.AperturePosition = @AperturePosition", AperturePosition);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "HalfFieldofView(HFOV)", @" AND a.HalfFieldofView = @HalfFieldofView", HalfFieldofView);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ImageHeight(ImgH)", @" AND a.ImageHeight = @ImageHeight", ImageHeight);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "L1RefractivePower", @" AND a.L1RefractivePower = @L1RefractivePower", L1RefractivePower);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "L1aSurfaceType", @" AND a.L1aSurfaceType = @L1aSurfaceType", L1aSurfaceType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "L1bSurfaceType", @" AND a.L1bSurfaceType = @L1bSurfaceType", L1bSurfaceType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "L2RefractivePower", @" AND a.L2RefractivePower = @L2RefractivePower", L2RefractivePower);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "L2aSurfaceType", @" AND a.L2aSurfaceType = @L2aSurfaceType", L2aSurfaceType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "L2bSurfaceType", @" AND a.L2bSurfaceType = @L2bSurfaceType", L2bSurfaceType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "L3RefractivePower", @" AND a.L3RefractivePower = @L3RefractivePower", L3RefractivePower);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "L3aSurfaceType", @" AND a.L3aSurfaceType = @L3aSurfaceType", L3aSurfaceType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "L3bSurfaceType", @" AND a.L3bSurfaceType = @L3bSurfaceType", L3bSurfaceType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "L4RefractivePower", @" AND a.L4RefractivePower = @L4RefractivePower", L4RefractivePower);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "L4aSurfaceType", @" AND a.L4aSurfaceType = @L4aSurfaceType", L4aSurfaceType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "L4bSurfaceType", @" AND a.L4bSurfaceType = @L4bSurfaceType", L4bSurfaceType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "L5RefractivePower", @" AND a.L5RefractivePower = @L5RefractivePower", L5RefractivePower);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "L5aSurfaceType", @" AND a.L5aSurfaceType = @L5aSurfaceType", L5aSurfaceType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "L5bSurfaceType", @" AND a.L5bSurfaceType = @L5bSurfaceType", L5bSurfaceType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "L6RefractivePower", @" AND a.L6RefractivePower = @L6RefractivePower", L6RefractivePower);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "L6aSurfaceType", @" AND a.L6aSurfaceType = @L6aSurfaceType", L6aSurfaceType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "L6bSurfaceType", @" AND a.L6bSurfaceType = @L6bSurfaceType", L6bSurfaceType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "L7RefractivePower", @" AND a.L7RefractivePower = @L7RefractivePower", L7RefractivePower);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "L7aSurfaceType", @" AND a.L7aSurfaceType = @L7aSurfaceType", L7aSurfaceType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "L7bSurfaceType", @" AND a.L7bSurfaceType = @L7bSurfaceType", L7bSurfaceType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "L8RefractivePower", @" AND a.L8RefractivePower = @L8RefractivePower", L8RefractivePower);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "L8aSurfaceType", @" AND a.L8aSurfaceType = @L8aSurfaceType", L8aSurfaceType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "L8bSurfaceType", @" AND a.L8bSurfaceType = @L8bSurfaceType", L8bSurfaceType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "L9RefractivePower", @" AND a.L9RefractivePower = @L9RefractivePower", L9RefractivePower);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "L9aSurfaceType", @" AND a.L9aSurfaceType = @L9aSurfaceType", L9aSurfaceType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "L9bSurfaceType", @" AND a.L9bSurfaceType = @L9bSurfaceType", L9bSurfaceType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "LXRefractivePower", @" AND a.LXRefractivePower = @LXRefractivePower", LXRefractivePower);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "LXaSurfaceType", @" AND a.LXaSurfaceType = @LXaSurfaceType", LXaSurfaceType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "LXbSurfaceType", @" AND a.LXbSurfaceType = @LXbSurfaceType", LXbSurfaceType);

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.DpId";
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

        #region //GetDesignPatentCount -- 取得專利Exce資料總比數 -- Ted 2024-01-10
        public string GetDesignPatentCount()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {

                    DynamicParameters dynamicParameters = new DynamicParameters();
                    int Num = -1;
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT Count(a.DpId) Num
                            FROM RFI.DesignPatent a
                            WHERE a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    var result = sqlConnection.Query(sql, dynamicParameters);
                    foreach (var item in result)
                    {
                        Num = item.Num;
                    }
                    

                    string[] setting = { Num.ToString() };

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = setting
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

        #region //GetModelDesign -- 取得鏡頭設計資料 -- Shintokuro 2024.01.02
        public string GetModelDesign(int MdId, string Model, string MdIdList,
            string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.MdId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @"  ,CompanyId, Model, Construction, EFL, FNo, TTL, CRA, OpticalDistortionFTan, 
                            OpticalDistortionF, MaxImageCircle, Sensor, VFOV, HFOV, DFOV, PartName, IR, MechanicalRetainer, MechanicalFBL, 
                            MechanicalThread, IPRating, Stage, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy
                                ";
                    sqlQuery.mainTables =
                        @"FROM RFI.ModelDesign a
                                 ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MdId", @" AND a.MdId = @MdId", MdId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Model", @" AND a.Model = @Model", Model);
                    if(MdIdList.Length> 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MdIdList", @" AND a.Model IN @MdIdList", MdIdList.Split(','));

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MdId";
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

        #region //GetLensFOV -- 取得鏡頭FOV資料 -- Shintokuro 2024.01.02
        public string GetLensFOV(string Model, int LfId, string ModelName, double FOV, double RealHeight
            , double GreaterMeets, double HFOV, double VFOV, double DFOV
            , string ButtonType, double FOVRangeD, double FOVRangeT
            , string ModelNoList
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    switch (Model)
                    {
                        case "General":
                            dynamicParameters = new DynamicParameters();
                            sqlQuery.mainKey = "a.LfId";
                            sqlQuery.auxKey = "";
                            sqlQuery.columns =
                                @"  ,a.Model, a.FOV, a.Realheight, FORMAT(a.CreateDate, 'yyyy-MM-dd') CreateDate
                                ";
                            sqlQuery.mainTables =
                                @"FROM RFI.LensFOV a
                                 ";
                            string queryTable = "";
                            sqlQuery.auxTables = queryTable;
                            string queryCondition = @"AND a.CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "LfId", @" AND a.LfId = @LfId", LfId);
                            BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModelName", @" AND a.Model LIKE '%' + @ModelName + '%'", ModelName);
                            BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "FOV", @" AND a.FOV = @FOV", FOV);
                            BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Realheight", @" AND a.Realheight = @RealHeight", RealHeight);
                            
                            sqlQuery.conditions = queryCondition;
                            sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.LfId";
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
                            break;
                        case "Mate":
                            if (HFOV < 0) throw new SystemException("【鏡頭FOV管理】 - 配對模式下條件HFOV不能為空!!");
                            if (VFOV < 0) throw new SystemException("【鏡頭FOV管理】 - 配對模式下條件VFOV不能為空!!");
                            if (DFOV < 0) throw new SystemException("【鏡頭FOV管理】 - 配對模式下條件DFOV不能為空!!");

                            var MdIdListArr = "";
                            if (ModelNoList.Length > 0)
                            {
                                foreach (var item in ModelNoList.Split(','))
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.MdId
                                            FROM RFI.ModelDesign a
                                            WHERE a.CompanyId = @CompanyId
                                            AND a.Model = @Model
                                    ";
                                    dynamicParameters.Add("CompanyId", CurrentCompany);
                                    dynamicParameters.Add("Model", item);
                                    var result3 = sqlConnection.Query(sql, dynamicParameters);
                                    if (result3.Count() > 0)
                                    {
                                        foreach (var item3 in result3)
                                        {
                                            MdIdListArr += item3.MdId + ",";
                                        }
                                    }
                                }
                                if (MdIdListArr.Length > 0)
                                {
                                    MdIdListArr = MdIdListArr.Substring(0, MdIdListArr.Length - 1);
                                }
                            }

                            dynamicParameters = new DynamicParameters();
                            sql = @"  SELECT 
                                        x5.Sensor, x5.Model, x5.Construction, x5.EFL, x5.FNo, x5.TTL, x5.CRA, x5.OpticalDistortionFTan
                                        ,x5.OpticalDistortionF, x5.MaxImageCircle, x5.Sensor, x5.VFOV, x5.HFOV, x5.DFOV, x5.PartName, x5.IR, x5.MechanicalRetainer, x5.MechanicalFBL
                                        ,x5.MechanicalThread, x5.IPRating, x5.Stage
                                        ,x1.Model
                                        ,x2.HFOVIH 
                                        ,x3.VFOVIH
                                        ,x4.DFOVIH
                                        FROM (
                                            SELECT DISTINCT  a.Model
                                            FROM RFI.LensFOV a 
                                            WHERE a.Realheight > @GreaterMeets
                                        ) x1
                                        OUTER APPLY(
                                            SELECT ISNULL(MAX(a.FOV), 0) * 2 HFOVIH, Model
                                            FROM RFI.LensFOV a
                                            WHERE x1.Model = a.Model
                                            AND a.Realheight < @HFOVIH
                                            GROUP BY Model
                                        ) x2
                                        OUTER APPLY(
                                            SELECT ISNULL(MAX(a.FOV), 0) * 2 VFOVIH, Model
                                            FROM RFI.LensFOV a
                                            WHERE x1.Model = a.Model
                                            AND a.Realheight < @VFOVIH
                                            GROUP BY Model
                                        ) x3
                                        OUTER APPLY(
                                            SELECT ISNULL(MAX(a.FOV) , 0) * 2 DFOVIH, Model
                                            FROM RFI.LensFOV a
                                            WHERE x1.Model = a.Model
                                            AND a.Realheight < @DFOVIH
                                            GROUP BY Model
                                        ) x4
                                        INNER JOIN RFI.ModelDesign x5 on x1.Model = x5.Model
                                        WHERE 1=1
                                        AND x5.CompanyId = @CompanyId
                                        AND( x2.HFOVIH is not null 
                                            OR x3.VFOVIH is not null
                                            OR x4.DFOVIH is not null)
                                    ";
                            switch (ButtonType) {
                                case "HFOV":
                                    sql += @"AND x2.HFOVIH >= @FOVRangeD
                                             AND x2.HFOVIH <= @FOVRangeT";
                                    break;
                                case "DFOV":
                                    sql += @"AND x4.DFOVIH >= @FOVRangeD
                                            AND x4.DFOVIH <= @FOVRangeT";
                                    break;
                                default:
                                    break;
                            }
                            if (MdIdListArr.Length > 0)
                            {
                                sql += " AND x5.MdId in ( " + MdIdListArr + @")";
                            }
                            dynamicParameters.Add("GreaterMeets", GreaterMeets);
                            dynamicParameters.Add("HFOVIH", HFOV);
                            dynamicParameters.Add("VFOVIH", VFOV);
                            dynamicParameters.Add("DFOVIH", DFOV);
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("FOVRangeD", FOVRangeD);
                            dynamicParameters.Add("FOVRangeT", FOVRangeT);
                            sql += " ORDER BY x5.MdId ";

                            var result1 = sqlConnection.Query(sql, dynamicParameters);
                            if (result1.Count() <= 0) throw new SystemException("【鏡頭FOV管理】 - 機種資料找不到請重新確認!!");

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                data = result1
                            });
                            #endregion
                            break;
                        case "MateDetail":
                            dynamicParameters = new DynamicParameters();
                            var MdIdListArr1 = "";
                            foreach (var item in ModelNoList.Split(','))
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MdId
                                    FROM RFI.ModelDesign a
                                    WHERE a.CompanyId = @CompanyId
                                    AND a.Model = @Model
                                    ";
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                dynamicParameters.Add("Model", item);
                                var result3 = sqlConnection.Query(sql, dynamicParameters);
                                if (result3.Count() > 0)
                                {
                                    foreach (var item3 in result3)
                                    {
                                        MdIdListArr1 += item3.MdId+",";
                                    }
                                }
                            }
                            if(MdIdListArr1.Length> 0)
                            {
                                MdIdListArr1 = MdIdListArr1.Substring(0, MdIdListArr1.Length - 1);
                            }

                            sql = @"SELECT a.Model, a.Construction, a.EFL, a.FNo, a.TTL, a.CRA, a.OpticalDistortionFTan, 
                                    a.OpticalDistortionF, a.MaxImageCircle, a.Sensor, a.VFOV, a.HFOV, a.DFOV, a.PartName, a.IR, a.MechanicalRetainer, a.MechanicalFBL, 
                                    a.MechanicalThread, a.IPRating, a.Stage,
                                     b.Realheight,b.FOV, FORMAT(b.CreateDate, 'yyyy-MM-dd') CreateDate
                                    FROM RFI.ModelDesign a  
                                    INNER JOIN RFI.LensFOV b on a.Model = b.Model
                                    WHERE a.MdId in ( " + MdIdListArr1 + @")
                                    AND b.LfId in ( " + ModelNoList + @")
                                    ORDER BY b.LfId
                                    ";
                            //dynamicParameters.Add("MdIdList", MdIdListArr);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() <= 0) throw new SystemException("【鏡頭FOV管理】 - 機種資料找不到請重新確認!!");

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                data = result2
                            });
                            #endregion
                            break;
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
        #endregion

        #region//Add
        #region //AddDesignPatentExcel -- 上傳專利Exce資料 -- Shintokuro 2023-12-27
        public string AddDesignPatentExcel(string DpIdList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        int rowsAffected = 0;
                        int Serial = 4;
                        int repeatNum = 0;

                        List<Dictionary<string, string>> DpJsonList = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(DpIdList);
                        string repeatList = "";

                        #region //解析Spreadsheet Data
                        foreach (var item in DpJsonList)
                        {
                            DateTime date;
                            DateTime? applicationDate = null;
                            DateTime? publicationDate = null;
                            DateTime? announcementDate = null;

                            string ApplicationNo = item["ApplicationNo"] != null ? item["ApplicationNo"].ToString() : throw new SystemException("【資料維護不完整】申請號欄位資料不可以為空,請重新確認~~");
                            string ApplicationDate = item["ApplicationDate"] != null ? item["ApplicationDate"].ToString() : null;
                            bool isValidApplicationDate = DateTime.TryParseExact(ApplicationDate, "yyyy/MM/dd", null, System.Globalization.DateTimeStyles.None, out date);
                            if (isValidApplicationDate) applicationDate = DateTime.ParseExact(ApplicationDate, "yyyy/MM/dd", null);
                            string Applicant = item["Applicant"] != null ? item["Applicant"].ToString() : null;

                            string PublicationNo = item["PublicationNo"] != null  ? item["PublicationNo"].ToString() : null;
                            if (PublicationNo == "-") PublicationNo = null;
                            string PublicationDate = item["PublicationDate"] != null ? item["PublicationDate"].ToString() : null;
                            bool isValidPublicationDate = DateTime.TryParseExact(PublicationDate, "yyyy/MM/dd", null, System.Globalization.DateTimeStyles.None, out date);
                            if (isValidPublicationDate) publicationDate = DateTime.ParseExact(PublicationDate, "yyyy/MM/dd", null);

                            string AnnouncementNo = item["AnnouncementNo"] != null ? item["AnnouncementNo"].ToString() : null;
                            if (AnnouncementNo == "-") AnnouncementNo = null;
                            string AnnouncementDate = item["AnnouncementDate"] != null ? item["AnnouncementDate"].ToString() : null;
                            bool isValidAnnouncementDate = DateTime.TryParseExact(AnnouncementDate, "yyyy/MM/dd", null, System.Globalization.DateTimeStyles.None, out date);
                            if (isValidAnnouncementDate) announcementDate = DateTime.ParseExact(AnnouncementDate, "yyyy/MM/dd", null);

                            string ImplementationExample = item["ImplementationExample"] != null ? item["ImplementationExample"].ToString() : throw new SystemException("【資料維護不完整】實施例欄位資料不可以為空,請重新確認~~");
                            string LensQuantity = item["LensQuantity"] != null ? item["LensQuantity"].ToString() : throw new SystemException("【資料維護不完整】透鏡數量欄位資料不可以為空,請重新確認~~");
                            string AperturePosition = item["AperturePosition"] != null ? item["AperturePosition"].ToString() : null;
                            string HalfFieldofView = item["HalfFieldofView"] != null ? item["HalfFieldofView"].ToString() : null;
                            string ImageHeight = item["ImageHeight"] != null ? item["ImageHeight"].ToString() : null;

                            string L1RefractivePower = item["L1RefractivePower"] != null ? item["L1RefractivePower"].ToString() : null;
                            string L1aSurfaceType = item["L1aSurfaceType"] != null ? item["L1aSurfaceType"].ToString() : null;
                            string L1bSurfaceType = item["L1bSurfaceType"] != null ? item["L1bSurfaceType"].ToString() : null;
                            if (L1RefractivePower == "-") L1RefractivePower = null;
                            if (L1aSurfaceType == "-") L1aSurfaceType = null;
                            if (L1bSurfaceType == "-") L1bSurfaceType = null;

                            string L2RefractivePower = item["L2RefractivePower"] != null ? item["L2RefractivePower"].ToString() : null;
                            string L2aSurfaceType = item["L2aSurfaceType"] != null ? item["L2aSurfaceType"].ToString() : null;
                            string L2bSurfaceType = item["L2bSurfaceType"] != null ? item["L2bSurfaceType"].ToString() : null;
                            if (L2RefractivePower == "-") L2RefractivePower = null;
                            if (L2aSurfaceType == "-") L2aSurfaceType = null;
                            if (L2bSurfaceType == "-") L2bSurfaceType = null;

                            string L3RefractivePower = item["L3RefractivePower"] != null ? item["L3RefractivePower"].ToString() : null;
                            string L3aSurfaceType = item["L3aSurfaceType"] != null ? item["L3aSurfaceType"].ToString() : null;
                            string L3bSurfaceType = item["L3bSurfaceType"] != null ? item["L3bSurfaceType"].ToString() : null;
                            if (L3RefractivePower == "-") L3RefractivePower = null;
                            if (L3aSurfaceType == "-") L3aSurfaceType = null;
                            if (L3bSurfaceType == "-") L3bSurfaceType = null;

                            string L4RefractivePower = item["L4RefractivePower"] != null ? item["L4RefractivePower"].ToString() : null;
                            string L4aSurfaceType = item["L4aSurfaceType"] != null ? item["L4aSurfaceType"].ToString() : null;
                            string L4bSurfaceType = item["L4bSurfaceType"] != null ? item["L4bSurfaceType"].ToString() : null;
                            if (L4RefractivePower == "-") L4RefractivePower = null;
                            if (L4aSurfaceType == "-") L4aSurfaceType = null;
                            if (L4bSurfaceType == "-") L4bSurfaceType = null;

                            string L5RefractivePower = item["L5RefractivePower"] != null ? item["L5RefractivePower"].ToString() : null;
                            string L5aSurfaceType = item["L5aSurfaceType"] != null ? item["L5aSurfaceType"].ToString() : null;
                            string L5bSurfaceType = item["L5bSurfaceType"] != null ? item["L5bSurfaceType"].ToString() : null;
                            if (L5RefractivePower == "-") L5RefractivePower = null;
                            if (L5aSurfaceType == "-") L5aSurfaceType = null;
                            if (L5bSurfaceType == "-") L5bSurfaceType = null;

                            string L6RefractivePower = item["L6RefractivePower"] != null ? item["L6RefractivePower"].ToString() : null;
                            string L6aSurfaceType = item["L6aSurfaceType"] != null ? item["L6aSurfaceType"].ToString() : null;
                            string L6bSurfaceType = item["L6bSurfaceType"] != null ? item["L6bSurfaceType"].ToString() : null;
                            if (L6RefractivePower == "-") L6RefractivePower = null;
                            if (L6aSurfaceType == "-") L6aSurfaceType = null;
                            if (L6bSurfaceType == "-") L6bSurfaceType = null;

                            string L7RefractivePower = item["L7RefractivePower"] != null ? item["L7RefractivePower"].ToString() : null;
                            string L7aSurfaceType = item["L7aSurfaceType"] != null ? item["L7aSurfaceType"].ToString() : null;
                            string L7bSurfaceType = item["L7bSurfaceType"] != null ? item["L7bSurfaceType"].ToString() : null;
                            if (L7RefractivePower == "-") L7RefractivePower = null;
                            if (L7aSurfaceType == "-") L7aSurfaceType = null;
                            if (L7bSurfaceType == "-") L7bSurfaceType = null;

                            string L8RefractivePower = item["L8RefractivePower"] != null ? item["L8RefractivePower"].ToString() : null;
                            string L8aSurfaceType = item["L8aSurfaceType"] != null ? item["L8aSurfaceType"].ToString() : null;
                            string L8bSurfaceType = item["L8bSurfaceType"] != null ? item["L8bSurfaceType"].ToString() : null;
                            if (L8RefractivePower == "-") L8RefractivePower = null;
                            if (L8aSurfaceType == "-") L8aSurfaceType = null;
                            if (L8bSurfaceType == "-") L8bSurfaceType = null;

                            string L9RefractivePower = item["L9RefractivePower"] != null ? item["L9RefractivePower"].ToString() : null;
                            string L9aSurfaceType = item["L9aSurfaceType"] != null ? item["L9aSurfaceType"].ToString() : null;
                            string L9bSurfaceType = item["L9bSurfaceType"] != null ? item["L9bSurfaceType"].ToString() : null;
                            if (L9RefractivePower == "-") L9RefractivePower = null;
                            if (L9aSurfaceType == "-") L9aSurfaceType = null;
                            if (L9bSurfaceType == "-") L9bSurfaceType = null;

                            string LXRefractivePower = item["LXRefractivePower"] != null ? item["LXRefractivePower"].ToString() : null;
                            string LXaSurfaceType = item["LXaSurfaceType"] != null ? item["LXaSurfaceType"].ToString() : null;
                            string LXbSurfaceType = item["LXbSurfaceType"] != null ? item["LXbSurfaceType"].ToString() : null;
                            if (LXRefractivePower == "-") LXRefractivePower = null;
                            if (LXaSurfaceType == "-") LXaSurfaceType = null;
                            if (LXbSurfaceType == "-") LXbSurfaceType = null;

                            #region //判斷專利資料是否存在重複
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                FROM RFI.DesignPatent
                                WHERE ApplicationNo = @ApplicationNo
                                AND ImplementationExample = @ImplementationExample
                                AND CompanyId =@CompanyId";
                            dynamicParameters.Add("ApplicationNo", ApplicationNo);
                            dynamicParameters.Add("ImplementationExample", ImplementationExample);
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() > 0)
                            {
                                repeatNum++;
                                repeatList += Serial+".【申請號:" + ApplicationNo + "】加【實例:" + ImplementationExample + "】!!<br>";
                            }
                            #endregion
                            Serial++;

                            if (repeatList.Length <= 0)
                            {
                                #region //INSERT PDM.DfmQiProcess
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO RFI.DesignPatent (CompanyId,
                                    ApplicationNo, ApplicationDate, Applicant, PublicationNo, PublicationDate, AnnouncementNo, AnnouncementDate, 
                                    ImplementationExample, LensQuantity, AperturePosition,HalfFieldofView, ImageHeight,
                                    L1RefractivePower, L1aSurfaceType, L1bSurfaceType, L2RefractivePower, L2aSurfaceType, L2bSurfaceType, 
                                    L3RefractivePower, L3aSurfaceType, L3bSurfaceType, L4RefractivePower, L4aSurfaceType, L4bSurfaceType, 
                                    L5RefractivePower, L5aSurfaceType, L5bSurfaceType, L6RefractivePower, L6aSurfaceType, L6bSurfaceType, 
                                    L7RefractivePower, L7aSurfaceType, L7bSurfaceType, L8RefractivePower, L8aSurfaceType, L8bSurfaceType, 
                                    L9RefractivePower, L9aSurfaceType, L9bSurfaceType, LXRefractivePower, LXaSurfaceType, LXbSurfaceType, 
                                    CreateDate, LastModifiedDate, CreateBy, LastModifiedBy
                                    )
                                    VALUES (@CompanyId,
                                    @ApplicationNo, @ApplicationDate, @Applicant, @PublicationNo, @PublicationDate, @AnnouncementNo, @AnnouncementDate,
                                    @ImplementationExample, @LensQuantity, @AperturePosition, @HalfFieldofView, @ImageHeight, 
                                    @L1RefractivePower, @L1aSurfaceType, @L1bSurfaceType, @L2RefractivePower, @L2aSurfaceType, @L2bSurfaceType, 
                                    @L3RefractivePower, @L3aSurfaceType, @L3bSurfaceType, @L4RefractivePower, @L4aSurfaceType, @L4bSurfaceType, 
                                    @L5RefractivePower, @L5aSurfaceType, @L5bSurfaceType, @L6RefractivePower, @L6aSurfaceType, @L6bSurfaceType, 
                                    @L7RefractivePower, @L7aSurfaceType, @L7bSurfaceType, @L8RefractivePower, @L8aSurfaceType, @L8bSurfaceType, 
                                    @L9RefractivePower, @L9aSurfaceType, @L9bSurfaceType, @LXRefractivePower, @LXaSurfaceType, @LXbSurfaceType,
                                    @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
)                                   ";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        CompanyId = CurrentCompany,
                                        ApplicationNo,
                                        ApplicationDate = isValidApplicationDate ? applicationDate : null,
                                        Applicant,
                                        PublicationNo,
                                        PublicationDate = isValidPublicationDate ? publicationDate : null,
                                        AnnouncementNo,
                                        AnnouncementDate = isValidAnnouncementDate ? announcementDate : null,
                                        ImplementationExample,
                                        LensQuantity,
                                        AperturePosition,
                                        HalfFieldofView,
                                        ImageHeight,
                                        L1RefractivePower,
                                        L1aSurfaceType,
                                        L1bSurfaceType,
                                        L2RefractivePower,
                                        L2aSurfaceType,
                                        L2bSurfaceType,
                                        L3RefractivePower,
                                        L3aSurfaceType,
                                        L3bSurfaceType,
                                        L4RefractivePower,
                                        L4aSurfaceType,
                                        L4bSurfaceType,
                                        L5RefractivePower,
                                        L5aSurfaceType,
                                        L5bSurfaceType,
                                        L6RefractivePower,
                                        L6aSurfaceType,
                                        L6bSurfaceType,
                                        L7RefractivePower,
                                        L7aSurfaceType,
                                        L7bSurfaceType,
                                        L8RefractivePower,
                                        L8aSurfaceType,
                                        L8bSurfaceType,
                                        L9RefractivePower,
                                        L9aSurfaceType,
                                        L9bSurfaceType,
                                        LXRefractivePower,
                                        LXaSurfaceType,
                                        LXbSurfaceType,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += 1;
                                #endregion
                            }
                        }
                        #endregion

                        if(repeatList.Length > 0)
                        {
                            throw new SystemException(repeatList+ "<br>以上" + repeatNum + "組組合已經存在,請重新確認!!");
                        }

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = "成功新增" + rowsAffected + " 筆資料!!"
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

        #region //AddModelDesign -- 新增鏡頭設計資料 -- SHintokuro 2024.01.08
        public string AddModelDesign(string Model, string Construction, double EFL, double FNo
            , double TTL, double CRA, string OpticalDistortionFTan, string OpticalDistortionF
            , double MaxImageCircle, string Sensor, double? VFOV, double? HFOV, double? DFOV
            , string PartName, string IR, string MechanicalRetainer, double? MechanicalFBL, string MechanicalThread
            , string IPRating, string Stage)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (Model.Length <= 0) throw new SystemException("【Model】不能為空!");
                        if (Construction.Length <= 0) throw new SystemException("【Construction】不能為空!");
                        if (EFL < 0) throw new SystemException("【EFL】不能為負!");
                        if (FNo < 0) throw new SystemException("【F/No】不能為負!");
                        if (TTL < 0) throw new SystemException("【TTL】不能為負!");
                        if (CRA < 0) throw new SystemException("【CRA】不能為負!");
                        if (MaxImageCircle < 0) throw new SystemException("【MaxImageCircle】不能為負!");

                        if (VFOV < 0) VFOV = null;
                        if (HFOV < 0) HFOV = null;
                        if (DFOV < 0) DFOV = null;
                        if (MechanicalFBL < 0) MechanicalFBL = null;
                       
                        #region //新增SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO RFI.ModelDesign (
                                    CompanyId, Model, Construction, EFL, FNo, TTL, CRA, OpticalDistortionFTan, 
                                    OpticalDistortionF, MaxImageCircle, Sensor, VFOV, HFOV, DFOV, PartName, IR, MechanicalRetainer, MechanicalFBL, 
                                    MechanicalThread, IPRating, Stage, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy
                                )
                                OUTPUT INSERTED.MdId
                                VALUES (
                                    @CompanyId, @Model, @Construction, @EFL, @FNo, @TTL, @CRA, @OpticalDistortionFTan, 
                                    @OpticalDistortionF, @MaxImageCircle, @Sensor, @VFOV, @HFOV, @DFOV, @PartName, @IR, @MechanicalRetainer, @MechanicalFBL, 
                                    @MechanicalThread, @IPRating, @Stage, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                )";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                Model,
                                Construction,
                                EFL,
                                FNo,
                                TTL,
                                CRA,
                                OpticalDistortionFTan,
                                OpticalDistortionF,
                                MaxImageCircle,
                                Sensor,
                                VFOV,
                                HFOV,
                                DFOV,
                                PartName,
                                IR,
                                MechanicalRetainer,
                                MechanicalFBL,
                                MechanicalThread,
                                IPRating,
                                Stage,
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
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
                        });
                        #endregion

                        transactionScope.Complete();
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

        #region //AddModelDesignExcel -- 上傳鏡頭設計Exce資料 -- Shintokuro 2024-01-10
        public string AddModelDesignExcel(string ModelDesignList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        int rowsAffected = 0;
                        List<Dictionary<string, string>> ModelDesignJsonList = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(ModelDesignList);


                        #region //解析Spreadsheet Data
                        foreach (var item in ModelDesignJsonList)
                        {

                            string Model = item["Model"] != null ? item["Model"].ToString() : throw new SystemException("【資料維護不完整】Model欄位資料不可以為空,請重新確認~~");
                            string Construction = item["Construction"] != null ? item["Construction"].ToString() : throw new SystemException("【資料維護不完整】Construction欄位資料不可以為空,請重新確認~~");
                            string EFL = item["EFL"] != null ? item["EFL"].ToString() : throw new SystemException("【資料維護不完整】EFL欄位資料不可以為空,請重新確認~~");
                            string FNo = item["FNo"] != null ? item["FNo"].ToString() : throw new SystemException("【資料維護不完整】FNo欄位資料不可以為空,請重新確認~~");
                            string TTL = item["TTL"] != null ? item["TTL"].ToString() : throw new SystemException("【資料維護不完整】TTL欄位資料不可以為空,請重新確認~~");
                            string CRA = item["CRA"] != null ? item["CRA"].ToString() : throw new SystemException("【資料維護不完整】CRA欄位資料不可以為空,請重新確認~~");
                            string MaxImageCircle = item["MaxImageCircle"] != null ? item["MaxImageCircle"].ToString() : null;

                            string OpticalDistortionFTan = item["OpticalDistortionFTan"] != null ? item["OpticalDistortionFTan"].ToString() :null;
                            string OpticalDistortionF = item["OpticalDistortionF"] != null ? item["OpticalDistortionF"].ToString() :null;
                            string Sensor = item["Sensor"] != null ? item["Sensor"].ToString() :null;
                            string VFOV = item["VFOV"] != null ?  item["VFOV"].ToString() :null;
                            string HFOV = item["HFOV"] != null ?  item["HFOV"].ToString() :null;
                            string DFOV = item["DFOV"] != null ? item["DFOV"].ToString() :null;
                            string PartName = item["PartName"] != null ? item["PartName"].ToString() :null;
                            string IR = item["IR"] != null ? item["IR"].ToString() :null;
                            string MechanicalRetainer = item["MechanicalRetainer"] != null ? item["MechanicalRetainer"].ToString() :null;
                            string MechanicalFBL = item["MechanicalFBL"] != null ? item["MechanicalFBL"].ToString() :null;
                            string MechanicalThread = item["MechanicalThread"] != null ? item["MechanicalThread"].ToString() :null;
                            string IPRating = item["IPRating"] != null ? item["IPRating"].ToString() :null;
                            string Stage = item["Stage"] != null ? item["Stage"].ToString() :null;
                            
                            #region //判斷專利資料是否存在重複
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                FROM RFI.ModelDesign
                                WHERE Model = @Model
                                AND CompanyId = @CompanyId";
                            dynamicParameters.Add("Model", Model);
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            var result = sqlConnection.Query(sql, dynamicParameters);
                            //if (result.Count() > 0) throw new SystemException("機種資料-【Model:" + Model + "】已經存在,請重新確認!!");
                            #endregion

                            //if (VFOV < 0) VFOV = null;
                            //if (HFOV < 0) HFOV = null;
                            //if (DFOV < 0) DFOV = null;
                            //if (MechanicalFBL < 0) MechanicalFBL = null;

                            #region //新增SQL
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO RFI.ModelDesign (
                                    CompanyId, Model, Construction, EFL, FNo, TTL, CRA, OpticalDistortionFTan, 
                                    OpticalDistortionF, MaxImageCircle, Sensor, VFOV, HFOV, DFOV, PartName, IR, MechanicalRetainer, MechanicalFBL, 
                                    MechanicalThread, IPRating, Stage, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy
                                )
                                OUTPUT INSERTED.MdId
                                VALUES (
                                    @CompanyId, @Model, @Construction, @EFL, @FNo, @TTL, @CRA, @OpticalDistortionFTan, 
                                    @OpticalDistortionF, @MaxImageCircle, @Sensor, @VFOV, @HFOV, @DFOV, @PartName, @IR, @MechanicalRetainer, @MechanicalFBL, 
                                    @MechanicalThread, @IPRating, @Stage, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                )";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    CompanyId = CurrentCompany,
                                    Model,
                                    Construction,
                                    EFL,
                                    FNo,
                                    TTL,
                                    CRA,
                                    OpticalDistortionFTan,
                                    OpticalDistortionF,
                                    MaxImageCircle,
                                    Sensor,
                                    VFOV,
                                    HFOV,
                                    DFOV,
                                    PartName,
                                    IR,
                                    MechanicalRetainer,
                                    MechanicalFBL,
                                    MechanicalThread,
                                    IPRating,
                                    Stage,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += 1;
                            #endregion
                        }
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = "成功新增" + rowsAffected + " 筆資料!!"
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

        #region //AddLensFOV -- 新增鏡頭FOV資料 -- SHintokuro 2024.01.02
        public string AddLensFOV(string Model, double FOV, double Realheight)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (Model.Length <= 0) throw new SystemException("【機種名稱】不能為空!");
                        if (FOV < 0) throw new SystemException("【FOV】不能為空!");
                        if (Realheight < 0) throw new SystemException("【Realheight】不能為空!");

                        #region //判斷FOV資料是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Model
                                    FROM RFI.LensFOV
                                    WHERE Model = @Model
                                    AND FOV = @FOV
                                    AND Realheight = @Realheight
                                    AND CompanyId = @CompanyId";
                        dynamicParameters.Add("Model", Model);
                        dynamicParameters.Add("FOV", FOV);
                        dynamicParameters.Add("Realheight", Realheight);
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("FOV資料-【Model:" + Model + "】+【FOV:" + FOV + "】+【Realheight:" + Realheight + "】該組合已經存在,請重新確認!!");
                        #endregion

                        #region //新增SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO RFI.LensFOV (CompanyId,Model, FOV, Realheight, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.LfId
                                VALUES (@CompanyId,@Model, @FOV, @Realheight, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                Model,
                                FOV,
                                Realheight,
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
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
                        });
                        #endregion

                        transactionScope.Complete();
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

        #region //AddLensFOVExcel -- 上傳LensFOVExce資料 -- Shintokuro 2024-01-11
        public string AddLensFOVExcel(string LensFOVList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        int rowsAffected = 0;
                        string repeatList = "";
                        int Serial = 2;
                        int repeatNum = 0;

                        List<Dictionary<string, string>> LensFOVLJsonist = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(LensFOVList);


                        #region //解析Spreadsheet Data
                        foreach (var item in LensFOVLJsonist)
                        {

                            string Model = item["Model"] != null ? item["Model"].ToString() : throw new SystemException("【資料維護不完整】Model欄位資料不可以為空,請重新確認~~");
                            string FOV = item["FOV"] != null ? item["FOV"].ToString() : throw new SystemException("【資料維護不完整】FOV欄位資料不可以為空,請重新確認~~");
                            string Realheight = item["Realheight"] != null ? item["Realheight"].ToString() : throw new SystemException("【資料維護不完整】Realheight欄位資料不可以為空,請重新確認~~");


                            #region //判斷FOV資料是否重複
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 Model
                                    FROM RFI.LensFOV
                                    WHERE Model = @Model
                                    AND FOV = @FOV
                                    AND Realheight = @Realheight
                                    AND CompanyId = @CompanyId";
                            dynamicParameters.Add("Model", Model);
                            dynamicParameters.Add("FOV", FOV);
                            dynamicParameters.Add("Realheight", Realheight);
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() > 0)
                            {
                                repeatNum++;
                                repeatList += Serial+ ".【Model:" + Model + "】+【FOV:" + FOV + "】+【Realheight:" + Realheight + "】<br>";
                            }
                            Serial++;
                            #endregion
                            if (repeatList.Length <= 0)
                            {
                                #region //新增SQL
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO RFI.LensFOV (CompanyId,Model, FOV, Realheight, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.LfId
                                        VALUES (@CompanyId,@Model, @FOV, @Realheight, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        CompanyId = CurrentCompany,
                                        Model,
                                        FOV,
                                        Realheight,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += 1;
                                #endregion
                            }
                        }
                        #endregion

                        if (repeatList.Length > 0)
                        {
                          throw new SystemException(repeatList+ "<br>以上"+ repeatNum + "組組合FOV資料已經存在,請重新確認!!");
                        }

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = "成功新增" + rowsAffected + " 筆資料!!"
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

        #region//Update
        #region //UpdateDesignPatentExcel -- 更新專利Exce資料 -- Shintokuro 2023-12-28
        public string UpdateDesignPatentExcel(string DpIdList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        int rowsAffected = 0;
                        List<Dictionary<string, string>> DpJsonList = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(DpIdList);

                        foreach(var item in DpJsonList)
                        {
                            DateTime date;
                            DateTime? applicationDate = null;
                            DateTime? publicationDate = null;
                            DateTime? announcementDate = null;

                            int DpId = Convert.ToInt32(item["DpId"]);
                            string ApplicationNo = item["ApplicationNo"] != null ? item["ApplicationNo"].ToString() : throw new SystemException("【資料維護不完整】申請號欄位資料不可以為空,請重新確認~~");
                            string ApplicationDate = item["ApplicationDate"] != null ? item["ApplicationDate"].ToString() : throw new SystemException("【資料維護不完整】申請日欄位資料不可以為空,請重新確認~~");
                            bool isValidApplicationDate = DateTime.TryParseExact(ApplicationDate, "yyyy/MM/dd", null, System.Globalization.DateTimeStyles.None, out date);
                            if (isValidApplicationDate) applicationDate = DateTime.ParseExact(ApplicationDate, "yyyy/MM/dd", null);

                            string Applicant = item["Applicant"] != null ? item["Applicant"].ToString() : throw new SystemException("【資料維護不完整】申請人欄位資料不可以為空,請重新確認~~");
                            string PublicationNo = item["PublicationNo"] != null && item["PublicationNo"].ToString() != "-" ? item["PublicationNo"].ToString() : null;
                            if (PublicationNo == "-") PublicationNo = null;

                            string PublicationDate = item["PublicationDate"] != null ? item["PublicationDate"].ToString() : null;
                            bool isValidPublicationDate = DateTime.TryParseExact(PublicationDate, "yyyy/MM/dd", null, System.Globalization.DateTimeStyles.None, out date);
                            if (isValidPublicationDate) publicationDate = DateTime.ParseExact(PublicationDate, "yyyy/MM/dd", null);

                            string AnnouncementNo = item["AnnouncementNo"] != null ? item["AnnouncementNo"].ToString() : null;
                            string AnnouncementDate = item["AnnouncementDate"] != null ? item["AnnouncementDate"].ToString() : null;
                            bool isValidAnnouncementDate = DateTime.TryParseExact(AnnouncementDate, "yyyy/MM/dd", null, System.Globalization.DateTimeStyles.None, out date);
                            if (isValidAnnouncementDate) announcementDate = DateTime.ParseExact(AnnouncementDate, "yyyy/MM/dd", null);
                            if (AnnouncementNo == "-") AnnouncementNo = null;

                            string ImplementationExample = item["ImplementationExample"] != null ? item["ImplementationExample"].ToString() : throw new SystemException("【資料維護不完整】實施例欄位資料不可以為空,請重新確認~~");
                            string LensQuantity = item["LensQuantity"] != null ? item["LensQuantity"].ToString() : throw new SystemException("【資料維護不完整】透鏡數量欄位資料不可以為空,請重新確認~~");
                            string AperturePosition = item["AperturePosition"] != null ? item["AperturePosition"].ToString() : null;
                            string HalfFieldofView = item["HalfFieldofView"] != null ? item["HalfFieldofView"].ToString() : null;
                            string ImageHeight = item["ImageHeight"] != null ? item["ImageHeight"].ToString() : null;

                            string L1RefractivePower = item["L1RefractivePower"] != null ? item["L1RefractivePower"].ToString() : null;
                            string L1aSurfaceType = item["L1aSurfaceType"] != null ? item["L1aSurfaceType"].ToString() : null;
                            string L1bSurfaceType = item["L1bSurfaceType"] != null ? item["L1bSurfaceType"].ToString() : null;
                            if (L1RefractivePower == "-") L1RefractivePower = null;
                            if (L1aSurfaceType == "-") L1aSurfaceType = null;
                            if (L1bSurfaceType == "-") L1bSurfaceType = null;

                            string L2RefractivePower = item["L2RefractivePower"] != null ? item["L2RefractivePower"].ToString() : null;
                            string L2aSurfaceType = item["L2aSurfaceType"] != null ? item["L2aSurfaceType"].ToString() : null;
                            string L2bSurfaceType = item["L2bSurfaceType"] != null ? item["L2bSurfaceType"].ToString() : null;
                            if (L2RefractivePower == "-") L2RefractivePower = null;
                            if (L2aSurfaceType == "-") L2aSurfaceType = null;
                            if (L2bSurfaceType == "-") L2bSurfaceType = null;

                            string L3RefractivePower = item["L3RefractivePower"] != null ? item["L3RefractivePower"].ToString() : null;
                            string L3aSurfaceType = item["L3aSurfaceType"] != null ? item["L3aSurfaceType"].ToString() : null;
                            string L3bSurfaceType = item["L3bSurfaceType"] != null ? item["L3bSurfaceType"].ToString() : null;
                            if (L3RefractivePower == "-") L3RefractivePower = null;
                            if (L3aSurfaceType == "-") L3aSurfaceType = null;
                            if (L3bSurfaceType == "-") L3bSurfaceType = null;

                            string L4RefractivePower = item["L4RefractivePower"] != null ? item["L4RefractivePower"].ToString() : null;
                            string L4aSurfaceType = item["L4aSurfaceType"] != null ? item["L4aSurfaceType"].ToString() : null;
                            string L4bSurfaceType = item["L4bSurfaceType"] != null ? item["L4bSurfaceType"].ToString() : null;
                            if (L4RefractivePower == "-") L4RefractivePower = null;
                            if (L4aSurfaceType == "-") L4aSurfaceType = null;
                            if (L4bSurfaceType == "-") L4bSurfaceType = null;

                            string L5RefractivePower = item["L5RefractivePower"] != null ? item["L5RefractivePower"].ToString() : null;
                            string L5aSurfaceType = item["L5aSurfaceType"] != null ? item["L5aSurfaceType"].ToString() : null;
                            string L5bSurfaceType = item["L5bSurfaceType"] != null ? item["L5bSurfaceType"].ToString() : null;
                            if (L5RefractivePower == "-") L5RefractivePower = null;
                            if (L5aSurfaceType == "-") L5aSurfaceType = null;
                            if (L5bSurfaceType == "-") L5bSurfaceType = null;

                            string L6RefractivePower = item["L6RefractivePower"] != null ? item["L6RefractivePower"].ToString() : null;
                            string L6aSurfaceType = item["L6aSurfaceType"] != null ? item["L6aSurfaceType"].ToString() : null;
                            string L6bSurfaceType = item["L6bSurfaceType"] != null ? item["L6bSurfaceType"].ToString() : null;
                            if (L6RefractivePower == "-") L6RefractivePower = null;
                            if (L6aSurfaceType == "-") L6aSurfaceType = null;
                            if (L6bSurfaceType == "-") L6bSurfaceType = null;

                            string L7RefractivePower = item["L7RefractivePower"] != null ? item["L7RefractivePower"].ToString() : null;
                            string L7aSurfaceType = item["L7aSurfaceType"] != null ? item["L7aSurfaceType"].ToString() : null;
                            string L7bSurfaceType = item["L7bSurfaceType"] != null ? item["L7bSurfaceType"].ToString() : null;
                            if (L7RefractivePower == "-") L7RefractivePower = null;
                            if (L7aSurfaceType == "-") L7aSurfaceType = null;
                            if (L7bSurfaceType == "-") L7bSurfaceType = null;

                            string L8RefractivePower = item["L8RefractivePower"] != null ? item["L8RefractivePower"].ToString() : null;
                            string L8aSurfaceType = item["L8aSurfaceType"] != null ? item["L8aSurfaceType"].ToString() : null;
                            string L8bSurfaceType = item["L8bSurfaceType"] != null ? item["L8bSurfaceType"].ToString() : null;
                            if (L8RefractivePower == "-") L8RefractivePower = null;
                            if (L8aSurfaceType == "-") L8aSurfaceType = null;
                            if (L8bSurfaceType == "-") L8bSurfaceType = null;

                            string L9RefractivePower = item["L9RefractivePower"] != null ? item["L9RefractivePower"].ToString() : null;
                            string L9aSurfaceType = item["L9aSurfaceType"] != null ? item["L9aSurfaceType"].ToString() : null;
                            string L9bSurfaceType = item["L9bSurfaceType"] != null ? item["L9bSurfaceType"].ToString() : null;
                            if (L9RefractivePower == "-") L9RefractivePower = null;
                            if (L9aSurfaceType == "-") L9aSurfaceType = null;
                            if (L9bSurfaceType == "-") L9bSurfaceType = null;

                            string LXRefractivePower = item["LXRefractivePower"] != null ? item["LXRefractivePower"].ToString() : null;
                            string LXaSurfaceType = item["LXaSurfaceType"] != null ? item["LXaSurfaceType"].ToString() : null;
                            string LXbSurfaceType = item["LXbSurfaceType"] != null ? item["LXbSurfaceType"].ToString() : null;
                            if (LXRefractivePower == "-") LXRefractivePower = null;
                            if (LXaSurfaceType == "-") LXaSurfaceType = null;
                            if (LXbSurfaceType == "-") LXbSurfaceType = null;

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE RFI.DesignPatent SET
                                    ApplicationNo = @ApplicationNo,
                                    ApplicationDate = @ApplicationDate,
                                    Applicant = @Applicant,
                                    PublicationNo = @PublicationNo,
                                    PublicationDate = @PublicationDate,
                                    AnnouncementNo = @AnnouncementNo,
                                    AnnouncementDate = @AnnouncementDate,
                                    ImplementationExample = @ImplementationExample,
                                    LensQuantity = @LensQuantity,
                                    AperturePosition = @AperturePosition,
                                    HalfFieldofView = @HalfFieldofView,
                                    ImageHeight = @ImageHeight,
                                    L1RefractivePower = @L1RefractivePower,
                                    L1aSurfaceType = @L1aSurfaceType,
                                    L1bSurfaceType = @L1bSurfaceType,
                                    L2RefractivePower = @L2RefractivePower,
                                    L2aSurfaceType = @L2aSurfaceType,
                                    L2bSurfaceType = @L2bSurfaceType,
                                    L3RefractivePower = @L3RefractivePower,
                                    L3aSurfaceType = @L3aSurfaceType,
                                    L3bSurfaceType = @L3bSurfaceType,
                                    L4RefractivePower = @L4RefractivePower,
                                    L4aSurfaceType = @L4aSurfaceType,
                                    L4bSurfaceType = @L4bSurfaceType,
                                    L5RefractivePower = @L5RefractivePower,
                                    L5aSurfaceType = @L5aSurfaceType,
                                    L5bSurfaceType = @L5bSurfaceType,
                                    L6RefractivePower = @L6RefractivePower,
                                    L6aSurfaceType = @L6aSurfaceType,
                                    L6bSurfaceType = @L6bSurfaceType,
                                    L7RefractivePower = @L7RefractivePower,
                                    L7aSurfaceType = @L7aSurfaceType,
                                    L7bSurfaceType = @L7bSurfaceType,
                                    L8RefractivePower = @L8RefractivePower,
                                    L8aSurfaceType = @L8aSurfaceType,
                                    L8bSurfaceType = @L8bSurfaceType,
                                    L9RefractivePower = @L9RefractivePower,
                                    L9aSurfaceType = @L9aSurfaceType,
                                    L9bSurfaceType = @L9bSurfaceType,
                                    LXRefractivePower = @LXRefractivePower,
                                    LXaSurfaceType = @LXaSurfaceType,
                                    LXbSurfaceType = @LXbSurfaceType,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE DpId = @DpId";
                            dynamicParameters.AddDynamicParams(
                              new
                              {
                                  ApplicationNo,
                                  ApplicationDate = isValidApplicationDate ? applicationDate : null,
                                  Applicant,
                                  PublicationNo,
                                  PublicationDate = isValidPublicationDate ? publicationDate : null,
                                  AnnouncementNo,
                                  AnnouncementDate = isValidAnnouncementDate ? announcementDate : null,
                                  ImplementationExample,
                                  LensQuantity,
                                  AperturePosition,
                                  HalfFieldofView,
                                  ImageHeight,
                                  L1RefractivePower,
                                  L1aSurfaceType,
                                  L1bSurfaceType,
                                  L2RefractivePower,
                                  L2aSurfaceType,
                                  L2bSurfaceType,
                                  L3RefractivePower,
                                  L3aSurfaceType,
                                  L3bSurfaceType,
                                  L4RefractivePower,
                                  L4aSurfaceType,
                                  L4bSurfaceType,
                                  L5RefractivePower,
                                  L5aSurfaceType,
                                  L5bSurfaceType,
                                  L6RefractivePower,
                                  L6aSurfaceType,
                                  L6bSurfaceType,
                                  L7RefractivePower,
                                  L7aSurfaceType,
                                  L7bSurfaceType,
                                  L8RefractivePower,
                                  L8aSurfaceType,
                                  L8bSurfaceType,
                                  L9RefractivePower,
                                  L9aSurfaceType,
                                  L9bSurfaceType,
                                  LXRefractivePower,
                                  LXaSurfaceType,
                                  LXbSurfaceType,
                                  LastModifiedDate,
                                  LastModifiedBy,
                                  DpId
                              });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        }


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

        #region //UpdateModelDesign -- 更新機種設計資料 -- SHintokuro 2024.01.08
        public string UpdateModelDesign(int MdId, string Model, string Construction, double EFL, double FNo
            , double TTL, double CRA, string OpticalDistortionFTan, string OpticalDistortionF
            , double MaxImageCircle, string Sensor, double? VFOV, double? HFOV, double? DFOV
            , string PartName, string IR, string MechanicalRetainer, double? MechanicalFBL, string MechanicalThread
            , string IPRating, string Stage)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (Model.Length <= 0) throw new SystemException("【Model】不能為空!");
                        if (Construction.Length <= 0) throw new SystemException("【Construction】不能為空!");
                        if (EFL < 0) throw new SystemException("【EFL】不能為負!");
                        if (FNo < 0) throw new SystemException("【F/No】不能為負!");
                        if (TTL < 0) throw new SystemException("【TTL】不能為負!");
                        if (CRA < 0) throw new SystemException("【CRA】不能為負!");
                        if (MaxImageCircle < 0) throw new SystemException("【MaxImageCircle】不能為負!");

                        if (VFOV < 0) VFOV = null;
                        if (HFOV < 0) HFOV = null;
                        if (DFOV < 0) DFOV = null;
                        if (MechanicalFBL < 0) MechanicalFBL = null;

                        #region //判斷機種設計資料是否存在正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM RFI.ModelDesign
                                WHERE MdId = @MdId";
                        dynamicParameters.Add("MdId", MdId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("機種資料找不到請重新確認!!");
                        #endregion

                        #region //更新SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE RFI.ModelDesign SET
                                Model= @Model,
                                Construction= @Construction,
                                EFL= @EFL,
                                FNo= @FNo,
                                TTL= @TTL,
                                CRA= @CRA,
                                OpticalDistortionFTan= @OpticalDistortionFTan,
                                OpticalDistortionF= @OpticalDistortionF,
                                MaxImageCircle= @MaxImageCircle,
                                Sensor= @Sensor,
                                VFOV= @VFOV,
                                HFOV= @HFOV,
                                DFOV= @DFOV,
                                PartName= @PartName,
                                IR= @IR,
                                MechanicalRetainer= @MechanicalRetainer,
                                MechanicalFBL= @MechanicalFBL,
                                MechanicalThread= @MechanicalThread,
                                IPRating= @IPRating,
                                Stage= @Stage,
                                CreateDate= @CreateDate,
                                LastModifiedDate= @LastModifiedDate,
                                CreateBy= @CreateBy,
                                LastModifiedBy= @LastModifiedBy
                                WHERE MdId = @MdId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Model,
                                Construction,
                                EFL,
                                FNo,
                                TTL,
                                CRA,
                                OpticalDistortionFTan,
                                OpticalDistortionF,
                                MaxImageCircle,
                                Sensor,
                                VFOV,
                                HFOV,
                                DFOV,
                                PartName,
                                IR,
                                MechanicalRetainer,
                                MechanicalFBL,
                                MechanicalThread,
                                IPRating,
                                Stage,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy,
                                MdId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
                        });
                        #endregion

                        transactionScope.Complete();
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

        #region //UpdateLensFOV -- 更新鏡頭FOV資料 -- SHintokuro 2024.01.02
        public string UpdateLensFOV(int LfId, string Model, double FOV, double Realheight)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (Model.Length <= 0) throw new SystemException("【機種名稱】不能為空!");

                        #region //判斷機種資料是否存在正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM RFI.LensFOV
                                WHERE LfId = @LfId";
                        dynamicParameters.Add("LfId", LfId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【鏡頭FOV管理】 - 機種資料找不到請重新確認!!");
                        #endregion

                        #region //判斷FOV資料是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Model
                                    FROM RFI.LensFOV
                                    WHERE Model = @Model
                                    AND FOV = @FOV
                                    AND Realheight = @Realheight
                                    AND LfId != @LfId
                                    AND CompanyId = @CompanyId";
                        dynamicParameters.Add("Model", Model);
                        dynamicParameters.Add("FOV", FOV);
                        dynamicParameters.Add("Realheight", Realheight);
                        dynamicParameters.Add("LfId", LfId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("FOV資料-【Model:" + Model + "】+【FOV:" + FOV + "】+【Realheight:" + Realheight + "】該組合已經存在,請重新確認!!");
                        #endregion
                        

                        #region //更新SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE RFI.LensFOV SET
                                Model = @Model,
                                FOV = @FOV,
                                Realheight = @Realheight,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE LfId = @LfId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Model,
                                FOV,
                                Realheight,
                                LastModifiedDate,
                                LastModifiedBy,
                                LfId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
                        });
                        #endregion

                        transactionScope.Complete();
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
        #endregion

        #region//Delete
        #region //DelteDesignPatentExcel -- 刪除專利Exce資料 -- Shintokuro 2023-12-28
        public string DelteDesignPatentExcel(string DpIdList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        int rowsAffected = 0;
                        List<int> dpIdList = DpIdList.Split(',').Select(int.Parse).ToList();
                        foreach (var DpId in dpIdList)
                        {
                            #region //判斷專利資料是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM RFI.DesignPatent
                                    WHERE DpId = @DpId";
                            dynamicParameters.Add("DpId", DpId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("專利資料查不到,請重新確認!!");
                            #endregion
                        }


                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE RFI.DesignPatent
                                WHERE DpId in ( "+ DpIdList + ")";
                        dynamicParameters.Add("DpId", DpIdList);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteModelDesign -- 刪除機種設計資料 -- SHintokuro 2024.01.08
        public string DeleteModelDesign(int MdId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷機種資料是否存在正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM RFI.ModelDesign
                                WHERE MdId = @MdId";
                        dynamicParameters.Add("MdId", MdId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("機種資料找不到請重新確認!!");
                        #endregion

                        

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE RFI.ModelDesign
                                WHERE MdId = @MdId";
                        dynamicParameters.Add("MdId", MdId);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                        });
                        #endregion

                        transactionScope.Complete();
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

        #region //DeleteLensFOV -- 刪除鏡頭FOV資料 -- SHintokuro 2024.01.02
        public string DeleteLensFOV(int LfId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        #region //判斷機種資料是否存在正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM RFI.LensFOV
                                WHERE LfId = @LfId";
                        dynamicParameters.Add("LfId", LfId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【鏡頭FOV管理】 - 機種資料找不到請重新確認!!");
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE RFI.LensFOV
                                WHERE LfId = @LfId";
                        dynamicParameters.Add("LfId", LfId);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                        });
                        #endregion

                        transactionScope.Complete();
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
        #endregion

        #region//PDF
        #endregion
    }
}
