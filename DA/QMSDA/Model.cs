using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QMSDA
{
    #region //量測項目
    public class QcItem
    {
        public int? QcItemId { get; set; }
        public int? QcClassId { get; set; }
        public string QcItemNo { get; set; }
        public string QcItemName { get; set; }
        public string QcItemDesc { get; set; }
        public string QcItemType { get; set; }
        public string QcType { get; set; }
        public string Status { get; set; }
        public string Remark { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //量測項目輔助表
    public class QcItemPrinciple
    {
        public int? PrincipleId { get; set; }
        public int? QcClassId { get; set; }
        public int? QmmDetailId { get; set; }
        public string PrincipleNo { get; set; }
        public string PrincipleDesc { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //量測項目輔助表 附加欄位
    public class PrincipleDetail
    {
        public int? PdId { get; set; }
        public int? PrincipleId { get; set; }
        public string PrincipleDesc { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //ResolveFile 解析檔案相關Modal

    #region //UA3P專用解析Model
    public class ResolveFileOfUA3PModel
    {
        public int FileId { get; set; }
        public string MachineNumber { get; set; }
        public string MachineName { get; set; }
        public string EffectiveDiameterR1 { get; set; }
        public string EffectiveDiameterR2 { get; set; }
        public string UserNo { get; set; }
    }
    #endregion

    #region //共用框架結果Modal
    public class ResolveFileResultModel
    {
        public string QcItemNo { get; set; }
        public string QcItemName { get; set; }
        public string LetteringNo { get; set; }
        public string MachineNumber { get; set; }
        public string MachineName { get; set; }
        public string MeasureValue { get; set; }
        public string QcItemDesc { get; set; }
        public double? DesignValue { get; set; }
        public double? UpperTolerance { get; set; }
        public double? LowerTolerance { get; set; }
        public string UserNo { get; set; }
        public double? ZAxis { get; set; }
        public string ErrorMessage { get; set; }
    }
    #endregion

    #region //QV量測項目上下限Model
    public class QvDesignValueModel
    {
        public string QcItemName { get; set; }
        public double? DesignValue { get; set; }
        public double? UpperTolerance { get; set; }
        public double? LowerTolerance { get; set; }
        public string ErrorMessage { get; set; }
    }
    #endregion

    #endregion

    #region //Spreadsheet Modal
    public class SpreadsheetModel
    {
        public List<Sheets> sheets { get; set; }
        public Formats formats { get; set; }
    }

    public class Sheets
    {
        public string name { get; set; }
        public List<Data> data { get; set; }
        public Cols cols { get; set; }
        public Rows rows { get; set; }
    }

    public class Data
    {
        public string cell { get; set; }
        public string css { get; set; }
        public string format { get; set; }
        public string value { get; set; }
    }

    public class Cols
    {
        public double? width { get; set; }
    }

    public class Rows
    {
        public double? height { get; set; }
    }

    public class Formats
    {
        public string name { get; set; }
        public string id { get; set; }
        public string mask { get; set; }
        public string example { get; set; }
    }

    public class QcMeasureDataTemp
    {
        public int? QcRecordId { get; set; }
        public string Cell { get; set; }
        public string Css { get; set; }
        public string Format { get; set; }
        public string Value { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //量測報表

    #region //上傳修改後excel解析用
    public class UploadModifyDataModel
    {
        public int? QmdId { get; set; }
        public int? QcRecordId { get;set;}
        public int? MoId { get; set; }
        public int? QcItemId { get; set; }
        public string QcItemDesc { get; set; }
        public string DesignValue { get; set; }
        public string UpperTolerance { get; set; }
        public string LowerTolerance { get; set; }
        public string ZAxis { get; set; }
        public string ModifyValue { get; set; }
        public int? BarcodeId { get; set; }
        public string LetteringSeq { get; set; }
        public string QmmDetailId { get; set; }
        public string BallMark { get; set; }
        public string Unit { get; set; }
        public string Surveyor { get; set; }
        public string Confirmer { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //點資料用
    public class QcMeasurePointData
    {
        public int? QmdId { get; set; }
        public int? QcRecordFileId { get; set; }
        public int? BarcodeId { get; set; }
        public int? QcItemId { get; set; }
        public string LotNumber { get; set; }
        public string Point { get; set; }
        public string PointValue { get; set; }
        public string Axis { get; set; }
        public string ZAxis { get; set; }
        public string Cavity { get; set; }
        public string PointSite { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #endregion

    #region //數據
    public class QcMeasurementData
    {
        public int QmdId { get; set; } // identity
        public int QcRecordId { get; set; } // NOT NULL
        public int? QcNoticeItemSpecId { get; set; }
        public int? QcItemId { get; set; }
        public string QcItemDesc { get; set; }
        public string DesignValue { get; set; }
        public string UpperTolerance { get; set; }
        public string LowerTolerance { get; set; }
        public double? ZAxis { get; set; }
        public int? BarcodeId { get; set; }
        public int QmmDetailId { get; set; } // NOT NULL
        public string MeasureValue { get; set; } = string.Empty; // NOT NULL
        public string QcResult { get; set; } = string.Empty; // NOT NULL
        public int? CauseId { get; set; }
        public string CauseDesc { get; set; }
        public string Cavity { get; set; }
        public string LotNumber { get; set; }
        public int? MakeCount { get; set; }
        public string CellHeader { get; set; }
        public string Row { get; set; }
        public string Remark { get; set; }
        public int? QcUserId { get; set; }
        public int? QcCycleTime { get; set; }
        public string BallMark { get; set; }
        public string Unit { get; set; }
        public DateTime CreateDate { get; set; } = DateTime.Now; // NOT NULL with default value
        public DateTime? LastModifiedDate { get; set; }
        public int CreateBy { get; set; } // NOT NULL
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //批次匯入EXCEL建立量測單據
    public class BmFileInfo
    {
        public int? FileId { get; set; }
        public string FileName { get; set; }
        public byte[] FileContent { get; set; }
        public string FileExtension { get; set; }
        public int FileSize { get; set; }
        public string ClientIP { get; set; }
        public string Source { get; set; }
        public string DeleteStatus { get; set; }
    }

    public class QcExcelFormat
    {
        public string QcType { get; set; }
        public string InputType { get; set; }
        public string WoErpFullNo { get; set; }
        public string MtlItemNo { get; set; }
        public string Remark { get; set; }
        public string QcItemFormatPath { get; set; }
    }

    public class QcExcelFormatMulti
    {
        public string Group { get; set; }
        public string QcType { get; set; }
        public string WoErpFullNo { get; set; }
        public string LotNumber { get; set; }
        public string OrderRemark { get; set; }
        public string QcItemTemplet { get; set; }
        public string QcItemNo { get; set; }
        public string QcItemRemark { get; set; }
        public string QcItemFormatPath { get; set; }
        public string MachineSetting { get; set; }

    }
    #endregion
}
