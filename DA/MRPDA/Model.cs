using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MRPDA
{
    #region //MRP需求明細 DemandLine
    public class DemandLine
    {
        public int? DemandId { get; set; }
        public int? DemandLineId { get; set; }
        public int? DemandPriority { get; set; }
        public string SourceType { get; set; }
        public int? SourceId { get; set; }
        public string SourceNo { get; set; }
        public int? SourceSeq { get; set; }
        public int? MtlItemId { get; set; }
        public string MtlItemNo { get; set; }
        public string MakeType { get; set; }
        public double? Quantity { get; set; }
        public int? CustomerId { get; set; }
        public DateTime? ScheduleDate { get; set; }
        public DateTime? ExpectDeliveryDate { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //MRP需求明細 DemandLineDtl
    public class DemandLineDtl
    {
        public int? DemandLineId { get; set; }
        public int? LineDtlId { get; set; }
        public int? BomLevel { get; set; }
        public string MtlItemNo { get; set; }
        public int? MtlItemId { get; set; }
        public int? ParentMtlItemId { get; set; }
        public string MakeType { get; set; }
        public double? Quantity { get; set; }
        public double? CompositionQuantity { get; set; }
        public double? Base { get; set; }
        public DateTime? ScheduleDate { get; set; }
        public int? WorkDays { get; set; }
        public DateTime? ExpectFinishDate { get; set; }
        public string Process { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
        public int? ProcessStatus { get; set; }
    }
    #endregion

    #region //MRP需求明細 MtlItemQuantity
    public class MtlItemQuantity
    {
        public string MtlItemNo { get; set; }
        public double? OpenPoQty { get; set; }
        public double? OpenWipQty { get; set; }
        public double? OnhandQty { get; set; }        
    }
    #endregion

    public class MtlItemUseQty
    {
        public int? DemandLineId { get; set; }
        public string MtlItemNo { get; set; }
        public double? UsePoQty { get; set; }
        public double? UseWipQty { get; set; }
        public double? UseOnhandQty { get; set; }
    }
}
