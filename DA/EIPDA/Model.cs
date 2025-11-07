using System.Collections.Generic;

namespace EIPDA
{
    public class RfqMainVM
    {
        int idx { get; set; }
        public string AssemblyName { get; set; }
        public int ProductUseId { get; set; }
        public string ProductUseName { get; set; }
        public int CustomerId { get; set; } 
        public List<RfqDetailVM> RfqDetail { get; set; }

        public RfqMainVM()
        {
            RfqDetail = new List<RfqDetailVM>();
        }
    }

    public class RfqDetailVM
    {
        public int idx { get; set; }
        public int RfqPkTypeId { get; set; }
        public int RfqProClassId { get; set; }
        public int RfqProTypeId { get; set; }
        public int LifeCycleQty { get; set; }
        public int PrototypeQty { get; set; }
        public string MtlName { get; set; }
        public string ProClassName { get; set; }
        public string ProTypeName { get; set; }
        public string CoatingFlag { get; set; }
        public string Description { get; set; }
        public string DemandDate { get; set; }
        public string ProdLifeCycleEnd { get; set; }
        public string ProdLifeCycleStart { get; set; }

        public List<RfqLineSolutionVM> RfqLineSolutionList { get; set; }

        public RfqDetailVM()
        {
            RfqLineSolutionList = new List<RfqLineSolutionVM>();
        }
    }

    public class RfqLineSolutionVM
    {
        public int idx { get; set; }
        public int RfqLineSolutionId { get; set; }
        public string PeriodicDemandType { get; set; }
        public int SolutionQty { get; set; }
    }

    public class CustWithCorpVM
    {
        public int CompanyId { get; set; }
        public string CompanyNo { get; set; }
        public string CompanyName { get; set; }
        public int LogoIcon { get; set; }
        public int CustomerId { get; set; }
        public string Status { get; set; }
        public int TotalCount { get; set; }
    }

    public class MemberCustCorpVM
    {
        public string CustomerIds { get; set; }
        public int MemberId { get; set; }
        public int CsCustId { get; set; }
        public string MemberEmail { get; set; }
        public string MemberName { get; set; }
        public string Gender { get; set; }
        public int MemberIcon { get; set; }
        public string OrgShortName { get; set; }
        public string EmailWithMember { get; set; }
        public string CustomerEnglishName { get; set; }
        public string CustomerName { get; set; }
        public string CNameAndEName { get; set; }
        public string Customers { get; set; }
        public int TotalCount { get; set; }
    }

    public class CorpCustVM
    {
        public string CustomerIds { get; set; }
        public int CsCustId { get; set; }
        public string StatusWithName { get; set; }
        public string CustomerEnglishName { get; set; }
        public string CustomerName { get; set; }
        public string CNameAndEName { get; set; }
        public string Customers { get; set; }
        public string Status { get; set; }
        public int TotalCount { get; set; }
    }
}