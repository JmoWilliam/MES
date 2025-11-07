using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EBPDA
{
    public class MonthUmo
    {
        public string UserNo { get; set; }
        public string UserName { get; set; }
        public string DepartmentNo { get; set; }
        public string DepartmentName { get; set; }
        public string Category { get; set; }
        public string Date { get; set; }
        public double TotalQty { get; set; }
        public double PriceAmount { get; set; }
        public double UmoDiscount { get; set; }
    }

    public class Pickup
    {
        public string UserNo { get; set; }
        public string UserName { get; set; }
        public string DepartmentName { get; set; }
        public string MealName { get; set; }
        public string Place { get; set; }
        public int Qty { get; set; }
    }

    public class TotalMealList
    {
        public DateTime Date { get; set; }
        public int? CompanyId { get; set; }
        public string CompanyName { get; set; }
        public string UserNo { get; set; }
        public string RestaurantName { get; set; }
        public double UmoDetailQty { get; set; }
        public double UmoAmount { get; set; }
        public double UmoDiscount { get; set; }
    }
    public class DailyReport
    {
        public DateTime Date { get; set; }
        public int? CompanyId { get; set; }
        public string CompanyName { get; set; }
        public string RestaurantName { get; set; }
        public string MealName { get; set; }
        public double UmoDetailQty { get; set; }
        public double UmoAmount { get; set; }
    }

    public class ReportData
    {
        public string UserNo { get; set; }

        public string UserName { get; set; }
        public string DepartmentNo { get; set; }
        public string DepartmentName { get; set; }
        public string DepartmentCategory { get; set; }
        public string UmoDate { get; set; }
        public string Date { get; set; }
        public double UmoDetailQty { get; set; }
        public double UmoAmount { get; set; }
        public double UmoDiscount { get; set; }
        public string MealName { get; set; }
        public string Pickup { get; set; }
        public int? CompanyId { get; set; }
        public string CompanyName { get; set; }
        public string RestaurantName { get; set; }
    }

    public class DetailReportData
    {
        public string UserNo { get; set; }

        public string UserName { get; set; }
        public string DepartmentNo { get; set; }
        public string DepartmentName { get; set; }
        public string DepartmentCategory { get; set; }
        public DateTime UmoDate { get; set; }
        public double UmoDetailQty { get; set; }
        public double UmoDetailPrice { get; set; }
        public string MealName { get; set; }
        public string Pickup { get; set; }
        public int? CompanyId { get; set; }
        public string CompanyName { get; set; }
        public string RestaurantName { get; set; }
    }

    public class CalendarData
    {
        public string Date { get; set; }
        public string Week { get; set; }
        public bool isHoliday { get; set; }
        public string Description { get; set; }
    }
}
