using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace ReadExcel.Models
{
    public class ExcelPackageExtensions
    {
        public string SiteId { get; set; }
        public string Sector { get; set; }
        public string UpgradeBatch { get; set; }
        public string UpgradeType { get; set; }
        public string PriorityArea { get; set; }
        public string PlanMonth { get; set; }
        public string PlanWeek { get; set; }
        public string DoneDate { get; set; }
        public string Status { get; set; }
        public string Reason { get; set; }
        public string Remarks { get; set; }
        public string QOS_Status { get; set; }
    }
}