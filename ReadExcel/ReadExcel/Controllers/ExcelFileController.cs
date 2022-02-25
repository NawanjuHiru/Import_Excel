using ReadExcel.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReadExcel.Controllers
{
    public class ExcelFileController : Controller
    {
        // GET: ExcelFile
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Import(HttpPostedFileBase excelFile)
        {
            if (excelFile.ContentLength == 0 || excelFile == null)
            {
                ViewBag.Error = "Please Select Excel File";
                return View("Index");
            }
            else
            {
                if (excelFile.FileName.EndsWith("xlsm"))
                {
                    Excel.Application application = new Excel.Application();
                    Excel.Workbook workbook = application.Workbooks.Open("C:/Users/User/Downloads/" + excelFile.FileName);
                    Excel.Worksheet worksheet = workbook.ActiveSheet;
                    Excel.Range range = worksheet.UsedRange;

                    List<ExcelPackageExtensions> excel = new List<ExcelPackageExtensions>();

                    for(int row=2; row <=range.Rows.Count; row++)
                    {
                        ExcelPackageExtensions exc = new ExcelPackageExtensions();
                        exc.SiteId = ((Excel.Range)range.Cells[row, 1]).Text;
                        exc.Sector = ((Excel.Range)range.Cells[row, 2]).Text;
                        exc.UpgradeBatch = ((Excel.Range)range.Cells[row, 3]).Text;
                        exc.UpgradeType = ((Excel.Range)range.Cells[row, 4]).Text;
                        exc.PriorityArea = ((Excel.Range)range.Cells[row, 5]).Text;
                        exc.PlanMonth = ((Excel.Range)range.Cells[row, 6]).Text;
                        exc.PlanWeek = ((Excel.Range)range.Cells[row, 7]).Text;
                        exc.DoneDate = ((Excel.Range)range.Cells[row, 8]).Text;
                        exc.Status = ((Excel.Range)range.Cells[row, 9]).Text;
                        exc.Reason = ((Excel.Range)range.Cells[row, 10]).Text;
                        exc.Remarks = ((Excel.Range)range.Cells[row, 11]).Text;
                        exc.QOS_Status = ((Excel.Range)range.Cells[row, 12]).Text;

                        excel.Add(exc);
                    }

                    if (excel != null)
                    {
                        ViewBag.ExcelList = excel;
                        return View("View");
                    }
                    else
                    {
                        ViewBag.Error = "Something went wrong";
                        return View("Index");
                    }
                }
                else
                {
                    ViewBag.Error = "Something went wrong";
                    return View("Index");
                }
            }
        }
    }
}