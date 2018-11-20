using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ExcelHelperCommon;
using CommonHelper;
using Models;
namespace DepositReceiptManagementSystem.Controllers
{
    public class IndexController : Controller
    {
        // GET: Index
        public ActionResult Index()
        {
            return View();
        }
        public FileResult Upload(HttpPostedFileBase fileUpload)
        {
            if (fileUpload == null)
            {
                throw new Exception("文件为空");
            }
            try
            {
                //将硬盘路径转化为服务器路径的文件流
                string fileName = Path.Combine(Request.MapPath("~/SaveFile"), Path.GetFileName(fileUpload.FileName));
                if (!Directory.Exists(Request.MapPath("~/SaveFile")))
                {
                    Directory.CreateDirectory(Request.MapPath("~/SaveFile"));
                }
                if (System.IO.File.Exists(fileName))
                {
                    System.IO.File.Delete(fileName);
                }
                //NPOI得到EXCEL的第一种方法              
                fileUpload.SaveAs(fileName);

                DataTable dtData = ImportHelper.RenderDataTableFromExcel(fileName, "订单信息表", 1);
                List<Deposit> list = ModelConvertHelper<Deposit>.DataTableToList(dtData, AppDomain.CurrentDomain.BaseDirectory + @"/ExcelModel/" + @"DepositReceiptExcelView.xml")
                    .OrderBy(p => p.Drugname)
                    .ThenBy(p => p.Number)
                    .ThenBy(p => p.Quantity)
                    .ThenBy(p => p.OrderTime)
                    .ToList();
                KeyValuePair<string, string> keyValuePair = new KeyValuePair<string, string>("订金信息表", "订金信息表");
                var fileNameNew = "订金信息表";
                var ms = ExportHelper<Deposit>.CreateExcelStreamByDatas(list, keyValuePair, AppDomain.CurrentDomain.BaseDirectory + @"/ExcelModel/" + @"DepositReceiptExcelView.xml", ref fileNameNew);

                return ExportToExcel(ms, fileNameNew);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public FileResult ExportToExcel(MemoryStream ms, string fileName)
        {
            return File(ms, "application/ms-excel", fileName);
        }
    }
}