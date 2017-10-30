using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ExcelUtility.Commons;
using ExcelUtility.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;

namespace ExcelUtility.Controllers
{
    public class HomeController : BaseController
    {
        private Interface1 inter1;
        private Interface inter;

        public HomeController()
        {
            inter = new Class();
            inter1 = new Class();
        }

        public IActionResult Index()
        {
            int val1 = inter.sum(10, 15);
            int val2 = inter1.sum(20, 25);
            return View();
        }

        public IActionResult About()
        {
            ViewData["Message"] = "Your application description page.";

            return View();
        }

        public IActionResult Contact()
        {
            ViewData["Message"] = "Your contact page.";

            return View();
        }

        public IActionResult DownloadTemplate()
        {
            string excelPath = "Template/Customer.xlsx";
            if (System.IO.File.Exists(Path.Combine(excelPath)))
            {
                string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                byte[] bytes = System.IO.File.ReadAllBytes(Path.Combine(excelPath));
                return this.File(bytes, contentType, Path.GetFileName(excelPath));
            }
            else
            {
                return this.Content("File doesn't exist.");
            }
        }

        public IActionResult Error()
        {
            return View();
        }

        public IActionResult ExcelImport()
        {
            return View();
        }

        [HttpPost]
        public IActionResult ExcelImport(IFormFile file)
        {
            var statesList = JsonConvert.DeserializeObject<List<MasterData>>(JsonData.StateJosn);
            var exitingCustomerList = JsonConvert.DeserializeObject<List<CustomerViewModel>>(JsonData.exitingCustomerJson);
            ExcelResponse excel = new ExcelResponse();
            Dictionary<string, IEnumerable<string>> cellExistin = new Dictionary<string, IEnumerable<string>>
            {
                { "BillingState", statesList.Select(x => x.Name) }
            };

            ExcelSheet<CustomerExcelModel> customerExcelSheet = this.FromExcel<CustomerExcelModel>(file.OpenReadStream(), "Customer", cellExistin, 2);
            if (!customerExcelSheet.Status)
            {
                excel.Message = customerExcelSheet.Message;
                return this.Json(excel);
            }
            return View();
        }

        public IActionResult ExcelExport()
        {
            List<CustomerExcelModel> exitingCustomerList = JsonConvert.DeserializeObject<List<CustomerExcelModel>>(JsonData.customerExportJson);
            byte[] bytes = this.ToExcel<CustomerExcelModel>(new FileInfo(Path.Combine(Constants.ExcelTemplateCustomer)), exitingCustomerList, "Customer", 2);
            string fileName = $"{Path.GetFileNameWithoutExtension(Constants.ExcelTemplateCustomer)}_{Convert.ToString(DateTime.Now.Ticks)}.xlsx";
            return this.File(bytes, "application/vnd.ms-excel", fileName);
        }
    }
}