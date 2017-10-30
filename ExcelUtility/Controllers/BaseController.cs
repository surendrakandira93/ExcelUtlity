using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ExcelUtility.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

// For more information on enabling MVC for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace ExcelUtility.Controllers
{
    public class BaseController : Controller
    {
        public dynamic FromExcel<TEntity>(System.IO.Stream excelfile, string sheetName, Dictionary<string, IEnumerable<string>> cellExistin = null, int startRow = 3)
              where TEntity : class
        {
            if (cellExistin == null)
            {
                cellExistin = new Dictionary<string, IEnumerable<string>>();
            }

            var excelSheet = new ExcelSheet<TEntity>();
            excelSheet.Name = sheetName;

            using (ExcelPackage package = new ExcelPackage(excelfile))
            {
                if (package.Workbook.Worksheets.Count == 0)
                {
                    excelSheet.Status = false;
                    excelSheet.Message = $"{sheetName} not Found in Excel File.";
                    return excelSheet;
                }

                ExcelWorksheet workSheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name.ToLower().Trim() == sheetName.ToLower().Trim());
                if (workSheet != null)
                {
                    startRow = workSheet.Row(1).RowEmpty(workSheet) ? startRow - 1 : startRow;
                    int rowindex = startRow == 0 ? 1 : startRow;

                    Reflections.ReturnPeriodId = 4;
                    Reflections.YearId = 4;
                    excelSheet.Records = cellExistin.Count > 0 ? workSheet.ToExcelSheet<TEntity>(cellExistin, rowindex) : workSheet.ToExcelSheet<TEntity>(rowindex);

                    excelSheet.Name = workSheet.Name;
                    excelSheet.Status = true;
                    excelSheet.Message = excelSheet.Records.Length > 0 ? string.Empty : $"No Data Found in {sheetName}.";
                    return excelSheet;
                }
                else
                {
                    excelSheet.Status = false;
                    excelSheet.Message = $"{sheetName} not Found in Excel File.";
                    return excelSheet;
                }
            }
        }

        public byte[] ToExcel<TEntity>(System.IO.FileInfo excelfile, dynamic entityList, string sheetName, int startRow)
          where TEntity : class
        {
            using (ExcelPackage package = new ExcelPackage(excelfile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name.ToLower().Trim() == sheetName.ToLower().Trim());
                worksheet = worksheet.ToEntityList<TEntity>(entityList as List<TEntity>, startRow);
                return package.GetAsByteArray();
            }
        }
    }
}