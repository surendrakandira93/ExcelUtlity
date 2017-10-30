using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelUtility.Models
{
    public class ExcelResponse
    {
        public bool Status { get; set; }

        public string Message { get; set; }

        public Dictionary<string, string> HtmlTables { get; set; } = new Dictionary<string, string>();

        public dynamic ExcelSheet { get; set; }
    }
}