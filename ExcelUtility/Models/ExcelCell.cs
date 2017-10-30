using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelUtility.Models
{
    public class ExcelCell
    {
        public bool CellValid { get; set; }

        public string ColorCode { get; set; }

        public string PropertyName { get; set; }

        public string Message { get; set; }
    }
}