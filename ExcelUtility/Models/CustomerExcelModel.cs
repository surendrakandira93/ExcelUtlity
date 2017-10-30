using System.Collections.Generic;

namespace ExcelUtility.Models
{
    public class CustomerExcelModel
    {
        [Cell(HeaderKey = "GSTIN", GSTIN = true, AllowEmpty = true)]
        public string GSTIN { get; set; }

        [Cell(HeaderKey = "Customer Name", AllowEmpty = false)]
        public string CustomerName { get; set; }

        [Cell(HeaderKey = "Address", AllowEmpty = false)]
        public string Address { get; set; }

        [Cell(HeaderKey = "City", AllowEmpty = false)]
        public string City { get; set; }

        [Cell(HeaderKey = "PIN", PIN = true, AllowEmpty = true)]
        public string PIN { get; set; }

        [Cell(HeaderKey = "Email Address", Email = true, AllowEmpty = true)]
        public string Email { get; set; }

        [Cell(HeaderKey = "Billing State", AllowEmpty = true)]
        public string BillingState { get; set; }

        public List<ExcelCell> CellInfo { get; set; } = new List<ExcelCell>();
    }
}