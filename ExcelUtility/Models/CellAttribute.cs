using System;

namespace ExcelUtility
{
    public class CellAttribute : Attribute
    {
        public bool AllowEmpty { get; set; }

        public string HeaderKey { get; set; }

        public bool Email { get; set; }

        public bool Date { get; set; }

        public bool FinancialDate { get; set; }

        public bool TransitionalDate { get; set; }

        public bool PIN { get; set; }

        public bool GSTIN { get; set; }

        public string ColorCode { get; set; }

        public bool Valid { get; set; } = true;

        public bool Unique { get; set; }

        public bool SameValue { get; set; }

        public string PropertyName { get; set; }

        public bool Numbers { get; set; }

        public string DefaultValue { get; set; }

        public bool Decimal { get; set; }

        public string Contains { get; set; }

        public string Spliter { get; set; }

        public bool HtmlVisible { get; set; } = true;

        public string MinValue { get; set; }

        public string MaxValue { get; set; } = "99999999999999.99";

        public int MinLength { get; set; }

        public int MaxLength { get; set; }
    }
}