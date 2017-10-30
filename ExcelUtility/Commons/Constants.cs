using System.Text.RegularExpressions;

namespace ExcelUtility
{
    public class Constants
    {
        public const string CellDuplicateColor = "#00c0ef";

        public const string CellEmptyColor = "#00a65a";

        public const string CellInvalidColor = "#dd4b39";

        public const string CellNotMatchedColor = "#f39c12";

        public const string DefaultColor = "#FFFFFF";

        public const string CellCodeExists = "#bbbd93";

        public static readonly Regex GstinRegex = new Regex(@"^([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[0-9]{1}[A-Z]{1}[0-9a-zA-Z]{1})|([0-9]{2}[0-9]{2}[A-Z]{3}[0-9]{5}[A-Z]{3})");

        public static readonly Regex EmailRegex = new Regex(@"^([a-zA-Z0-9_\.\-])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+$");

        public const string ExcelTemplateCustomer = @"Template/Customer.xlsx";
    }
}