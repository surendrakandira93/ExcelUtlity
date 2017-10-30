using System.Linq.Expressions;

namespace ExcelUtility
{
    public class ExpressionFilter
    {
        public string PropertyName { get; set; }

        public ExpressionType Operation { get; set; }

        public object Value { get; set; }
    }
}