using System;
using System.Linq.Expressions;

namespace ExcelUtility
{
    public class CellEmptyIfAttribute : Attribute
    {
        public string Property { get; set; }

        public ExpressionType ExpressionType { get; set; }

        public string Value { get; set; }

        public Type DataType { get; set; }
    }
}