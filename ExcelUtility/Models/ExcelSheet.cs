using System;

namespace ExcelUtility
{
    public class ExcelSheet<TEntity>
        where TEntity : class
    {
        public string Name { get; set; }

        public int StartIndex { get; set; }

        public Type Type { get; set; }

        public TEntity[] Records { get; set; }

        public bool Status { get; set; }

        public string Message { get; set; }
    }
}