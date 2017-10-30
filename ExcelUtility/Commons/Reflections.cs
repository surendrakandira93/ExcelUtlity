using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using ExcelUtility.Models;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using static ExcelUtility.Enums;

namespace ExcelUtility
{
    public static class Reflections
    {
        private static readonly Dictionary<Type, List<PropertyInfo>> PropertyDictionary = new Dictionary<Type, List<PropertyInfo>>();
        private static readonly Dictionary<PropertyInfo, List<Attribute>> CellAttributesDictionary = new Dictionary<PropertyInfo, List<Attribute>>();
        private static readonly Dictionary<PropertyInfo, List<Attribute>> CellEmptyAttributreDictionary = new Dictionary<PropertyInfo, List<Attribute>>();
        private static readonly Dictionary<string, IEnumerable<string>> FilterByDictionary = new Dictionary<string, IEnumerable<string>>();

        public static int ReturnPeriodId { get; set; }

        public static int YearId { get; set; }

        public static string ToCellKey(this PropertyInfo propertyInfo)
        {
            return GetPropertyInfo<CellAttribute>(propertyInfo, "HeaderKey");
        }

        public static bool RowEmpty(this ExcelRow excelRow, ExcelWorksheet ws)
        {
            bool empty = true;
            if (excelRow != null || ws != null)
            {
                for (int i = 1; i < ws.Dimension.Columns; i++)
                {
                    if (!string.IsNullOrEmpty(ws.Cells[1, i].Value as string))
                    {
                        empty = false;
                        break;
                    }
                }
            }
            return empty;
        }

        public static void ToCellValidate<TEntity>(this TEntity[] entities)
                    where TEntity : class
        {
            foreach (var entity in entities)
            {
                PropertyInfo[] properties = entity.GetType().GetProperties(true);
                List<ExcelCell> cells = new List<ExcelCell>(properties.Length);
                foreach (PropertyInfo propertyInfo in properties)
                {
                    var cellEmptyIf = (CellEmptyIfAttribute)propertyInfo.GetCellEmptyAttibute().FirstOrDefault();

                    foreach (Attribute item in propertyInfo.GetCellAttributes())
                    {
                        var value = propertyInfo.GetValue(entity) == null ? string.Empty : (string)propertyInfo.GetValue(entity);
                        var attribute = item as CellAttribute;
                        if (attribute != null)
                        {
                            if (string.IsNullOrEmpty(value) && !string.IsNullOrEmpty(attribute.DefaultValue))
                            {
                                value = attribute.DefaultValue;
                                propertyInfo.SetValue(entity, value);
                            }

                            if ((attribute.Numbers || attribute.Decimal) && !attribute.AllowEmpty && string.IsNullOrEmpty(value))
                            {
                                value = "0";
                                if (!string.IsNullOrEmpty(attribute.DefaultValue))
                                {
                                    value = attribute.DefaultValue;
                                }
                                propertyInfo.SetValue(entity, value);
                            }

                            if (cellEmptyIf != null)
                            {
                                CellEmptyIf(cellEmptyIf, properties, ref attribute, entity);
                            }

                            ExpressionFilter filterExpression = new ExpressionFilter { PropertyName = propertyInfo.Name, Value = value };

                            if (!attribute.AllowEmpty && string.IsNullOrEmpty(value))
                            {
                                cells.AddCellInfo(false, propertyInfo.Name, Constants.CellEmptyColor, "Cell should not be empty");
                            }
                            else if (attribute.Date && !string.IsNullOrEmpty(value))
                            {
                                DateTime dt;
                                DateTime.TryParse(value, out dt);
                                if (dt == DateTime.MinValue)
                                {
                                    cells.AddCellInfo(false, propertyInfo.Name, Constants.CellInvalidColor, "Date is not valid");
                                }
                                else if (attribute.Date && attribute.FinancialDate && dt != DateTime.MinValue)
                                {
                                    if (!CheckFinancialDate(ReturnPeriodId, YearId, dt))
                                    {
                                        cells.AddCellInfo(false, propertyInfo.Name, Constants.CellInvalidColor, "Date Should be financial Date.");
                                    }
                                }
                                else if (attribute.Date && attribute.TransitionalDate && dt != DateTime.MinValue)
                                {
                                    if (!CheckTransitionalDate(dt))
                                    {
                                        cells.AddCellInfo(false, propertyInfo.Name, Constants.CellInvalidColor, "Date Should be before 1st July 2017.");
                                    }
                                }
                            }
                            else if ((attribute.Numbers || attribute.Decimal) && !string.IsNullOrEmpty(value))
                            {
                                int convertToInteger;
                                decimal convertToDecimal;

                                decimal minValue = Convert.ToDecimal(attribute.MinValue);
                                decimal maxValue = Convert.ToDecimal(attribute.MaxValue);

                                if (attribute.Numbers && !int.TryParse(value, out convertToInteger))
                                {
                                    cells.AddCellInfo(false, propertyInfo.Name, Constants.CellInvalidColor, "value is not valid");
                                }

                                if (attribute.Decimal && !decimal.TryParse(value, out convertToDecimal))
                                {
                                    cells.AddCellInfo(false, propertyInfo.Name, Constants.CellInvalidColor, "value is not valid");
                                }

                                if (attribute.Decimal && decimal.TryParse(value, out convertToDecimal) && (convertToDecimal < minValue || convertToDecimal > maxValue))
                                {
                                    cells.AddCellInfo(false, propertyInfo.Name, Constants.CellInvalidColor, "value is not valid");
                                }

                                if (attribute.Numbers && int.TryParse(value, out convertToInteger) && (convertToInteger < minValue || convertToInteger > maxValue))
                                {
                                    cells.AddCellInfo(false, propertyInfo.Name, Constants.CellInvalidColor, "value is not valid");
                                }
                            }
                            else if (attribute.GSTIN && !string.IsNullOrEmpty(value) && (!Constants.GstinRegex.IsMatch(value) || value.Length > 15 || value.Length < 15))
                            {
                                cells.AddCellInfo(false, propertyInfo.Name, Constants.CellInvalidColor, "Value is not valid GSTIN");
                            }
                            else if (attribute.Email && !string.IsNullOrEmpty(value) && !Constants.EmailRegex.IsMatch(value))
                            {
                                cells.AddCellInfo(false, propertyInfo.Name, Constants.CellInvalidColor, "value is not valid email address");
                            }
                            else if (attribute.PIN && !string.IsNullOrEmpty(value) && (value.Length > 6 || value.Length < 6))
                            {
                                cells.AddCellInfo(false, propertyInfo.Name, Constants.CellInvalidColor, "PIN Number should be 6 Digit");
                            }
                            else if (attribute.Unique && !string.IsNullOrEmpty(value))
                            {
                                filterExpression.Operation = ExpressionType.Equal;
                                if (entities.Count(GetExpression<TEntity>(new List<ExpressionFilter> { filterExpression }).Compile()) > 1)
                                {
                                    cells.AddCellInfo(false, propertyInfo.Name, Constants.CellDuplicateColor, "Duplicate Value");
                                }
                            }
                            else if (attribute.SameValue && !string.IsNullOrEmpty(value))
                            {
                                filterExpression.Operation = ExpressionType.Equal;
                                if (entities.Count(GetExpression<TEntity>(new List<ExpressionFilter> { filterExpression }).Compile()) == 0 && entities.Length > 0)
                                {
                                    cells.AddCellInfo(false, propertyInfo.Name, Constants.CellInvalidColor, "Value Should Be Same in Column");
                                }
                            }
                            else if (!string.IsNullOrEmpty(attribute.Contains) && !string.IsNullOrEmpty(value))
                            {
                                string spliter = !string.IsNullOrEmpty(attribute.Spliter) ? attribute.Spliter : ",";

                                var splitedValues = attribute.Contains.Split(new[] { spliter }, StringSplitOptions.RemoveEmptyEntries);
                                if (splitedValues.Count(x => x.ToLower().Contains(value.ToLower())) == 0)
                                {
                                    cells.AddCellInfo(false, propertyInfo.Name, Constants.CellNotMatchedColor, "value is not valid");
                                }
                            }
                            if (attribute.MinLength > 0 || attribute.MaxLength > 0)
                            {
                                if (!string.IsNullOrEmpty(value) && value.Length < attribute.MinLength)
                                {
                                    cells.AddCellInfo(false, propertyInfo.Name, Constants.CellInvalidColor, $"Value should be minimum {attribute.MinLength} Digit");
                                }

                                if (!string.IsNullOrEmpty(value) && value.Length > attribute.MaxLength)
                                {
                                    cells.AddCellInfo(false, propertyInfo.Name, Constants.CellInvalidColor, $"Value should not be max {attribute.MinLength} Digit");
                                }
                            }

                            IEnumerable<string> filterIn;
                            if (FilterByDictionary.TryGetValue(propertyInfo.Name, out filterIn))
                            {
                                if (filterIn.Where(x => !string.IsNullOrEmpty(x)).Count(x => x.ToLower().Contains(value.ToLower())) == 0)
                                {
                                    cells.AddCellInfo(false, propertyInfo.Name, Constants.CellInvalidColor, "value is not valid");
                                }
                            }
                        }
                    }
                }

                if (cells.Count > 0)
                {
                    var cellInfo = properties.FirstOrDefault(p => p.PropertyType == typeof(List<ExcelCell>));
                    cellInfo?.SetValue(entity, cells);
                }
            }
        }

        public static void AddCellError<TEntity>(this IEnumerable<TEntity> entities, bool condition, string property, string message, CellType cellType = CellType.Empty)
            where TEntity : ExcelCell
        {
            if (entities == null)
            {
                entities = new List<TEntity>();
            }

            if (condition)
            {
                var cell = new ExcelCell { PropertyName = property, Message = message };
                switch (cellType)
                {
                    case CellType.Empty: cell.ColorCode = Constants.CellEmptyColor; break;
                    case CellType.Invalid: cell.ColorCode = Constants.CellInvalidColor; break;
                    case CellType.Exist: cell.ColorCode = Constants.CellCodeExists; break;
                    case CellType.Duplicate: cell.ColorCode = Constants.CellDuplicateColor; break;
                    case CellType.NotMatched: cell.ColorCode = Constants.CellNotMatchedColor; break;
                }

                ((List<ExcelCell>)entities).Add(cell);
            }
        }

        public static string ToHtmlTable<TEntity>(this IEnumerable<TEntity> entities, bool isErrors)
        {
            StringBuilder tblBody = new StringBuilder();
            StringBuilder tblHead = new StringBuilder();
            bool isHead = false;
            int td = 0;
            if (isErrors)
            {
                foreach (TEntity entity in entities)
                {
                    List<ExcelCell> infoCells = null;
                    var cellinfoType = entity.GetType().GetProperties(true).FirstOrDefault(c => c.PropertyType == typeof(List<ExcelCell>));
                    if (cellinfoType != null)
                    {
                        infoCells = (List<ExcelCell>)cellinfoType.GetValue(entity) ?? new List<ExcelCell>();
                    }

                    if (!isHead)
                    {
                        tblHead.Append("<tr>");
                        foreach (PropertyInfo propertyInfo in entity.GetType().GetProperties(true))
                        {
                            var key = propertyInfo.ToCellKey();
                            if (!string.IsNullOrEmpty(key))
                            {
                                tblHead.Append($"<th>{propertyInfo.ToCellKey()}</th>");
                            }
                        }

                        tblHead.Append("</tr>");
                    }

                    tblBody.Append("<tr>");
                    foreach (var propertyInfo in entity.GetType().GetProperties(true))
                    {
                        if (propertyInfo.PropertyType != typeof(List<ExcelCell>) && propertyInfo.GetCellAttributes().Length > 0)
                        {
                            var infoCell = infoCells.FirstOrDefault(c => c.PropertyName == propertyInfo.Name);
                            var colorCode = infoCell != null ? infoCell.ColorCode : "#FFFFFF";
                            var message = infoCell != null ? infoCell.Message : string.Empty;

                            var key = propertyInfo.ToCellKey();
                            if (!string.IsNullOrEmpty(key))
                            {
                                string value = propertyInfo.GetValue(entity) == null ? string.Empty : (string)propertyInfo.GetValue(entity);
                                value = string.IsNullOrEmpty(value) ? "&nbsp;" : value;

                                tblBody.Append($"<td style='background:{colorCode}'><a style='display:block;' data-toggle='tooltip' data-original-title='{message}'>{value}</a></td>");
                                td++;
                            }
                        }
                    }

                    tblBody.Append($"</tr>");
                    isHead = true;
                }
            }
            if (td == 0)
            {
                return string.Empty;
            }

            return $"<table id='{Guid.NewGuid()}' class='table table-standard table-bordered table-striped table-striped-min-width dataTable no-footer'><thead>{tblHead}</thead><tbody>{tblBody}</tbody></table>";
        }

        public static IEnumerable<TEntity> Batch<TEntity>(this IEnumerable<TEntity> entities, int skip, int take)
        {
            return entities.Skip(skip).Take(take);
        }

        public static object AddBatchItem(this Dictionary<string, object> dic, string key, object data)
        {
            object existData;
            if (dic.TryGetValue(key, out existData))
            {
                dic[key] = data;
            }
            else
            {
                dic.Add(key, data);
            }

            return dic[key];
        }

        public static object GetBatchItem<TEntity>(this Dictionary<string, object> dic, string key)
            where TEntity : class
        {
            object existData;
            if (dic.TryGetValue(key, out existData))
            {
                return dic[key];
            }

            dic.Add(key, (TEntity)Activator.CreateInstance(typeof(TEntity)));

            return dic[key];
        }

        public static TEntity[] ToExcelSheet<TEntity>(this ExcelWorksheet worksheet, int startRow)
            where TEntity : class
        {
            return WorkSheetToArray<TEntity>(worksheet, startRow);
        }

        public static TEntity[] ToExcelSheet<TEntity>(this ExcelWorksheet worksheet, Dictionary<string, IEnumerable<string>> filterKeyValues, int startRow = 2)
            where TEntity : class
        {
            foreach (var item in filterKeyValues)
            {
                IEnumerable<string> values;
                if (FilterByDictionary.TryGetValue(item.Key, out values) == false)
                {
                    FilterByDictionary.Add(item.Key, item.Value);
                }
            }

            return WorkSheetToArray<TEntity>(worksheet, startRow);
        }

        public static ExcelWorksheet ToEntityList<TEntity>(this ExcelWorksheet worksheet, List<TEntity> entityList, int startRow)
            where TEntity : class
        {
            return ListToWorkSheet<TEntity>(worksheet, entityList, startRow);
        }

        private static void CellEmptyIf<TAttribute, TEntity>(TAttribute cellEmptyIfAttribute, PropertyInfo[] properties, ref CellAttribute cellAttribute, TEntity entity)
            where TAttribute : CellEmptyIfAttribute
            where TEntity : class
        {
            if (cellEmptyIfAttribute != null && !string.IsNullOrEmpty(cellEmptyIfAttribute.Property))
            {
                var validateTo = properties.FirstOrDefault(x => x.Name == cellEmptyIfAttribute.Property);
                if (validateTo != null)
                {
                    var validateToValue = validateTo.GetValue(entity) == null ? string.Empty : (string)validateTo.GetValue(entity);
                    switch (cellEmptyIfAttribute.ExpressionType)
                    {
                        case ExpressionType.NotEqual:
                            cellAttribute.AllowEmpty = validateToValue.ToLower() != cellEmptyIfAttribute.Value.ToLower();
                            break;

                        case ExpressionType.Equal:
                            cellAttribute.AllowEmpty = validateToValue.ToLower() == cellEmptyIfAttribute.Value.ToLower();
                            break;
                    }
                }
            }
        }

        private static TEntity[] WorkSheetToArray<TEntity>(ExcelWorksheet worksheet, int startRow)
                    where TEntity : class
        {
            TEntity entity = (TEntity)Activator.CreateInstance(typeof(TEntity));
            PropertyInfo[] propertyInfo = entity.GetType().GetProperties(true);

            Properties[] properties = propertyInfo.GetProperties();
            Dictionary<string, int> columns = new Dictionary<string, int>();
            var jArray = new JArray();
            try
            {
                long maxColumns = worksheet.Dimension.Columns;
                long maxRows = worksheet.Dimension.Rows;
                foreach (Properties property in properties)
                {
                    for (int index = 1; index <= maxColumns; index++)
                    {
                        var cellKey = property.CellKey;
                        var noOfMerge = 0;

                        validateCell:

                        var rowindex = (startRow - 1 > 0 ? startRow - 1 : 1) - noOfMerge;
                        ExcelRange cell = worksheet.Cells[rowindex == 0 ? 1 : rowindex, index];
                        if (cell.Merge && noOfMerge == 0)
                        {
                            noOfMerge = 1;
                            goto validateCell;
                        }

                        var cellValue = cell.Value as string;
                        if ((!string.IsNullOrEmpty(cellValue) ? cellValue : string.Empty).TrimEnd('*').Trim().ToLower() == cellKey.ToLower() && !string.IsNullOrEmpty(cellKey.TrimEnd('*').Trim()))
                        {
                            columns.Add(cellKey, index);
                        }
                    }
                }

                for (int row = startRow; row <= maxRows; row++)
                {
                    int counter = 0;
                    int blank = 1;
                    JObject jObject = new JObject();
                    foreach (var property in properties)
                    {
                        var dicKey = property.Name;
                        var dicValue = string.Empty;
                        if (property.Type != typeof(List<ExcelCell>))
                        {
                            for (int col = 1; col <= columns.Count; col++)
                            {
                                var column = columns.FirstOrDefault(c => c.Key == property.CellKey);
                                if (column.Key == property.CellKey && column.Value > 0)
                                {
                                    var excelCell = worksheet.Cells[row, column.Value];
                                    var cellAttribute = property.Attribute as CellAttribute;
                                    if (cellAttribute != null && cellAttribute.Date)
                                    {
                                        if (excelCell.Style.Numberformat.Format.ToLower().Contains("yy"))
                                        {
                                            excelCell.Style.Numberformat.Format = "dd/MM/yyyy";
                                            dicValue = Convert.ToString(excelCell.Text);
                                            break;
                                        }

                                        try
                                        {
                                            dicValue = DateTimeExtensions.FromOADate(Convert.ToDouble(worksheet.Cells[row, column.Value].Value)).ToString("dd/MM/yyyy");
                                            break;
                                        }
                                        catch (FormatException ex)
                                        {
                                            Console.WriteLine(ex.Message);
                                        }
                                    }

                                    if (cellAttribute != null && cellAttribute.Decimal)
                                    {
                                        ////excelCell.Style.Numberformat.Format = "0.00";
                                        dicValue = Convert.ToString(excelCell.Text).Replace("%", string.Empty);
                                        break;
                                    }

                                    try
                                    {
                                        dicValue = Convert.ToString(worksheet.Cells[row, column.Value].Text);
                                        break;
                                    }
                                    catch (NullReferenceException)
                                    {
                                        dicValue = Convert.ToString(worksheet.Cells[row, column.Value].Value);
                                        break;
                                    }
                                }
                            }

                            if (string.IsNullOrEmpty(dicValue.Trim()))
                            {
                                blank++;
                            }

                            jObject.Add(dicKey, dicValue.Trim());
                        }
                        else
                        {
                            jObject.Add(dicKey, new JArray());
                        }

                        if (string.IsNullOrEmpty(dicValue))
                        {
                            counter++;
                        }
                    }

                    if (counter > 0 && counter == properties.Length)
                    {
                        //// skip statement
                    }
                    else
                    {
                        if (blank > 1 && blank > (columns.Count - 2))
                        {
                            Console.WriteLine("s");
                        }
                        else
                        {
                            jArray.Add(jObject);
                        }
                    }
                }

                TEntity[] entities = Newtonsoft.Json.JsonConvert.DeserializeObject<TEntity[]>(jArray.ToString());
                entities.ToCellValidate();
                return entities;
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }

        private static ExcelWorksheet ListToWorkSheet<TEntity>(ExcelWorksheet worksheet, List<TEntity> entityList, int startRow)
                   where TEntity : class
        {
            Type entity = typeof(TEntity);
            PropertyInfo[] propertyInfo = entity.GetProperties();

            Properties[] properties = propertyInfo.GetProperties();
            Dictionary<string, int> columns = new Dictionary<string, int>();
            int rowNo = startRow;

            long maxColumns = worksheet.Dimension.Columns;
            foreach (Properties property in properties)
            {
                for (int index = 1; index <= maxColumns; index++)
                {
                    var cellKey = property.CellKey;
                    var noOfMerge = 0;

                    validateCell:

                    var rowindex = (startRow - 1 > 0 ? startRow - 1 : 1) - noOfMerge;
                    ExcelRange cell = worksheet.Cells[rowindex == 0 ? 1 : rowindex, index];
                    if (cell.Merge && noOfMerge == 0)
                    {
                        noOfMerge = 1;
                        goto validateCell;
                    }

                    var cellValue = cell.Value as string;
                    if ((!string.IsNullOrEmpty(cellValue) ? cellValue : string.Empty).TrimEnd('*').Trim().ToLower() == cellKey.ToLower() && !string.IsNullOrEmpty(cellKey.TrimEnd('*').Trim()))
                    {
                        columns.Add(property.Name, index);
                    }
                }
            }

            foreach (TEntity entityItem in entityList)
            {
                try
                {
                    foreach (var item in columns)
                    {
                        var prop = propertyInfo.FirstOrDefault(x => x.Name.ToLower() == item.Key.ToLower());
                        object value = prop.GetValue(entityItem);
                        worksheet.Cells[rowNo, item.Value].Value = value;
                        worksheet.Cells[rowNo, item.Value].Style.Numberformat.Format = GetValueFromEntity(prop, value);
                    }
                }
                catch (System.Exception ex)
                {
                    throw ex;
                }

                rowNo++;
            }

            return worksheet;
        }

        private static string GetValueFromEntity(PropertyInfo propertyInfo, object value)
        {
            string valueType = string.Empty;

            var custom = propertyInfo.GetCustomAttribute<CellAttribute>();
            if (custom.Date)
            {
                valueType = "dd/MM/yyyy";
            }
            else if (custom.Decimal)
            {
                if (value.ToString().Contains("%"))
                {
                    valueType = "0%";
                }
                else
                {
                    valueType = "0.00";
                }
            }
            else if (custom.Numbers)
            {
                valueType = "0";
            }
            else
            {
                valueType = "0.00";
            }

            return valueType;
        }

        private static PropertyInfo[] GetProperties(this Type type, bool cached)
        {
            List<PropertyInfo> properties;
            if (PropertyDictionary.TryGetValue(type, out properties) == false)
            {
                if (type != null)
                {
                    properties = type.GetProperties().Where(x => x.PropertyType == typeof(string) || x.PropertyType == typeof(List<ExcelCell>)).ToList();
                    if (cached)
                    {
                        PropertyDictionary.Add(type, properties);
                    }
                }
            }

            return properties?.ToArray();
        }

        private static Properties[] GetProperties(this PropertyInfo[] propertyInfo)
        {
            if (propertyInfo.Length == 0)
            {
                return new List<Properties>().ToArray();
            }

            Properties[] properties = new Properties[propertyInfo.Length];

            for (int index = 0; index < propertyInfo.Length; index++)
            {
                properties[index] = new Properties
                {
                    Name = propertyInfo[index].Name,
                    CellKey = propertyInfo[index].ToCellKey(),
                    Type = propertyInfo[index].PropertyType,
                    Attribute = propertyInfo[index].GetCellAttributes().FirstOrDefault()
                };
            }

            return properties;
        }

        private static string GetPropertyInfo<TAttributre>(PropertyInfo propertyInfo, string keyName)
            where TAttributre : Attribute
        {
            string value = string.Empty;
            foreach (Attribute attribute in propertyInfo.GetCustomAttributes(true).Where(x => x.GetType() == typeof(TAttributre)))
            {
                var property = attribute.GetType().GetRuntimeProperties().FirstOrDefault(p => string.Equals(p.Name, keyName, StringComparison.OrdinalIgnoreCase));
                if (property != null)
                {
                    value = property.GetValue(attribute) as string;
                }
            }

            return value;
        }

        private static void AddCellInfo<TEntity>(this List<TEntity> entities, bool isValid, string propertyname, string colorCode, string message)
        {
            if (typeof(TEntity) == typeof(ExcelCell))
            {
                bool isvalid = false;

                TEntity entity = (TEntity)Activator.CreateInstance(typeof(TEntity));

                foreach (PropertyInfo propertyInfo in entity.GetType().GetProperties(true))
                {
                    switch (propertyInfo.Name)
                    {
                        case "CellValid":
                            propertyInfo.SetValue(entity, isValid); isvalid = true;
                            break;

                        case "ColorCode":
                            propertyInfo.SetValue(entity, colorCode); isvalid = true;
                            break;

                        case "PropertyName":
                            propertyInfo.SetValue(entity, propertyname); isvalid = true;
                            break;

                        case "Message":
                            propertyInfo.SetValue(entity, message); isvalid = true;
                            break;
                    }
                }

                if (isvalid)
                {
                    entities.Add(entity);
                }
            }
        }

        private static Attribute[] GetCellAttributes(this PropertyInfo propertyInfo)
        {
            List<Attribute> properties;
            if (CellAttributesDictionary.TryGetValue(propertyInfo, out properties) == false)
            {
                properties = propertyInfo.GetCustomAttributes(true).Where(x => x.GetType() == typeof(CellAttribute)).ToList();
                CellAttributesDictionary.Add(propertyInfo, properties);
            }

            return properties.ToArray();
        }

        private static Attribute[] GetCellEmptyAttibute(this PropertyInfo propertyInfo)
        {
            List<Attribute> properties;
            if (CellEmptyAttributreDictionary.TryGetValue(propertyInfo, out properties) == false)
            {
                properties = propertyInfo.GetCustomAttributes(true).Where(x => x.GetType() == typeof(CellEmptyIfAttribute)).ToList();
                CellEmptyAttributreDictionary.Add(propertyInfo, properties);
            }

            return properties.ToArray();
        }

        private static Expression<Func<TEntity, bool>> GetExpression<TEntity>(IList<ExpressionFilter> filters)
        {
            if (filters.Count == 0)
            {
                return null;
            }

            ParameterExpression param = Expression.Parameter(typeof(TEntity), "t");
            Expression exp = null;

            if (filters.Count == 1)
            {
                exp = GetExpression<TEntity>(param, filters[0]);
            }
            else if (filters.Count == 2)
            {
                exp = GetExpression<TEntity>(param, filters[0], filters[1]);
            }
            else
            {
                while (filters.Count > 0)
                {
                    var f1 = filters[0];
                    var f2 = filters[1];

                    if (exp == null)
                    {
                        exp = GetExpression<TEntity>(param, filters[0], filters[1]);
                    }
                    else
                    {
                        exp = Expression.AndAlso(exp, GetExpression<TEntity>(param, filters[0], filters[1]));
                    }

                    filters.Remove(f1);
                    filters.Remove(f2);

                    if (filters.Count == 1)
                    {
                        exp = Expression.AndAlso(exp, GetExpression<TEntity>(param, filters[0]));
                        filters.RemoveAt(0);
                    }
                }
            }

            return Expression.Lambda<Func<TEntity, bool>>(exp, param);
        }

        private static Expression GetExpression<TEntity>(ParameterExpression param, ExpressionFilter filter)
        {
            MemberExpression member = Expression.Property(param, filter.PropertyName);
            ConstantExpression constant = Expression.Constant(filter.Value);

            switch (filter.Operation)
            {
                case ExpressionType.Equal:
                    return Expression.Equal(member, constant);

                case ExpressionType.GreaterThan:
                    return Expression.GreaterThan(member, constant);

                case ExpressionType.GreaterThanOrEqual:
                    return Expression.GreaterThanOrEqual(member, constant);

                case ExpressionType.LessThan:
                    return Expression.LessThan(member, constant);

                case ExpressionType.LessThanOrEqual:
                    return Expression.LessThanOrEqual(member, constant);
            }

            return null;
        }

        private static BinaryExpression GetExpression<TEntity>(ParameterExpression param, ExpressionFilter filter1, ExpressionFilter filter2)
        {
            Expression bin1 = GetExpression<TEntity>(param, filter1);
            Expression bin2 = GetExpression<TEntity>(param, filter2);

            return Expression.AndAlso(bin1, bin2);
        }

        private static bool CheckFinancialDate(int returnPeriodId, int yearid, DateTime datetime)
        {
            Enums.PeriodType periodTypeText = (Enums.PeriodType)returnPeriodId;
            Enums.FinancialYearText finacialYearTexts = (Enums.FinancialYearText)yearid;
            var rest = finacialYearTexts.ToString().Substring(2, finacialYearTexts.ToString().Length - 4);
            int startYear = Convert.ToInt32(rest);
            int endYear = startYear;
            int startMonth = 0;
            int endMonth = 0;
            DateTime startDate = new DateTime(startYear, 4, 1);

            switch (periodTypeText)
            {
                case Enums.PeriodType.Q1:
                    startMonth = 4;
                    endMonth = 6;
                    break;

                case Enums.PeriodType.Q2:
                    startMonth = 7;
                    endMonth = 9;
                    break;

                case Enums.PeriodType.Q3:
                    startMonth = 10;
                    endMonth = 12;
                    break;

                case Enums.PeriodType.Q4:
                    startMonth = 1;
                    endMonth = 3;
                    endYear = startYear + 1;
                    break;

                case Enums.PeriodType.April:
                    startMonth = 4;
                    endMonth = 4;
                    break;

                case Enums.PeriodType.May:
                    startMonth = 5;
                    endMonth = 5;
                    break;

                case Enums.PeriodType.June:
                    startMonth = 6;
                    endMonth = 6;
                    break;

                case Enums.PeriodType.July:
                    startMonth = 7;
                    endMonth = 7;
                    break;

                case Enums.PeriodType.August:
                    startMonth = 8;
                    endMonth = 8;
                    break;

                case Enums.PeriodType.September:
                    startMonth = 9;
                    endMonth = 9;
                    break;

                case Enums.PeriodType.October:
                    startMonth = 10;
                    endMonth = 10;
                    break;

                case Enums.PeriodType.November:
                    startMonth = 11;
                    endMonth = 11;
                    break;

                case Enums.PeriodType.December:
                    startMonth = 12;
                    endMonth = 12;
                    break;

                case Enums.PeriodType.January:
                    startMonth = 1;
                    endMonth = 1;
                    endYear = startYear + 1;
                    break;

                case Enums.PeriodType.February:
                    startMonth = 2;
                    endMonth = 2;
                    endYear = startYear + 1;
                    break;

                case Enums.PeriodType.March:
                    startMonth = 3;
                    endMonth = 3;
                    endYear = startYear + 1;
                    break;
            }

            if (startMonth >= 4)
            {
                startDate = new DateTime(startYear, startMonth, 1);
            }
            else
            {
                startDate = new DateTime(endYear, startMonth, 1);
            }

            var lastDate = new DateTime(endYear, endMonth, 1);
            DateTime endDate = lastDate.AddMonths(1).AddDays(-1);

            return datetime.Date <= endDate.Date && datetime.Date >= startDate.Date;
        }

        private static bool CheckTransitionalDate(DateTime datetime)
        {
            return datetime.Date <= new DateTime(2017, 6, 30).Date;
        }
    }
}