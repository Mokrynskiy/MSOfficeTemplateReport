using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace MSOfficeTemplateReport.Extensions
{
    public static class ExcelExtensions
    {
        public static XLCellValue ConvertToXLValue(this object obj)
        {
            var type = obj.GetType().Name;

            XLCellValue value = new XLCellValue();

            switch (type)
            {
                case "Boolean":
                    value = (bool)obj;
                    break;
                case "Byte":
                    value = (byte)obj;
                    break;
                case "SByte":
                    value = (sbyte)obj;
                    break;
                case "Int16":
                    value = (short)obj;
                    break;
                case "UInt16":
                    value = (ushort)obj;
                    break;
                case "Int32":
                    value = (int)obj;
                    break;
                case "UInt32":
                    value = (uint)obj;
                    break;
                case "Int64":
                    value = (long)obj;
                    break;
                case "UInt64":
                    value = (ulong)obj;
                    break;
                case "Single":
                    value = (float)obj;
                    break;
                case "Double":
                    value = (double)obj;
                    break;
                case "Decimal":
                    value = (decimal)obj;
                    break;
                case "Char":
                    value = (char)obj;
                    break;
                case "String":
                    value = (string)obj;
                    break;
                case "DateTime":
                    value = (DateTime)obj;
                    break;
                case "TimeSpan":
                    value = (TimeSpan)obj;
                    break;
                default:
                    value = $"Не удалось преобразовать тип {obj.GetType().Name} к типу XLCellValue";
                    break;
            }
            return value;
        }

    }
}
