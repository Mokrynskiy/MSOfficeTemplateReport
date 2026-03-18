using ClosedXML.Excel;
using System;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace MSOfficeTemplateReport.Extensions
{
    internal static class ExcelExtensions
    {
        internal static XLCellValue GetExcelValue(this JsonNode node, string fieldName)
        {
            var field = node.AsObject().FirstOrDefault(x => x.Key == fieldName);

            XLCellValue value = new XLCellValue();

            switch (field.Value.GetValueKind())
            {
                case JsonValueKind.String:

                    if (DateTime.TryParse(field.Value.ToString(), out _))
                    {
                        value = DateTime.Parse(field.Value.ToString());
                    }
                    else
                    {
                        value = field.Value.ToString();
                    }

                    break;

                case JsonValueKind.Number:

                    value = long.TryParse(field.Value.ToString(), out _) ? long.Parse(field.Value.ToString()) : double.Parse(field.Value.ToString().Replace('.', ','));

                    break;

                case JsonValueKind.True:

                    value = bool.Parse(field.Value.ToString());

                    break;

                case JsonValueKind.False:

                    value = bool.Parse(field.Value.ToString());

                    break;

                default:

                    value = "";

                    break;
            }

            return value;
        }
    }    
}
