using System;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace MSOfficeTemplateReport.Extensions
{
    internal static class StringExtensions
    {
        internal static string[] SplitTag(this string tag)
        {
            return tag.Replace("{", "").Replace("}", "").Split('.');
        }

        internal static string GetValue(this JsonNode node, string fieldName)
        {
            var field = node.AsObject().FirstOrDefault(x => x.Key == fieldName);

            if (field.Key == null)
                return null;

            if (field.Value == null)
                return "";

            if (field.Value.GetValueKind() == JsonValueKind.Number)
                return field.Value.ToString().Replace('.', ',');

            return field.Value?.ToString();
        }
    }
}
