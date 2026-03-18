using MSOfficeTemplateReport.Abstract;
using MSOfficeTemplateReport.ExcelReport;
using MSOfficeTemplateReport.Extensions;
using MSOfficeTemplateReport.Models;
using MSOfficeTemplateReport.WordReport;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace MSOfficeTemplateReport
{
    public static class Template
    {
        internal static readonly Regex TagRegex = new Regex("[\\{]{2}(?!(Item))+[a-zA-Z]+[.]+[a-zA-Z]+[\\}]{2}");

        internal static readonly Regex ItemRegex = new Regex ("[\\{]{2}Item+[.]+[a-zA-Z]+[\\}]{2}");

        internal static readonly Regex SummaryRegex = new Regex("<<Sum>>");

        internal static readonly Regex RowNumberRegex = new Regex("<<RowNumber>>");

        public static ITemplate Create(string filePath, Dictionary<string, object> variables = null)
        {
            var byteArray = File.ReadAllBytes(filePath);

            return Create(byteArray, variables);
        }

        public static ITemplate Create(byte[] byteArray, Dictionary<string, object> variables = null)
        {
            var format = byteArray.GetFormat();

            if(format == null)
                throw new Exception("Формат файла шаблона не поддерживается");

            Dictionary<string, JsonNode> jsonVariables = new Dictionary<string, JsonNode>();

            if(variables != null)
            {
                foreach (var variable in variables)
                {
                    AddJsonVariable(ref jsonVariables, variable.Key, variable.Value);
                }
            }

            switch (format)
            {                
                case FileFormates.Docx:
                    return new WordTemplate(byteArray, jsonVariables);
                    
                case FileFormates.Xlsx:
                    return new ExcelTemplate(byteArray, jsonVariables);
                    
                default:
                    throw new Exception("Формат файла шаблона не поддерживается");
            }            
        }
        
        internal static void AddJsonVariable(ref Dictionary<string, JsonNode> jsonVariables, string name, object variable)
        {
            var typeName = variable.GetType().Name;

            var element = typeName == "JsonElement" ? variable
                : typeName == "String" ? JsonElement.Parse(variable.ToString())
                : JsonElement.Parse(JsonSerializer.Serialize(variable, new JsonSerializerOptions { PropertyNameCaseInsensitive = true }));

            var jObject = JsonNode.Parse(element?.ToString());

            if (jObject?.GetValueKind() == JsonValueKind.Object)
            {
                foreach (var field in jObject.AsObject())
                {
                    if (field.Value?.GetValueKind() == JsonValueKind.Array)
                    {
                        if (jsonVariables.TryGetValue(field.Key, out _))
                            continue;

                        jsonVariables.Add(field.Key, JsonObject.Parse(field.Value?.ToString()));

                        continue;
                    }
                }

                jsonVariables.Add(name, jObject);
            }
            else if (jObject?.GetValueKind() == JsonValueKind.Array)
            {
                if (jsonVariables.TryGetValue(name, out _))
                    return;

                jsonVariables.Add(name, jObject);
            }
        }
    }
}
