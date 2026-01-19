using MSOfficeTemplateReport.Abstract;
using MSOfficeTemplateReport.ExcelReport;
using MSOfficeTemplateReport.Extensions;
using MSOfficeTemplateReport.WordReport;
using System;
using System.Collections.Generic;
using System.IO;

namespace MSOfficeTemplateReport
{
    public static class Template
    {
        public static ITemplate Create(byte[] byteArray)
        {
            var format = byteArray.GetFormat();

            switch (format)
            {
                case Models.FileFormat.NotDefined:
                    throw new Exception("Формат файла шаблона не поддерживается");
                case Models.FileFormat.Docx:
                    return new WordTemplate(byteArray);
                case Models.FileFormat.Xlsx:
                    return new ExcelTemplate(byteArray);
                default:
                    throw new Exception("Формат файла шаблона не поддерживается");                    
            }
        }

        public static ITemplate Create(string templatePath)
        {
            var format = Path.GetExtension(templatePath);

            switch (format)
            {                
                case ".docx":
                    return new WordTemplate(templatePath);
                case ".xlsx":
                    return new ExcelTemplate(templatePath);
                default:
                    throw new Exception("Формат файла шаблона не поддерживается");
            }
        }

        public static ITemplate Create(byte[] byteArray, Dictionary<string, object> variables)
        {
            var format = byteArray.GetFormat();

            ITemplate template;

            switch (format)
            {
                case Models.FileFormat.NotDefined:
                    throw new Exception("Формат файла шаблона не поддерживается");
                case Models.FileFormat.Docx:
                    template = new WordTemplate(byteArray);
                    break;
                case Models.FileFormat.Xlsx:
                    template =  new ExcelTemplate(byteArray);
                    break;
                default:
                    throw new Exception("Формат файла шаблона не поддерживается");
            }
            
            template.AddVariables(variables);

            return template;
        }

        public static ITemplate Create(string templatePath, Dictionary<string, object> variables)
        {
            var format = Path.GetExtension(templatePath);

            ITemplate template;

            switch (format)
            {
                case ".docx":
                    template =  new WordTemplate(templatePath);
                    break;
                case ".xlsx":
                    template =  new ExcelTemplate(templatePath);
                    break;
                default:
                    throw new Exception("Формат файла шаблона не поддерживается");
            }

            template.AddVariables(variables);

            return template;
        }
    }
}
