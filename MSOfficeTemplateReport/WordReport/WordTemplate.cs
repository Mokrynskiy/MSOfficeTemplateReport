using DocumentFormat.OpenXml.Packaging;
using MSOfficeTemplateReport.Abstract;
using MSOfficeTemplateReport.Extensions;
using MSOfficeTemplateReport.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;
using Path = System.IO.Path;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableProperties = DocumentFormat.OpenXml.Wordprocessing.TableProperties;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace MSOfficeTemplateReport.WordReport
{
    internal sealed class WordTemplate : ITemplate
    {
        private readonly byte[] byteArray;

        private Dictionary<string, JsonNode> variables;

        public WordTemplate(byte[] byteArray, Dictionary<string, JsonNode> variables)
        {
            this.byteArray = byteArray;

            this.variables = variables;
        }

        public void AddVariable(string name, object variable)
        {
            Template.AddJsonVariable(ref variables, name, variable);
        }

        public ReportResultModel Generate(string fileName = null)
        {
            try
            {
                using (var ms = new MemoryStream())
                {
                    ms.Write(this.byteArray, 0, this.byteArray.Length);

                    using (var document = WordprocessingDocument.Open(ms, true))
                    {
                        document.CleanRun();

                        FillText(document.GetAllText().Where(x => Template.TagRegex.IsMatch(x.Text)));

                        FillTables(document.GetAllTables());

                        string file = string.IsNullOrWhiteSpace(fileName) ?
                            $"{DateTime.Now.Ticks}{FileFormates.Docx}" :
                            $"{Path.GetFileNameWithoutExtension(fileName)}{FileFormates.Docx}";

                        document.Save();

                        var byteArray = ms.ToArray();

                        return new ReportResultModel(file, byteArray);
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void FillText(IEnumerable<Text> text)
        {
            foreach (var txt in text)
            {
                var tag = Template.TagRegex.Match(txt.Text?.ToString())?.Value;

                string[] strArray = tag?.SplitTag();

                string key = strArray?[0];

                string name = strArray?[1];

                if (variables.TryGetValue(key, out var value))
                {
                    var val = value.GetValue(name);

                    if (val != null)
                        txt.Text = txt.Text.Replace(tag, value.GetValue(name));
                }
            }
        }

        private void FillTables(IEnumerable<TableProperties> tablesProps)
        {
            foreach (TableProperties tableProp in tablesProps)
            {
                string tableCaption = tableProp.TableCaption?.Val?.ToString();

                var variable = variables.FirstOrDefault(x => x.Key == tableCaption);

                if (variable.Key != null && variable.Value != null && variable.Value.GetValueKind() == JsonValueKind.Array)
                {
                    var array = variable.Value.AsArray();

                    Table table = (Table)tableProp?.Parent;

                    var rows = table.Descendants<TableRow>();

                    TableRow tableRow = null;

                    foreach (TableRow r in rows)
                    {
                        var text = r.Descendants<Text>().Where(x => Template.ItemRegex.IsMatch(x?.Text)).ToList();

                        if (text.Count > 0)
                        {
                            tableRow = r;
                            break;
                        }
                    }

                    if (tableRow != null)
                    {
                        foreach (var item in array)
                        {
                            int rowNumber = 1;

                            var clonedRow = (TableRow)tableRow.CloneNode(true);

                            tableRow.InsertBeforeSelf(clonedRow);

                            foreach (var text in clonedRow.Descendants<Text>().Where(x => Template.ItemRegex.IsMatch(x.Text)))
                            {
                                var tag = Template.ItemRegex.Match(text.Text)?.Value;

                                var propName = text.Text.ToString()?.SplitTag()[1];

                                string str = item.GetValue(propName);

                                if (str != null)
                                    text.Text = text.Text.Replace(tag, str);
                            }
                        }

                        tableRow.Remove();
                    }
                }
            }
        }
    }
}
