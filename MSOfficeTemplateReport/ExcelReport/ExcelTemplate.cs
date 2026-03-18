using ClosedXML.Excel;
using MSOfficeTemplateReport.Abstract;
using MSOfficeTemplateReport.Extensions;
using MSOfficeTemplateReport.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace MSOfficeTemplateReport.ExcelReport
{
    internal class ExcelTemplate : ITemplate
    {
        private Dictionary<string, JsonNode> variables;

        private readonly byte[] byteArray;

        public ExcelTemplate(byte[] byteArray, Dictionary<string, JsonNode> variables)
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
                string file = string.IsNullOrWhiteSpace(fileName) ?
                    $"{DateTime.Now.Ticks}{FileFormates.Xlsx}" :
                    $"{Path.GetFileNameWithoutExtension(fileName)}{FileFormates.Xlsx}";

                using (var ms = new MemoryStream())
                {
                    ms.Write(this.byteArray, 0, this.byteArray.Length);

                    using (var workbook = new XLWorkbook(ms))
                    {
                        FillDocument(workbook);

                        workbook.Save();

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

        private void FillDocument(XLWorkbook workbook)
        {
            try
            {
                var worsheets = workbook?.Worksheets;

                foreach (var worksheet in worsheets)
                {
                    var cells = worksheet?.RangeUsed()?.CellsUsed().Where(x => x.Value.IsText && Template.TagRegex.IsMatch(x.GetText()));

                    foreach (var variable in variables)
                    {
                        if (variable.Value.GetValueKind() == JsonValueKind.Array)
                        {
                            var ranges = worksheet?.Ranges(variable.Key);

                            if (ranges != null)
                            {
                                foreach (var range in ranges)
                                {
                                    var data = variable.Value.AsArray();

                                    range.FirstRow().InsertRowsBelow(data.Count - 1);

                                    int rowNumber = 1;

                                    foreach (var row in range.Rows())
                                    {
                                        if (range.LastRow() == row)
                                        {
                                            var summaryFields = row.Cells().Where(x => x.Value.IsText && Template.SummaryRegex.IsMatch(x.GetText())).ToList();

                                            if (summaryFields.Count > 0)
                                            {
                                                foreach (var field in summaryFields)
                                                {
                                                    var formula = "SUM(" + worksheet.Cell(range.FirstRow()
                                                        .RowNumber(), field.Address.ColumnNumber).Address.ToString() +
                                                        ":" + worksheet.Cell(field.Address.RowNumber - 1, field.Address.ColumnNumber)
                                                        .Address.ToString() + ")";

                                                    field.FormulaA1 = formula;
                                                }
                                            }

                                            continue;
                                        }

                                        if (rowNumber < data.Count)
                                            row.CopyTo(range.Row(rowNumber + 1));

                                        var values = data[rowNumber - 1];

                                        var tableCells = row.Cells().Where(x => x.Value.IsText && (Template.ItemRegex.IsMatch(x.GetText()) || Template.RowNumberRegex.IsMatch(x.GetText())));

                                        foreach (var cell in tableCells)
                                        {
                                            if (Template.RowNumberRegex.IsMatch(cell.GetText()))
                                            {
                                                cell.Value = rowNumber;

                                                continue;
                                            }

                                            var fieldName = cell.GetText().SplitTag()[1];

                                            cell.Value = values.GetExcelValue(fieldName);
                                        }

                                        rowNumber++;
                                    }
                                }
                            }
                        }
                        else
                        {
                            Regex celreg = new Regex("{{" + variable.Key);

                            var currentCells = cells?.Where(x => x.Value.IsText && celreg.IsMatch(x.Value.ToString()));

                            foreach (var cell in currentCells)
                            {
                                var matches = Template.TagRegex.Matches(cell.GetText());

                                var text = cell.GetText();

                                if (matches.Count == 1 && text == matches[0].Value)
                                {
                                    var fieldName = matches[0].Value.SplitTag()[1];

                                    cell.Value = variable.Value.GetExcelValue(fieldName);
                                }
                                else
                                {
                                    foreach (var match in matches)
                                    {
                                        var fieldName = match?.ToString()?.SplitTag()[1];

                                        var value = variable.Value.GetValue(fieldName);

                                        text = text.Replace(match.ToString(), value);
                                    }

                                    cell.SetValue(text);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }
    }
}
