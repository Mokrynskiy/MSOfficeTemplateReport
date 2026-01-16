using ClosedXML.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace MSOfficeTemplateReport.ExcelReport
{
    public class ExcelTemplate : ITemplate
    {
        private string _path;
        private byte[] _byteArray;
        private Dictionary<string, object> _variables = new Dictionary<string, object>();
        private XLWorkbook _workbook;
        private readonly Regex _regex = new Regex("\\{\\{.*?\\}\\}");
        private readonly Regex _itemRegex = new Regex("Item");
        private MemoryStream _ms;
        public ExcelTemplate(string path)
        {
            _path = path;
        }

        public ExcelTemplate(byte[] byteArray)
        {
            _byteArray = byteArray;
        }

        public void AddVariable(string name, object data) => _variables.Add(name, data);

        public void Generate()
        {
            try
            {
                if(!_byteArray.Any())
                    _byteArray = File.ReadAllBytes(_path);
                _ms = new MemoryStream();
                _ms.Write(_byteArray, 0, _byteArray.Length);
                _workbook = new XLWorkbook(_ms);
                FillDocument();
            }
            catch (Exception ex)
            {
                _ms.Close();
                _workbook.Dispose();
                throw ex;
            }
        }

        public string SaveAs(string path)
        {
            try
            {
                _workbook.SaveAs(path);
                _workbook.Dispose();
                return path;
            }
            catch (Exception ex)
            {
                _ms.Close();
                _workbook.Dispose();
                throw ex;
            }
        }

        public byte[] ToByteArray()
        {
            _workbook.Save();
            var byteArray = _ms.ToArray();
            _ms.Close();
            _workbook.Dispose();
            return byteArray;
        }

        private void FillDocument()
        {
            try
            {
                var worsheets = _workbook.Worksheets;
                foreach (var worksheet in worsheets)
                {
                    var cells = worksheet.RangeUsed().CellsUsed().Where(x => _regex.IsMatch(x.GetText()));
                    foreach (var variable in _variables)
                    {
                        var type = variable.Value.GetType();
                        if (type.IsArray || type.IsGenericType)
                        {
                            var ranges = worksheet.Ranges(variable.Key);
                            foreach (var range in ranges)
                            {
                                int startRow = range.LastRow().RowNumber();
                                var data = (IList)variable.Value;
                                range.InsertRowsBelow(data.Count - 1);
                                foreach (var item in (IList)variable.Value)
                                {
                                    foreach (var cell in worksheet.Row(startRow).Cells())
                                    {
                                        worksheet.Cell(startRow + 1, cell.WorksheetColumn().ColumnNumber()).SetValue(cell.Value);
                                        if (!cell.Value.IsBlank && _regex.IsMatch(cell.Value.GetText()) && _itemRegex.IsMatch(cell.Value.GetText()))
                                        {
                                            var fieldName = cell.GetText().Replace("{", "").Replace("}", "").Split('.')[1];                                            
                                            var value = item.GetType().GetProperty(fieldName)?.GetValue(item);                                          
                                            cell.Value = value.ConvertToXLValue();
                                        }
                                    }
                                    startRow++;                                    
                                }
                                worksheet.Row(startRow).Delete();
                            }
                        }
                        else
                        {
                            Regex celreg = new Regex(variable.Key);
                            var currentCells = cells.Where(x => celreg.IsMatch(x.GetText()));
                            foreach (var cell in currentCells)
                            {
                                var fieldName = cell.GetText().Replace("{", "").Replace("}", "").Split('.')[1];
                                object obj = variable.Value;
                                string str = obj.GetType().GetProperty(fieldName)?.GetValue(obj)?.ToString();
                                cell.SetValue(str);
                            }
                        }
                    }
                }                
            }
            catch (Exception ex)
            {
                _ms.Close();
                _workbook.Dispose();
                throw ex;
            }
        }

    }
}
