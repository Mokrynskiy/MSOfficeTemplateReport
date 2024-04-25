using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OpenXmlPowerTools;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;

namespace WordTemplateReport.WordReport
{
    public sealed class WordTemplate
    {
        private string _path;
        private Dictionary<string, object> _variables;
        private MemoryStream _ms;
        private WordprocessingDocument _document;
        private readonly Regex _regex = new Regex("\\{\\{.*?\\}\\}");
        private readonly Regex _tagRegex = new Regex("[\\{]{2}[a-zA-Z.]+[\\}]{2}");
        private readonly Regex _itemRegex = new Regex("Item");

        public WordTemplate(string path)
        {
            _path = path;
            _variables = new Dictionary<string, object>();
        }

        public void AddVariable(string name, object data) => this._variables.Add(name, data);

        public void Generate()
        {
            try
            {
                byte[] buffer = File.ReadAllBytes(this._path);
                MemoryStream memoryStream = new MemoryStream();
                memoryStream.Write(buffer, 0, buffer.Length);
                _document = WordprocessingDocument.Open((Stream)memoryStream, true);
                CleanDoc();
                FillHeader();
                FillFooter();
                FillBody();
            }
            catch (Exception ex)
            {
                _document.Dispose();
                throw ex;
            }
        }
        private void FillBody()
        {
            var body = _document.MainDocumentPart.Document.Body;
            if (body == null) return;
            var text = body.Descendants<Text>();
            var tables = body.Descendants<TableProperties>();
            if (text != null) FillText(text);
            if (tables != null) FillTables(tables);
        }
        private void FillFooter()
        {
            var footer = _document.MainDocumentPart.FooterParts.FirstOrDefault();
            if (footer == null) return;
            var text = footer.Footer.Descendants<Text>();
            var tables = footer.Footer.Descendants<TableProperties>();
            if (text != null) FillText(text);
            if (tables != null) FillTables(tables);
        }
        private void FillHeader()
        {
            var header = _document.MainDocumentPart.HeaderParts.FirstOrDefault();
            if (header == null) return;
            var text = header.Header.Descendants<Text>();
            var tables = header.Header.Descendants<TableProperties>();
            if(text != null) FillText(text);
            if (tables != null) FillTables(tables);
        }
        private void FillText (IEnumerable<Text> text)
        {
            foreach (var txt in text)
            {
                if (_regex.IsMatch(txt.Text))
                {
                    foreach (object match in this._tagRegex.Matches(txt.Text))
                    {
                        string[] strArray = match.ToString().Replace("{", "").Replace("}", "").Split('.');
                        if (strArray.Length == 2)
                        {
                            string key = strArray[0];
                            if (!(key == "Item"))
                            {
                                string name = strArray[1];
                                object obj = _variables.Where<KeyValuePair<string, object>>((Func<KeyValuePair<string, object>, bool>)(v => v.Key == key)).Select<KeyValuePair<string, object>, object>((Func<KeyValuePair<string, object>, object>)(v => v.Value)).FirstOrDefault<object>();
                                if (obj != null)
                                {
                                    string str = obj.GetType().GetProperty(name)?.GetValue(obj)?.ToString();
                                    if (str != null)
                                    {
                                        if (str.Contains("rtf1"))
                                        {
                                            string id = name + (new Random(1000000).Next());
                                            using (MemoryStream sourceStream = new MemoryStream(Encoding.ASCII.GetBytes(str)))
                                                _document.MainDocumentPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Rtf, id).FeedData((Stream)sourceStream);
                                            txt.InsertAfterSelf<AltChunk>(new AltChunk()
                                            {
                                                Id = (StringValue)id
                                            });
                                            txt.Text = txt.Text.Replace(match.ToString(), "");
                                        }
                                        else if (str.Contains("<HTML"))
                                        {
                                            string id = name + (new Random(1000000).Next());
                                            using (MemoryStream sourceStream = new MemoryStream(Encoding.UTF8.GetBytes(str)))
                                                _document.MainDocumentPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html, id).FeedData((Stream)sourceStream);
                                            txt.InsertAfterSelf<AltChunk>(new AltChunk()
                                            {
                                                Id = (StringValue)id
                                            });
                                            txt.Text = txt.Text.Replace(match.ToString(), "");
                                        }
                                        else
                                            txt.Text = txt.Text.Replace(match.ToString(), str);
                                    }
                                    else
                                        txt.Text = txt.Text.Replace(match.ToString(), "[" + match.ToString() + " - значение было null]");
                                }
                            }
                        }
                    }
                }
            }
        }
        private void FillTables(IEnumerable<TableProperties> tables)
        {
            foreach (TableProperties table in tables)
            {
                string tableCaption = table.TableCaption?.Val?.ToString();
                if (tableCaption != null)
                {
                    IEnumerable<KeyValuePair<string, object>> variable = _variables.Where<KeyValuePair<string, object>>((Func<KeyValuePair<string, object>, bool>)(v => v.Key == tableCaption));
                    if (variable.Any<KeyValuePair<string, object>>())
                    {
                        var values = variable.FirstOrDefault<KeyValuePair<string, object>>().Value;
                        values.GetType();
                        if (values != null)
                        {
                            IList list = (IList)values;
                            Table parent = (Table)table.Parent;
                            TableRow tableRow = null;
                            foreach (var row in parent.Descendants<TableRow>())
                            {
                                IEnumerable<Text> textList = row.Descendants<Text>().Where((Func<Text, bool>)(t => _regex.IsMatch(t.Text)));
                                if (textList.Any() && this._itemRegex.IsMatch(textList.FirstOrDefault().Text))
                                {
                                    tableRow = row;
                                    break;
                                }
                            }
                            if (tableRow != null)
                            {
                                int valueIndex = list.Count - 1;
                                for (int i = 1; i < list.Count; ++i)
                                {
                                    var clonedRow = (TableRow)tableRow.CloneNode(true);
                                    var value = list[valueIndex];
                                    tableRow.InsertAfterSelf(clonedRow);
                                    foreach (var txt in clonedRow.Descendants<Text>())
                                    {
                                        foreach (object match in _tagRegex.Matches(txt.Text))
                                        {
                                            string[] tagValue = match.ToString().Replace("{", "").Replace("}", "").Split('.');
                                            if (((IEnumerable<string>)tagValue).Count<string>() == 2 && tagValue[0] == "Item")
                                            {
                                                string name = tagValue[1];
                                                string str = value.GetType().GetProperty(name)?.GetValue(value)?.ToString();
                                                if (str != null)
                                                {
                                                    if (str.Contains("rtf1"))
                                                    {
                                                        string id = name + (new Random(10000000).Next());
                                                        using (MemoryStream sourceStream = new MemoryStream(Encoding.ASCII.GetBytes(str)))
                                                            _document.MainDocumentPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Rtf, id).FeedData((Stream)sourceStream);
                                                        txt.InsertAfterSelf<AltChunk>(new AltChunk()
                                                        {
                                                            Id = (StringValue)id
                                                        });
                                                        txt.Text = txt.Text.Replace(match.ToString(), "");
                                                    }
                                                    else if (str.Contains("<HTML"))
                                                    {
                                                        string id = name + (new Random(10000000).Next());
                                                        using (MemoryStream sourceStream = new MemoryStream(Encoding.UTF8.GetBytes(str)))
                                                            _document.MainDocumentPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html, id).FeedData((Stream)sourceStream);
                                                        AltChunk newElement2 = new AltChunk();
                                                        newElement2.Id = (StringValue)id;
                                                        txt.InsertAfterSelf<AltChunk>(newElement2);
                                                        txt.Remove();
                                                    }
                                                    else
                                                        txt.Text = txt.Text.Replace(match.ToString(), str);
                                                }
                                                else
                                                    txt.Text = txt.Text.Replace(match.ToString(), "[" + match.ToString() + " - значение было null]");
                                            }
                                        }
                                    }
                                    --valueIndex;
                                }
                                foreach (var txt in tableRow.Descendants<Text>())
                                {
                                    var listIndex = list[valueIndex];
                                    foreach (object match in this._tagRegex.Matches(txt.Text))
                                    {
                                        string[] source5 = match.ToString().Replace("{", "").Replace("}", "").Split('.');
                                        if (((IEnumerable<string>)source5).Count<string>() == 2 && source5[0] == "Item")
                                        {
                                            string name = source5[1];
                                            string newValue = listIndex.GetType().GetProperty(name)?.GetValue(listIndex)?.ToString();
                                            if (newValue != null)
                                            {
                                                if (newValue.Contains("rtf1"))
                                                {
                                                    string id = name + +(new Random(10000000).Next());
                                                    using (MemoryStream sourceStream = new MemoryStream(Encoding.ASCII.GetBytes(newValue)))
                                                        _document.MainDocumentPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Rtf, id).FeedData((Stream)sourceStream);
                                                    txt.InsertAfterSelf<AltChunk>(new AltChunk()
                                                    {
                                                        Id = (StringValue)id
                                                    });
                                                    txt.Text = txt.Text.Replace(match.ToString(), "");
                                                }
                                                else if (newValue.Contains("<HTML"))
                                                {
                                                    string id = name + +(new Random(10000000).Next());
                                                    using (MemoryStream sourceStream = new MemoryStream(Encoding.UTF8.GetBytes(newValue)))
                                                        _document.MainDocumentPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html, id).FeedData((Stream)sourceStream);
                                                    AltChunk newElement2 = new AltChunk();
                                                    newElement2.Id = (StringValue)id;
                                                    txt.InsertAfterSelf<AltChunk>(newElement2);
                                                    txt.Remove();
                                                }
                                                else
                                                    txt.Text = txt.Text.Replace(match.ToString(), newValue);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        private void CleanDoc()
        {
            MarkupSimplifier.SimplifyMarkup(_document, new SimplifyMarkupSettings()
            {
                RemoveComments = true,
                RemoveContentControls = true,
                RemoveEndAndFootNotes = true,
                RemoveFieldCodes = false,
                RemoveLastRenderedPageBreak = true,
                RemovePermissions = true,
                RemoveProof = true,
                RemoveRsidInfo = true,
                RemoveSmartTags = true,
                RemoveSoftHyphens = true,
                ReplaceTabsWithSpaces = true
            });
            _document.CleanRun();
        }
        public string SaveAs(string outputFilePath)
        {
            try
            {
                OpenXmlPackage openXmlPackage = _document.Clone(outputFilePath);
                _document.Dispose();
                openXmlPackage.Dispose();
                return outputFilePath;
            }
            catch (Exception ex)
            {
                _document.Dispose();
                throw ex;
            }
        }
    }
}
