using MSOfficeTemplateReport.Models;
using System;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace MSOfficeTemplateReport.Extensions
{
    internal static class FileExtensions
    {
        internal static readonly byte[] DocxXlsxSignature = new byte[] { 0x50, 0x4B, 0x03, 0x04 };

        internal static string GetFormat(this byte[] data)
        {
            if (data.Length < DocxXlsxSignature.Length || !data.Take(DocxXlsxSignature.Length).SequenceEqual(DocxXlsxSignature))
                return null;

            try
            {
                using (var stream = new MemoryStream(data))
                {
                    using (var archive = new ZipArchive(stream, ZipArchiveMode.Read))
                    {
                        if (archive.GetEntry("word/document.xml") != null && archive.GetEntry("[Content_Types].xml") != null)
                            return FileFormates.Docx;
                        else if (archive.GetEntry("xl/workbook.xml") != null && archive.GetEntry("[Content_Types].xml") != null)
                            return FileFormates.Xlsx;
                        else
                            return null;
                    }
                }
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}