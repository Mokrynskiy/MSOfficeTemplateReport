using System.IO;

namespace MSOfficeTemplateReport.Models
{
    public sealed class ReportResultModel
    {
        public string FileName { get; set; }
        public byte[] ByteArray { get; set; }

        public ReportResultModel(string fileName, byte[] byteArray)
        {
            FileName = fileName;

            ByteArray = byteArray;
        }

        public void SaveAs(string filePath)
        {
            File.WriteAllBytes(filePath, ByteArray);
        }
    }
}