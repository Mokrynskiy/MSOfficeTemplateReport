using MSOfficeTemplateReport.Models;
using System.Collections.Generic;

namespace MSOfficeTemplateReport.Abstract
{
    public interface ITemplate
    {
        void AddVariable(string name, object data);

        void AddVariables(Dictionary<string, object> variables);

        void Generate();

        GenerateResultModel Generate(string fileName);

        string SaveAs(string path);

        byte[] ToByteArray();
    }
}
