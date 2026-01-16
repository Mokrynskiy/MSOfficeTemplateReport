using System;
using System.Collections.Generic;
using System.Text;

namespace MSOfficeTemplateReport
{
    public interface ITemplate
    {
        void AddVariable(string name, object data);
        void Generate();

        string SaveAs(string path);
    }
}
