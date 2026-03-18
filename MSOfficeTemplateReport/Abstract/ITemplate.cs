using MSOfficeTemplateReport.Models;
using System.Text.Json.Nodes;

namespace MSOfficeTemplateReport.Abstract
{
    public interface ITemplate
    {
        /// <summary>
        /// Добавить переменную отчета
        /// </summary>
        /// <param name="name">Псевдоним</param>
        /// <param name="variable">Object или Json</param>
        void AddVariable(string name, object variable);

        /// <summary>
        /// Сгенерировать отчет
        /// </summary>
        /// <param name="fileName">Имя файла</param>
        /// <returns></returns>
        ReportResultModel Generate(string fileName = null);
    }
}
