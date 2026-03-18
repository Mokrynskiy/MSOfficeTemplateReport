using MSOfficeTemplateReport;
using TestProject;

const string excelTemplatePath = "testexceltemplate.xlsx";

const string wordTemplatePath = "testwordtemplate.docx";

Dictionary<string, object> variables = new Dictionary<string, object>();

variables.Add("Order", new Order());

var template = Template.Create(wordTemplatePath);

template.AddVariable("Order", new Order());

var result = template.Generate("result");

result.SaveAs("result.xlsx");
