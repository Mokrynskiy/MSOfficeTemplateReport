using MSOfficeTemplateReport;
using TestProject;

const string wordTemplatePath = "testwordtemplate.docx";

var template = Template.Create(wordTemplatePath);

template.AddVariable("Order", new Order());

var result = template.Generate();

result.SaveAs("result.xlsx");
