using MSOfficeTemplateReport.ExcelReport;
using MSOfficeTemplateReport.WordReport;

Header header = new()
{
    TitleA = "Заголовок1",
    TitleB = "Заголовок2",
    TitleC = "Заголовок3"   
};
List<Positions> pos = new List<Positions>
{
    new Positions {PropA = 1.1, PropB = "1.2", PropC = "1.3"},
    new Positions {PropA = 2.1, PropB = "2.2", PropC = "2.3"},
    new Positions {PropA = 3.1, PropB = "3.2", PropC = "3.3"}
};


string resultFileName = "Result.xlsx";
var data = File.ReadAllBytes("ExcelTest.xlsx");
var template = new ExcelTemplate(data);
template.AddVariable("Header", header);
template.AddVariable("Prod", pos);
template.Generate();
template.SaveAs(resultFileName);


class Header
{
    public string TitleA { get; set; }
    public string TitleB { get; set; }
    public string TitleC { get; set; }
}

class Positions
{
    public double PropA { get; set; }
    public string PropB { get; set; }
    public string PropC { get; set; }
}