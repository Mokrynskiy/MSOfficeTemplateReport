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
    new Positions {PropA = 1.1, PropB = 1.2, PropC = DateTime.UtcNow},
    new Positions {PropA = 2.1, PropB = 2.2, PropC = DateTime.Now},
    new Positions {PropA = 3.1, PropB = 3.2, PropC = DateTime.Now}
};


string resultFileName = "Result.xlsx";
var data = File.ReadAllBytes("ExcelTest.xlsx");
var template = new ExcelTemplate(data);
template.AddVariable("Header", header);
template.AddVariable("Prod", pos);
template.Generate();
File.WriteAllBytes("1.xlsx", template.ToByteArray());
//template.SaveAs(resultFileName);


class Header
{
    public string TitleA { get; set; }
    public string TitleB { get; set; }
    public string TitleC { get; set; }
}

class Positions
{
    public double PropA { get; set; }
    public double PropB { get; set; }
    public DateTime PropC { get; set; }
}