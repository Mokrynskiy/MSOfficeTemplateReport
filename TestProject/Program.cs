using MSOfficeTemplateReport;

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


Dictionary<string, object> variables = new Dictionary<string, object>();
variables.Add("Header", header);
variables.Add("Prod", pos);

string resultFileName = "Result.docx";
var data = File.ReadAllBytes("ExcelTest.xlsx");
var template = Template.Create("ExcelTest.xlsx", variables);
var result = template.Generate(null);

File.WriteAllBytes(result.FileName, result.ByteArray);


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