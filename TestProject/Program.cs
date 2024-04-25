using WordTemplateReport.WordReport;

Header header = new()
{
    TitleA = "Заголовок1",
    TitleB = "Заголовок2",
    TitleC = @"Заголовок3",
    Positions = new()
    {
        new() {PropA = "1.1", PropB = "1.2", PropC = "<HTML lang=\"ru\"><HEAD><meta charset=\"utf-8\"/></HEAD><BODY><DIV style=\"color:red\"> <b><u>2</u></b> наб., сер.2738, годен до 12.07.24, 21.31%; </DIV><DIV style=\"color:red\"> <b><u>16</u></b> наб., сер.2821, годен до 17.11.24, 56.28%; </DIV></BODY></HTML>"},
        new() {PropA = "2.1", PropB = "2.2", PropC = "2.3"},
        new() {PropA = "3.1", PropB = "3.2", PropC = "<HTML lang=\"ru\"><HEAD><meta charset=\"utf-8\"/></HEAD><BODY><DIV style=\"color:red\"> <b><u>2</u></b> наб., сер.2738, годен до 12.07.24, 21.31%; </DIV><DIV style=\"color:red\"> <b><u>16</u></b> наб., сер.2821, годен до 17.11.24, 56.28%; </DIV></BODY></HTML>"}
    },
    Users = new()
    {
        new(){Position = "Должность1", Name="Имя1", Role = 1},
        new(){Position = "Должность2", Name="Имя2", Role = 1},
        new(){Position = "Должность3", Name="Имя3", Role = 1},
        new(){Position = "Должность4", Name="Имя4", Role = 1},
        new(){Position = "Должность5", Name="Имя5", Role = 2},
        new(){Position = "Должность6", Name="Имя6", Role = 2},
    }
};
string resultFileName = "Result.docx";
var template = new WordTemplate("TestTemplate.docx");
template.AddVariable("Header", header);
template.AddVariable("Prod", header.Positions);
template.AddVariable("First", header.Users.Where(x => x.Role == 1).ToList());
template.AddVariable("Second", header.Users.Where(x => x.Role == 2).ToList());
template.Generate();
template.SaveAs(resultFileName);


class Header
{
    public string TitleA { get; set; }
    public string TitleB { get; set; }
    public string TitleC { get; set; }
    public List<Positions> Positions { get; set; }
    public List<Users> Users { get; set; }
}

class Positions
{
    public string PropA { get; set; }
    public string PropB { get; set; }
    public string PropC { get; set; }
}
class Users
{
    public string Position { get; set; }
    public string Name { get; set; }
    public int Role { get; set; }
}