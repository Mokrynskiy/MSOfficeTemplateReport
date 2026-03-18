# Генерация отчетов по шаблонам EXCEL и WORD

# ПРАВИЛА ФОРМИРОВАНИЯ ШАБЛОНОВ
## Правила формирования шаблона Word
>Тэги, за исключением тэгов в табличной части (которая заполняется из коллекций), состоят из пары {{Ключ.Значение}}, где ключ - key элемента переданной в Dictionary<string, object>, а значение - это наименование поля объекта.

<img width="1349" height="384" alt="wordtemplate" src="https://github.com/user-attachments/assets/849bc2ee-4e0b-42e5-bb2c-0c3ebe93c78e" />


>Тэги в табличной части (которая заполняеися из коллекций) имеют следующую структуру ***{{Item.Value}}***, где "Item" - это обязателное слово, а 'Value' - это имя поля из элемента коллекции. 
Для таблиц, тек же, необходимым условием является задание в свойствах таблицы заголовка замещающего текста который должен совпадать с наименованием поля с коллекцией в объекте или, если коллекция была передана как отдельная переменная, псевданиму данной переменной.

<img width="386" height="449" alt="wordtemplatetablesettings" src="https://github.com/user-attachments/assets/34fcbb96-f174-4d46-90f4-625c6d553602" />

## Правила формирования шаблона Excel
>Оформление тэгов в шаблонах Excel аналогично Word за исключением наличия специальных тэгов:
**<\<Sum>>** - Выводит сумму значений колонки
**<\<RowNumber>>** - Нумерует строки

<img width="1181" height="314" alt="exceltemplate" src="https://github.com/user-attachments/assets/27dd2d40-f264-4788-857f-3f1891e36869" />

>Для таблиц в Excel, которые заполняются из коллекций, необходимо создавать именованный диапазон ячеек который должен состоять из двухстрок.

<img width="1177" height="160" alt="excelnamedarray" src="https://github.com/user-attachments/assets/8ee5b4db-28e5-4c53-b464-f97769562420" />


<img width="559" height="446" alt="exceltemplatetablesettings" src="https://github.com/user-attachments/assets/305faf8c-949d-4d9a-90fe-7c1a42777086" />

<hr>

# ПРИМЕР КОНСОЛЬНОГО ПРИЛОЖЕНИЯ НА С#
Для выполнения примера необходимо в корень проекта добавить файлы шаблонов по ссылкам выше и в свойствах файла в VisualStudio выставить параметры - действе при сборке - содержание, копировать в выходной каталог - копировать более позднюю версию.

* **класс Order.cs**
```cs
internal class Order
{
    public int Number { get; set; }

    public string Date { get; set; }

    public string Customer { get; set; }

    public decimal Sum => Positions.Sum(x => x.Sum);

    public string SumString => Sum.ToString("N2");

    public List<OrderPosition> Positions { get; set; }
}
```

* **класс OrderPosition.cs**
```cs
internal class OrderPosition
{
    public int RowNumber { get; set; }
    
    public string Name { get; set; }
    
    public int Amount { get; set; }
    
    public decimal Price { get; set; }

    public decimal Sum => Amount * Price;

    public string PriceString => Price.ToString("N2");

    public string SumString => Sum.ToString("N2");
}
```
* **класс Program.cs**
```cs
using MSOfficeTemplateReport;
using TestProject;

const string wordTemplatePath = "testwordtemplate.docx";

var template = Template.Create(wordTemplatePath);

template.AddVariable("Order", new Order());

var result = template.Generate();

result.SaveAs("result.xlsx");
```
