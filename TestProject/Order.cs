namespace TestProject;

public class Order
{
    public int Number { get; set; }

    public string Date { get; set; }

    public string Customer { get; set; }

    public decimal Sum => Positions.Sum(x => x.Sum);

    public string SumString => Sum.ToString("N2");

    public List<OrderPosition> Positions { get; set; }

    public Order()
    {
        Date = DateTime.Now.ToShortDateString();
        Number = 156;
        Customer = "Иванов Иван Иванович";
        Positions = new List<OrderPosition>
        {
            new OrderPosition{RowNumber = 1, Name = "Товар 1", Amount = 1, Price = 120500},
            new OrderPosition{RowNumber = 2, Name = "Товар 2", Amount = 10, Price = 1505},
            new OrderPosition{RowNumber = 3, Name = "Товар 3", Amount = 5, Price = 125.99M},
            new OrderPosition{RowNumber = 4, Name = "Товар 4", Amount = 3, Price = 2564}
        };
    }
}
