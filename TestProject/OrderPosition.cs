namespace TestProject;

public class OrderPosition
{
    public int RowNumber { get; set; }

    public string Name { get; set; }

    public int Amount { get; set; }

    public decimal Price { get; set; }

    public decimal Sum => Amount * Price;

    public string PriceString => Price.ToString("N2");

    public string SumString => Sum.ToString("N2");
}
