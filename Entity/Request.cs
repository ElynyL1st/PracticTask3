namespace PracticTask3.Entity
{
    public class Request : BaseTable
    {
        public Goods? Product { get; set; }
        public Client? Client { get; set; }
        public double RequestNumber { get; set; }
        public double Quantity { get; set; }
        public DateTime Date { get; set; }

        public override string ToString()
        {
            string str = $"--------- Информация о заявке --------- \n" +
                $"* Товар: {Product.Name} ({Product.Id})\n" +
                $"* Номер заявки: {RequestNumber} от {Date.ToString("D")}\n" +
                $"* Колличество: {Quantity}\n" +
                $"* Цена за единицу: {Product.Price}\n" +
                $"* Компания: {Client.CompanyName}\n" +
                $"* Адрес: {Client.Adress} \n" +
                $"* Контактное лицо: {Client.ClientName}";

            return str;
        }
    }
}
