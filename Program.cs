using PracticTask3.Entity;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using ConsoleTables;
using DocumentFormat.OpenXml.Bibliography;

namespace PracticTask3
{
    internal class Program
    {
        private static string _filePath;
        static void Main(string[] args)
        {
            try
            {

                Console.Clear();
                Console.Write("******** РАБОТА С ТАБЛИЦАМИ ******** \n\n" +
                    "Для выхода из программы в любой момент введите \"exit\" \n\n" +
                    "Введите путь к файлу Excel: ");

                //@"C:\Users\max11\source\repos\Практическое задание для кандидата.xlsx";
                _filePath = Console.ReadLine();
                if (!File.Exists(_filePath))
                    throw new Exception("Файл не обнаружен. Перезапустите программу");

                using (var workBook = new XLWorkbook(_filePath))
                {
                    //Создание сущностей
                    var goods = new Dictionary<double, Goods>();
                    var requests = new Dictionary<double, Request>();
                    var clients = new Dictionary<double, Client>();

                    //Создание словаря Товаров
                    foreach (var row in workBook.Worksheet(1).RowsUsed())
                    {
                        if (row.RowNumber() == 1) continue;
                        goods[row.Cell(1).Value.GetNumber()] = new Goods()
                        {
                            Id = row.Cell(1).Value.GetNumber(),
                            Name = row.Cell(2).Value.GetText(),
                            Units = row.Cell(3).Value.GetText(),
                            Price = row.Cell(4).Value.GetNumber()
                        };
                    }
                    //Создание словаря Клиентов
                    foreach (var row in workBook.Worksheet(2).RowsUsed())
                    {
                        if (row.RowNumber() == 1) continue;
                        clients[row.Cell(1).Value.GetNumber()] = new Client()
                        {
                            Id = row.Cell(1).Value.GetNumber(),
                            CompanyName = row.Cell(2).Value.GetText(),
                            Adress = row.Cell(3).Value.GetText(),
                            ClientName = row.Cell(4).Value.GetText()
                        };
                    }

                    //Создание словаря Заявок
                    foreach (var row in workBook.Worksheet(3).RowsUsed())
                    {
                        if (row.RowNumber() == 1) continue;
                        requests[row.Cell(1).Value.GetNumber()] = new Request()
                        {
                            Id = row.Cell(1).Value.GetNumber(),
                            Product = goods[row.Cell(2).Value.GetNumber()],
                            Client = clients[row.Cell(3).Value.GetNumber()],
                            RequestNumber = row.Cell(4).Value.GetNumber(),
                            Quantity = row.Cell(5).Value.GetNumber(),
                            Date = row.Cell(6).Value.GetDateTime()
                        };
                    }

                    do
                    {
                        Console.Clear();

                        Console.WriteLine($"******** Работа с Excel ********\n\n" +
                            $"Доступные действия: \n" +
                            $"[1] - Получить информацию о заявках по наимеованию товара\n" +
                            $"[2] - Изменение контактной информации клиентов\n" +
                            $"[3] - Поиск золотого клинента в указаный период времени\n");
                        Console.Write("Выберите действие: ");
                        var userStr = Console.ReadLine();
                        if (userStr == "exit")
                            break;

                        switch (userStr)
                        {
                            case "1":
                                Console.Clear();
                                Console.Write("Введите наименование товара: ");
                                var productName = Console.ReadLine();
                                if (productName == "exit") continue;
                                var key = goods.FirstOrDefault(p => p.Value.Name == productName).Key;
                                var req = requests.Where(p => p.Value.Product.Id == key).ToList();
                                ProductInfo(req);
                                Console.ReadKey();
                                break;

                            case "2":
                                ChangeClientInfo(workBook);
                                break;
                            case "3":
                                Console.Clear();
                                Console.WriteLine("* Введите даты в промежутки между которыми будет произведен поиск\n\n" +
                                    "* Даты вводить в формате \"День.Месяц.Год\" или \"Год.Месяц.День\"\n\n" +
                                    "* Пример: 1.3.2023 или 2023.3.1\n");

                                Console.Write("Введите начальную дату: ");
                                var userInput = Console.ReadLine();
                                if (userInput == "exit") continue;
                                var start = DateTime.Parse(userInput);

                                Console.Write("Введите конечную дату: ");
                                userInput = Console.ReadLine();
                                if (userInput == "exit") continue;
                                var end = DateTime.Parse(userInput);

                                GoldenClient(requests, start, end);
                                break;
                        }
                    } while (true);





                    //string s = "1.4.2023";
                    //string s2 = "1.7.2023";

                    //var dt = DateTime.Parse(s);
                    //var dt2 = DateTime.Parse(s2);

                    //

                }
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        //Вывод информации о товаре
        public static void ProductInfo(List<KeyValuePair<double, Request>> table)
        {
            table.ForEach(e => Console.WriteLine(e.Value.ToString()));
        }

        //Изменение информации о клиенте
        public static void ChangeClientInfo(XLWorkbook wb)
        {
            Console.Clear();
            Console.WriteLine("************* Редактирование контактной информации *************\n");
            var ws = wb.Worksheet(2);
            foreach (var row in ws.RowsUsed())
            {
                Console.WriteLine($"{row.Cell(1).GetString(),-11} | {row.Cell(2).GetString(),-25} | {row.Cell(3).GetString(),-60} | {row.Cell(4).GetString(),-5}");
            }

            Console.Write("\nВведите код клиента: ");
            var clientIndex = Console.ReadLine();

            var rows = ws.RowsUsed(r => r.FirstCellUsed().GetString() == clientIndex);

            foreach (var row in rows)
            {
                Console.Write("\nВведите наимнование организации: ");
                var companyName = Console.ReadLine();
                Console.Write("Введите адрес: ");
                var companyAdress = Console.ReadLine();
                Console.Write("Введите контактное лицо (ФИО): ");
                var companyClient = Console.ReadLine();

                if (companyName == "")
                    companyName = row.Cell(2).Value.GetText();
                if (companyAdress == "")
                    companyAdress = row.Cell(3).Value.GetText();
                if (companyClient == "")
                    companyClient = row.Cell(4).Value.GetText();

                row.Cell(2).Value = companyName;
                row.Cell(3).Value = companyAdress;
                row.Cell(4).Value = companyClient;

                Console.Write("\nДанные изменены");
                Console.ReadKey();

                wb.Save();

            }
        }

        public static void GoldenClient(Dictionary<double, Request> requests, DateTime start, DateTime end)
        {
            var clientsSearch = requests
                        .Where(d => d.Value.Date >= start && d.Value.Date <= end)
                        .GroupBy(p => p.Value.Client.CompanyName)
                        .Select(g => new { Name = g.Key, Count = g.Count() })
                        .ToList();
            var maxCount = clientsSearch.Max(p => p.Count);
            var goldenClients = clientsSearch.Where(p => p.Count == maxCount).ToList();

            foreach (var g in goldenClients)
            {
                Console.WriteLine($"{g.Name}");
            }
            Console.ReadKey();
        }
    }

}
