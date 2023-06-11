using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace Akelon3Test
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Введите путь до файла с данными:");
            string filePath = Console.ReadLine();

            using (var workbook = new XLWorkbook(filePath))
            {
                var productsSheet = workbook.Worksheet("Товары");
                var customersSheet = workbook.Worksheet("Клиенты");
                var ordersSheet = workbook.Worksheet("Заявки");
                var separator = "---------------------------------------------------------------"; //Разделитель блоков 


                while (true)
                {
                    Console.WriteLine();
                    Console.WriteLine("Выберите команду:");
                    Console.WriteLine("1. Поиск клиентов по наименованию товара");
                    Console.WriteLine("2. Изменение контактного лица клиента");
                    Console.WriteLine("3. Поиск золотого клиента");
                    Console.WriteLine("4. Выйти");
                    Console.WriteLine(separator);

                    int choice;
                    if (!int.TryParse(Console.ReadLine(), out choice)) //Проверка на число
                    {
                        Console.WriteLine("Некорректный ввод. Попробуйте снова.");
                        Console.WriteLine(separator);
                        continue;
                    }

                    switch (choice)
                    {
                        case 1:
                            Console.WriteLine("Введите наименование товара:");
                            string productName = Console.ReadLine();
                            SearchCustomersByProduct(productsSheet, customersSheet, ordersSheet, productName);
                            Console.WriteLine(separator);
                            break;

                        case 2:
                            Console.WriteLine("Введите название организации клиента:");
                            string customerOrganization = Console.ReadLine();
                            Console.WriteLine("Введите ФИО нового контактного лица организации:");
                            string newContactPerson = Console.ReadLine();
                            UpdateContactPerson(customersSheet, customerOrganization, newContactPerson);
                            Console.WriteLine(separator);
                            break;

                        case 3:
                            Console.WriteLine("Введите год:");
                            int year;
                            if (!int.TryParse(Console.ReadLine(), out year))
                            {
                                Console.WriteLine("Некорректный ввод года. Попробуйте снова.");
                                Console.WriteLine(separator);
                                break;
                            }
                            Console.WriteLine("Введите месяц:");
                            int month;
                            if (!int.TryParse(Console.ReadLine(), out month))
                            {
                                Console.WriteLine("Некорректный ввод месяца. Попробуйте снова.");
                                Console.WriteLine(separator);
                                break;
                            }
                            FindGoldenCustomer(ordersSheet,customersSheet, year, month);
                            Console.WriteLine(separator);
                            break;

                        case 4:
                            return;

                        default:
                            Console.WriteLine("Некорректный выбор. Попробуйте снова.");
                            Console.WriteLine(separator);
                            break;
                    }
                }
            }
        }

        static void SearchCustomersByProduct(IXLWorksheet productsSheet, IXLWorksheet customersSheet, IXLWorksheet ordersSheet, string productName)
        {
            var productRows = productsSheet.RowsUsed().Skip(1); // Skip(1) -Пропуск первой строки(с заголовком)
            var customers = new List<string>();

            foreach (var row in productRows)
            {
                if (row.Cell(2).GetString() == productName)
                {
                    var productCode = row.Cell(1).GetString();
                    var orderRows = ordersSheet.RowsUsed().Skip(1); 

                    foreach (var orderRow in orderRows)
                    {
                        if (orderRow.Cell(2).GetString() == productCode)
                        {
                            var customerCode = orderRow.Cell(3).GetString();
                            var customerRow = customersSheet.RowsUsed().Skip(1) 
                                .FirstOrDefault(r => r.Cell(1).GetString() == customerCode); //Сразу выбираем строку по номеру

                            if (customerRow != null)
                            {
                                var customerName = customerRow.Cell(2).GetString();
                                var orderQuantity = Convert.ToInt32(orderRow.Cell(5).GetString());
                                var orderPrice = row.Cell(4).GetDouble();
                                var orderDate = orderRow.Cell(6).GetDateTime();

                                var customerInfo = $"{customerName}, Количество: {orderQuantity}, Цена: {orderPrice}, Дата заказа: {orderDate.ToShortDateString()}";
                                customers.Add(customerInfo);
                            }
                        }
                    }
                }
            }

            if (customers.Count > 0)
            {
                Console.WriteLine("Клиенты, заказавшие данный товар:");
                foreach (var customerInfo in customers)
                {
                    Console.WriteLine(customerInfo);
                }
            }
            else
            {
                Console.WriteLine("Нет клиентов, заказавших данный товар.");
            }
        }

        static void UpdateContactPerson(IXLWorksheet customersSheet, string customerOrganization, string newContactPerson)
        {
            var customerRow = customersSheet.RowsUsed().Skip(1) 
                .FirstOrDefault(r => r.Cell(2).GetString() == customerOrganization); //Сразу выбираем строку с нужной организацией

            if (customerRow != null)
            {
                var oldContactPerson = customerRow.Cell(4).GetString();
                customerRow.Cell(4).Value = newContactPerson;

                Console.WriteLine($"Контактное лицо клиента '{customerOrganization}' успешно изменено. " +
                    $"Старое контактное лицо: {oldContactPerson}, Новое контактное лицо: {newContactPerson}");
            }
            else
            {
                Console.WriteLine($"Клиент с названием организации '{customerOrganization}' не найден.");
            }
        }

        static void FindGoldenCustomer(IXLWorksheet ordersSheet, IXLWorksheet customersSheet, int year, int month)
        {
            var orderRows = ordersSheet.RowsUsed().Skip(1); 

            var customerOrders = new Dictionary<string, int>();

            foreach (var row in orderRows)
            {
                var orderDate = row.Cell(6).GetDateTime();

                if (orderDate.Year == year && orderDate.Month == month)
                {
                    var customerCode = row.Cell(3).GetString();
                    if (customerOrders.ContainsKey(customerCode))
                    {
                        customerOrders[customerCode]++;
                    }
                    else
                    {
                        customerOrders.Add(customerCode, 1);
                    }
                }
            }

            if (customerOrders.Count > 0)
            {
                var goldenCustomerCode = customerOrders.OrderByDescending(c => c.Value).First().Key;
                var customerRow = customersSheet.RowsUsed().Skip(1)
                .FirstOrDefault(r => r.Cell(1).GetString() == goldenCustomerCode);
                if (customerRow != null)
                { 
                    Console.WriteLine($"Золотой клиент за {month}.{year}: {customerRow.Cell(2).GetString()}");
                }
            }
            else
            {
                Console.WriteLine($"Заказов за {month}.{year} не найдено.");
            }
        }
    }
}
