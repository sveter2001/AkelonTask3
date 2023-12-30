using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace PracticTask3
{
    internal class DataProccesor
    {
        private SpreadsheetDocument _MyDocument;

        public DataProccesor(string path)
        {
            try
            {
                _MyDocument = SpreadsheetDocument.Open(path, true);
                Console.Clear();
                Console.WriteLine("Открыт файл " + $"{path}");
            }
            catch
            {
                Console.WriteLine("!@#$%^&   Неполучилось открыть файл   !@#$%^&");
                Environment.Exit(0);
            }
        }

        public void GetGoldenClient(string month, string year)
        {
            DateTime dateOfOrder;
            Dictionary<string, int> OrdersClientsDic = new Dictionary<string, int>();
            WorkbookPart workbookPart = _MyDocument.WorkbookPart;
            WorksheetPart worksheetPart1 = workbookPart.WorksheetParts.First();
            WorksheetPart worksheetPart2 = workbookPart.WorksheetParts.ElementAt(1);

            var sheetData1 = worksheetPart1.Worksheet.Elements<SheetData>().First();
            var sheetData2 = worksheetPart2.Worksheet.Elements<SheetData>().First();

            SharedStringTablePart sharedStringTablePart = workbookPart.SharedStringTablePart;
            SharedStringTable sharedStringTable = sharedStringTablePart.SharedStringTable;

            foreach (Row row1 in sheetData1.Elements<Row>().Skip(1))//перебор таблицы Заявки
            {
                if (row1.Elements<Cell>().ElementAt(5).InnerText != "")
                {
                    dateOfOrder = DateTime.FromOADate(double.Parse(row1.Elements<Cell>().ElementAt(5).InnerText));
                    if (month == "")
                    {
                        if (year == dateOfOrder.ToString("yyyy"))
                        {
                            if (OrdersClientsDic.ContainsKey(row1.Elements<Cell>().ElementAt(2).InnerText))
                            {
                                OrdersClientsDic[row1.Elements<Cell>().ElementAt(2).InnerText] += 1;
                            }
                            else
                            {
                                OrdersClientsDic.Add(row1.Elements<Cell>().ElementAt(2).InnerText, 1);
                            }
                        }
                    }
                    else if (month == dateOfOrder.ToString("MM") && year == dateOfOrder.ToString("yyyy"))
                    {
                        if (OrdersClientsDic.ContainsKey(row1.Elements<Cell>().ElementAt(2).InnerText))
                        {
                            OrdersClientsDic[row1.Elements<Cell>().ElementAt(2).InnerText] += 1;
                        }
                        else
                        {
                            OrdersClientsDic.Add(row1.Elements<Cell>().ElementAt(2).InnerText, 1);
                        }
                    }
                }
            }
            if (OrdersClientsDic.Count != 0)
            {
                int maxValue = OrdersClientsDic.Max(pair => pair.Value);
                Console.WriteLine($"Большее количество заказов от одного клиента: {maxValue}");
                Console.WriteLine("Клиент(ы):");

                var maxElements = OrdersClientsDic.Where(x => x.Value == maxValue).Select(x => x.Key);

                foreach (Row row2 in sheetData2.Elements<Row>().Skip(1))//перебор таблицы Клиенты
                {
                    if (maxElements.Any(element => element == row2.Elements<Cell>().ElementAt(0).InnerText))
                    {
                        Console.WriteLine($"{sharedStringTable.ElementAt(int.Parse(row2.Elements<Cell>().ElementAt(1).InnerText)).InnerText}");
                    }
                }
            }
            else
            {
                Console.WriteLine("Заказов за данный период не найдено");
            }
        }

        public void GetCustomers()
        {
            WorkbookPart workbookPart = _MyDocument.WorkbookPart;
            WorksheetPart worksheetPart2 = workbookPart.WorksheetParts.ElementAt(1);
            var sheetData2 = worksheetPart2.Worksheet.Elements<SheetData>().First();

            SharedStringTablePart sharedStringTablePart = workbookPart.SharedStringTablePart;
            SharedStringTable sharedStringTable = sharedStringTablePart.SharedStringTable;

            foreach (Row row2 in sheetData2.Elements<Row>().Skip(1))//перебор таблицы Клиенты
            {
                Cell cell = row2.Elements<Cell>().ElementAt(1);

                string value = cell.InnerText;
                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                {
                    int index = int.Parse(value);
                    value = sharedStringTable.ElementAt(index).InnerText;
                }
                Console.WriteLine(value);
            }
        }

        public void GetGoods()
        {
            WorkbookPart workbookPart = _MyDocument.WorkbookPart;
            WorksheetPart worksheetPart3 = workbookPart.WorksheetParts.ElementAt(2);
            var sheetData3 = worksheetPart3.Worksheet.Elements<SheetData>().First();

            SharedStringTablePart sharedStringTablePart = workbookPart.SharedStringTablePart;
            SharedStringTable sharedStringTable = sharedStringTablePart.SharedStringTable;

            foreach (Row row3 in sheetData3.Elements<Row>().Skip(1))//перебор таблицы Товары
            {
                Cell cell = row3.Elements<Cell>().ElementAt(1);

                string value = cell.InnerText;
                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                {
                    int index = int.Parse(value);
                    value = sharedStringTable.ElementAt(index).InnerText;
                }
                Console.WriteLine(value);
            }
        }

        public void ChangeContact(string item)
        {
            string contact = "";
            WorkbookPart workbookPart = _MyDocument.WorkbookPart;
            WorksheetPart worksheetPart2 = workbookPart.WorksheetParts.ElementAt(1);
            var sheetData2 = worksheetPart2.Worksheet.Elements<SheetData>().First();

            SharedStringTablePart sharedStringTablePart = workbookPart.SharedStringTablePart;
            SharedStringTable sharedStringTable = sharedStringTablePart.SharedStringTable;

            foreach (Row row2 in sheetData2.Elements<Row>().Skip(1))//перебор таблицы Клиенты
            {
                Cell cell = row2.Elements<Cell>().ElementAt(1);

                string value = cell.InnerText;
                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                {
                    int index = int.Parse(value);
                    value = sharedStringTable.ElementAt(index).InnerText;
                }
                if (value == item)
                {
                    contact = sharedStringTable.ElementAt(int.Parse(row2.Elements<Cell>()
                                        .ElementAt(3).InnerText)).InnerText;
                    Console.WriteLine("Текущее контактное лицо (ФИО): " + $"{contact}");
                    Console.WriteLine("Для отмены изменений оставте поле пустым (нажмите Enter)");
                    Console.Write("Новое контактное лицо (ФИО): ");
                    contact = Console.ReadLine();
                    if (contact != "")
                    {
                        sharedStringTable.ChildElements[int.Parse(row2.Elements<Cell>()//вставка значения в ячейку через Xml
                                        .ElementAt(3).InnerText)].InnerXml =
                                        $"<x:t xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">{contact}</x:t>";
                        _MyDocument.Save();
                    }

                    if (contact == sharedStringTable.ElementAt(int.Parse(row2.Elements<Cell>()
                                        .ElementAt(3).InnerText)).InnerText)
                    {
                        Console.WriteLine("\nУспешное изменение");
                    }
                    else
                    {
                        Console.WriteLine("\nИзменений нет");
                        contact = sharedStringTable.ElementAt(int.Parse(row2.Elements<Cell>()
                                        .ElementAt(3).InnerText)).InnerText;
                    }
                    Console.WriteLine("Текущее контактное лицо (ФИО): " + $"{contact}");
                }
            }
        }

        public void GetInfo(string item)
        {
            Console.Clear();
            Console.WriteLine(item + "\n");
            string itemCode = "";
            string unit = "";
            string price = "";
            string clientCode = "";
            string amount = "";
            string dateOfOrder = "";
            string orgName = "";
            string addres = "";
            string contact = "";
            bool orderExists = false;
            WorkbookPart workbookPart = _MyDocument.WorkbookPart;
            WorksheetPart worksheetPart1 = workbookPart.WorksheetParts.First();
            WorksheetPart worksheetPart2 = workbookPart.WorksheetParts.ElementAt(1);
            WorksheetPart worksheetPart3 = workbookPart.WorksheetParts.ElementAt(2);
            var sheetData1 = worksheetPart1.Worksheet.Elements<SheetData>().First();
            var sheetData2 = worksheetPart2.Worksheet.Elements<SheetData>().First();
            var sheetData3 = worksheetPart3.Worksheet.Elements<SheetData>().First();

            SharedStringTablePart sharedStringTablePart = workbookPart.SharedStringTablePart;
            SharedStringTable sharedStringTable = sharedStringTablePart.SharedStringTable;

            foreach (Row row3 in sheetData3.Elements<Row>().Skip(1))//перебор таблицы Товары
            {
                Cell cell = row3.Elements<Cell>().ElementAt(1);

                string value = cell.InnerText;
                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                {
                    int index = int.Parse(value);
                    value = sharedStringTable.ElementAt(index).InnerText;
                }
                if (value == item)
                {
                    itemCode = row3.Elements<Cell>().ElementAt(0).InnerText;
                    unit = sharedStringTable.ElementAt(int.Parse(row3.Elements<Cell>()
                        .ElementAt(2).InnerText)).InnerText;
                    price = row3.Elements<Cell>().ElementAt(3).InnerText;
                    foreach (Row row1 in sheetData1.Elements<Row>().Skip(1))//перебор таблицы заявки
                    {
                        if (itemCode == row1.Elements<Cell>().ElementAt(1).InnerText)
                        {
                            orderExists = true;
                            clientCode = row1.Elements<Cell>().ElementAt(2).InnerText;
                            amount = row1.Elements<Cell>().ElementAt(4).InnerText;
                            dateOfOrder = DateTime.FromOADate(double.Parse(row1.Elements<Cell>().ElementAt(5).InnerText)).ToString("dd-MM-yyyy");
                            foreach (Row row2 in sheetData2.Elements<Row>().Skip(1))//перебор таблицы Клиенты
                            {
                                if (clientCode == row2.Elements<Cell>().ElementAt(0).InnerText && clientCode != "")
                                {
                                    orgName = sharedStringTable.ElementAt(int.Parse(row2.Elements<Cell>()
                                        .ElementAt(1).InnerText)).InnerText;
                                    addres = sharedStringTable.ElementAt(int.Parse(row2.Elements<Cell>()
                                        .ElementAt(2).InnerText)).InnerText;
                                    contact = sharedStringTable.ElementAt(int.Parse(row2.Elements<Cell>()
                                        .ElementAt(3).InnerText)).InnerText;
                                }
                            }
                            if (orderExists)
                            {
                                Console.WriteLine("Наименование организации: " + $"{orgName}");
                                Console.WriteLine("Адрес: " + $"{addres}");
                                Console.WriteLine("Контактное лицо (ФИО): " + $"{contact}");
                                Console.WriteLine("Дата заказа: " + $"{dateOfOrder}");
                                Console.WriteLine("Количество: " + $"{amount}");
                                Console.WriteLine("Цена: " + $"{price}/{unit}");
                                Console.WriteLine();
                                amount = "";
                                dateOfOrder = "";
                                orgName = "";
                                addres = "";
                                contact = "";
                            }
                        }
                    }
                    if (!orderExists)
                    {
                        Console.WriteLine("Заявок нет");
                    }
                }
            }
        }
    }
}