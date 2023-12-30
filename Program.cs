using System;

namespace PracticTask3
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            string action = "";
            DataProccesor dataProccesor;
            Console.WriteLine("Пожайлуйста введите путь к файлу");
            Console.WriteLine(@"Пример: C:\Users\Имя пользователя\Documents\Data.xlsx");
            Console.WriteLine("Для поиска в текущей папке введите только название файла");
            string filepath = Console.ReadLine();
            if (!filepath.Contains(":"))
            {
                dataProccesor = new DataProccesor($"{Environment.CurrentDirectory}" + "\\" + filepath);
                //dataProccesor = new DataProccesor($"{Environment.CurrentDirectory}" + "\\" + "Data.xlsx");
            }
            else
            {
                dataProccesor = new DataProccesor(filepath);
            }
            while (true)//Основной цикл программы
            {
                Console.WriteLine("1 - Получить информацию о клиентах");
                Console.WriteLine("2 - Изменение контактного лица клиента");
                Console.WriteLine("3 - Определение золотого клиента");
                Console.WriteLine("4 - Выход");
                action = Console.ReadLine();
                switch (action)
                {
                    case "1":
                        Info(dataProccesor);
                        break;

                    case "2":
                        Change(dataProccesor);
                        break;

                    case "3":
                        Golden(dataProccesor);
                        break;

                    case "4":
                        Environment.Exit(0);
                        break;
                }
            }
        }

        public static void Info(DataProccesor dataProccesor)
        {
            Console.Clear();
            dataProccesor.GetGoods();
            Console.WriteLine("Введите название товара для получения информации о клиенте:");
            string item = Console.ReadLine();
            Console.WriteLine();
            dataProccesor.GetInfo(item);
            Console.WriteLine("Для продолжения нажмите любую клавишу");
            Console.ReadKey();
            Console.Clear();
        }

        private static void Change(DataProccesor dataProccesor)
        {
            Console.Clear();
            dataProccesor.GetCustomers();
            Console.WriteLine("Введите название организации для смены контактного лица:");
            string item = Console.ReadLine();
            Console.WriteLine();
            dataProccesor.ChangeContact(item);
            Console.WriteLine("\nДля продолжения нажмите любую клавишу");
            Console.ReadKey();
            Console.Clear();
        }

        private static void Golden(DataProccesor dataProccesor)
        {
            Console.Clear();
            Console.Write("Какой год вас интересует? (пример: 2023): ");
            string year = Console.ReadLine();
            Console.Write("Если вас интересует информация за весь год оставьте поле месяц пустым (нажмите Enter)");
            Console.Write("\nКакой месяц вас интересует? (пример: 04): ");
            string month = Console.ReadLine();
            dataProccesor.GetGoldenClient(month, year);
            Console.WriteLine("\nДля продолжения нажмите любую клавишу");
            Console.ReadKey();
            Console.Clear();
        }
    }
}