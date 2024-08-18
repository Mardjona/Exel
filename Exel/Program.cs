using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeOpenXml;

public class Event
{
    public string Name { get; set; }
    public int EventNumber { get; set; }
    public string Latitude { get; set; }
    public string Longitude { get; set; }
    public string Height { get; set; }
}

public partial class ExcelReader
{
    public static List<Event> ReadEventsFromExcel(string filePath, int startRow)
    {
        // Проверяем существование файла
        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException("Файл не найден", filePath);
        }

        var eventsList = new List<Event>();

        // Используем EPPlus для чтения данных из Excel
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0]; // читаем первый лист

            // Проверка на пустой лист
            if (worksheet.Dimension == null)
            {
                return eventsList; // возвращаем пустой список, если лист пуст
            }

            int rowCount = worksheet.Dimension.Rows;

            for (int row = startRow; row <= rowCount; row++) // начинаем с заданной строки
            {
                var eventName = worksheet.Cells[row, 1].Text; // столбец Name

                // Извлекаем числовую часть (номер события) из имени события с помощью регулярного выражения
                var match = MyRegex().Match(eventName); // ищем первую последовательность цифр
                var eventNumber = match.Success ? int.Parse(match.Value) : 0;

                var latitude = worksheet.Cells[row, 12].Text; // столбец Latitude
                var longitude = worksheet.Cells[row, 16].Text;// столбец Longitude
                var height = worksheet.Cells[row, 19].Text; // столбец Height

                // Добавляем событие в список
                eventsList.Add(new Event
                {
                    Name = eventName,
                    EventNumber = eventNumber,
                    Latitude = latitude,
                    Longitude = longitude,
                    Height = height
                });
            }
        }

        return eventsList;
    }

    [GeneratedRegex(@"\d+")]
    private static partial Regex MyRegex();
}

 class Program
{
    static void Main(string[] args)
    {
        try
        {
            //  номер строки, с которой начинаются данные
            var events = ExcelReader.ReadEventsFromExcel("/Users/admin/Downloads/Topograph15_12_23_EventKinematic_23.12.20_111111.xlsx", 23);

            foreach (var e in events)
            {
                Console.WriteLine($"{e.Name}, EventNumber: {e.EventNumber}, Latitude: {e.Latitude}, Longitude: {e.Longitude}, Height: {e.Height}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Произошла ошибка: {ex.Message}");
        }
    }
}


