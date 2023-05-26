using System;
using System.IO;
using OfficeOpenXml;

class Program
{
    static void Main()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        string filePath = "C:\\Users\\roman\\source\\repos\\Kompik\\Kompik\\Res.xlsx";
        FileInfo file = new FileInfo(filePath);
        
        using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Result");
            // Поля
            worksheet.Cells[1, 1].Value = "Версия ОС";
            worksheet.Cells[1, 2].Value = "Имя ПК";
            worksheet.Cells[1, 3].Value = "Тип процессора";
            worksheet.Cells[1, 4].Value = "Объем ОЗУ";
            worksheet.Cells[1, 5].Value = "Разрешение экрана";

            // ДАнные о пК
            string osVersion = Environment.OSVersion.Version.ToString();
            string pcName = Environment.MachineName;
            string processorType = Environment.GetEnvironmentVariable("PROCESSOR_IDENTIFIER");
            string ram = (Environment.WorkingSet / (1024)).ToString(); // Объем ОЗУ 
            string screenResolution = $"{Console.WindowWidth}x{Console.WindowHeight}";

          
            worksheet.Cells[2, 1].Value = osVersion;
            worksheet.Cells[2, 2].Value = pcName;
            worksheet.Cells[2, 3].Value = processorType;
            worksheet.Cells[2, 4].Value = ram;
            worksheet.Cells[2, 5].Value = screenResolution;

            worksheet.Cells.AutoFitColumns();
            package.Save();

            Console.WriteLine($"Файл Excel успешно создан: {filePath}");
        }
    }
}
