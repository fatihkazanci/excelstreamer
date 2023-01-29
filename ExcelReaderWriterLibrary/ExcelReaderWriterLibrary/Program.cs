using ExcelStreamerLibrary;
using ExcelStreamerLibrary.Models;
using System;
using System.Collections.Generic;

namespace ExcelReaderWriterLibrary
{
    internal class Program
    {
        static void Main(string[] args)
        {
            ExampleReadExcel();
            //ExampleCreateExcel();
            Console.ReadLine();
        }

        private static void ExampleReadExcel()
        {
            string excelPath = $"{AppDomain.CurrentDomain.BaseDirectory}ExampleExcel.xlsx";
            if (!string.IsNullOrEmpty(excelPath))
            {
                using (ExcelStreamer excelStreamer = new(excelPath))
                {
                    List<ExampleExcelWorkSheetModel> exampleListNoLimit = excelStreamer.WorkSheet<ExampleExcelWorkSheetModel>("Page1");

                    List<ExampleExcelWorkSheetModel> exampleListNoLimitOnlyColumns = excelStreamer.WorkSheet<ExampleExcelWorkSheetModel>("Page1", "a");

                    List<ExampleExcelWorkSheetModel> exampleListNoLimitOnlyColumns2 = excelStreamer.WorkSheet<ExampleExcelWorkSheetModel>("Page1", nameof(ExampleExcelWorkSheetModel.Name));

                    List<ExampleExcelWorkSheetModel> exampleList = excelStreamer.WorkSheet<ExampleExcelWorkSheetModel>("Page1", 1, 4, nameof(ExampleExcelWorkSheetModel.Name), nameof(ExampleExcelWorkSheetModel.Surname));

                    ExampleExcelModel exampleLetterList = excelStreamer.WorkSheets<ExampleExcelModel>(1, 4, "a", "b");

                    ExampleExcelModel exampleLetterListNoLimit = excelStreamer.WorkSheets<ExampleExcelModel>();

                    ExcelStreamerCountResponse allSheetCount = excelStreamer.Count();

                    ExcelStreamerWorkSheetCountResponse exampleSheetCount = excelStreamer.Count("Page1");

                    ExampleExcelWorkSheetModel exampleSheetData = excelStreamer.Get<ExampleExcelWorkSheetModel>("Page1", 1, nameof(ExampleExcelWorkSheetModel.Name));

                    string exampleSheetDataName = excelStreamer.Get<ExampleExcelWorkSheetModel, string>("Page1", nameof(ExampleExcelWorkSheetModel.Name), 1);

                    string exampleSheetDataSurname = excelStreamer.Get<string>("Page1", "b", 1);

                    foreach (var item in exampleList)
                    {
                        Console.WriteLine($"{item.Name} {item.Surname}");
                    }

                    Console.WriteLine("-------------------------------");
                    Console.WriteLine($"Total Sheets => {allSheetCount.TotalSheet}");
                    Console.WriteLine("-------------------------------");
                    foreach (ExcelStreamerWorkSheetCountResponse item in allSheetCount.Sheets)
                    {
                        Console.WriteLine($"{item.SheetName.ToUpper()} COUNT");
                        Console.WriteLine("******");
                        Console.WriteLine($"Row Count => {item.RowCount}");
                        Console.WriteLine($"Column Count => {item.ColumnCount}");
                        Console.WriteLine("******");
                    }
                    Console.WriteLine("-------------------------------");

                    //exampleList[1].Name = "Osman";
                    //excelStreamer.Update(exampleList[1]);

                    //excelStreamer.UpdateWorkSheetName("Page1", "ExampleSheetName");

                    //excelStreamer.Update("Kazım", "Page1", "a", 1);
                    //excelStreamer.Update<ExampleExcelWorkSheetModel>("Kazım", nameof(ExampleExcelWorkSheetModel.Name), 1);
                    //excelStreamer.SaveChanges();
                }
            }
            else
            {
                Console.WriteLine("Excel Dosyası okunamadı");
            }
        }
        public static void ExampleCreateExcel()
        {
            using (ExcelStreamer excelStreamer = new())
            {
                string excelPath = $"{AppDomain.CurrentDomain.BaseDirectory}CreatedExampleExcel.xlsx";
                excelStreamer.CreateExcelFile(excelPath, "Page1");
                excelStreamer.Update("Name", "a", 1);
                excelStreamer.Update("Surname", "b", 1);
                excelStreamer.Update<ExampleExcelWorkSheetModel>("Fatih", nameof(ExampleExcelWorkSheetModel.Name), 2);
                excelStreamer.Update<ExampleExcelWorkSheetModel>("Kazancı", nameof(ExampleExcelWorkSheetModel.Surname), 2);
                excelStreamer.SaveChanges();

            }
        }

    }
}
