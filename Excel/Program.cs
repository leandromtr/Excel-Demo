using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;


namespace ExcelDemo
{
    partial class Program
    {
        static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var path = @"C:\ExcelDemo\";
            bool exists = System.IO.Directory.Exists(path);
            if (!exists)
                System.IO.Directory.CreateDirectory(path);

            var file = new FileInfo(path + @"\ExcelFile.xlsx");


            var people = GetSetupData();


            await SaveExcelFile(people, file);

            List<PersonModel> personFromExcel = await LoadExcelFile(file);

            foreach (var person in personFromExcel)
            {
                Console.WriteLine($"{person.Id} {person.FistName} {person.LastName}");
            }
        }

        private static async Task<List<PersonModel>> LoadExcelFile(FileInfo file)
        {
            List<PersonModel> output = new();

            using var package = new ExcelPackage(file);
            await package.LoadAsync(file);
            var workSheet = package.Workbook.Worksheets[0];

            int row = 3;
            int col = 1;

            while (string.IsNullOrWhiteSpace(workSheet.Cells[row, col].Value?.ToString()) == false)
            {
                PersonModel person = new();
                person.Id = int.Parse(workSheet.Cells[row, col].Value.ToString());
                person.FistName = workSheet.Cells[row, col + 1].Value.ToString();
                person.LastName = workSheet.Cells[row, col + 2].Value.ToString();
                output.Add(person);
                row += 1;
            }

            return output;
        }

        private static async Task SaveExcelFile(List<PersonModel> people, FileInfo file)
        {
            DeleteIfExists(file);

            using var package = new ExcelPackage(file);
            var workSheet = package.Workbook.Worksheets.Add("MainReport");

            var range = workSheet.Cells["A2"].LoadFromCollection(people, true);
            range.AutoFitColumns();

            // Format the Header
            workSheet.Cells["A1"].Value = "Excel Report";
            workSheet.Cells["A1:C1"].Merge = true;
            workSheet.Column(1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            workSheet.Row(1).Style.Font.Size = 24;
            workSheet.Row(1).Style.Font.Color.SetColor(Color.Blue);

            workSheet.Column(2).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            workSheet.Row(2).Style.Font.Bold = true;

            workSheet.Column(3).Width = 20;

            await package.SaveAsync();
        }

        private static void DeleteIfExists(FileInfo file)
        {
            if (file.Exists)
            {
                file.Delete();
            }
        }

        private static List<PersonModel> GetSetupData()
        {
            List<PersonModel> output = new()
            {
                new() { Id = 1, FistName = "Leandro", LastName = "Reis" },
                new() { Id = 2, FistName = "Lucas", LastName = "Souza" },
                new() { Id = 3, FistName = "Teteus", LastName = "Silva" }
            };
            return output;
        }
    }
}