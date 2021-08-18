using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;

namespace ExcelDemo
{
    class Program
    {
        static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var file = new FileInfo(@"C:\Demos\YouTubeDemo.xlsx");

            var people = GetSetupData();

            await SaveExcelFile(people, file);

            List<PersonModel> peopleFromExcel = await LoadExcelFile(file);
            foreach (var person in peopleFromExcel)
                Console.WriteLine($"{person.Id} {person.FirstName} {person.LastName}");
        }

        private static List<PersonModel> GetSetupData()
        {
            var output = new List<PersonModel>
            {
                new() {Id = 1, FirstName = "Filipe", LastName = "Silva"},
                new() {Id = 2, FirstName = "Sue", LastName = "Storm"},
                new() {Id = 3, FirstName = "Jane", LastName = "Smith"}
            };

            return output;
        }

        private static void DeleteIfExists(FileInfo file)
        {
            if (file.Exists)
                file.Delete();
        }

        private static async Task SaveExcelFile(List<PersonModel> people, FileInfo file)
        {
            DeleteIfExists(file);

            using var package = new ExcelPackage(file);

            var workSheet = package.Workbook.Worksheets.Add("MainReport");

            var range = workSheet.Cells["A2"].LoadFromCollection(people, true);
            range.AutoFitColumns();

            // Formats the header
            workSheet.Cells["A1"].Value = "Report";
            workSheet.Cells["A1:C1"].Merge = true;
            workSheet.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Row(1).Style.Font.Size = 24;
            workSheet.Row(1).Style.Font.Color.SetColor(Color.Blue);

            workSheet.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Row(2).Style.Font.Bold = true;
            workSheet.Column(3).Width = 20;

            await package.SaveAsync();
        }

        private static async Task<List<PersonModel>> LoadExcelFile(FileInfo file)
        {
            var personModel = new List<PersonModel>();

            using var package = new ExcelPackage(file);

            await package.LoadAsync(file);

            var workSheet = package.Workbook.Worksheets[0];

            int row = 3;
            int col = 1;

            while(string.IsNullOrEmpty(workSheet.Cells[row, col].Value?.ToString()) == false)
            {
                var person = new PersonModel();
                person.Id = int.Parse(workSheet.Cells[row, col].Value.ToString());
                person.FirstName = workSheet.Cells[row, col + 1].Value.ToString();
                person.LastName = workSheet.Cells[row, col + 2].Value.ToString();
                personModel.Add(person);

                row++;
            };

            return personModel;
        }
    }
}
