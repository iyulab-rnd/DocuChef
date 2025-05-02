using ClosedXML.Excel;
using DocuChef;
using DocuChef.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

class Program
{
    static async Task Main(string[] args)
    {
        Console.WriteLine("DocuChef Basic Test");

        // Setup directories
        string templatesDir = Path.Combine(AppContext.BaseDirectory, "Templates");
        string outputsDir = Path.Combine(AppContext.BaseDirectory, "Outputs");

        Directory.CreateDirectory(templatesDir);
        Directory.CreateDirectory(outputsDir);

        // Create test template
        string templatePath = Path.Combine(templatesDir, "BasicTemplate.xlsx");
        CreateBasicTemplate(templatePath);
        Console.WriteLine($"Created template: {templatePath}");

        // Create test data
        var data = new
        {
            ReportTitle = "Product Inventory Report",
            Company = "Test Company Ltd.",
            ReportDate = DateTime.Now,
            Contact = new { Name = "John Smith", Email = "john@example.com" },
            Products = new List<dynamic>
            {
                new { Id = 101, Name = "Product A", Category = "Category 1", Price = 15000, InStock = 10 },
                new { Id = 102, Name = "Product B", Category = "Category 1", Price = 20000, InStock = 5 },
                new { Id = 103, Name = "Product C", Category = "Category 2", Price = 18000, InStock = 8 }
            }
        };

        // Create DocuChef and process template
        try
        {
            // Create recipe options
            var options = new RecipeOptions
            {
                Excel = new ExcelOptions { AutoFitColumns = true }
            };

            // Create DocuChef instance
            var chef = new Chef();

            // Load Excel recipe
            using var recipe = chef.LoadExcelRecipe(templatePath, options);

            // Add data
            recipe.AddData(data);

            // Save output
            string outputPath = Path.Combine(outputsDir, "BasicOutput.xlsx");
            await recipe.SaveAsync(outputPath);

            Console.WriteLine($"Document generated: {outputPath}");

            // Open output file
            OpenFile(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }

        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }

    static void CreateBasicTemplate(string templatePath)
    {
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("Report");

            // Add header
            worksheet.Cell("A1").Value = "{{ReportTitle}}";
            worksheet.Range("A1:E1").Merge().Style.Font.SetBold().Font.FontSize = 16;

            // Add company info
            worksheet.Cell("A3").Value = "Company:";
            worksheet.Cell("B3").Value = "{{Company}}";
            worksheet.Cell("A4").Value = "Date:";
            worksheet.Cell("B4").Value = "{{ReportDate}}";
            worksheet.Cell("A5").Value = "Contact:";
            worksheet.Cell("B5").Value = "{{Contact.Name}}";
            worksheet.Cell("A6").Value = "Email:";
            worksheet.Cell("B6").Value = "{{Contact.Email}}";

            // Add products table header
            worksheet.Cell("A8").Value = "ID";
            worksheet.Cell("B8").Value = "Name";
            worksheet.Cell("C8").Value = "Category";
            worksheet.Cell("D8").Value = "Price";
            worksheet.Cell("E8").Value = "Stock";

            // Format header
            var headerRange = worksheet.Range("A8:E8");
            headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
            headerRange.Style.Font.SetBold();

            // Add template row
            worksheet.Cell("A9").Value = "{{item.Id}}";
            worksheet.Cell("B9").Value = "{{item.Name}}";
            worksheet.Cell("C9").Value = "{{item.Category}}";
            worksheet.Cell("D9").Value = "{{item.Price}}";
            worksheet.Cell("E9").Value = "{{item.InStock}}";

            // Format price as currency
            worksheet.Cell("D9").Style.NumberFormat.Format = "#,##0";

            // Add ClosedXML.Report tags
            worksheet.Cell("A10").Value = "<<Range Products>>";
            worksheet.Cell("D10").Value = "<<Sum>>";
            worksheet.Cell("E10").Value = "<<Sum>>";

            // Define named range
            var productsRange = worksheet.Range("A9:E10");
            workbook.DefinedNames.Add("Products", productsRange);

            // Adjust column widths
            worksheet.Columns().AdjustToContents();

            // Save template
            workbook.SaveAs(templatePath);
        }
    }

    static void OpenFile(string filePath)
    {
        try
        {
            if (File.Exists(filePath))
            {
                var process = new System.Diagnostics.Process();
                process.StartInfo.FileName = filePath;
                process.StartInfo.UseShellExecute = true;
                process.Start();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error opening file: {ex.Message}");
        }
    }
}