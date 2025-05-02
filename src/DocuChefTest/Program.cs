//using ClosedXML.Excel;
//using ClosedXML.Report;

//namespace ClosedXmlReportTest;

//class Program
//{
//    static void Main(string[] args)
//    {
//        Console.WriteLine("ClosedXML.Report Direct Test");
//        string templatesDir = Path.Combine(AppContext.BaseDirectory, "Templates");
//        string outputsDir = Path.Combine(AppContext.BaseDirectory, "Outputs");

//        // Ensure directories exist
//        Directory.CreateDirectory(templatesDir);
//        Directory.CreateDirectory(outputsDir);

//        // Create template path
//        string templatePath = Path.Combine(templatesDir, "TestTemplate.xlsx");

//        // Create test template
//        CreateTemplate(templatePath);
//        Console.WriteLine($"Created template at: {templatePath}");

//        // Prepare test data
//        var data = new
//        {
//            ReportTitle = "Product Inventory Report",
//            ReportDate = DateTime.Now,
//            Company = "Test Company Ltd.",
//            Contact = new
//            {
//                Name = "John Smith",
//                Email = "john.smith@example.com"
//            },
//            Products = new List<object>
//                {
//                    new { Id = 101, Name = "Product A", Category = "Category 1", Price = 10000, InStock = 50 },
//                    new { Id = 102, Name = "Product B", Category = "Category 1", Price = 20000, InStock = 30 },
//                    new { Id = 103, Name = "Product C", Category = "Category 2", Price = 15000, InStock = 0 },
//                    new { Id = 104, Name = "Product D", Category = "Category 2", Price = 25000, InStock = 10 },
//                    new { Id = 105, Name = "Product E", Category = "Category 3", Price = 30000, InStock = 5 }
//                }
//        };

//        // Process template with test data
//        string outputPath = Path.Combine(outputsDir, "TestOutput.xlsx");
//        ProcessTemplate(templatePath, data, outputPath);
//        Console.WriteLine($"Created output at: {outputPath}");

//        Console.WriteLine("Done! Press any key to exit...");
//        Console.ReadKey();

//        // Open both files
//        OpenExcelFile(templatePath);
//        OpenExcelFile(outputPath);
//    }

//    // Add this new method to open Excel files
//    static void OpenExcelFile(string filePath)
//    {
//        try
//        {
//            // Check if file exists
//            if (File.Exists(filePath))
//            {
//                Console.WriteLine($"Opening file: {filePath}");

//                // Use Process.Start to open the file with the default application
//                var process = new System.Diagnostics.Process();
//                process.StartInfo.FileName = filePath;
//                process.StartInfo.UseShellExecute = true;
//                process.Start();
//            }
//            else
//            {
//                Console.WriteLine($"File not found: {filePath}");
//            }
//        }
//        catch (Exception ex)
//        {
//            Console.WriteLine($"Error opening file: {ex.Message}");
//        }
//    }

//    static void CreateTemplate(string templatePath)
//    {
//        using (var workbook = new XLWorkbook())
//        {
//            var worksheet = workbook.Worksheets.Add("Report");

//            // Add header
//            worksheet.Cell("A1").Value = "{{ReportTitle}}";
//            worksheet.Range("A1:E1").Merge().Style.Font.SetBold().Font.FontSize = 16;

//            // Add company info
//            worksheet.Cell("A3").Value = "Company:";
//            worksheet.Cell("B3").Value = "{{Company}}";
//            worksheet.Cell("A4").Value = "Date:";
//            worksheet.Cell("B4").Value = "{{ReportDate}}";
//            worksheet.Cell("A5").Value = "Contact:";
//            worksheet.Cell("B5").Value = "{{Contact.Name}}";
//            worksheet.Cell("A6").Value = "Email:";
//            worksheet.Cell("B6").Value = "{{Contact.Email}}";

//            // Setup products table
//            worksheet.Cell("A8").Value = "Product ID";
//            worksheet.Cell("B8").Value = "Name";
//            worksheet.Cell("C8").Value = "Category";
//            worksheet.Cell("D8").Value = "Price";
//            worksheet.Cell("E8").Value = "In Stock";

//            // Format header
//            var headerRange = worksheet.Range("A8:E8");
//            headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
//            headerRange.Style.Font.SetBold();

//            // Add template row for products
//            worksheet.Cell("A9").Value = "{{item.Id}}";
//            worksheet.Cell("B9").Value = "{{item.Name}}";
//            worksheet.Cell("C9").Value = "{{item.Category}}";
//            worksheet.Cell("D9").Value = "{{item.Price}}";
//            worksheet.Cell("E9").Value = "{{item.InStock}}";

//            // Format price as currency
//            worksheet.Cell("D9").Style.NumberFormat.Format = "#,##0";

//            // Add service row with tags
//            worksheet.Cell("A10").Value = "<<Range Products>>";
//            worksheet.Cell("D10").Value = "<<Sum>>";
//            worksheet.Cell("E10").Value = "<<Sum>>";

//            // Define named range for products
//            var productsRange = worksheet.Range("A9:E10");
//            workbook.DefinedNames.Add("Products", productsRange);

//            // Auto-fit columns
//            worksheet.Columns().AdjustToContents();

//            // Save the template
//            workbook.SaveAs(templatePath);
//        }
//    }

//    static void ProcessTemplate(string templatePath, object data, string outputPath)
//    {
//        try
//        {
//            // Create a template instance from the template file
//            var template = new XLTemplate(templatePath);

//            // Add data variables to the template
//            template.AddVariable(data);

//            // Generate the report by processing the template
//            template.Generate();

//            // Save the result to the output path
//            template.SaveAs(outputPath);

//            Console.WriteLine("Template processed successfully!");
//        }
//        catch (Exception ex)
//        {
//            Console.WriteLine($"Error processing template: {ex.Message}");
//            Console.WriteLine(ex.StackTrace);
//        }
//    }
//}