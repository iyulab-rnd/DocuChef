# DocuChef

The Master Chef for Document Templates - Cook delicious documents with your data and templates.

## Overview

DocuChef is a powerful and flexible templating engine for document generation. With this library, you can create professional documents by simply defining templates with placeholders and binding your data. Currently, DocuChef supports Word (.docx), Excel (.xlsx), and PowerPoint (.pptx) formats.

## Installation

```bash
dotnet add package DocuChef
```

## Key Features

- **Unified API**: Consistent API across all document types for a smoother learning curve
- **Multi-Format Support**: Generate various document formats with the same engine
- **Flexible Template Syntax**: Use the optimal syntax for each document type
- **Rich Data Binding**: Bind data using dictionaries, anonymous objects, or strongly-typed models
- **Collection Rendering**: Easily populate tables, lists, and charts with collection data
- **Conditional Content**: Include or exclude content based on data values
- **Direct Document Access**: Access the underlying document models for pixel-perfect customization
- **Extensibility**: Customize every aspect of the document generation process with callbacks and plugins
- **Format Control**: Apply formatting to your data with built-in or custom format specifiers

## Core Concepts

### Common Pattern

All document templates follow a consistent workflow:

```csharp
// 1. Create the template engine
var chef = new DocuChef();

// 2. Load a template file
var recipe = chef.LoadRecipe("template.extension"); // or LoadExcelRecipe, LoadWordRecipe, LoadPowerPointRecipe

// 3. Bind data to the template
recipe.AddIngredients(data); // Why just "AddData" when you can add INGREDIENTS! 🧑‍🍳

// 4. Generate the output document
await recipe.ServeAsync("output.extension"); // Voilà! Your document is SERVED! 🍽️
```

> **Kitchen Secret**: Behind the scenes, `AddIngredients()` calls boring old `AddData()`, and `ServeAsync()` calls plain `SaveAsync()`. We just thought the culinary names made coding more delicious! 😉

### Template Options

Configure the template behavior with options:

```csharp
var cookingOptions = new RecipeOptions {
    CultureInfo = CultureInfo.GetCultureInfo("en-US"),
    NullDisplayString = "N/A",
    VariableResolver = CustomVariableResolver,
    CustomFormatters = customFormatters,
    // Document-specific options can be set via specific properties
    Excel = new ExcelOptions { /* Excel-specific settings */ },
    PowerPoint = new PowerPointOptions { /* PowerPoint-specific settings */ },
    Word = new WordOptions { /* Word-specific settings */ }
};

var chef = new DocuChef();
var recipe = chef.LoadRecipe("template.xlsx", cookingOptions);
```

### Advanced Configuration

Configure global settings for all templates:

```csharp
// Create global settings
var globalSettings = new Dictionary<string, object> {
    { "DefaultCulture", CultureInfo.GetCultureInfo("en-US") },
    { "DefaultNullDisplay", "N/A" }
};

// Create DocuChef with settings and logging
var chef = new DocuChef(globalSettings, (level, message, ex) => {
    Console.WriteLine($"[{level}] {message}");
    if (ex != null) Console.WriteLine(ex.ToString());
});

// Load multiple templates with the same configuration
var excelRecipe = chef.LoadExcelRecipe("financial.xlsx");
var wordRecipe = chef.LoadWordRecipe("report.docx");
```

## Excel Templates

Excel templates use a placeholder syntax: `{{variable}}` or `{{expression}}` in cell values.

### Basic Usage

```csharp
// Create the template engine
var chef = new DocuChef();

// Load an Excel template
var recipe = chef.LoadExcelRecipe("template.xlsx");

// Bind data to the template
var ingredients = new { 
    CustomerName = "Acme Corp", 
    OrderDate = DateTime.Now,
    TotalAmount = 1250.50
};
recipe.AddIngredients(ingredients);

// Generate the document
await recipe.ServeAsync("output.xlsx");
```

### Object Binding in Excel

```csharp
// Templates can access properties of nested objects
// Template cells containing:
// {{Company.Name}}
// {{Company.Address.Street}}, {{Company.Address.City}}
// {{Contact.Email}}

// Create a complex object with nested properties
var customerData = new {
    OrderNumber = "ORD-9876",
    Company = new {
        Name = "Global Enterprises Inc.",
        Industry = "Technology",
        Address = new {
            Street = "789 Corporate Blvd",
            City = "Enterprise City",
            State = "CA",
            ZipCode = "94105"
        }
    },
    Contact = new {
        Name = "Sarah Johnson",
        Title = "Procurement Manager",
        Email = "sjohnson@globalent.example"
    }
};

recipe.AddIngredients(customerData);
// Template cells will be populated with corresponding object properties
```

### Collection Binding in Excel

```csharp
// Define a template region for collection data
// First define a named range in Excel called "ProductItems"
// In the first row of that range, include:
// {{#each Products}} {{Name}} {{Category}} {{Price:C2}} {{Stock}} {{#if InStock}}In Stock{{else}}Out of Stock{{/if}}

// Set up collection data
var inventory = new {
    StoreLocation = "Main Warehouse",
    Products = new List<object> {
        new { 
            Name = "Deluxe Widget",
            Category = "Widgets",
            Price = 29.99,
            Stock = 42,
            InStock = true
        },
        new { 
            Name = "Premium Gadget",
            Category = "Gadgets",
            Price = 49.95,
            Stock = 13, 
            InStock = true
        },
        new { 
            Name = "Super Tool",
            Category = "Tools",
            Price = 199.50,
            Stock = 0,
            InStock = false
        }
    }
};

// Bind the data to the template
recipe.AddIngredients(inventory);
```

### Direct Excel Manipulation

```csharp
// Access the underlying workbook for custom modifications
var chef = new DocuChef();
var recipe = chef.LoadExcelRecipe("report_template.xlsx");

// Set callback to access the workbook before saving
recipe.OnWorkbookCreated = (workbook) => {
    // Get a specific worksheet
    var sheet = workbook.Worksheet("Summary");
    
    // Apply custom formatting
    sheet.Range("A1:E1").Merge().AddToNamed("Title");
    sheet.Cell("A1").Style.Font.Bold = true;
    sheet.Cell("A1").Style.Fill.BackgroundColor = XLColor.LightBlue;
    
    // Add a custom chart
    var chartData = sheet.Range("B3:C7");
    var chart = sheet.AddChart("SalesChart", XLChartType.Pie);
    chart.SetSourceData(chartData);
    
    // Set print options
    sheet.PageSetup.PaperSize = XLPaperSize.A4Paper;
    sheet.PageSetup.FitToPages(1, 1);
    sheet.PageSetup.Footer.Left.AddText("Confidential");
    sheet.PageSetup.Footer.Right.AddText(DateTime.Now.ToString("yyyy-MM-dd"));
};

// Data will be bound as usual, and then your custom modifications will be applied
recipe.AddIngredients(data);
await recipe.ServeAsync("custom_report.xlsx");
```

## Word Templates

Word templates use a simple placeholder syntax: `${variable}` or `${expression}` in text elements.

### Basic Usage

```csharp
// Create the template engine
var chef = new DocuChef();

// Load a Word template
var recipe = chef.LoadWordRecipe("template.docx");

// Bind data to the template
var ingredients = new { 
    Title = "Quarterly Report",
    ReportDate = DateTime.Now,
    Author = "Finance Team"
};
recipe.AddIngredients(ingredients);

// Generate the document
await recipe.ServeAsync("output.docx");
```

### Object Binding in Word

```csharp
// Templates can access properties of nested objects
// Template text containing:
// Project: ${Project.Name}
// Status: ${Project.Status}
// Manager: ${Project.Manager.Name}, ${Project.Manager.Department}

// Create a complex object with nested properties
var projectData = new {
    Company = "Acme Innovations",
    Project = new {
        Name = "Next Generation Platform",
        Code = "NGP-2023",
        Status = "In Progress",
        Manager = new {
            Name = "John Smith",
            Department = "R&D",
            Email = "jsmith@acme.example"
        }
    }
};

recipe.AddIngredients(projectData);
// Template text will be populated with corresponding object properties
```

### Collection Binding in Word

Word templates use special comment markers to define template regions for collections:

```
<!--#begin:TableRows:Employees-->
Name: ${Name}
Position: ${Position}
Department: ${Department}
<!--#end:TableRows:Employees-->
```

```csharp
// Set up collection data
var companyData = new {
    CompanyName = "Acme Corp",
    Employees = new List<object> {
        new { 
            Name = "John Doe",
            Position = "Software Engineer",
            Department = "Engineering"
        },
        new { 
            Name = "Jane Smith",
            Position = "Product Manager",
            Department = "Product"
        },
        new { 
            Name = "Mike Johnson",
            Position = "UX Designer",
            Department = "Design"
        }
    }
};

// Bind the data to the template
recipe.AddIngredients(companyData);
```

### Conditional Content in Word

```
<!--#if:HasFinancialData-->
Financial Summary:
Revenue: ${Revenue:C2}
Expenses: ${Expenses:C2}
Profit: ${Profit:C2}
<!--#else:HasFinancialData-->
Financial data not available for this period.
<!--#endif:HasFinancialData-->
```

### Direct Word Manipulation

```csharp
// Access the underlying document for custom modifications
var chef = new DocuChef();
var recipe = chef.LoadWordRecipe("report_template.docx");

// Set callback to access the document before saving
recipe.OnDocumentCreated = (document) => {
    // Custom document modifications using the Word API
    // Add a header
    var header = document.Sections[0].Headers.Primary;
    var paragraph = header.InsertParagraph("Confidential Report");
    paragraph.Alignment = Alignment.Right;
    
    // Add a footer with page numbers
    var footer = document.Sections[0].Footers.Primary;
    var pageNumber = footer.InsertParagraph();
    pageNumber.Alignment = Alignment.Center;
    pageNumber.AppendPageNumber(PageNumberFormat.Decimal);
};

recipe.AddIngredients(data);
await recipe.ServeAsync("enhanced_report.docx");
```

## PowerPoint Templates

PowerPoint templates use `${expression}` syntax in text elements and slide notes for collection iterations.

### Basic Usage

```csharp
// Create the template engine
var chef = new DocuChef();

// Load a PowerPoint template
var recipe = chef.LoadPowerPointRecipe("presentation.pptx");

// Bind data to the template
var ingredients = new { 
    Title = "Quarterly Report",
    ReportDate = DateTime.Now,
    Author = "Finance Team"
};
recipe.AddIngredients(ingredients);

// Generate the document
await recipe.ServeAsync("output.pptx");
```

### Creating PowerPoint Templates

#### 1. Single Variable Binding

Use `${variable}` syntax in text boxes to bind data:

```
Title: ${Title}
Date: ${ReportDate:yyyy-MM-dd}
Author: ${Author}
```

#### 2. Object Property Binding

Access properties of nested objects using dot notation:

```
Project: ${Project.Name}
Status: ${Project.Status}
Manager: ${Project.Manager.Name}, ${Project.Manager.Department}
```

```csharp
// Create a complex object with nested properties
var projectData = new {
    Company = "Acme Innovations",
    Project = new {
        Name = "Next Generation Platform",
        Code = "NGP-2023",
        Status = "In Progress",
        Manager = new {
            Name = "John Smith",
            Department = "R&D",
            Email = "jsmith@acme.example"
        }
    }
};

// Bind the complex object to the template
recipe.AddIngredients(projectData);
```

#### 3. Collection Data Binding (Slide Repetition)

To bind collection data in PowerPoint, add special directives in slide notes:

**Add directive in slide notes:**
```
@repeat: TeamMembers
```

This slide will be duplicated for each item in the `TeamMembers` collection. Within the slide, you can reference properties of the current item directly:

```
Name: ${Name}
Role: ${Role}
Experience: ${YearsExperience} years
```

```csharp
// Define collection data
var teamPresentation = new {
    ProjectTitle = "System Redesign Initiative",
    TeamMembers = new List<object> {
        new { 
            Name = "John Smith", 
            Role = "Project Lead", 
            YearsExperience = 8
        },
        new { 
            Name = "Jane Doe", 
            Role = "Backend Developer", 
            YearsExperience = 5
        },
        new { 
            Name = "Mike Johnson", 
            Role = "Frontend Developer", 
            YearsExperience = 6
        }
    }
};

// Bind collection data to the template
recipe.AddIngredients(teamPresentation);
```

#### 4. Conditional Slides (Optional Inclusion)

To conditionally include a slide, add the following directive in slide notes:

```
@if: HasFinancialData
```

This slide will only be included in the final presentation if the `HasFinancialData` property is true.

#### 5. Using Format Specifiers

You can use standard C# format specifiers to format your data:

```
Date: ${Date:yyyy-MM-dd}
Amount: ${Amount:C2}
Percentage: ${Percentage:P1}
Number: ${Value:N0}
```

### Direct PowerPoint Manipulation

```csharp
// Set callback for additional manipulation after base template processing
var chef = new DocuChef();
var recipe = chef.LoadPowerPointRecipe("presentation.pptx");

recipe.OnPresentationCreated = (presentation) => {
    // Access specific slide
    var titleSlide = presentation.Slides[0];
    
    // Add logo to title slide
    var logoShape = titleSlide.Shapes.AddPicture(
        "company_logo.png",
        LinkToFile: false,
        SaveWithDocument: true,
        Left: 500,
        Top: 50,
        Width: 100,
        Height: 50
    );
    
    // Set presentation properties
    presentation.BuiltInDocumentProperties.Title = "Quarterly Report";
    presentation.BuiltInDocumentProperties.Subject = "Financial Analysis";
    presentation.BuiltInDocumentProperties.Company = "Acme Corp";
};

recipe.AddIngredients(data);
await recipe.ServeAsync("enhanced_presentation.pptx");
```

## Advanced Features

### Custom Variable Resolvers

```csharp
// Create custom resolver for special variables
var cookingOptions = new RecipeOptions {
    VariableResolver = (expression, parameter) => {
        if (expression == "CurrentUser")
            return Environment.UserName;
        if (expression == "ServerEnvironment")
            return GetEnvironmentName();
        if (expression == "BuildVersion")
            return Assembly.GetExecutingAssembly().GetName().Version.ToString();
        
        return null; // Fall back to standard resolution
    }
};

var chef = new DocuChef();
var recipe = chef.LoadRecipe("template.xlsx", cookingOptions);
```

### Custom Formatting Handlers

```csharp
// Register custom formatters for special data representation
var cookingOptions = new RecipeOptions {
    CustomFormatters = new Dictionary<string, Func<object, string, string>> {
        ["phone"] = (value, format) => {
            // Format phone numbers based on region code
            var number = value.ToString();
            if (format == "US")
                return $"({number.Substring(0, 3)}) {number.Substring(3, 3)}-{number.Substring(6)}";
            if (format == "UK")
                return $"+44 {number.Substring(0, 4)} {number.Substring(4, 6)} {number.Substring(10)}";
            return number;
        },
        ["highlight"] = (value, format) => {
            // Marker to be interpreted as special command in Excel
            return $"__HIGHLIGHT_{format?.ToUpper()}_{value}__";
        }
    }
};

var chef = new DocuChef();
var recipe = chef.LoadRecipe("template.xlsx", cookingOptions);
```

### Document-Specific Rendering Plugins

```csharp
// Excel-specific plugin for custom chart generation
var cookingOptions = new RecipeOptions {
    Excel = new ExcelOptions {
        Plugins = new List<IExcelPlugin> {
            new ChartGenerationPlugin(),
            new ConditionalFormattingPlugin()
        }
    }
};

var chef = new DocuChef();
var recipe = chef.LoadExcelRecipe("template.xlsx", cookingOptions);
```

### Batch Document Generation

```csharp
// Generate multiple documents from the same template
var chef = new DocuChef();
var recipe = chef.LoadRecipe("employee_profile.pptx");

// List of employees to process
var employees = GetEmployeesFromDatabase(); // Returns list of employee objects

// Process each employee
foreach (var employee in employees)
{
    // Bind current employee data
    recipe.AddIngredients(employee);
    
    // Save with unique filename
    await recipe.ServeAsync($"Profile_{employee.Id}.pptx");
    
    // Reset template for next employee
    recipe.Reset();
}
```

### Combining Multiple Templates

```csharp
// Create a DocuChef instance
var chef = new DocuChef();

// Load multiple templates
var reportRecipe = chef.LoadExcelRecipe("financial_report.xlsx");
var presentationRecipe = chef.LoadPowerPointRecipe("report_presentation.pptx");

// Process both with the same data
var quarterlyData = GetQuarterlyData();
reportRecipe.AddIngredients(quarterlyData);
presentationRecipe.AddIngredients(quarterlyData);

// Save both outputs
await reportRecipe.ServeAsync("Q2_Financial_Report.xlsx");
await presentationRecipe.ServeAsync("Q2_Financial_Presentation.pptx");
```

## Implementation Details

DocuChef is designed with a pluggable architecture that allows for flexible implementation of document processing engines. The library is structured to support different document formats through specialized handlers, with a unified API layer that provides a consistent experience regardless of the underlying implementation.

## Compatibility

- .NET 8.0+