# DocuChef

The Master Chef for Document Templates - Cook delicious documents with your data and templates.

## Overview

DocuChef provides a unified interface for document generation across multiple formats. It supports Excel document generation using ClosedXML.Report.XLCustom and PowerPoint document generation using DollarSignEngine, with future plans to integrate additional template engines for Word documents.

In the spirit of its culinary name, DocuChef offers both standard API methods and fun cooking-themed extension methods that make template processing feel like preparing a delicious dish!

## Current Features

- **Excel Template Processing**: Generate Excel documents from templates using ClosedXML.Report.XLCustom
- **PowerPoint Template Processing**: Generate PowerPoint presentations from templates with embedded variables and functions
- **Flexible Variable Binding**: Add variables, complex objects, collections to your templates
- **Global Variables**: Access system information and date/time within your templates
- **Custom Function Support**: Register custom functions for Excel cell processing and PowerPoint shape processing
- **Error Handling**: Clear error reporting with specialized exception types
- **Culinary API Theme**: Optional cooking-themed extension methods for a more enjoyable API experience

## Planned Features

- Word document support
- Additional built-in functions for Excel and PowerPoint templates
- Enhanced PowerPoint chart and table functionality
- Enhanced formatting options

## Installation

```
Install-Package DocuChef
```

Or via .NET CLI:

```
dotnet add package DocuChef
```

## Quick Start

### Standard API Usage

```csharp
// Create document processor
var docuChef = new Chef();

// Load your template (Excel or PowerPoint)
var template = docuChef.LoadTemplate("template.xlsx"); // or "template.pptx"

// Add your data
template.AddVariable("Title", "Sales Report");
template.AddVariable("Products", productList);
template.AddVariable("Date", DateTime.Now);

// Generate the document
if (template is ExcelRecipe excelRecipe)
{
    var document = excelRecipe.Generate();
    document.SaveAs("result.xlsx");
}
else if (template is PowerPointRecipe pptRecipe)
{
    var document = pptRecipe.Generate();
    document.SaveAs("result.pptx");
}
```

### Fun Culinary API (Extension Methods)

```csharp
// Create your chef
var chef = new Chef();

// Load a recipe (template)
var recipe = chef.LoadRecipe("template.xlsx"); // or "template.pptx"

// Add ingredients (data)
recipe.AddIngredient("Title", "Sales Report"); 
recipe.AddIngredient("Products", productList);
recipe.AddIngredient("Date", DateTime.Now);

// Cook your dish (generate document)
if (recipe is ExcelRecipe excelRecipe)
{
    var dish = excelRecipe.Cook(); // Extension for Generate
    dish.Serve("result.xlsx"); // Extension for SaveAs
}
else if (recipe is PowerPointRecipe pptRecipe)
{
    var dish = pptRecipe.Cook();
    dish.Serve("result.pptx");
}
```

## Template Syntax

### Excel Templates

For Excel templates, DocuChef leverages ClosedXML.Report.XLCustom with its expression capabilities:

```
{{VariableName}}                     // Simple variable
{{Object.Property}}                  // Object property
{{Value:F2}}                         // Format expression
{{Value|function(param1,param2)}}    // Function expression
```

### PowerPoint Templates

For PowerPoint templates, DocuChef uses DollarSignEngine for expression evaluation and processing:

#### 1. Value Binding (Within Slide Elements)

```
${PropertyName}                // Simple property binding
${Object.PropertyName}         // Nested property binding
${Value:FormatSpecifier}       // Format specifier usage
${Condition ? Value1 : Value2} // Conditional expression
```

#### 2. Special Functions (Within Slide Elements)

```
${ppt.Image("ImageProperty")}   // Image binding
${ppt.Chart("ChartData")}       // Chart data binding (in development)
${ppt.Table("TableData")}       // Table data binding (in development)
```

#### 3. Control Directives (In Slide Notes)

```
#if: Condition, target: "ShapeName", visibleWhenFalse: "OtherShapeName"   // Conditional visibility
#foreach: CollectionName, target: "ShapeName" (in development)            // Repetition processing
#slide-foreach: CollectionName (in development)                           // Slide duplication
```

## Advanced Usage

### Custom Functions (Excel)

```csharp
// Register custom functions for Excel templates
var excelTemplate = docuChef.LoadExcelTemplate("template.xlsx");

// Function with parameters
excelTemplate.RegisterFunction("color", (cell, value, parameters) => {
    cell.SetValue(value);
    if (parameters.Length > 0)
    {
        var colorName = parameters[0];
        // Apply styling based on parameters
    }
});

// Usage in template: {{Value|color(Red)}}
```

### Custom Functions (PowerPoint)

```csharp
// Register custom functions for PowerPoint templates
var pptTemplate = docuChef.LoadPowerPointTemplate("template.pptx");

// Create a custom PowerPoint function
var customFunction = new PowerPointFunction(
    "customFunc", 
    "Custom function description",
    (context, value, parameters) => {
        // Process PowerPoint shape
        return value;
    }
);

// Register the function
pptTemplate.RegisterFunction(customFunction);

// Usage in template: ${ppt.customFunc("text")}
```

### Dynamic Data

```csharp
// Register dynamic variables that are evaluated at runtime
template.RegisterGlobalVariable("RandomId", () => Guid.NewGuid().ToString());
template.RegisterGlobalVariable("CurrentUser", () => Environment.UserName);

// In Excel template: {{RandomId}}
// In PowerPoint template: ${RandomId}
```

### Batch Processing

```csharp
// Process multiple Excel invoices
var excelTemplate = docuChef.LoadExcelTemplate("invoice.xlsx");

foreach (var invoice in invoices)
{
    excelTemplate.ClearVariables();
    excelTemplate.AddVariable("Invoice", invoice);
    var document = excelTemplate.Generate();
    document.SaveAs($"invoice_{invoice.Id}.xlsx");
}

// Process multiple PowerPoint presentations
var pptTemplate = docuChef.LoadPowerPointTemplate("presentation.pptx");

foreach (var department in departments)
{
    pptTemplate.ClearVariables();
    pptTemplate.AddVariable("Department", department);
    var document = pptTemplate.Generate();
    document.SaveAs($"presentation_{department.Id}.pptx");
}
```

## Built-in Features

### Global Variables

DocuChef automatically provides these global variables in both Excel and PowerPoint templates:

| Name | Description | Excel Example | PowerPoint Example |
|------|-------------|---------------|-------------------|
| `Today` | Current date | `{{Today:d}}` | `${Today:d}` |
| `Now` | Current date and time | `{{Now:f}}` | `${Now:f}` |
| `Year` | Current year | `{{Year}}` | `${Year}` |
| `Month` | Current month | `{{Month}}` | `${Month}` |
| `Day` | Current day | `{{Day}}` | `${Day}` |
| `MachineName` | Computer name | `{{MachineName}}` | `${MachineName}` |
| `UserName` | User name | `{{UserName}}` | `${UserName}` |
| `OSVersion` | Operating system version | `{{OSVersion}}` | `${OSVersion}` |
| `ProcessorCount` | Number of processors | `{{ProcessorCount}}` | `${ProcessorCount}` |

### PowerPoint Special Functions

DocuChef provides these special functions for PowerPoint templates:

#### Image Function
```
${ppt.Image("imageProperty")}
${ppt.Image("imagePath", width:300, height:200, preserveAspectRatio:true)}
```

#### Chart Function (In Development)
```
${ppt.Chart("salesData")}
${ppt.Chart("salesData", series:"series", categories:"categories", title:"Monthly Sales")}
```

#### Table Function (In Development)
```
${ppt.Table("employeeData")}
${ppt.Table("employeeData", headers:true, startRow:1, endRow:10)}
```

### PowerPoint Control Directives

Place these directives in slide notes to control PowerPoint slide behavior:

#### Conditional Directive
```
#if: hasData, target:"dataShape"
```

#### Repetition Directive (In Development)
```
#foreach: products, target:"productShape", maxItems:10
```

#### Slide Duplication Directive (In Development)
```
#slide-foreach: categories
```

## Culinary Extension Methods

DocuChef provides these fun cooking-themed extension methods that map to the standard API:

| Standard API | Culinary Extension |
|--------------|-------------------|
| `docuChef.LoadTemplate()` | `chef.LoadRecipe()` |
| `docuChef.LoadExcelTemplate()` | `chef.LoadExcelRecipe()` |
| `docuChef.LoadPowerPointTemplate()` | `chef.LoadPresentationRecipe()` |
| `template.AddVariable()` | `recipe.AddIngredient()` |
| `template.AddVariable(object)` | `recipe.AddIngredients()` |
| `template.ClearVariables()` | `recipe.ClearIngredients()` |
| `excelTemplate.RegisterFunction()` | `excelRecipe.RegisterTechnique()` |
| `pptTemplate.RegisterFunction()` | `pptRecipe.RegisterTechnique()` |
| `excelTemplate.Generate()` | `excelRecipe.Cook()` |
| `pptTemplate.Generate()` | `pptRecipe.Cook()` |
| `document.SaveAs()` | `dish.Serve()` |

Enabling the culinary extensions:

```csharp
using DocuChef; // Includes extension methods

var chef = new Chef();
var recipe = chef.LoadRecipe("template.xlsx");
recipe.AddIngredient("Title", "Sales Report");
var dish = recipe.Cook();
dish.Serve("result.xlsx");
```

## Error Handling

```csharp
try
{
    var document = template.Generate();
    // Or with extensions: var dish = recipe.Cook();
}
catch (DocuChefException ex)
{
    Console.WriteLine($"Template processing error: {ex.Message}");
    
    // Get more detailed information about the error
    if (ex is TemplateProcessingException)
    {
        Console.WriteLine("Error occurred during template processing");
    }
    else if (ex is InvalidTemplateFormatException)
    {
        Console.WriteLine("Invalid template format");
    }
    else if (ex is VariableOperationException)
    {
        Console.WriteLine("Error in variable operations");
    }
}
```

## Implementation Details

DocuChef supports multiple document formats with a unified API:

- **Excel (.xlsx)**: Using ClosedXML.Report.XLCustom for powerful template processing
- **PowerPoint (.pptx)**: Using DollarSignEngine for C# expression interpolation and slide processing
- **Word (.docx)**: Planned for future release

This architecture allows DocuChef to provide a consistent API while leveraging the strengths of each specialized engine.

## PowerPoint Template Processing Details

### Preparing Templates

1. **Design your presentation** in PowerPoint with placeholders for dynamic content
2. **Add shape names** to elements you want to control (select shape → right-click → name)
3. **Insert expressions** in text elements using the `${...}` syntax
4. **Add control directives** in slide notes using the `#directive: ...` syntax

### Processing Rules

- **Text Replacement**: Any `${...}` expression in text will be evaluated and replaced
- **Image Replacement**: Use `${ppt.Image("...")}` in a text or image shape
- **Shape Visibility**: Use `#if` directives in slide notes to control shape visibility
- **Chart and Table Support**: In development with initial implementations available

## Configuration Options

DocuChef provides several configuration options through the `RecipeOptions` class:

```csharp
// Create options
var options = new RecipeOptions
{
    CultureInfo = CultureInfo.GetCultureInfo("en-US"),
    EnableVerboseLogging = true,
    ThrowOnMissingVariable = false,
    MaxIterationItems = 500
};

// Excel-specific options
options.Excel.ThrowOnMissingVariable = true;
options.Excel.EnableVerboseLogging = true;

// PowerPoint-specific options
options.PowerPoint.DefaultImageWidth = 400;
options.PowerPoint.DefaultImageHeight = 300;
options.PowerPoint.PreserveImageAspectRatio = true;

// Create chef with options
var chef = new Chef(options);
```

## Utility Extensions

DocuChef includes several utility extension methods for working with files and objects:

```csharp
// Ensure a directory exists before saving
filePath.EnsureDirectoryExists();

// Get content type from file extension
string contentType = ".jpg".GetContentType();

// Get a unique file path
string uniquePath = filePath.GetUniquePath();

// Get properties of an object as a dictionary
var properties = myObject.GetProperties();
```

## Roadmap

- Complete PowerPoint chart and table processing implementation
- Improve handling of PowerPoint repetition directives
- Add support for Word document templates
- Add more built-in functions for all document types
- Enhance documentation with comprehensive examples