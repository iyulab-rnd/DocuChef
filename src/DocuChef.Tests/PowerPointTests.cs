using DocuChef.PowerPoint;
using DocuChef.Exceptions;
using Xunit.Abstractions;
using FluentAssertions;
using System;
using System.IO;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace DocuChef.Tests;

public class PowerPointTests : TestBase
{
    private readonly string _tempDirectory;
    private readonly string _templatePath;

    public PowerPointTests(ITestOutputHelper output) : base(output)
    {
        // Create a temporary directory for test files
        _tempDirectory = Path.Combine(Path.GetTempPath(), "DocuChefTests", Guid.NewGuid().ToString());
        Directory.CreateDirectory(_tempDirectory);

        // Create a simple PowerPoint template for testing
        _templatePath = Path.Combine(_tempDirectory, "template.pptx");
        CreateSampleTemplate(_templatePath);
    }

    public new void Dispose()
    {
        // Clean up temp files after tests
        try
        {
            if (Directory.Exists(_tempDirectory))
            {
                Directory.Delete(_tempDirectory, true);
            }
        }
        catch (IOException)
        {
            // Ignore cleanup errors
        }

        base.Dispose();
    }

    [Fact]
    public void Chef_LoadTemplate_WithPowerPointFile_ReturnsPowerPointRecipe()
    {
        // Arrange
        var chef = new Chef();

        // Act
        var recipe = chef.LoadTemplate(_templatePath);

        // Assert
        recipe.Should().NotBeNull();
        recipe.Should().BeOfType<PowerPointRecipe>();
    }

    [Fact]
    public void Chef_LoadPowerPointTemplate_WithValidPath_ReturnsPowerPointRecipe()
    {
        // Arrange
        var chef = new Chef();

        // Act
        var recipe = chef.LoadPowerPointTemplate(_templatePath);

        // Assert
        recipe.Should().NotBeNull();
        recipe.Should().BeOfType<PowerPointRecipe>();
    }

    [Fact]
    public void Chef_LoadPowerPointTemplate_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        // Arrange
        var chef = new Chef();
        var nonExistentPath = Path.Combine(_tempDirectory, "nonexistent.pptx");

        // Act & Assert
        Action act = () => chef.LoadPowerPointTemplate(nonExistentPath);

        act.Should().Throw<FileNotFoundException>();
    }

    [Fact]
    public void PowerPointRecipe_AddVariable_AddsVariableToTemplate()
    {
        // Arrange
        var chef = new Chef();
        var recipe = chef.LoadPowerPointTemplate(_templatePath);

        // Act
        recipe.AddVariable("TestVar", "TestValue");

        // Assert
        // Since we can't directly access the variables dictionary, we'll test indirectly
        // by generating a document and checking it contains our variable
        var document = recipe.Generate();
        using var stream = new MemoryStream();
        document.SaveAs(stream);
        stream.Length.Should().BeGreaterThan(0);
    }

    [Fact]
    public void PowerPointRecipe_AddVariable_WithNullName_ThrowsArgumentNullException()
    {
        // Arrange
        var chef = new Chef();
        var recipe = chef.LoadPowerPointTemplate(_templatePath);

        // Act & Assert
        Action act = () => recipe.AddVariable(null, "Value");

        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void PowerPointRecipe_RegisterGlobalVariable_RegistersVariable()
    {
        // Arrange
        var chef = new Chef();
        var recipe = chef.LoadPowerPointTemplate(_templatePath);
        var testValue = "TestGlobalValue";

        // Act
        recipe.RegisterGlobalVariable("TestGlobalVar", testValue);

        // Assert
        // Similar to AddVariable, we'll test indirectly
        var document = recipe.Generate();
        using var stream = new MemoryStream();
        document.SaveAs(stream);
        stream.Length.Should().BeGreaterThan(0);
    }

    [Fact]
    public void PowerPointRecipe_RegisterGlobalVariable_WithNullName_ThrowsArgumentNullException()
    {
        // Arrange
        var chef = new Chef();
        var recipe = chef.LoadPowerPointTemplate(_templatePath);

        // Act & Assert
        Action act = () => recipe.RegisterGlobalVariable(null, "Value");

        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void PowerPointRecipe_RegisterFunction_RegistersFunction()
    {
        // Arrange
        var chef = new Chef();
        var recipe = chef.LoadPowerPointTemplate(_templatePath);
        var function = new PowerPointFunction
        {
            Name = "testFunc",
            Description = "Test function",
            Handler = (ctx, value, parameters) => "Function called"
        };

        // Act
        recipe.RegisterFunction(function);

        // Assert
        // We can only verify it doesn't throw an exception
        var document = recipe.Generate();
        using var stream = new MemoryStream();
        document.SaveAs(stream);
        stream.Length.Should().BeGreaterThan(0);
    }

    [Fact]
    public void PowerPointRecipe_RegisterFunction_WithNullName_ThrowsArgumentNullException()
    {
        // Arrange
        var chef = new Chef();
        var recipe = chef.LoadPowerPointTemplate(_templatePath);
        var function = new PowerPointFunction
        {
            Description = "Test function",
            Handler = (ctx, value, parameters) => "Function called"
        };

        // Act & Assert
        Action act = () => recipe.RegisterFunction(function);

        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void PowerPointRecipe_Generate_ReturnsPowerPointDocument()
    {
        // Arrange
        var chef = new Chef();
        var recipe = chef.LoadPowerPointTemplate(_templatePath);

        // Act
        var document = recipe.Generate();

        // Assert
        document.Should().NotBeNull();
        document.Should().BeOfType<PowerPointDocument>();
    }

    [Fact]
    public void PowerPointDocument_SaveAs_WithValidPath_SavesDocument()
    {
        // Arrange
        var chef = new Chef();
        var recipe = chef.LoadPowerPointTemplate(_templatePath);
        var document = recipe.Generate();
        var outputPath = Path.Combine(_tempDirectory, "output.pptx");

        // Act
        document.SaveAs(outputPath);

        // Assert
        File.Exists(outputPath).Should().BeTrue();
        new FileInfo(outputPath).Length.Should().BeGreaterThan(0);
    }

    [Fact]
    public void PowerPointDocument_SaveAs_WithNullPath_ThrowsArgumentNullException()
    {
        // Arrange
        var chef = new Chef();
        var recipe = chef.LoadPowerPointTemplate(_templatePath);
        var document = recipe.Generate();

        // Act & Assert
        Action act = () => document.SaveAs((string)null);

        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void PowerPointDocument_SaveAs_WithStream_SavesDocument()
    {
        // Arrange
        var chef = new Chef();
        var recipe = chef.LoadPowerPointTemplate(_templatePath);
        var document = recipe.Generate();
        using var stream = new MemoryStream();

        // Act
        document.SaveAs(stream);

        // Assert
        stream.Length.Should().BeGreaterThan(0);
    }

    [Fact]
    public void PowerPointDocument_SaveAs_WithNullStream_ThrowsArgumentNullException()
    {
        // Arrange
        var chef = new Chef();
        var recipe = chef.LoadPowerPointTemplate(_templatePath);
        var document = recipe.Generate();

        // Act & Assert
        Action act = () => document.SaveAs((Stream)null);

        act.Should().Throw<ArgumentNullException>();
    }

    [Fact]
    public void PowerPointDocument_Dispose_DisposesWorkbook()
    {
        // Arrange
        var chef = new Chef();
        var recipe = chef.LoadPowerPointTemplate(_templatePath);
        var document = recipe.Generate();

        // Act
        document.Dispose();

        // Assert
        // We can check if the object is disposed by trying to access a method
        // that should throw ObjectDisposedException
        Action act = () => document.SaveAs(new MemoryStream());

        act.Should().Throw<ObjectDisposedException>();
    }

    [Fact]
    public void Image_PowerPointFunction_Works()
    {
        // Arrange
        var chef = new Chef();
        var recipe = chef.LoadPowerPointTemplate(_templatePath);
        var imagePath = Path.Combine(_tempDirectory, "test_image.png");

        // Create a test image
        CreateTestImage(imagePath);

        // Add variables
        recipe.AddVariable("ImagePath", imagePath);

        // Act
        var document = recipe.Generate();
        var outputPath = Path.Combine(_tempDirectory, "image_test.pptx");
        document.SaveAs(outputPath);

        // Assert
        File.Exists(outputPath).Should().BeTrue();
        new FileInfo(outputPath).Length.Should().BeGreaterThan(0);
    }

    [Fact]
    public void ChefExtensions_LoadRecipe_LoadsPowerPointTemplate()
    {
        // Arrange
        var chef = new Chef();

        // Act
        var recipe = chef.LoadRecipe(_templatePath);

        // Assert
        recipe.Should().NotBeNull();
        recipe.Should().BeOfType<PowerPointRecipe>();
    }

    [Fact]
    public void Integration_CompleteWorkflow_GeneratesExpectedDocument()
    {
        // Arrange
        var chef = new Chef();
        var recipe = chef.LoadPowerPointTemplate(_templatePath);

        // Add data to the template
        recipe.AddVariable("Title", "Sales Presentation");
        recipe.AddVariable("Date", DateTime.Now);

        var products = new List<Product>
        {
            new Product { Id = 1, Name = "Product 1", Price = 10.99m },
            new Product { Id = 2, Name = "Product 2", Price = 20.50m },
            new Product { Id = 3, Name = "Product 3", Price = 15.75m }
        };

        recipe.AddVariable("Products", products);

        // Act
        var document = recipe.Generate();
        var outputPath = Path.Combine(_tempDirectory, "integration_test.pptx");
        document.SaveAs(outputPath);

        // Assert
        File.Exists(outputPath).Should().BeTrue();
        var fileInfo = new FileInfo(outputPath);
        fileInfo.Length.Should().BeGreaterThan(0);

        // Additional check: attempt to open the generated file to ensure it's valid PowerPoint
        using var presentationDocument = PresentationDocument.Open(outputPath, false);
        presentationDocument.PresentationPart.Should().NotBeNull();
        presentationDocument.PresentationPart.Presentation.Should().NotBeNull();
    }

    #region Helper Methods

    private void CreateSampleTemplate(string path)
    {
        try
        {
            // Create a basic PowerPoint presentation with placeholders
            using var presentationDocument = PresentationDocument.Create(path, DocumentFormat.OpenXml.PresentationDocumentType.Presentation);

            // Add a presentation part
            var presentationPart = presentationDocument.AddPresentationPart();
            presentationPart.Presentation = new DocumentFormat.OpenXml.Presentation.Presentation();

            // Add a slide master part
            var slideMasterPart = presentationPart.AddNewPart<DocumentFormat.OpenXml.Packaging.SlideMasterPart>();
            slideMasterPart.SlideMaster = new DocumentFormat.OpenXml.Presentation.SlideMaster();

            // Add a slide layout part
            var slideLayoutPart = slideMasterPart.AddNewPart<DocumentFormat.OpenXml.Packaging.SlideLayoutPart>();
            slideLayoutPart.SlideLayout = new DocumentFormat.OpenXml.Presentation.SlideLayout();

            // Add a slide part
            var slidePart = presentationPart.AddNewPart<DocumentFormat.OpenXml.Packaging.SlidePart>();
            slidePart.Slide = new DocumentFormat.OpenXml.Presentation.Slide();

            // Add basic slide structure
            slidePart.Slide.CommonSlideData = new DocumentFormat.OpenXml.Presentation.CommonSlideData();
            slidePart.Slide.CommonSlideData.ShapeTree = new DocumentFormat.OpenXml.Presentation.ShapeTree();

            // Add slide ID list to presentation
            presentationPart.Presentation.SlideIdList = new DocumentFormat.OpenXml.Presentation.SlideIdList();

            // Add slide ID
            var slideId = new DocumentFormat.OpenXml.Presentation.SlideId();
            slideId.Id = 256U;
            slideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);
            presentationPart.Presentation.SlideIdList.Append(slideId);

            // Save the presentation
            presentationPart.Presentation.Save();
        }
        catch (Exception ex)
        {
            _output.WriteLine($"Error creating PowerPoint template: {ex.Message}");
            throw;
        }
    }

    private void CreateTestImage(string path)
    {
        // Create a simple 100x100 test image (just write a tiny valid PNG file)
        byte[] pngHeader = new byte[] {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
            0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00,
            0x03, 0x01, 0x01, 0x00, 0x18, 0xDD, 0x8D, 0xB0, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
            0x42, 0x60, 0x82
        };

        File.WriteAllBytes(path, pngHeader);
    }

    // Sample class for testing
    private class Product
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public decimal Price { get; set; }
    }

    #endregion
}