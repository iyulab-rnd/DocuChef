using Xunit.Abstractions;

namespace DocuChef.Tests;

/// <summary>
/// Base class for XLCustom test classes with common helper methods
/// </summary>
public abstract class TestBase : IDisposable
{
    protected readonly ITestOutputHelper _output;

    protected TestBase(ITestOutputHelper output)
    {
        _output = output ?? throw new ArgumentNullException(nameof(output));
    }

    public void Dispose()
    {
        GC.SuppressFinalize(this);
    }
}