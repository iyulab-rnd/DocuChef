using System.Globalization;
using System.Text;
using ClosedXML.Excel;
using ClosedXML.Report;
using DocuChef.Exceptions;
using DocuChef.Utils;
using Xunit.Abstractions;

namespace DocuChef.Tests;

[CollectionDefinition("Excel Tests Collection")]
public class ExcelTestsCollection : ICollectionFixture<ExcelTestsFixture>
{
}

public class ExcelTestsFixture : IDisposable
{
    public string TemplatesDir { get; }
    public string OutputDir { get; }

    public ExcelTestsFixture()
    {
        // 테스트별 고유 디렉토리 생성 (GUID 사용)
        var testRunId = Guid.NewGuid().ToString("N");
        TemplatesDir = Path.Combine(AppContext.BaseDirectory, "TestData", "Templates", testRunId);
        OutputDir = Path.Combine(AppContext.BaseDirectory, "TestOutput", testRunId);

        // 디렉토리 생성
        Directory.CreateDirectory(TemplatesDir);
        Directory.CreateDirectory(OutputDir);
    }

    public void Dispose()
    {
        // 테스트 종료 후 정리 (선택적)
        // 주석 해제하면 테스트 후 생성된 파일 삭제
        /*
        try
        {
            if (Directory.Exists(TemplatesDir))
                Directory.Delete(TemplatesDir, true);
                
            if (Directory.Exists(OutputDir))
                Directory.Delete(OutputDir, true);
        }
        catch
        {
            // 정리 중 오류 무시
        }
        */
    }
}

[Collection("Excel Tests Collection")]
public partial class ExcelTests : TestBase, IDisposable
{
    private readonly ExcelTestsFixture _fixture;
    private readonly string _templatesDir;
    private readonly string _outputDir;
    private readonly Chef _chef;

    public ExcelTests(ITestOutputHelper output, ExcelTestsFixture fixture) : base(output)
    {
        _fixture = fixture;
        _templatesDir = _fixture.TemplatesDir;
        _outputDir = _fixture.OutputDir;

        // 안전한 로깅 콜백으로 DocuChef 인스턴스 생성
        _chef = new Chef();
        _chef.SetLogCallback(CreateSafeLogCallback());

        // 각 테스트 메서드 실행 시 새 템플릿 생성
        // ExcelTests.CreateForms.cs
        CreateTestTemplates();
    }

    private LogCallback CreateSafeLogCallback()
    {
        return (level, message, ex) =>
        {
            try
            {
                var logMessage = $"[{level}] {message}";
                if (ex != null)
                    logMessage += $" Error: {ex.Message}";

                _output.WriteLine(logMessage);
            }
            catch (InvalidOperationException)
            {
                // 테스트 컨텍스트 밖에서 호출된 경우 콘솔에 출력
                Console.WriteLine($"[{level}] {message}");
            }
        };
    }

}