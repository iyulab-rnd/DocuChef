using DocuChef;
using DocuChef.PowerPoint;
using System.Diagnostics;

Console.WriteLine("DocuChef PowerPoint 템플릿 테스트 - 다중 슬라이드 및 데이터 바인딩");
Console.WriteLine("=======================================================");

// 파일 경로 설정
string basePath = AppDomain.CurrentDomain.BaseDirectory;
string templatePath = Path.Combine(basePath, "files", "ppt", "template_1.pptx");
string logoPath = Path.Combine(basePath, "files", "logo.png");
string outputPath = Path.Combine(basePath, "output_multi_slides.pptx");

// 템플릿 파일 존재 확인
if (!File.Exists(templatePath))
{
    Console.WriteLine($"템플릿 파일을 찾을 수 없습니다: {templatePath}");
    return;
}

// 로고 파일 존재 확인
if (!File.Exists(logoPath))
{
    Console.WriteLine($"로고 파일을 찾을 수 없습니다: {logoPath}");
    Console.WriteLine("계속 진행하지만 로고가 표시되지 않을 수 있습니다.");
}

Console.WriteLine($"템플릿 파일: {templatePath}");
Console.WriteLine($"로고 파일: {logoPath}");

try
{
    // Chef 인스턴스 생성
    using var chef = new Chef(new RecipeOptions()
    {
        EnableVerboseLogging = true,
        PowerPoint = new PowerPointOptions
        {
            MaxSlidesFromTemplate = 5,  // 최대 5개 슬라이드 생성 허용
            CreateNewSlidesWhenNeeded = true  // 필요시 새 슬라이드 생성
        }
    });

    // PowerPoint 템플릿 로드
    Console.WriteLine("템플릿 로드 중...");
    var recipe = chef.LoadPowerPointTemplate(templatePath);

    // 기본 변수 추가
    Console.WriteLine("변수 추가 중...");
    recipe.AddVariable("Title", "DocuChef 테스트");
    recipe.AddVariable("Subtitle", "다중 슬라이드 및 데이터 바인딩 테스트");
    recipe.AddVariable("Date", DateTime.Now);
    recipe.AddVariable("LogoPath", logoPath);
    recipe.AddVariable("CompanyName", "DocuChef 기술 연구소");

    // Items 배열 생성
    var items = new List<Item>();
    for (int i = 1; i <= 25; i++)
    {
        items.Add(new Item
        {
            Id = i,
            Name = $"상품 {i}",
            Description = $"상품 {i}에 대한 설명입니다.",
            Price = 10000 * i
        });
    }

    // Items 변수 추가
    recipe.AddVariable("Items", items);
    Console.WriteLine($"총 {items.Count}개의 상품 항목이 추가되었습니다.");

    // 문서 생성
    Console.WriteLine("문서 생성 중...");
    var document = recipe.Generate();

    // 문서 저장
    Console.WriteLine($"문서 저장 중: {outputPath}");
    document.SaveAs(outputPath);
    Console.WriteLine("문서 생성 완료!");

    // 자동으로 생성된 문서 열기
    Console.WriteLine("생성된 문서를 열고 있습니다...");
    Process.Start(new ProcessStartInfo
    {
        FileName = outputPath,
        UseShellExecute = true
    });
}
catch (Exception ex)
{
    Console.WriteLine($"오류 발생: {ex.Message}");
    Console.WriteLine(ex.StackTrace);
}

Console.WriteLine("프로그램이 완료되었습니다. 아무 키나 누르세요...");
Console.ReadKey();

// 상품 항목 클래스
public class Item
{
    public int Id { get; set; }
    public string Name { get; set; }
    public string Description { get; set; }
    public decimal Price { get; set; }
}