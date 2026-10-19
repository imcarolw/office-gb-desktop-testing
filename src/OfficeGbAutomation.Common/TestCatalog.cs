namespace OfficeGbAutomation.Common;

public sealed class CatalogCase
{
    public string Application { get; init; } = "";
    public int CaseId { get; init; }
    public string Title { get; init; } = "";
    public string[] Aliases { get; init; } = [];
    public string Project { get; init; } = "";
    public string TestName { get; init; } = "";
    public int ExpectedSteps { get; init; }
}
