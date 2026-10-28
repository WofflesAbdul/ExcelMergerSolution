using System;

public static class OperationRegistry
{
    public static readonly OperationInfo Merge = new OperationInfo(
        OperationRequested.Merge,
        "Status: Merging files...",
        "Completed: Files merged successfully!",
        "Error: Failed to merge files!"
    );

    public static readonly OperationInfo Sort = new OperationInfo(
        OperationRequested.Sort,
        "Status: Sorting worksheets...",
        "Completed: Worksheets sorted successfully!",
        "Error: Failed to sort worksheets!"
    );

    public static readonly OperationInfo MergeAndSort = new OperationInfo(
        OperationRequested.MergeAndSort,
        "Status: Merging files and sorting worksheets...",
        "Completed: Files merged and worksheets sorted successfully!",
        "Error: Failed to merge files and sort worksheets!"
    );

    public static readonly OperationInfo CreateNewFile = new OperationInfo(
        OperationRequested.CreateNewFile,
        "Status: Creating a new workbook...",
        "Completed: New workbook created successfully!",
        "Error: Failed to create a new workbook!"
    );

    public static readonly OperationInfo AddCoverPage = new OperationInfo(
        OperationRequested.AddCoverPage,
        "Status: Adding cover page...",
        "Completed: Cover page added successfully!",
        "Error: Failed to add cover page!"
    );

    public static readonly OperationInfo AddSummaryPage = new OperationInfo(
        OperationRequested.AddSummaryPage,
        "Status: Adding summary page...",
        "Completed: Summary page added successfully!",
        "Error: Failed to add summary page!"
    );

    public static readonly OperationInfo DownloadTemplateFile = new OperationInfo(
        OperationRequested.DownloadTemplateFile,
        "Status: Downloading template file...",
        "Completed: Template file downloaded successfully!",
        "Error: Failed to download template file!"
    );

    public static readonly OperationInfo MergeCoverPage = new OperationInfo(
        OperationRequested.MergeCoverPage,
        "Status: Merging cover page...",
        "Completed: Cover page merged successfully!",
        "Error: Failed to merge cover page!"
    );

    public static readonly OperationInfo MergeSummaryPage = new OperationInfo(
        OperationRequested.MergeSummaryPage,
        "Status: Merging summary page...",
        "Completed: Summary page merged successfully!",
        "Error: Failed to merge summary page!"
    );

    public static readonly OperationInfo MergeReportWorkbook = new OperationInfo(
        OperationRequested.MergeReportWorkbook,
        "Status: Merging report workbook...",
        "Completed: Report workbook merged successfully!",
        "Error: Failed to merge report workbook!"
    );

    public static OperationInfo Get(OperationRequested requested)
    {
        switch (requested)
        {
            case OperationRequested.Merge: return Merge;
            case OperationRequested.Sort: return Sort;
            case OperationRequested.MergeAndSort: return MergeAndSort;
            case OperationRequested.CreateNewFile: return CreateNewFile;
            case OperationRequested.AddCoverPage: return AddCoverPage;
            case OperationRequested.AddSummaryPage: return AddSummaryPage;
            case OperationRequested.DownloadTemplateFile: return DownloadTemplateFile;
            case OperationRequested.MergeCoverPage: return MergeCoverPage;
            case OperationRequested.MergeSummaryPage: return MergeSummaryPage;
            case OperationRequested.MergeReportWorkbook: return MergeReportWorkbook;
            default: throw new ArgumentOutOfRangeException(nameof(requested), requested, null);
        }
    }
}