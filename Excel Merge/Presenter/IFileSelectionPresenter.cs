using System;
using System.Threading.Tasks;

public interface IFileSelectionPresenter
{
    TargetFileMode TargetFileMode { get; }

    void SelectBaseFile();

    void SelectDirectoryPath();

    void SelectTargetFiles();

    void OnFilenameSet(string filename);

    void OnReset(object sender, EventArgs e);

    void OnTargetFileModeChanged(object sender, TargetFileMode mode);

    void OnOpenFileClicked(object sender, EventArgs e);

    void OnOpenFolderClicked(object sender, EventArgs e);

    void OnModelStateChanged(object sender, ModelStateChangedEventArgs e);

    Task RunOperationAsync(OperationRequested op, Func<Task> action);

    Task MergeAction();

    Task SortAction();

    Task MergeAndSortAction();

    Task CreateNewFileAction();

    void LoadLastFile();

    Task AddCoverPageAction();

    Task AddSummaryPageAction();

    string SelectTemplateDestination();

    Task DownloadTemplateFileAsync(string destinationPath);

    string SelectSourceWorkbook();

    Task MergeCoverPageAction(string sourceFile);

    Task MergeSummaryPageAction(string sourceFile);

    Task MergeReportWorkbookAction(string sourceFile);
}
