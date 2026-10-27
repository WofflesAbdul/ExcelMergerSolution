using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel_Handling;
using Excel_Merge.Properties;

public class FileSelectionPresenter : IFileSelectionPresenter
{
    private readonly IFileSelectionView view;
    private readonly FileSelectionModel model;
    private bool isRunning = false;
    private OperationRequested currentOperation = OperationRequested.None;
    private CancellationTokenSource progressAnimationCts;

    public FileSelectionPresenter(FileSelectionModel model, IFileSelectionView view)
    {
        this.model = model ?? throw new ArgumentNullException(nameof(model));
        this.view = view ?? throw new ArgumentNullException(nameof(view));

        model.ModelStateChanged += OnModelStateChanged;

        // GUI triggers → presenter
        view.OpenFileClicked += OnOpenFileClicked;
        view.OpenFolderClicked += OnOpenFolderClicked;
        view.TargetFileModeChanged += OnTargetFileModeChanged;
        view.ReportMenuOpening += OnReportMenuOpening;

        view.ResetRequested += OnReset;
    }

    public TargetFileMode TargetFileMode { get; private set; }

    public void SelectBaseFile()
    {
        using (OpenFileDialog dlg = new OpenFileDialog())
        {
            dlg.Filter = "Excel Files|*.xlsx;*.xls";
            dlg.Title = "Select Base Excel File";

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                model.ExistingBaseFilePath = dlg.FileName;
            }
        }
    }

    public void SelectDirectoryPath()
    {
        using (FolderBrowserDialog dlg = new FolderBrowserDialog())
        {
            dlg.Description = "Select directory for new base file";

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                model.DirectoryPath = dlg.SelectedPath;
            }
        }
    }

    public void SelectTargetFiles()
    {
        using (OpenFileDialog dlg = new OpenFileDialog())
        {
            dlg.Filter = "Excel Files|*.xlsx;*.xls";
            dlg.Title = "Select Target Excel Files";
            dlg.Multiselect = true;

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                model.AddTargetFiles(dlg.FileNames);
            }
        }
    }

    public void OnFilenameSet(string filename)
    {
        model.NewFileName = filename;
    }

    public void OnReset(object sender, EventArgs e)
    {
        model.ExistingBaseFilePath = null;
        model.NewFileName = null;
        model.DirectoryPath = null;
        model.TargetFilePaths.Clear();
    }

    public void OnTargetFileModeChanged(object sender, TargetFileMode mode)
    {
        model.ExistingBaseFilePath = null;
        model.NewFileName = null;
        model.DirectoryPath = null;
        TargetFileMode = mode;
    }

    public void OnOpenFileClicked(object sender, EventArgs e)
    {
        if (!string.IsNullOrEmpty(model.ExistingBaseFilePath) && File.Exists(model.ExistingBaseFilePath))
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = model.ExistingBaseFilePath,
                UseShellExecute = true,
            });
        }
    }

    public void OnOpenFolderClicked(object sender, EventArgs e)
    {
        string folderPath;

        if (!string.IsNullOrEmpty(model.DirectoryPath) && Directory.Exists(model.DirectoryPath))
        {
            folderPath = model.DirectoryPath;
        }
        else
        {
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = folderPath,
            UseShellExecute = true,
        });
    }

    public void OnModelStateChanged(object sender, ModelStateChangedEventArgs e)
    {
        (view as Control)?.SafeInvoke(() => UpdateView(e));
    }

    private void OnReportMenuOpening(object sender, EventArgs e)
    {
        bool hasValidBaseFile = TargetFileMode == TargetFileMode.ExistingFile
            && !string.IsNullOrWhiteSpace(model.ExistingBaseFilePath)
            && File.Exists(model.ExistingBaseFilePath);

        view.SetReportActionsEnabled(hasValidBaseFile);
    }

    public async Task RunOperationAsync(OperationRequested op, Func<Task> action)
    {
        var info = OperationRegistry.Get(op);

        // Check if another operation is running
        if (isRunning)
        {
            view.SetCompletionStatus($"{info.ErrorMessage} {OperationRegistry.Get(currentOperation).OngoingMessage} is already running.");
            return;
        }

        isRunning = true;
        currentOperation = op;

        // Display ongoing status
        view.SetOngoingStatus(info.OngoingMessage);

        // Lock UI controls
        view.LockControls(true);
        view.SetProgress(0);

        try
        {
            // Execute the actual task
            await action();

            // Success message
            view.SetCompletionStatus(info.CompletedMessage);

            if (op == OperationRequested.CreateNewFile)
            {
                // Auto-close after 10s
                _ = MessageBoxHelper.ShowMessageAsync("Success", info.CompletedMessage, 10000);
            }
            else
            {
                // Non-blocking message box for Merge and Sort
                _ = MessageBoxHelper.ShowMessageAsync("Success", info.CompletedMessage);
            }

            // Special handling after merge
            if (op == OperationRequested.Merge || op == OperationRequested.MergeAndSort)
            {
                model.ClearTargetFiles();
            }

            Settings.Default.LastSavedFile = model.ExistingBaseFilePath;
            Settings.Default.Save();
        }
        catch (Exception ex)
        {
            view.SetCompletionStatus($"{info.ErrorMessage} {ex.Message}");
            view.ShowError(info.ErrorMessage, ex);
        }
        finally
        {
            // Unlock controls and reset state
            view.LockControls(false);
            view.SetOngoingStatus("Status: Standby");
            isRunning = false;
            currentOperation = OperationRequested.None;
            _ = Task.Delay(3000).ContinueWith(_ => { view.SetProgress(0); });
        }
    }

    public Task MergeAction()
    {
        return RunOnStaThreadAsync(() =>
        {
            var merger = new ExcelMerger();

            merger.MergeFiles(baseFile: model.ExistingBaseFilePath, targets: model.TargetFilePaths.ToList(), reportProgress: percent => view.SetProgress(percent), reportStatus: message => view.SetOngoingStatus(message));
        });
    }

    public Task MergeAndSortAction()
    {
        return RunOnStaThreadAsync(() =>
        {
            var merger = new ExcelMerger();

            merger.MergeAndSortFiles(baseFile: model.ExistingBaseFilePath, targets: model.TargetFilePaths.ToList(), reportProgress: percent => view.SetProgress(percent), reportStatus: message => view.SetOngoingStatus(message));
        });
    }

    public Task SortAction()
    {
        return RunOnStaThreadAsync(() =>
        {
            var sorter = new ExcelSorter();

            sorter.SortSheets(filePath: model.ExistingBaseFilePath, reportProgress: percent => view.SetProgress(percent), reportStatus: message => view.SetOngoingStatus(message));
        });
    }

    public Task CreateNewFileAction()
    {
        return RunOnStaThreadAsync(() =>
        {
            string createdFilePath = ExcelFileCreator.CreateNewExcel(directoryPath: model.DirectoryPath, fileName: model.NewFileName);

            (view as Control)?.SafeInvoke(() => view.NewFileCreated());

            model.NewFileName = null;
            model.DirectoryPath = null;
            model.ExistingBaseFilePath = createdFilePath;
        });
    }

    public void LoadLastFile()
    {
        string lastSavedFile = Settings.Default.LastSavedFile;

        if (string.IsNullOrEmpty(lastSavedFile))
        {
            return;
        }

        switch (TargetFileMode)
        {
            case TargetFileMode.ExistingFile:
                model.ExistingBaseFilePath = lastSavedFile;
                break;

            /// <remarks> DO NOT USE, unless placeholder font is reverted to default upon use. </remarks>
            case TargetFileMode.NewFile:
                // model.NewFileName = Path.GetFileName(lastSavedFile);
                model.DirectoryPath = Path.GetDirectoryName(lastSavedFile);
                break;
        }
    }

    public Task AddCoverPageAction()
    {
        return RunOnStaThreadAsync(() =>
        {
            var templateHandler = new ExcelTemplateHandler();

            templateHandler.AddCoverPage(
                baseFile: model.ExistingBaseFilePath,
                templateFile: Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resource", "DvtReportTemplate.xltx"),
                reportProgress: percent => view.SetProgress(percent),
                reportStatus: message => view.SetOngoingStatus(message));
        });
    }

    public void DownloadTemplateFile()
    {
        using (SaveFileDialog dialog = new SaveFileDialog())
        {
            dialog.Title = "Download Report Template";
            dialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
            dialog.FileName = "DvtReportTemplate.xlsx";
            dialog.DefaultExt = "xlsx";
            dialog.AddExtension = true;
            dialog.OverwritePrompt = true;

            if (dialog.ShowDialog() != DialogResult.OK)
                return;

            string destinationPath = dialog.FileName;
            string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resource", "DvtReportTemplate.xltx");

            if (!File.Exists(templatePath))
            {
                view.ShowError("Template Not Found", new FileNotFoundException("The report template could not be found.", templatePath));
                return;
            }

            try
            {
                RunOnStaThreadAsync(() =>
                {
                    Microsoft.Office.Interop.Excel.Application excel = null;
                    Microsoft.Office.Interop.Excel.Workbook workbook = null;

                    try
                    {
                        excel = new Microsoft.Office.Interop.Excel.Application
                        {
                            Visible = false,
                            DisplayAlerts = false
                        };

                        workbook = excel.Workbooks.Open(templatePath);
                        workbook.SaveAs(destinationPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
                    }
                    finally
                    {
                        if (workbook != null)
                        {
                            workbook.Close(false);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                        }

                        if (excel != null)
                        {
                            excel.Quit();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                        }
                    }
                }).GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                view.ShowError("Download Template Failed", ex);
            }
        }
    }

    private static Task RunOnStaThreadAsync(Action action)
    {
        var completion = new TaskCompletionSource<object>();

        var thread = new Thread(() =>
        {
            try
            {
                action();
                completion.SetResult(null);
            }
            catch (Exception ex)
            {
                completion.SetException(ex);
            }
        });

        thread.IsBackground = true;
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();

        return completion.Task;
    }

    private void UpdateView(ModelStateChangedEventArgs e)
    {
        view.DisplayFileName(TargetFileMode == TargetFileMode.ExistingFile ? Path.GetFileName(e.ExistingBaseFilePath) : e.NewFileName);
        view.DisplayDirectoryPath(e.DirectoryPath);
        view.DisplayTargetFilePaths(e.TargetFilePaths.Select(p => Path.GetFileName(p)));

        view.SetMergeButtonEnabled(e.CanMerge);
        view.SetSortButtonEnabled(e.HasValidBaseFile);
        view.SetOpenFileButtonEnabled(e.HasValidBaseFile);
        view.SetOpenFolderButtonEnabled(e.HasDirectoryPath);
    }

    private string BuildExceptionMessage(Exception ex)
    {
        var messages = new List<string>();
        var current = ex;
        int level = 0;

        while (current != null)
        {
            messages.Add(
                $"Level {level}: {current.GetType().FullName}\n" +
                $"Message: {current.Message}\n" +
                $"HRESULT: 0x{current.HResult:X8}\n" +
                $"Stack Trace:\n{current.StackTrace}");

            current = current.InnerException;
            level++;
        }

        return string.Join(Environment.NewLine + Environment.NewLine, messages);
    }
}