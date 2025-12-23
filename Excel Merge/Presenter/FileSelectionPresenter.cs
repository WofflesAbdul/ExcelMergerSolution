using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel_Handling;

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

            // Special handling after merge
            if (op == OperationRequested.Merge)
            {
                model.ClearTargetFiles();
            }
        }
        catch (Exception ex)
        {
            view.SetCompletionStatus($"{info.ErrorMessage} {ex.Message}");
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
        return Task.Run(() =>
        {
            var merger = new ExcelMerger();

            // Pass a lambda to report progress
            merger.MergeFiles(
                model.ExistingBaseFilePath,
                model.TargetFilePaths,
                percent => view.SetProgress(percent)
            );

            ExcelFileCreator.RemovePlaceholderSheets(model.ExistingBaseFilePath);
        });
    }

    public async Task SortAction()
    {
        progressAnimationCts = new CancellationTokenSource();
        var animationTask = view.AnimateProgressBarAsync(20, 200, progressAnimationCts.Token);

        var sorterTask = Task.Run(() =>
        {
            var sorter = new FunctionalTestSorter();
            sorter.SortSheets(model.ExistingBaseFilePath);
        });

        await Task.WhenAll(sorterTask, animationTask);
        view.SetProgress(100);
        progressAnimationCts.Cancel();
    }

    public Task CreateNewFileAction()
    {
        return Task.Run(() =>
        {
            string createdFilePath = ExcelFileCreator.CreateNewExcel(directoryPath: model.DirectoryPath, fileName: model.NewFileName);

            // Marshal UI updates to the UI thread safely
            (view as Control)?.SafeInvoke(() => view.NewFileCreated());

            // Reset new file info
            model.NewFileName = null;
            model.DirectoryPath = null;
            model.ExistingBaseFilePath = createdFilePath;
        });
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
}
