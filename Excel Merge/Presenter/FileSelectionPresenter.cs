using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel_Handling;

public class FileSelectionPresenter : IFileSelectionPresenter
{
    private readonly IFileSelectionView view;
    private readonly FileSelectionModel model;
    private CancellationTokenSource progressAnimationCts;

    public FileSelectionPresenter(FileSelectionModel model, IFileSelectionView view)
    {
        this.model = model ?? throw new ArgumentNullException(nameof(model));
        this.view = view ?? throw new ArgumentNullException(nameof(view));

        // GUI triggers → presenter
        view.OpenFileClicked += (s, e) => OnOpenFileClicked();
        view.OpenFolderClicked += (s, e) => OnOpenFolderClicked();
        view.InputFileModeChanged += (s, mode) => OnInputFileModeChanged();

        view.MergeRequested += (s, e) => RunMerge();
        view.SortRequested += (s, e) => RunSort();
        view.CreateNewFileRequested += (s, e) => RunCreateNewBaseFile();
        view.ResetRequested += (s, e) => OnReset();
    }

    public event EventHandler<TaskCompletedEventArgs> TaskCompleted;

    public string BaseFilePath => model.ExistingBaseFilePath;

    public string NewBaseDirectoryPath => model.DirectoryPath;

    public string NewBaseFilename => model.NewFileName;

    public IReadOnlyList<string> TargetFilePaths => model.TargetFilePaths.AsReadOnly();

    public void SelectBaseFile(string newlyCreatedBaseFile = null)
    {
        using (OpenFileDialog dlg = new OpenFileDialog())
        {
            dlg.Filter = "Excel Files|*.xlsx;*.xls";
            dlg.Title = "Select Base Excel File";

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                SetBaseFile(dlg.FileName);
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
                model.TargetFilePaths.Clear();
                model.TargetFilePaths.AddRange(dlg.FileNames);

                List<string> quotedNames = new List<string>();
                foreach (string filePath in model.TargetFilePaths)
                {
                    quotedNames.Add("\"" + Path.GetFileName(filePath) + "\"");
                }

                string displayText = string.Join(", ", quotedNames);
                view.UpdateTargetFilenames(displayText);
            }
        }
    }

    public void OnFilenameSet(string filename) => model.NewFileName = filename;

    public void OnDirectorySet(string directoryPath) => model.DirectoryPath = directoryPath;

    public void OnReset()
    {
        model.ExistingBaseFilePath = null;
        model.NewFileName = null;
        model.DirectoryPath = null;
        model.TargetFilePaths.Clear();
    }

    public void OnInputFileModeChanged()
    {
        model.ExistingBaseFilePath = null;
        model.NewFileName = null;
        model.DirectoryPath = null;
    }

    public void OnOpenFileClicked()
    {
        if (!string.IsNullOrEmpty(BaseFilePath) && File.Exists(BaseFilePath))
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = BaseFilePath,
                UseShellExecute = true,
            });
        }
    }

    public void OnOpenFolderClicked()
    {
        string folderPath;

        if (!string.IsNullOrEmpty(BaseFilePath) && File.Exists(BaseFilePath))
        {
            folderPath = Path.GetDirectoryName(BaseFilePath);
        }
        else if (!string.IsNullOrEmpty(NewBaseDirectoryPath) && Directory.Exists(NewBaseDirectoryPath))
        {
            folderPath = NewBaseDirectoryPath;
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
        view.SetMergeButtonEnabled(e.CanMerge);
        view.SetSortButtonEnabled(e.HasValidBaseFile);
        view.SetOpenFileButtonEnabled(e.HasValidBaseFile);
        view.SetOpenFolderButtonEnabled(e.HasDirectoryPath);
    }

    public void RunMerge()
    {
        view.LockControls(true);
        view.SetProgress(0);

        try
        {
            var merger = new ExcelMerger();
            merger.MergeFiles(BaseFilePath, TargetFilePaths);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Merge failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            view.LockControls(false);
            view.
        }
    }

    public void RunSort()
    {
        view.LockControls(false);
        progressAnimationCts = new CancellationTokenSource();

        try
        {
            view.SetProgress(0);
            var animationTask = AnimateProgressBarAsync(toolStripProgressBar1, 20, 200, progressAnimationCts.Token);

            var sorter = new Excel_Handling.FunctionalTestSorter();
            await Task.Run(() => sorter.SortSheets(presenter.BaseFilePath));

            toolStripProgressBar1.Value = toolStripProgressBar1.Maximum;
            presenter.NotifySortCompleted();

            progressAnimationCts.Cancel();
            await animationTask;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Sort failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            progressAnimationCts?.Cancel();
            view.LockControls(true);
            view.SetProgress(0);
        }
    }

    public void RunCreateNewBaseFile()
    {
        throw new NotImplementedException();
    }

    public void SetBaseFile(string filename)
    {
        model.ExistingBaseFilePath = filename;
        view.UpdateFilename(Path.GetFileName(filename));
        view.UpdateDirectory(Path.GetDirectoryName(filename));
    }
}
