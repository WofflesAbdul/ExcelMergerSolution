using Excel_Handling;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

public class FileSelectionPresenter : IFileSelectionPresenter
{
    private readonly IFileSelectionView view;
    private readonly FileSelectionModel model;
    private bool isRunning;

    public FileSelectionPresenter(FileSelectionModel model, IFileSelectionView view)
    {
        this.model = model ?? throw new ArgumentNullException(nameof(model));
        this.view = view ?? throw new ArgumentNullException(nameof(view));

        model.ModelStateChanged += OnModelStateChanged;

        // GUI triggers → presenter
        view.OpenFileClicked += OnOpenFileClicked;
        view.OpenFolderClicked += OnOpenFolderClicked;
        view.InputFileModeChanged += OnInputFileModeChanged;

        view.MergeRequested += async (s, e) => await RunMergeAsync(s, e);
        view.SortRequested += async (s, e) => await RunSortAsync(s, e);
        view.CreateNewFileRequested += async (s, e) => await RunCreateNewBaseFileAsync(s, e);
        view.ResetRequested += OnReset;
    }

    public InputFileMode InputFileMode { get; private set; }

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

    public void OnInputFileModeChanged(object sender, InputFileMode mode)
    {
        model.ExistingBaseFilePath = null;
        model.NewFileName = null;
        model.DirectoryPath = null;
        InputFileMode = mode;
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
        view.DisplayFileName(InputFileMode == InputFileMode.ExistingFile ? Path.GetFileName(e.ExistingBaseFilePath) : e.NewFileName);
        view.DisplayDirectoryPath(e.DirectoryPath);
        view.DisplayTargetFilePaths(e.TargetFilePaths.Select(p => Path.GetFileName(p)));

        view.SetMergeButtonEnabled(e.CanMerge);
        view.SetSortButtonEnabled(e.HasValidBaseFile);
        view.SetOpenFileButtonEnabled(e.HasValidBaseFile);
        view.SetOpenFolderButtonEnabled(e.HasDirectoryPath);
    }

    public async Task RunMergeAsync(object sender, EventArgs e)
    {
        if (CheckRunning()) return;

        isRunning = true;
        view.LockControls(true);
        view.SetProgress(0);

        try
        {
            await Task.Run(() =>
            {
                var merger = new ExcelMerger();
                merger.MergeFiles(model.ExistingBaseFilePath, model.TargetFilePaths);
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Merge failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            view.LockControls(false);
            model.ClearTargetFiles();
        }
    }

    public async Task RunSortAsync(object sender, EventArgs e)
    {
        if (CheckRunning()) return;

        isRunning = true;
        view.LockControls(true);
        view.SetProgress(0);

        try
        {
            await Task.Run(() =>
            {
                var sorter = new FunctionalTestSorter();
                sorter.SortSheets(model.ExistingBaseFilePath);
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Sort failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            isRunning = false;
            view.LockControls(false);
        }
    }

    public async Task RunCreateNewBaseFileAsync(object sender, EventArgs e)
    {

        if (CheckRunning()) return;

        isRunning = true;
        view.LockControls(true);
        view.SetProgress(0);

        try
        {
            await Task.Run(() =>
            {
                string incomingTargetFile = Path.Combine(model.DirectoryPath, model.NewFileName);

                ExcelFileCreator.CreateNewExcel(incomingTargetFile);

                model.NewFileName = null;
                model.DirectoryPath = null;
                model.ExistingBaseFilePath = incomingTargetFile;
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Create file failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            isRunning = false;
        }
    }

    public bool CheckRunning()
    {
        if (isRunning)
        {
            view.
            return true;
        }

        return false;
    }
}
