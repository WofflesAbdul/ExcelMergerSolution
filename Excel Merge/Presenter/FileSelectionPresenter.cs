using Excel_Handling;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

/// <summary>
/// Handles the selection of base and target Excel files.
/// Acts as the Presenter in the MVP pattern.
/// </summary>
public class FileSelectionPresenter
{
    private readonly IFileSelectionView view;
    private readonly FileSelectionModel model;

    public FileSelectionPresenter(IFileSelectionView view)
    {
        this.view = view ?? throw new ArgumentNullException(nameof(view));
        model = new FileSelectionModel();

        view.OpenFileClicked += (s, e) => OnOpenFileClicked();
        view.OpenFolderClicked += (s, e) => OnOpenFolderClicked();
        view.BaseFileModeChanged += (s, mode) => OnBaseFileModeSelectionChanged(mode);
    }

    public event EventHandler MergeRequested;

    public event EventHandler MergeCompleted;

    public event EventHandler SortRequested;

    public event EventHandler SortCompleted;

    public event EventHandler ResetRequested;

    public string BaseFilePath => model.BaseFilePath;

    public bool UseDvtReportTemplateExcel => false; // TODO: Form1 Checkbox check state

    public IReadOnlyList<string> TargetFilePaths => model.TargetFilePaths.AsReadOnly();

    public void SelectBaseFile(string newlyCreatedBaseFile = null)
    {
        if (string.IsNullOrEmpty(newlyCreatedBaseFile))
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "Excel Files|*.xlsx;*.xls";
                dlg.Title = "Select Base Excel File";

                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    model.BaseFilePath = dlg.FileName;
                    view.UpdateBaseFileName(Path.GetFileName(model.BaseFilePath));
                    view.UpdateBaseFileFolderName(Path.GetDirectoryName(model.BaseFilePath));
                }
            }
        }
        else
        {
            model.BaseFilePath = newlyCreatedBaseFile;
            view.UpdateBaseFileName(Path.GetFileName(model.BaseFilePath));
            view.UpdateBaseFileFolderName(Path.GetDirectoryName(model.BaseFilePath));
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
                view.UpdateTargetFileNames(displayText);
            }
        }
    }

    public void CheckMergeButtonState()
    {
        bool enabled = false;

        switch (view.CurrentBaseFileMode)
        {
            case BaseFileModeSelection.UseExistingFile:
                enabled = !string.IsNullOrEmpty(model.BaseFilePath)
                          && model.TargetFilePaths.Count > 0;
                break;

            case BaseFileModeSelection.CreateNewFile:
                enabled = !string.IsNullOrEmpty(model.NewBaseFilename)
                          && !string.IsNullOrEmpty(model.NewBaseDirectoryPath)
                          && model.TargetFilePaths.Count > 0;
                break;
        }

        view.SetMergeButtonEnabled(enabled);
    }

    public void CheckSortButtonState()
    {
        bool enabled = !string.IsNullOrEmpty(model.BaseFilePath);
        view.SetSortButtonEnabled(enabled);
        view.SetOpenFileButtonEnabled(enabled);
    }

    public void ClearTargetSelection()
    {
        model.TargetFilePaths.Clear();
        view.UpdateTargetFileNames(string.Empty);
        CheckMergeButtonState(); // disable merge button
    }

    public void OnBaseFileModeSelectionChanged(BaseFileModeSelection mode)
    {
        // Reset any value stored prior to selection change
        model.NewBaseDirectoryPath = null;
        model.NewBaseDirectoryPath = null;
        model.BaseFilePath = null;
        view.UpdateUIForBaseFileSelectionMode();
        CheckMergeButtonState();
    }

    public void OnMergeClicked() => MergeRequested?.Invoke(this, EventArgs.Empty);

    public void OnSortClicked() => SortRequested?.Invoke(this, EventArgs.Empty);

    public void OnResetClicked() => ResetRequested?.Invoke(this, EventArgs.Empty);

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
        if (!string.IsNullOrEmpty(BaseFilePath) && File.Exists(BaseFilePath))
        {
            string folderPath = Path.GetDirectoryName(BaseFilePath);

            if (!string.IsNullOrEmpty(folderPath) && Directory.Exists(folderPath))
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = folderPath,
                    UseShellExecute = true
                });
            }
        }
    }

    public void OnNewBaseFilenameEntered(string filename)
    {
        model.NewBaseFilename = filename;
        CheckMergeButtonState();
    }

    public void OnNewBaseFileTargetLocationChanged()
    {
        using (var dlg = new FolderBrowserDialog())
        {
            dlg.Description = "Select folder for new Base Excel file";
            dlg.ShowNewFolderButton = true;

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                model.NewBaseDirectoryPath = dlg.SelectedPath;
                view.UpdateBaseFileFolderName(model.NewBaseDirectoryPath);
                // Excel file to be created upon OnMergeClicked
            }
        }

        CheckMergeButtonState();
    }

    public void NotifyMergeCompleted()
    {
        MergeCompleted?.Invoke(this, EventArgs.Empty);

        var result = view.ShowPrompt("Merge completed! Do you want to open the merged file?", "Merge Completed");
        if (result == DialogResult.Yes && !string.IsNullOrEmpty(BaseFilePath) && File.Exists(BaseFilePath))
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = BaseFilePath,
                UseShellExecute = true
            });
        }
    }

    public void NotifySortCompleted()
    {
        SortCompleted?.Invoke(this, EventArgs.Empty);

        var result = view.ShowPrompt("Sort completed! Do you want to open the sorted file?", "Sort Completed");
        if (result == DialogResult.Yes && !string.IsNullOrEmpty(BaseFilePath) && File.Exists(BaseFilePath))
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = BaseFilePath,
                UseShellExecute = true
            });
        }
    }

    public void ResetSelection()
    {
        model.BaseFilePath = null;
        model.TargetFilePaths.Clear();
        model.NewBaseFilename = null;
        model.NewBaseDirectoryPath = null;
        view.UpdateBaseFileName(string.Empty);
        view.UpdateUIForBaseFileSelectionMode();
        view.UpdateBaseFileFolderName(string.Empty);
        view.UpdateTargetFileNames(string.Empty);
        CheckMergeButtonState();
        CheckSortButtonState();
    }

    public void SetProgress(int percent) => view.SetProgress(percent);

    public string CreateNewBaseFile()
    {
        if (string.IsNullOrWhiteSpace(model.NewBaseDirectoryPath) ||
            string.IsNullOrWhiteSpace(model.NewBaseFilename))
        {
            throw new InvalidOperationException("New base file path or filename is not set.");
        }

        string fullPath = Path.Combine(
            model.NewBaseDirectoryPath,
            model.NewBaseFilename.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                ? model.NewBaseFilename
                : model.NewBaseFilename + ".xlsx"
        );

        var creator = new ExcelFileCreator();
        bool success = creator.CreateNewExcel(fullPath);

        if (!success)
        {
            throw new IOException($"Failed to create new Excel file at {fullPath}");
        }

        // Track placeholder sheet for later removal
        bool hasPlaceholderSheet = creator.HasPlaceholderSheet;

        // Keep model consistent
        model.BaseFilePath = fullPath;

        // Optionally store the creator if you want to remove the placeholder later
        _lastExcelCreator = hasPlaceholderSheet ? creator : null;

        return fullPath;
    }

    // New field in presenter
    private ExcelFileCreator _lastExcelCreator;

    // Call this after merge if needed
    public void RemovePlaceholderIfNeeded()
    {
        _lastExcelCreator?.RemovePlaceholderSheet(model.BaseFilePath);
        _lastExcelCreator = null;
    }

}
