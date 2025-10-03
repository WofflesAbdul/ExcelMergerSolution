using Excel_Handling;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

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

    public bool UseDvtReportTemplateExcel => false;

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
                    CheckSortButtonState();
                    CheckMergeButtonState();
                }
            }
        }
        else
        {
            model.BaseFilePath = newlyCreatedBaseFile;
            view.UpdateBaseFileName(Path.GetFileName(model.BaseFilePath));
            view.UpdateBaseFileFolderName(Path.GetDirectoryName(model.BaseFilePath));
            CheckSortButtonState();
            CheckMergeButtonState();
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

                CheckMergeButtonState();
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
        CheckMergeButtonState();
    }

    public void OnBaseFileModeSelectionChanged(BaseFileModeSelection mode)
    {
        // Reset any value stored prior to selection change
        model.NewBaseDirectoryPath = null;
        model.NewBaseFilename = null;
        model.BaseFilePath = null;
        view.UpdateUIForBaseFileSelectionMode();
        CheckMergeButtonState();
        CheckSortButtonState();
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
                UseShellExecute = true,
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
        if (string.IsNullOrEmpty(model.NewBaseDirectoryPath) ||
            string.IsNullOrEmpty(model.NewBaseFilename))
        {
            throw new InvalidOperationException("New base file path or filename is not set.");
        }

        string fullPath = Path.Combine(model.NewBaseDirectoryPath, model.NewBaseFilename + ".xlsx");
        ExcelFileCreator.CreateNewExcel(fullPath);
        model.BaseFilePath = fullPath;
        return fullPath;
    }

    public void RemovePlaceholderIfNeeded()
    {
        ExcelFileCreator.RemovePlaceholderSheets(model.BaseFilePath);
    }

}
