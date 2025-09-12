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
    }

    public event EventHandler MergeRequested;

    public event EventHandler MergeCompleted;

    public event EventHandler SortRequested;

    public event EventHandler SortCompleted;

    public event EventHandler ResetRequested;

    public string BaseFilePath => model.BaseFilePath;

    public IReadOnlyList<string> TargetFilePaths => model.TargetFilePaths.AsReadOnly();

    public void SelectBaseFile()
    {
        using (OpenFileDialog dlg = new OpenFileDialog())
        {
            dlg.Filter = "Excel Files|*.xlsx;*.xls";
            dlg.Title = "Select Base Excel File";

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                model.BaseFilePath = dlg.FileName;
                view.UpdateBaseFileName(Path.GetFileName(model.BaseFilePath));
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
                view.UpdateTargetFileNames(displayText);
            }
        }
    }

    public void CheckMergeButtonState()
    {
        bool enabled = !string.IsNullOrEmpty(model.BaseFilePath) && model.TargetFilePaths.Count > 0;
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
        view.UpdateBaseFileName(string.Empty);
        view.UpdateTargetFileNames(string.Empty);
        CheckMergeButtonState();
        CheckSortButtonState();
    }

    public void SetProgress(int percent) => view.SetProgress(percent);
}
