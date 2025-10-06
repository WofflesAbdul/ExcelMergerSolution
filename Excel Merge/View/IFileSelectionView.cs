using System;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;

public interface IFileSelectionView
{
    event EventHandler MergeRequested;

    event EventHandler SortRequested;

    event EventHandler CreateNewFileRequested;

    event EventHandler ResetRequested;

    event EventHandler OpenFileClicked;

    event EventHandler OpenFolderClicked;

    event EventHandler<InputFileMode> InputFileModeChanged; // new event for radio button toggle

    void DisplayFileName(string fileName);

    void DisplayDirectoryPath(string directoryPath);

    void DisplayTargetFilePaths(IEnumerable<string> targetFileNames);

    void SetMergeButtonEnabled(bool enabled);

    void SetSortButtonEnabled(bool enabled);

    void SetOpenFileButtonEnabled(bool enabled);

    void SetOpenFolderButtonEnabled(bool enabled);

    void LockControls(bool enable);

    void SetProgress(int percent);

    DialogResult ShowPrompt(string message, string title);
}
