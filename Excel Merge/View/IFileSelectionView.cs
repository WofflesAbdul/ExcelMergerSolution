using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

public interface IFileSelectionView
{
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

    void SetOngoingStatus(string message);

    void SetCompletionStatus(string message, bool isError = false);

    void NewFileCreated();

    Task AnimateProgressBarAsync(int steps, int delayMs, CancellationToken token);
}
