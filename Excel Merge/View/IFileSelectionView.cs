using System;
using System.Windows.Forms;

public interface IFileSelectionView
{
    event EventHandler MergeClicked;

    event EventHandler SortClicked;

    event EventHandler ResetClicked;

    event EventHandler OpenFileClicked;

    event EventHandler OpenFolderClicked;

    event EventHandler<InputFileMode> InputFileModeChanged;// new event for radio button toggle

    InputFileMode CurrentBaseFileMode { get; } // new property to let presenter know current mode

    void UpdateFilename(string name);

    void UpdateDirectory(string name);

    void UpdateTargetFilenames(string names);

    void SetMergeButtonEnabled(bool enabled);

    void SetSortButtonEnabled(bool enabled);

    void SetOpenFileButtonEnabled(bool enabled);

    void LockControls(bool enable);

    void SetProgress(int percent);

    void UpdateUIForBaseFileSelectionMode();

    DialogResult ShowPrompt(string message, string title);
}
