using System;
using System.Windows.Forms;

public interface IFileSelectionView
{
    event EventHandler MergeClicked;

    event EventHandler SortClicked;

    event EventHandler ResetClicked;

    event EventHandler OpenFileClicked;

    event EventHandler OpenFolderClicked;

    event EventHandler<BaseFileModeSelection> BaseFileModeChanged;// new event for radio button toggle

    BaseFileModeSelection CurrentBaseFileMode { get; } // new property to let presenter know current mode

    void UpdateBaseFileName(string name);

    void UpdateBaseFileFolderName(string name);

    void UpdateTargetFileNames(string names);

    void SetMergeButtonEnabled(bool enabled);

    void SetSortButtonEnabled(bool enabled);

    void SetOpenFileButtonEnabled(bool enabled);

    void ApplyTaskControlLock(bool disableForTask);

    void SetProgress(int percent);

    void UpdateUIForBaseFileSelectionMode();

    DialogResult ShowPrompt(string message, string title);
}
