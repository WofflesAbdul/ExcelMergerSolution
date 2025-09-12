using System;
using System.Windows.Forms;

public interface IFileSelectionView
{
    event EventHandler MergeClicked;

    event EventHandler SortClicked;

    event EventHandler ResetClicked;

    event EventHandler OpenFileClicked;

    void UpdateBaseFileName(string name);

    void UpdateTargetFileNames(string names);

    void SetMergeButtonEnabled(bool enabled);

    void SetSortButtonEnabled(bool enabled);

    void SetOpenFileButtonEnabled(bool enabled);

    void ApplyTaskControlLock(bool disableForTask);

    void SetProgress(int percent);

    DialogResult ShowPrompt(string message, string title);
}
