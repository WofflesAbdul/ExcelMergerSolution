using System;

public interface IFileSelectionView
{
    event EventHandler MergeClicked;

    event EventHandler SortClicked;

    event EventHandler ResetClicked;

    void UpdateBaseFileName(string name);

    void UpdateTargetFileNames(string names);

    void SetMergeButtonEnabled(bool enabled);

    void SetSortButtonEnabled(bool enabled);

    void ApplyTaskControlLock(bool disableForTask);

    void SetProgress(int percent);
}
