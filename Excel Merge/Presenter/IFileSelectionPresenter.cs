using System;
using System.Collections.Generic;

public interface IFileSelectionPresenter
{
    event EventHandler<TaskCompletedEventArgs> TaskCompleted;

    string BaseFilePath { get; }

    string NewBaseDirectoryPath { get; }

    string NewBaseFilename { get; }

    IReadOnlyList<string> TargetFilePaths { get; }

    void SelectBaseFile(string newlyCreatedBaseFile = null);

    void SelectTargetFiles();

    void OnFilenameSet(string filename);

    void OnDirectorySet(string directoryPath);

    void OnReset();

    void OnInputFileModeChanged();

    void OnOpenFileClicked();

    void OnOpenFolderClicked();

    void OnModelStateChanged(object sender, ModelStateChangedEventArgs e);

    void RunMerge();

    void RunSort();

    void RunCreateNewBaseFile();

    void SetBaseFile(string filename);
}
