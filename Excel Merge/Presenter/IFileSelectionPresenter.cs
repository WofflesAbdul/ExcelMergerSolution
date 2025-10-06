using System;
using System.Threading.Tasks;

public interface IFileSelectionPresenter
{
    InputFileMode InputFileMode { get; }

    void SelectBaseFile();

    void SelectDirectoryPath();

    void SelectTargetFiles();

    void OnFilenameSet(string filename);

    void OnReset(object sender, EventArgs e);

    void OnInputFileModeChanged(object sender, InputFileMode mode);

    void OnOpenFileClicked(object sender, EventArgs e);

    void OnOpenFolderClicked(object sender, EventArgs e);

    void OnModelStateChanged(object sender, ModelStateChangedEventArgs e);

    Task RunMergeAsync(object sender, EventArgs e);

    Task RunSortAsync(object sender, EventArgs e);

    Task RunCreateNewBaseFileAsync(object sender, EventArgs e);

    bool CheckRunning();


}
