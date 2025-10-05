using System;
using System.Collections.Generic;

public class FileSelectionModel
{
    private string baseFilePath;
    private string newBaseDirectoryPath;
    private string newBaseFilename;
    private List<string> targetFilePaths = new List<string>();

    public event EventHandler<ModelStateChangedEventArgs> ModelStateChanged;

    public string BaseFilePath
    {
        get => baseFilePath;
        set
        {
            baseFilePath = value;
            NotifyStateChanged();
        }
    }

    public string NewBaseDirectoryPath
    {
        get => newBaseDirectoryPath;
        set
        {
            newBaseDirectoryPath = value;
            NotifyStateChanged();
        }
    }

    public string NewBaseFilename
    {
        get => newBaseFilename;
        set
        {
            newBaseFilename = value;
            NotifyStateChanged();
        }
    }

    public List<string> TargetFilePaths
    {
        get => targetFilePaths;
        set
        {
            targetFilePaths = value ?? new List<string>();
            NotifyStateChanged();
        }
    }

    private void NotifyStateChanged()
    {
        ModelStateChanged?.Invoke(this, new ModelStateChangedEventArgs
        {
            HasDirectoryPath = !string.IsNullOrWhiteSpace(NewBaseDirectoryPath),
            HasValidBaseFile = !string.IsNullOrWhiteSpace(BaseFilePath),
            CanMerge = (TargetFilePaths.Count > 0) &&
                       (!string.IsNullOrWhiteSpace(BaseFilePath) ||
                        (!string.IsNullOrWhiteSpace(NewBaseFilename) && !string.IsNullOrWhiteSpace(NewBaseDirectoryPath))),
        });
    }
}
