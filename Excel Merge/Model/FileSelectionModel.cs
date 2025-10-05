using System;
using System.Collections.Generic;
using System.IO;


public class FileSelectionModel : IFileSelectionModel
{
    private string existingBaseFilePath;
    private string directoryPath;
    private string newFileName;
    private List<string> targetFilePaths = new List<string>();

    public event EventHandler<ModelStateChangedEventArgs> ModelStateChanged;

    public string ExistingBaseFilePath
    {
        get => existingBaseFilePath;
        set
        {
            existingBaseFilePath = value;
            NotifyStateChanged();

            DirectoryPath = Path.GetDirectoryName(value);
        }
    }

    public string DirectoryPath
    {
        get => directoryPath;
        set
        {
            directoryPath = value;
            NotifyStateChanged();
        }
    }

    public string NewFileName
    {
        get => newFileName;
        set
        {
            newFileName = value;
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
            HasDirectoryPath = !string.IsNullOrWhiteSpace(DirectoryPath),
            HasValidBaseFile = !string.IsNullOrWhiteSpace(ExistingBaseFilePath),
            CanMerge = (TargetFilePaths.Count > 0) &&
                   (!string.IsNullOrWhiteSpace(ExistingBaseFilePath) ||
                    (!string.IsNullOrWhiteSpace(NewFileName) && !string.IsNullOrWhiteSpace(DirectoryPath))),
        });
    }
}
