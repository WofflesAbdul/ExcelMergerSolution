using System;
using System.Collections.Generic;
using System.IO;

public class FileSelectionModel : IFileSelectionModel
{
    private string existingBaseFilePath;
    private string directoryPath;
    private string newFileName;

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

    public List<string> TargetFilePaths { get; } = new List<string>();

    public void AddTargetFiles(string[] targetFilePaths)
    {
        TargetFilePaths.Clear();
        TargetFilePaths.AddRange(targetFilePaths);
        NotifyStateChanged();
    }

    public void ClearTargetFiles()
    {
        TargetFilePaths.Clear();
        NotifyStateChanged();
    }

    private void NotifyStateChanged()
    {
        var handler = ModelStateChanged;
        if (handler != null)
        {
            handler(this, new ModelStateChangedEventArgs
            {
                ExistingBaseFilePath = this.ExistingBaseFilePath,
                NewFileName = this.NewFileName,
                DirectoryPath = this.DirectoryPath,
                TargetFilePaths = new List<string>(this.TargetFilePaths),

                HasDirectoryPath = !string.IsNullOrWhiteSpace(DirectoryPath),
                HasValidBaseFile = !string.IsNullOrWhiteSpace(ExistingBaseFilePath),
                CanMerge = (TargetFilePaths.Count > 0) &&
                       (!string.IsNullOrWhiteSpace(ExistingBaseFilePath) ||
                        (!string.IsNullOrWhiteSpace(NewFileName) && !string.IsNullOrWhiteSpace(DirectoryPath))),
            });
        }
    }
}
