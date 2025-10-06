using System;
using System.Collections.Generic;

public class ModelStateChangedEventArgs : EventArgs
{
    public string ExistingBaseFilePath { get; set; }

    public string DirectoryPath { get; set; }

    public string NewFileName { get; set; }

    public IReadOnlyList<string> TargetFilePaths { get; set; }

    public bool HasDirectoryPath { get; set; }

    public bool HasValidBaseFile { get; set; }

    public bool CanMerge { get; set; }
}
