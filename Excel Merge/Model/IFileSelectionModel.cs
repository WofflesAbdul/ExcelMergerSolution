using System;
using System.Collections.Generic;

public interface IFileSelectionModel
{
    event EventHandler<ModelStateChangedEventArgs> ModelStateChanged;

    string ExistingBaseFilePath { get; set; }

    string DirectoryPath { get; set; }

    string NewFileName { get; set; }

    List<string> TargetFilePaths { get; }
}
