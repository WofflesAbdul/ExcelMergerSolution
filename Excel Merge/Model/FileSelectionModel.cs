using System.Collections.Generic;

public class FileSelectionModel
{
    // Base file full path
    public string BaseFilePath { get; set; }

    // List of target file full paths
    public List<string> TargetFilePaths { get; } = new List<string>();
}