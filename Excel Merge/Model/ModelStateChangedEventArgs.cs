using System;

public class ModelStateChangedEventArgs : EventArgs
{
    public bool HasDirectoryPath { get; set; }

    public bool HasValidBaseFile { get; set; }

    public bool CanMerge { get; set; }
}
