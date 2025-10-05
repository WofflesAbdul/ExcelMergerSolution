using System;

public class TaskCompletedEventArgs : EventArgs
{
    public string Message { get; set; } // e.g., "Merge completed!"

    public string FilePath { get; set; } // Optional: path of file involved

    public TaskType Type { get; set; } // Enum to differentiate Merge/Sort/Create
}