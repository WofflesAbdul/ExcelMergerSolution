using System;

public static class OperationRegistry
{
    public static readonly OperationInfo Merge = new OperationInfo(
        OperationRequested.Merge,
        "Status: Merging in process...",
        "Completed: Merge action is completed!",
        "Error: Unable to perform Merge action!"
    );

    public static readonly OperationInfo Sort = new OperationInfo(
        OperationRequested.Sort,
        "Status: Sorting in process...",
        "Completed: Sort action is completed!",
        "Error: Unable to perform Sort action!"
    );

    public static readonly OperationInfo CreateNewFile = new OperationInfo(
        OperationRequested.CreateNewFile,
        "Status: Creating new file in process...",
        "Completed: Create New File action is completed!",
        "Error: Unable to perform Create New File action!"
    );

    public static OperationInfo Get(OperationRequested requested)
    {
        switch (requested)
        {
            case OperationRequested.Merge: return Merge;
            case OperationRequested.Sort: return Sort;
            case OperationRequested.CreateNewFile: return CreateNewFile;
            default: throw new ArgumentOutOfRangeException(nameof(requested), requested, null);
        }
    }
}
