public class OperationInfo
{
    public OperationInfo(OperationRequested requested, string ongoing, string completed, string error)
    {
        Requested = requested;
        OngoingMessage = ongoing;
        CompletedMessage = completed;
        ErrorMessage = error;
    }

    public OperationRequested Requested { get; }

    public string OngoingMessage { get; }

    public string ErrorMessage { get; }

    public string CompletedMessage { get; }
}
