namespace epplus_memory_streams.IO;

internal sealed class MemoryStreamException : Exception
{
    public override string StackTrace { get; }

    public MemoryStreamException(string message, string stackTrace) : base(message)
    {
        StackTrace = stackTrace;
    }
}