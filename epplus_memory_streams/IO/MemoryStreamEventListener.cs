using System.Diagnostics.Tracing;

namespace epplus_memory_streams.IO;

public sealed class MemoryStreamEventListener : EventListener
{
    private const string MemoryStreamCreated = "MemoryStreamCreated";
    private const string MemoryStreamFinalized = "MemoryStreamFinalized";

    private static readonly HashSet<string> Events = new HashSet<string>(StringComparer.Ordinal)
    {
        MemoryStreamCreated,
        "MemoryStreamDisposed",
        MemoryStreamFinalized,
    };

    protected override void OnEventSourceCreated(EventSource eventSource)
    {
        if (eventSource.Name == "Microsoft-IO-RecyclableMemoryStream")
        {
            var args = new Dictionary<string, string>
            {
                ["EventCounterIntervalSec"] = "1"
            };

            EnableEvents(eventSource, EventLevel.Verbose, EventKeywords.All, args);
        }
    }

    protected override void OnEventWritten(EventWrittenEventArgs eventWrittenEventArgs)
    {
        if (!Events.Contains(eventWrittenEventArgs.EventName))
        {
            return;
        }

        var args = eventWrittenEventArgs.Payload?.Where(x => x != null).ToList() ?? new List<object>();

        var message = string.Join(", ", args.Where(x => !(x is string s && s.StartsWith("   at System.Environment.get_StackTrace()", StringComparison.Ordinal))));

        if (string.Equals(MemoryStreamFinalized, eventWrittenEventArgs.EventName, StringComparison.Ordinal))
        {
            var stacktrace = args.FirstOrDefault(x => x is string s && s.StartsWith("   at System.Environment.get_StackTrace()", StringComparison.Ordinal)) as string;

            var exception = new MemoryStreamException(eventWrittenEventArgs.EventName, stacktrace);

            Console.WriteLine("[{0}] {1}, {2}", eventWrittenEventArgs.EventName, message, exception);

        }
        else
        {
            Console.WriteLine("[{0}] {1}", eventWrittenEventArgs.EventName, message);
        }
    }
}