using Microsoft.IO;

namespace epplus_memory_streams.IO;

/// <summary>
/// http://www.philosophicalgeek.com/2015/02/06/announcing-microsoft-io-recycablememorystream/
/// </summary>
public sealed class MemoryStreamManager
{
    private static volatile RecyclableMemoryStreamManager _defaultManager;

    static MemoryStreamManager()
    {
        _defaultManager = new RecyclableMemoryStreamManager(new RecyclableMemoryStreamManager.Options
        {
            // todo: can we collect stats and set reasonable limits?
        });
    }

    public static RecyclableMemoryStreamManager GetManager()
    {
        return _defaultManager;
    }

    public static void GenerateCallStacks()
    {
        var recyclableMemoryStreamManager = new RecyclableMemoryStreamManager(new RecyclableMemoryStreamManager.Options
        {
            GenerateCallStacks = true
        });

        _defaultManager = recyclableMemoryStreamManager;
    }

    /// <summary>
    /// Retrieve a new MemoryStream object with the given tag and a default initial capacity.
    /// </summary>
    /// <returns>A MemoryStream.</returns>
    public static MemoryStream GetStream()
    {
        return _defaultManager.GetStream("Default");
    }

    /// <summary>
    /// Retrieve a new MemoryStream object with the given tag and with contents copied from the provided
    /// buffer. The provided buffer is not wrapped or used after construction.
    /// </summary>
    /// <remarks>The new stream's position is set to the beginning of the stream when returned.</remarks>
    /// <param name="buffer">The byte buffer to copy data from.</param>
    /// <returns>A MemoryStream.</returns>
    public static MemoryStream GetStream(byte[] buffer)
    {
        return _defaultManager.GetStream("Default", buffer, 0, buffer.Length);
    }

    /// <summary>
    /// Retrieve a new MemoryStream object with the given tag and with contents copied from the provided
    /// buffer. The provided buffer is not wrapped or used after construction.
    /// </summary>
    /// <remarks>The new stream's position is set to the beginning of the stream when returned.</remarks>
    /// <param name="buffer">The byte buffer to copy data from.</param>
    /// <param name="offset">The offset from the start of the buffer to copy from.</param>
    /// <param name="count">The number of bytes to copy from the buffer.</param>
    /// <returns>A MemoryStream.</returns>
    public static MemoryStream GetStream(byte[] buffer, int offset, int count)
    {
        return _defaultManager.GetStream("Default", buffer, offset, count);
    }
}