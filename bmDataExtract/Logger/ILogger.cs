namespace bmDataExtract
{
    public interface ILogger
    {
        void Log(string message, EventType type = EventType.None, bool finishLine = true, bool includeDate = false);
    }

    public enum EventType
    {
        Error, Info, Warning, Success, None
    }
}