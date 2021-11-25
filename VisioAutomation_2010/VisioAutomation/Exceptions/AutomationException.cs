namespace VisioAutomation.Exceptions;

[System.Serializable]
public class AutomationException : System.Exception
{
    public AutomationException()
    {
    }

    public AutomationException(string message) : base(message)
    {
    }

    public AutomationException(string message, System.Exception inner)
        : base(message, inner)
    {
    }

    protected AutomationException(
        System.Runtime.Serialization.SerializationInfo info,
        System.Runtime.Serialization.StreamingContext context)
        : base(info, context)
    {
    }
}