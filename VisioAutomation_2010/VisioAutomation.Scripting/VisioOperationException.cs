namespace VisioAutomation.Scripting
{
    [System.Serializable]
    public class VisioOperationException : System.Exception
    {
        public VisioOperationException() { }
        public VisioOperationException(string message) : base(message) { }
        public VisioOperationException(string message, System.Exception inner) : base(message, inner) { }
    }
}