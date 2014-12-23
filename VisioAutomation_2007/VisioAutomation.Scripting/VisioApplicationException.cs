namespace VisioAutomation.Scripting
{
    [System.Serializable]
    public class VisioApplicationException : ScriptingException
    {
        public VisioApplicationException() { }
        public VisioApplicationException(string message) : base(message) { }
        public VisioApplicationException(string message, System.Exception inner) : base(message, inner) { }
    }
}