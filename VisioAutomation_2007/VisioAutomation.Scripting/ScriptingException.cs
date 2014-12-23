namespace VisioAutomation.Scripting
{
    [System.Serializable]
    public class ScriptingException : System.Exception
    {
        public ScriptingException() { }
        public ScriptingException(string message) : base(message) { }
        public ScriptingException(string message, System.Exception inner) : base(message, inner) { }
    }
}