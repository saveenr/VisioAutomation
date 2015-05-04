namespace VisioAutomation
{
    [System.Serializable]
    public class AutomationException : System.Exception
    {
        // For guidelines regarding the creation of new exception types, see
        //    http://msdn.microsoft.com/library/default.asp?url=/library/en-us/cpgenref/html/cpconerrorraisinghandlingguidelines.asp
        // and
        //    http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dncscol/html/csharp07192001.asp
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
}