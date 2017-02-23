namespace VisioAutomation.Exceptions
{
    [System.Serializable]
    public class VisioOperationException : AutomationException
    {
        // for those cases when we ask Visio to do something and it didn't do it

        public VisioOperationException()
        {
        }

        public VisioOperationException(string message) : base(message)
        {
        }

        public VisioOperationException(string message, System.Exception inner)
            : base(message, inner)
        {
        }

        protected VisioOperationException(
            System.Runtime.Serialization.SerializationInfo info,
            System.Runtime.Serialization.StreamingContext context)
            : base(info, context)
        {
        }
    }
}