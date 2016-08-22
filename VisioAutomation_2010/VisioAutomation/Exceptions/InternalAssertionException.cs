namespace VisioAutomation.Exceptions
{
    [System.Serializable]
    public class InternalAssertionException : AutomationException
    {
        public InternalAssertionException()
        {
        }

        public InternalAssertionException(string message) : base(message)
        {
        }

        public InternalAssertionException(string message, System.Exception inner)
            : base(message, inner)
        {
        }

        protected InternalAssertionException(
            System.Runtime.Serialization.SerializationInfo info,
            System.Runtime.Serialization.StreamingContext context)
            : base(info, context)
        {
        }
    }
}