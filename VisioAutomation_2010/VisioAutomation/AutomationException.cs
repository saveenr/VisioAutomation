namespace VisioAutomation
{
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

    [System.Serializable]
    public class QueryFrozenException : AutomationException
    {
        public QueryFrozenException()
        {
        }

        public QueryFrozenException(string message) : base(message)
        {
        }

        public QueryFrozenException(string message, System.Exception inner)
            : base(message, inner)
        {
        }

        protected QueryFrozenException(
            System.Runtime.Serialization.SerializationInfo info,
            System.Runtime.Serialization.StreamingContext context)
            : base(info, context)
        {
        }
    }

    [System.Serializable]
    public class DuplicateQueryColumnException : AutomationException
    {
        public DuplicateQueryColumnException()
        {
        }

        public DuplicateQueryColumnException(string message) : base(message)
        {
        }

        public DuplicateQueryColumnException(string message, System.Exception inner)
            : base(message, inner)
        {
        }

        protected DuplicateQueryColumnException(
            System.Runtime.Serialization.SerializationInfo info,
            System.Runtime.Serialization.StreamingContext context)
            : base(info, context)
        {
        }
    }

}