namespace GenTreeOps
{
    /// <summary>
    /// Assists in performing a depth-first traversal of nodes for some Node type T. 
    /// T need not be of any specific type.
    /// </summary>
    /// 
    public enum WalkEventType
    {
        EventEnter,
        EventExit
    }

    public struct WalkEvent<T>
    {
 
        public readonly WalkEventType Type;
        public readonly T Node;

        public WalkEvent(T node, WalkEventType event_type)
        {
            this.Node = node;
            this.Type = event_type;
        }

        public static WalkEvent<T> CreateEnterEvent(T node)
        {
            var we = new WalkEvent<T>(node, WalkEventType.EventEnter);
            return we;
        }

        public static WalkEvent<T> CreateExitEvent(T node)
        {
            var we = new WalkEvent<T>(node, WalkEventType.EventExit);
            return we;
        }
    }
}