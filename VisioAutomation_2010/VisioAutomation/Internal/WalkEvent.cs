using System.Collections.Generic;
using System.Linq;
using VA=VisioAutomation;

namespace VisioAutomation.Internal
{
    /// <summary>
    /// Assists in performing a depth-first traversal of nodes for some Node type T. 
    /// T need not be of any specific type.
    /// </summary>
    internal struct WalkEvent<T>
    {
        public enum WalkEventType
        {
            Enter,
            Exit
        };
 
        public readonly WalkEventType Type;
        public readonly T Node;

        public WalkEvent(T node, WalkEventType event_type)
        {
            this.Node = node;
            this.Type = event_type;
        }

        public static WalkEvent<T> CreateEnterEvent(T node)
        {
            var we = new WalkEvent<T>(node,WalkEventType.Enter);
            return we;
        }

        public static WalkEvent<T> CreateExitEvent(T node)
        {
            var we = new WalkEvent<T>(node, WalkEventType.Exit);
            return we;
        }

        public bool HasEnteredNode
        {
            get { return this.Type == WalkEventType.Enter; }
        }

        public bool HasExitedNode
        {
            get { return this.Type == WalkEventType.Exit; }
        }
    }
}