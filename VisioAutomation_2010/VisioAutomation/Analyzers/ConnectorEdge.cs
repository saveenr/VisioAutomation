using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Analyzers
{
    public readonly struct ConnectorEdge
    {
        public IVisio.Shape Connector { get; }
        public IVisio.Shape From { get; }
        public IVisio.Shape To { get; }

        public ConnectorEdge(IVisio.Shape connectingshape, IVisio.Shape fromshape, IVisio.Shape toshape) : this()
        {
            this.Connector = connectingshape;
            this.From = fromshape ?? throw new System.ArgumentNullException(nameof(fromshape));
            this.To = toshape ?? throw new System.ArgumentNullException(nameof(toshape));
        }

        public override string ToString()
        {
            string from_name = this.From !=null ? this.From.NameU : "null";
            string to_name = this.To != null ? this.To.NameU : "null";

            if (this.Connector != null)
            {
                var connector_name = this.Connector.NameU;
                return string.Format("({0}:{1}->{2})", connector_name, from_name, to_name);                
            }
            else
            {
                return string.Format("({0}->{1})", from_name, to_name);
            }
        }
    }
}