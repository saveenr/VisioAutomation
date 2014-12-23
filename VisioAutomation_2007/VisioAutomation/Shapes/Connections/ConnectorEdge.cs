using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes.Connections
{
    public struct ConnectorEdge
    {
        public IVisio.Shape Connector { get; private set; }
        public IVisio.Shape From { get; private set; }
        public IVisio.Shape To { get; private set; }

        public ConnectorEdge(IVisio.Shape connectingshape, IVisio.Shape fromshape, IVisio.Shape toshape) : this()
        {
            if (fromshape == null)
            {
                throw new System.ArgumentNullException("fromshape");
            }

            if (toshape == null)
            {
                throw new System.ArgumentNullException("toshape");
            }

            this.Connector = connectingshape;
            this.From = fromshape;
            this.To = toshape;
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