using IVisio=Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;


namespace VisioAutomation.Models.Grid
{
    public class Node
    {
        public IVisio.Master Master { get; set; }
        public string Text { get; set; }
        public IVisio.Shape Shape { get; set; }
        public VA.Drawing.Rectangle Rectangle { get; set; }
        public short ShapeID { get; set; }
        public int Column { get; set; }
        public int Row { get; set; }
        public object Data { get; set; }
        public bool Draw { get; set; }

        public VA.DOM.ShapeCells Cells { get; set; }

        public Node()
        {
            this.ShapeID = -1;
        }

    }
}