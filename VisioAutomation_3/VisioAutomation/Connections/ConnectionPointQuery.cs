using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Connections
{
    class ConnectionPointQuery : VA.ShapeSheet.Query.SectionQuery
    {
        public VA.ShapeSheet.Query.SectionQueryColumn DirX { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn DirY { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Type { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn X { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Y { get; set; }

        public ConnectionPointQuery() :
            base(IVisio.VisSectionIndices.visSectionConnectionPts)
        {
            DirX = this.AddColumn(VA.ShapeSheet.SRCConstants.Connections_DirX,"DirX");
            DirY = this.AddColumn(VA.ShapeSheet.SRCConstants.Connections_DirY, "DirY");
            Type = this.AddColumn(VA.ShapeSheet.SRCConstants.Connections_Type, "Type");
            X = this.AddColumn(VA.ShapeSheet.SRCConstants.Connections_X, "X");
            Y = this.AddColumn(VA.ShapeSheet.SRCConstants.Connections_Y, "Y");
        }
    }
}