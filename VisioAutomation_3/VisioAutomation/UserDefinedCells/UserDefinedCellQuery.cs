using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.UserDefinedCells
{
    class UserDefinedCellQuery : VA.ShapeSheet.Query.SectionQuery
    {
        public VA.ShapeSheet.Query.SectionQueryColumn Value { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Prompt { get; set; }
        
        public UserDefinedCellQuery() :
            base(IVisio.VisSectionIndices.visSectionUser)
        {
            Value = this.AddColumn(VA.ShapeSheet.SRCConstants.User_Value, "Value");
            Prompt = this.AddColumn(VA.ShapeSheet.SRCConstants.User_Prompt, "Prompt");
        }
    }
}