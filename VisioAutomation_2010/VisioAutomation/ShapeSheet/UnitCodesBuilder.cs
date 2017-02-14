using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet
{
    public class UnitCodesBuilder : ShapeSheetArrayBuilder<IVisio.VisUnitCodes>
    {

        public UnitCodesBuilder() : base()
        {
        }

        public UnitCodesBuilder(int capacity) : base(capacity)
        {
        }
    }
}