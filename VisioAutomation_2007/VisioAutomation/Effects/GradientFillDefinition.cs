using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Effects
{
    public class GradientFillDefinition
    {
        public VA.ShapeSheet.FormulaLiteral FillPattern;
        public VA.ShapeSheet.FormulaLiteral StartColor;
        public VA.ShapeSheet.FormulaLiteral EndColor;
        public VA.ShapeSheet.FormulaLiteral StartTransparency;
        public VA.ShapeSheet.FormulaLiteral EndTransparency;

        public void Apply(VA.ShapeSheet.Update.SIDSRCUpdate update, short shapeid)
        {
            update.SetFormula(shapeid, VA.ShapeSheet.SRCConstants.FillPattern, this.FillPattern);
            update.SetFormula(shapeid, VA.ShapeSheet.SRCConstants.FillForegnd, this.StartColor);
            update.SetFormula(shapeid, VA.ShapeSheet.SRCConstants.FillBkgnd, this.EndColor);
            update.SetFormula(shapeid, VA.ShapeSheet.SRCConstants.FillForegndTrans, this.StartTransparency);
            update.SetFormula(shapeid, VA.ShapeSheet.SRCConstants.FillBkgndTrans, this.EndTransparency);
        }
    }
}