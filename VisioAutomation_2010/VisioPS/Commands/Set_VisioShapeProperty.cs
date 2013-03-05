using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;
using System.Linq;
using VA=VisioAutomation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioShapeProperty")]
    public class Set_VisioShapeProperty: VisioPSCmdlet
    {
        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string Width { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string Height { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string PinX { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string PinY { get; set; }
        
        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string FillPattern { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string FillForegroundColor { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string FillbackgroundColor { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string LinePattern { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string LineWeight{ get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)]
        public string LineColor { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter BlastGuards;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TestCircular;


        [SMA.Parameter(Mandatory = false)]
        public IList<IVisio.Shape> Shapes;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            var update = new VisioAutomation.ShapeSheet.Update();
            update.BlastGuards = this.BlastGuards;
            update.TestCircular= this.TestCircular;

            var target_shapes = this.Shapes ?? scriptingsession.Selection.GetShapes();

            foreach (var shape in target_shapes)
            {
                var id = shape.ID16;
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.Width, this.Width);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.Height, this.Height);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PinX, this.PinX);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.PinY, this.PinY);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.FillForegnd, this.FillForegroundColor);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.FillPattern, this.FillPattern);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.FillBkgnd, this.FillbackgroundColor);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LineColor, this.LineColor);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LinePattern, this.LinePattern);
                update.SetFormulaIgnoreNull(id, VisioAutomation.ShapeSheet.SRCConstants.LineWeight, this.LineWeight);
            }

            var page = scriptingsession.Page.Get();
            update.Execute(page);
        }
    }
}