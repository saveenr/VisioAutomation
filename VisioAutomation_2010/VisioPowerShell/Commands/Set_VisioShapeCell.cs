using System.Collections;
using VisioScripting.Models;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, VisioPowerShell.Nouns.VisioShapeCell)]
    public class Set_VisioShapeCell : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false, Position = 0)]
        public Hashtable Hashtable { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter BlastGuards { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TestCircular { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes { get; set; }

        protected override void ProcessRecord()
        {
            var target_shapes = this.Shapes ?? this.Client.Selection.GetShapes();
            var targets = new TargetShapes(target_shapes);

            var dic = Set_VisioPageCell.CellHashtableToDictionary(this.Hashtable);
            this.Client.ShapeSheet.SetShapeCells(targets, dic, this.BlastGuards, this.TestCircular);
        }
    }
}