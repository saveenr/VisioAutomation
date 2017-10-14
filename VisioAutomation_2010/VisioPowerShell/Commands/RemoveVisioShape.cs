﻿using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Remove, VisioPowerShell.Commands.Nouns.VisioShape)]
    public class RemoveVisioShape : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            if (this.Shapes == null)
            {
                this.Client.Selection.DeleteShapesInSelection();                
            }
            else
            {
                foreach (var shape in this.Shapes)
                {
                    shape.Delete();
                }
            }
        }
    }
}