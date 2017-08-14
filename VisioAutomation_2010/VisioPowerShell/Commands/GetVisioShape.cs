using System;
using System.Linq;
using System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Get, VisioPowerShell.Commands.Nouns.VisioShape)]
    public class GetVisioShape : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = false)]
        public object[] Name;

        [Parameter(Mandatory = false)]
        public SwitchParameter Recursive;

        [Parameter(Mandatory = false)]
        public SwitchParameter SubSelected;

        protected override void ProcessRecord()
        {
            if (this.Name == null)
            {
                // return selected shapes

                if (this.Recursive)
                {
                    this.WriteVerbose("Returning selected shapes (nested)");
                    var shapes = this.Client.Selection.GetShapesRecursive();
                    this.WriteObject(shapes, false);
                }
                if (this.SubSelected)
                {
                    this.WriteVerbose("Returning selected shapes (subselecte)");
                    var shapes = this.Client.Selection.GetSubSelectedShapes();
                    this.WriteObject(shapes, false);
                }
                else
                {
                    this.WriteVerbose("Returning selected shapes ");
                    var shapes = this.Client.Selection.GetShapes();
                    this.WriteObject(shapes, false);
                }                
            }
            else
            {
                if (this.Name.Contains("*"))
                {
                    var shapes = this.Client.Draw.GetAllShapes();
                    this.WriteObject(shapes, false);
                }
                else
                {
                    bool all_ints = this.Name.All(i => i is int);
                    bool all_strings = this.Name.All(i => i is string);

                    if (!all_ints && !all_strings)
                    {
                        throw new ArgumentOutOfRangeException("must be array of only ints or only strings");
                    }

                    if (all_ints)
                    {
                        var ints = this.Name.Where(i => i is int).Cast<int>().ToArray();
                        var shapes = this.Client.Page.GetShapesByID(ints);
                        this.WriteObject(shapes, false);
                    }
                    else if (all_strings)
                    {
                        var strings = this.Name.Where(i => i is string).Cast<string>().ToArray();
                        var shapes = this.Client.Page.GetShapesByName(strings);
                        this.WriteObject(shapes, false);
                    }
                    else
                    {
                        throw new ArgumentOutOfRangeException("Should never get here");
                    }                    
                }
            }
        }
    }
}