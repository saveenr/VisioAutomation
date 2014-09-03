using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using VA = VisioAutomation;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioShape")]
    public class Get_VisioShape : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = false)]
        public object[] NameOrID;


        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Recursive;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter SubSelected;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;


            if (this.NameOrID == null)
            {
                // return selected shapes

                this.WriteVerboseEx("NUll or *");
                this.WriteVerboseEx("a {0}", this.NameOrID == null);
                this.WriteVerboseEx("b {0}", this.NameOrID == null || this.NameOrID.Contains("*"));

                if (this.Recursive)
                {
                    this.WriteVerboseEx("Returning selected shapes (nested)");
                    var shapes = scriptingsession.Selection.GetShapesRecursive();
                    this.WriteObject(shapes, false);
                }
                if (this.SubSelected)
                {
                    this.WriteVerboseEx("Returning selected shapes (subselecte)");
                    var shapes = scriptingsession.Selection.GetSubSelectedShapes();
                    this.WriteObject(shapes, false);
                }
                else
                {
                    this.WriteVerboseEx("Returning selected shapes ");
                    var shapes = scriptingsession.Selection.GetShapes();
                    this.WriteObject(shapes, false);
                }                
            }
            else
            {
                if (this.NameOrID.Contains("*"))
                {
                    var shapes = scriptingsession.Draw.GetAllShapes();
                    this.WriteObject(shapes, false);
                }
                else
                {
                    bool all_ints = this.NameOrID.All(i => i is int);
                    bool all_strings = this.NameOrID.All(i => i is string);

                    if (!all_ints && !all_strings)
                    {
                        throw new System.ArgumentOutOfRangeException("must be array of only ints or only strings");
                    }

                    if (all_ints)
                    {
                        var ints = this.NameOrID.Where(i => i is int).Cast<int>().ToArray();
                        var shapes = scriptingsession.Page.GetShapesByID(ints);
                        this.WriteObject(shapes, false);
                    }
                    else if (all_strings)
                    {
                        var strings = this.NameOrID.Where(i => i is string).Cast<string>().ToArray();
                        var shapes = scriptingsession.Page.GetShapesByName(strings);
                        this.WriteObject(shapes, false);
                    }
                    else
                    {
                        throw new System.ArgumentOutOfRangeException("Should never get here");
                    }                    
                }
            }
        }
    }
}