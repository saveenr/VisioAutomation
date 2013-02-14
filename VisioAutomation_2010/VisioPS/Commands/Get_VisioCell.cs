using System.Collections.Generic;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioCell")]
    public class Get_VisioCell : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true, ParameterSetName = "cell")]
        public VisioAutomation.ShapeSheet.SRC [] SRC;

        [SMA.Parameter(Position = 0, Mandatory = false, ParameterSetName = "section")]
        public IVisio.VisSectionIndices SectionIndex;

        [SMA.Parameter(Position = 0, Mandatory = false, ParameterSetName = "section")]
        public IVisio.VisCellIndices [] CellIndices;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter GetResults;

        [SMA.Parameter(Mandatory = false)]
        public ResultType ResultType= ResultType.Double;

        [SMA.Parameter(Mandatory = false)]
        public IList<Microsoft.Office.Interop.Visio.Shape> Shapes;

        protected override void ProcessRecord()
        {

            var scriptingsession = this.ScriptingSession;

            if (!this.GetResults)
            {
                if (this.SRC != null)
                {
                    var formulas = scriptingsession.ShapeSheet.QueryFormulas(this.Shapes, this.SRC);
                    this.WriteObject(formulas);
                }
                else
                {
                    var formulas = scriptingsession.ShapeSheet.QueryFormulas(this.Shapes, this.SectionIndex, this.CellIndices);
                    this.WriteObject(formulas);
                }
            }
            {
                object results;

                if (this.SRC != null)
                {
                    if (this.ResultType == ResultType.Double)
                    {
                        results = scriptingsession.ShapeSheet.QueryResults<double>(this.Shapes, this.SRC);
                    }
                    else if (this.ResultType == ResultType.Integer)
                    {
                        results = scriptingsession.ShapeSheet.QueryResults<int>(this.Shapes, this.SRC);
                    }
                    else if (this.ResultType == ResultType.Boolean)
                    {
                        results = scriptingsession.ShapeSheet.QueryResults<bool>(this.Shapes, this.SRC);
                    }
                    else if (this.ResultType == ResultType.String)
                    {
                        results = scriptingsession.ShapeSheet.QueryResults<string>(this.Shapes, this.SRC);
                    }
                    else
                    {
                        results = scriptingsession.ShapeSheet.QueryResults<double>(this.Shapes, this.SRC);
                    }
                    this.WriteObject(results);
                }
                else
                {

                    if (this.ResultType == ResultType.Double)
                    {
                        results = scriptingsession.ShapeSheet.QueryResults<double>(this.Shapes, this.SectionIndex, this.CellIndices);
                    }
                    else if (this.ResultType == ResultType.Integer)
                    {
                        results = scriptingsession.ShapeSheet.QueryResults<int>(this.Shapes, this.SectionIndex, this.CellIndices);
                    }
                    else if (this.ResultType == ResultType.Boolean)
                    {
                        results = scriptingsession.ShapeSheet.QueryResults<bool>(this.Shapes, this.SectionIndex, this.CellIndices);
                    }
                    else if (this.ResultType == ResultType.String)
                    {
                        results = scriptingsession.ShapeSheet.QueryResults<string>(this.Shapes, this.SectionIndex, this.CellIndices);
                    }
                    else
                    {
                        results = scriptingsession.ShapeSheet.QueryResults<double>(this.Shapes, this.SectionIndex, this.CellIndices);
                    }
                    this.WriteObject(results);

                }
                
            }

        }
    }
}