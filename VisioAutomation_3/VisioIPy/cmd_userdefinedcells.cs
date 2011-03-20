using VAS = VisioAutomation.Scripting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioIPy
{
    public partial class VisioIPySession
    {
        public IDictionary<IVisio.Shape, IList<VA.UserDefinedCells.UserDefinedCell>> GetUserDefinedCells()
        {
            var prop_dic = this.ScriptingSession.UserDefinedCell.GetUserDefinedCells();
            return prop_dic;
        }

        public void SetUserDefinedCell(string name, string value)
        {
            var prop = new VA.UserDefinedCells.UserDefinedCell(name, value);
            this.ScriptingSession.UserDefinedCell.SetUserDefinedCell(prop);
        }

        public void DeleteUserDefinedCell(string name)
        {
            this.ScriptingSession.UserDefinedCell.DeleteUserDefinedCell(name);
        }
    }
}