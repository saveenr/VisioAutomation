using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Shapes;
using VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class UserDefinedCellCommands : CommandSet
    {
        internal UserDefinedCellCommands(Client client) :
            base(client)
        {

        }

        public Dictionary<IVisio.Shape, Dictionary<string,UserDefinedCellCells>> GetUserDefinedCells(Models.TargetShapes targets, CellValueType cvt)
        {
            var cmdtarget = this._client.GetCommandTargetPage();
            var prop_dic = new Dictionary<IVisio.Shape, Dictionary<string, UserDefinedCellCells>>();

            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return prop_dic;
            }

            var page = cmdtarget.ActivePage;
            var list_user_props = UserDefinedCellHelper.GetCellsAsDictionary((IVisio.Page) page , targets.Shapes, cvt);

            for (int i = 0; i < targets.Shapes.Count; i++)
            {
                var shape = targets.Shapes[i];
                var props = list_user_props[i];
                prop_dic[shape] = props;
            }

            return prop_dic;
        }

        public List<bool> ShapesContainUserDefinedCellsWithName(Models.TargetShapes targets, string name)
        {
            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return new List<bool>();
            }

            var all_shapes = this._client.Selection.GetShapesInSelection();
            var results = all_shapes.Select(s => UserDefinedCellHelper.Contains(s, name)).ToList();

            return results;
        }
       
        public void DeleteUserDefinedCellsByName(Models.TargetShapes targets, string name)
        {
            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return;
            } 

            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            if (name.Length < 1)
            {
                throw new System.ArgumentException(nameof(name),"name cannot be empty");
            }

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DeleteUserDefinedCellsByName)))
            {
                foreach (var shape in targets.Shapes)
                {
                    UserDefinedCellHelper.Delete(shape, name);
                }
            }
        }

        public void SetUserDefinedCell(Models.TargetShapes targets, Models.UserDefinedCell userdefinedcell)
        {
            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return;
            }

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(SetUserDefinedCell)))
            {
                foreach (var shape in targets.Shapes)
                {
                    UserDefinedCellHelper.Set(shape, userdefinedcell.Name, userdefinedcell.Cells);
                }
            }
        }
    }
}