using System.Collections.Generic;
using System.Linq;
using VA=VisioAutomation;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class UserDefinedCellCommands : CommandSet
    {
        internal UserDefinedCellCommands(Client client) :
            base(client)
        {

        }

        public Dictionary<IVisio.Shape, VA.Shapes.UserDefinedCellDictionary> GetUserDefinedCells(Models.TargetShapes targets, VASS.CellValueType cvt)
        {
            var cmdtarget = this._client.GetCommandTargetPage();
            var dicof_shape_to_udcelldic = new Dictionary<IVisio.Shape, VA.Shapes.UserDefinedCellDictionary>();

            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return dicof_shape_to_udcelldic;
            }

            var page = cmdtarget.ActivePage;
            var shapeidpairs = VisioAutomation.ShapeIDPairs.FromShapes(targets.Shapes);
            var listof_udcelldic = VA.Shapes.UserDefinedCellHelper.GetCellsAsDictionary((IVisio.Page) page , shapeidpairs, cvt);

            for (int i = 0; i < targets.Shapes.Count; i++)
            {
                var shape = targets.Shapes[i];
                var props = listof_udcelldic[i];
                dicof_shape_to_udcelldic[shape] = props;
            }

            return dicof_shape_to_udcelldic;
        }

        public List<bool> ContainsUserDefinedCellsWithName(Models.TargetShapes targets, string name)
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
            var results = all_shapes.Select(s => VA.Shapes.UserDefinedCellHelper.Contains(s, name)).ToList();

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
                throw new System.ArgumentException("name cannot be empty", nameof(name));
            }

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DeleteUserDefinedCellsByName)))
            {
                foreach (var shape in targets.Shapes)
                {
                    VA.Shapes.UserDefinedCellHelper.Delete(shape, name);
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
                    VA.Shapes.UserDefinedCellHelper.Set(shape, userdefinedcell.Name, userdefinedcell.Cells);
                }
            }
        }
    }
}