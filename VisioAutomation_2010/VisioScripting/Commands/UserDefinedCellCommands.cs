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

        public Dictionary<IVisio.Shape, VA.Shapes.UserDefinedCellDictionary> GetUserDefinedCells_ShapeDictionary(TargetShapes targetshapes, VASS.CellValueType cvt)
        {
            targetshapes = targetshapes.Resolve(this._client);
            var listof_udcelldic = GetUserDefinedCells_List(targetshapes, cvt);

            var dicof_shape_to_udcelldic = new Dictionary<IVisio.Shape, VA.Shapes.UserDefinedCellDictionary>();
            for (int i = 0; i < targetshapes.Shapes.Count; i++)
            {
                var shape = targetshapes.Shapes[i];
                var props = listof_udcelldic[i];
                dicof_shape_to_udcelldic[shape] = props;
            }

            return dicof_shape_to_udcelldic;
        }

        public List<VA.Shapes.UserDefinedCellDictionary> GetUserDefinedCells_List(TargetShapes targetshapes, VASS.CellValueType cvt)
        {
            var cmdtarget = this._client.GetCommandTargetPage();

            targetshapes = targetshapes.Resolve(this._client);

            if (targetshapes.Shapes.Count < 1)
            {
                return new List<VA.Shapes.UserDefinedCellDictionary>(0);
            }

            var page = cmdtarget.ActivePage;
            var shapeidpairs = targetshapes.ToShapeIDPairs();
            var listof_udcelldic = VA.Shapes.UserDefinedCellHelper.GetDictionary((IVisio.Page)page, shapeidpairs, cvt);

            return listof_udcelldic;
        }

        public List<bool> ContainsUserDefinedCellsWithName(TargetShapes targetshapes, string name)
        {
            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            targetshapes = targetshapes.Resolve(this._client);

            if (targetshapes.Shapes.Count < 1)
            {
                return new List<bool>();
            }

            var all_shapes = this._client.Selection.GetShapes(new TargetSelection());
            var results = all_shapes.Select(s => VA.Shapes.UserDefinedCellHelper.Contains(s, name)).ToList();

            return results;
        }
       
        public void DeleteUserDefinedCellsByName(TargetShapes targetshapes, string name)
        {
            targetshapes = targetshapes.Resolve(this._client);

            if (targetshapes.Shapes.Count < 1)
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
                foreach (var shape in targetshapes.Shapes)
                {
                    VA.Shapes.UserDefinedCellHelper.Delete(shape, name);
                }
            }
        }

        public void SetUserDefinedCell(TargetShapes targetshapes, string name, VA.Shapes.UserDefinedCellCells udcellcells)
        {
            targetshapes = targetshapes.Resolve(this._client);

            if (targetshapes.Shapes.Count < 1)
            {
                return;
            }

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(SetUserDefinedCell)))
            {
                foreach (var shape in targetshapes.Shapes)
                {
                    VA.Shapes.UserDefinedCellHelper.Set(shape, name, udcellcells);
                }
            }
        }
    }
}