using System.Collections.Generic;
using System.Linq;
using VA_UDC = VisioAutomation.Shapes.UserDefinedCells;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting.Commands
{
    public class UserDefinedCellCommands : CommandSet
    {
        internal UserDefinedCellCommands(Client client) :
            base(client)
        {

        }

        public IDictionary<IVisio.Shape, IList<VA_UDC.UserDefinedCell>> Get(TargetShapes targets)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var prop_dic = new Dictionary<IVisio.Shape, IList<VA_UDC.UserDefinedCell>>();

            var shapes = targets.ResolveShapesEx(this._client);

            if (shapes.Shapes.Count < 1)
            {
                return prop_dic;
            }

            var application = this._client.Application.Get();
            var page = application.ActivePage;
            var list_user_props = VA_UDC.UserDefinedCellHelper.Get(page, shapes.Shapes);

            for (int i = 0; i < shapes.Shapes.Count; i++)
            {
                var shape = shapes.Shapes[i];
                var props = list_user_props[i];
                prop_dic[shape] = props;
            }

            return prop_dic;
        }

        public IList<bool> Contains(TargetShapes targets, string name)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            var shapes = targets.ResolveShapesEx(this._client);

            if (shapes.Shapes.Count < 1)
            {
                return new List<bool>();
            }

            var all_shapes = this._client.Selection.GetShapes();
            var results = all_shapes.Select(s => VA_UDC.UserDefinedCellHelper.Contains(s, name)).ToList();

            return results;
        }
       
        public void Delete(TargetShapes targets, string name)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = targets.ResolveShapesEx(this._client);

            if (shapes.Shapes.Count < 1)
            {
                return;
            } 

            if (name == null)
            {
                throw new System.ArgumentNullException("name cannot be null","name");
            }

            if (name.Length < 1)
            {
                throw new System.ArgumentException("name cannot be empty", nameof(name));
            }

            using (var undoscope = this._client.Application.NewUndoScope("Delete User-Defined Cell"))
            {
                foreach (var shape in shapes.Shapes)
                {
                    VA_UDC.UserDefinedCellHelper.Delete(shape, name);
                }
            }
        }

        public void Set(TargetShapes targets, VA_UDC.UserDefinedCell userdefinedcell)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = targets.ResolveShapesEx(this._client);

            if (shapes.Shapes.Count < 1)
            {
                return;
            }

            using (var undoscope = this._client.Application.NewUndoScope("Set User-Defined Cell"))
            {
                foreach (var shape in shapes.Shapes)
                {
                    VA_UDC.UserDefinedCellHelper.Set(shape, userdefinedcell.Name, userdefinedcell.Value.Formula.Value, userdefinedcell.Prompt.Formula.Value);
                }
            }
        }
    }
}