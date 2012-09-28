using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class UserDefinedCellCommands : CommandSet
    {
        public UserDefinedCellCommands(Session session) :
            base(session)
        {

        }
        
        public IDictionary<IVisio.Shape, IList<VA.UserDefinedCells.UserDefinedCell>> GetUserDefinedCells()
        {
            var prop_dic = new Dictionary<IVisio.Shape, IList<VA.UserDefinedCells.UserDefinedCell>>();
            if (!this.Session.HasSelectedShapes())
            {
                return prop_dic;
            }

            var shapes = this.Session.Selection.EnumShapes().ToList();
            var application = this.Session.VisioApplication;
            var page = application.ActivePage;
            var list_user_props = VA.UserDefinedCells.UserDefinedCellsHelper.GetUserDefinedCells(page, shapes);

            for (int i = 0; i < shapes.Count; i++)
            {
                var shape = shapes[i];
                var props = list_user_props[i];
                prop_dic[shape] = props;
            }

            return prop_dic;
        }

        public IList<bool> Contains(string name)
        {
            if (name == null)
            {
                throw new System.ArgumentNullException("name");
            }

            if (!this.Session.HasSelectedShapes())
            {
                return new List<bool>();
            }

            var results = (from s in this.Session.Selection.EnumShapes().ToList()
                           select VA.UserDefinedCells.UserDefinedCellsHelper.HasUserDefinedCell(s, name))
                .ToList();

            return results;
        }

        public void Delete(string name)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            if (name == null)
            {
                throw new System.ArgumentNullException("name");
            }

            if (name.Length < 1)
            {
                throw new System.ArgumentException("name");
            }

            var shapes = this.Session.Selection.EnumShapes().ToList();

            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                foreach (var shape in shapes)
                {
                    VA.UserDefinedCells.UserDefinedCellsHelper.DeleteUserDefinedCell(shape, name);
                }
            }
        }

        public void Set(VA.UserDefinedCells.UserDefinedCell userdefinedcell)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            if (userdefinedcell == null)
            {
                throw new System.ArgumentNullException("userdefinedcell");
            }

            var shapes = this.Session.Selection.EnumShapes().ToList();

            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                foreach (var shape in shapes)
                {
                    VA.UserDefinedCells.UserDefinedCellsHelper.SetUserDefinedCell(shape, userdefinedcell.Name, userdefinedcell.Value, userdefinedcell.Prompt);
                }
            }
        }

    }
}