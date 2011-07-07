using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Linq;

namespace TestVisioAutomation
{
    public class TestHelper
    {
        public readonly string OutputPath;

        public TestHelper(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                throw new System.ArgumentException("name");
            }

            this.OutputPath = GetOutputPathEx(name);

            PrepareOutputPath();
        }

        public string GetTestMethodOutputFilename(string ext)
        {
            string abs_path = this.OutputPath;
            string abs_filename = System.IO.Path.Combine(abs_path, GetMethodName(2) + ext);
            return abs_filename;
        }

        public string GetTestMethodOutputFilename()
        {
            string abs_path = this.OutputPath;
            string abs_filename = System.IO.Path.Combine(abs_path, GetMethodName(2));
            return abs_filename;
        }

        private void PrepareOutputPath()
        {
            if (!System.IO.Directory.Exists(this.OutputPath))
            {
                System.IO.Directory.CreateDirectory(this.OutputPath);
            }
        }

        private static string GetOutputPathEx(string name)
        {
            string path = System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments);
            return System.IO.Path.Combine(path, name);
        }

        private static string GetMethodName(int depth)
        {
            var stackTrace = new System.Diagnostics.StackTrace();
            var stackFrame = stackTrace.GetFrame(depth);
            var methodBase = stackFrame.GetMethod();

            string n = methodBase.DeclaringType.Name + "." + methodBase.Name;
            return n;
        }

        private static string GetMethodName()
        {
            return GetMethodName(2);
        }

        private static string GetMethodName(string ext)
        {
            return GetMethodName(2) + ext;
        }

        public static void AreEqual(double x, double y, VA.Drawing.Point p, double delta)
        {
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(x, p.X, delta);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(y, p.Y, delta);
        }

        public static void AreEqual(double x, double y, VA.Drawing.Size p, double delta)
        {
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(x, p.Width, delta);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(y, p.Height, delta);
        }

        public static void SetFormulasU<T>(
            IVisio.Shape shape,
            IEnumerable<T> items,
            System.Func<T, bool> hasformula,
            System.Func<T, VA.ShapeSheet.SRC> getsrc,
            System.Func<T, string> getformula,
            IVisio.VisGetSetArgs flags)
        {
            if (items == null)
            {
                throw new System.ArgumentNullException("items");
            }

            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            var update = new VA.ShapeSheet.Update.SRCUpdate();
            update.BlastGuards = ((short) flags & (short) IVisio.VisGetSetArgs.visSetBlastGuards) != 0;
            update.TestCircular = ((short) flags & (short) IVisio.VisGetSetArgs.visSetTestCircular) != 0;

            foreach (var item in items.Where(hasformula))
            {
                update.SetFormula(getsrc(item), getformula(item));
            }

            update.Execute(shape);
        }

        public static VA.ShapeSheet.Query.CellQuery BuildCellQuery(IList<VA.ShapeSheet.SRC> srcs)
        {
            var query = new VA.ShapeSheet.Query.CellQuery();
            foreach (var src in srcs)
            {
                query.AddColumn(src);
            }
            return query;
        }

        public static void setformulas(VA.DOM.ShapeCells shapecells, IVisio.Page page, IVisio.Shape shape)
        {
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();
            shapecells.Apply(update, shape.ID16);
            update.Execute(page);
        }
    }
}