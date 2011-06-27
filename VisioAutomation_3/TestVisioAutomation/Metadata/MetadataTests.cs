using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{


    [TestClass]
    public class MetadataTests : VisioAutomationTest
    {
    
        [TestMethod]
        public void VerifyMetadaDBCreation()
        {
            var db = new VA.Metadata.MetadataDB();

            var allcells = db.GetCells();

            var dupe_cell_names = get_dupes(allcells.Select(c => c.Name));
            Assert.IsTrue( dupe_cell_names.Contains("Tabs"));
            Assert.IsTrue(dupe_cell_names.Contains("HideForApply"));

            Assert.AreEqual(346,allcells.Count);

            var visio_2007_cells = allcells.Where(c => c.MinVersion.Contains("Visio2007")).ToList();
            Assert.AreEqual(344, visio_2007_cells.Count());

        }

        public List<T> get_dupes<T>(IEnumerable<T> items)
        {
            var set = new HashSet<T>();
            var dupes = new List<T>();

            foreach (var item in items)
            {
                if (set.Contains(item))
                {
                    dupes.Add(item);
                }
                else
                {
                    set.Add(item);
                }
            }

            return dupes;
        }

        public void no_dupes<T>(IEnumerable<T> items)
        {
            var set = new HashSet<T>();
            var dupes = new List<T>();

            foreach (var item in items)
            {
                if (set.Contains(item))
                {
                    dupes.Add(item);
                }
                else
                {
                    set.Add(item);
                }
            }

            if (dupes.Count > 0)
            {
                Assert.Fail(string.Format("Duplicated {0}",dupes.Count));
            }
        }

    }
}
