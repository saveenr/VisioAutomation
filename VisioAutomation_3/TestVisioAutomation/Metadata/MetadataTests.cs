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
        public void Verify_Shape_GetResults_For_Multiple_Types()
        {
            var converter = new ExcelUtil.ExcelXmlToDataSetConverter();

            converter.Parse(VisioAutomation.Metadata.Properties.Resources.cells);
            var cells_table = converter.DataSet.Tables[0];
            converter.Parse(VisioAutomation.Metadata.Properties.Resources.cellvalues);
            var cellvalues_table = converter.DataSet.Tables[0];
            converter.Parse(VisioAutomation.Metadata.Properties.Resources.sections);
            var sections_table = converter.DataSet.Tables[0];
            converter.Parse(VisioAutomation.Metadata.Properties.Resources.automationenums);
            var automationenums_table = converter.DataSet.Tables[0];
        }

    }
}
