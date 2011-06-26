using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Data;
using ExcelUtil;

namespace VisioAutomation.Metadata
{
    public class MetadataDB
    {
        ExcelXmlToDataSetConverter converter = new ExcelUtil.ExcelXmlToDataSetConverter();

        public System.Data.DataTable GetCells()
        {
            converter.Parse(VisioAutomation.Metadata.Properties.Resources.cells);
            var cells_table = converter.DataSet.Tables[0];
            return cells_table;
        }

        public System.Data.DataTable GetCellValues()
        {
            converter.Parse(VisioAutomation.Metadata.Properties.Resources.cellvalues);
            var cellvalues_table = converter.DataSet.Tables[0];
            return cellvalues_table;
        }

        public System.Data.DataTable GetSections()
        {
            converter.Parse(VisioAutomation.Metadata.Properties.Resources.sections);
            var sections_table = converter.DataSet.Tables[0];
            return sections_table;
        }

        public System.Data.DataTable GetAutomationEnums()
        {
            converter.Parse(VisioAutomation.Metadata.Properties.Resources.automationenums);
            var automationenums_table = converter.DataSet.Tables[0];
            return automationenums_table;
        }

    }
}
