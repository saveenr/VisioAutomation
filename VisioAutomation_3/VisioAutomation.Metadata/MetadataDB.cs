using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using ExcelUtil;

namespace VisioAutomation.Metadata
{
    public class MetadataDB
    {
        private List<Cell> _cells;
        private List<CellValue> _cellvals;
        private List<Section> _sections;
        private List<AutomationConstant> _constants;
        private ExcelXmlToDataSetConverter converter;
 
        public MetadataDB()
        {
            this.converter = new ExcelUtil.ExcelXmlToDataSetConverter();
            initconstants();
            initcellvalues();
            initcells();
            initsections();
        }

        public List<Cell> Cells
        {
            get { return _cells; }
        }

        public List<Section> Sections
        {
            get { return this._sections; }
        }

        public List<AutomationConstant> AutomationEnums
        {
            get { return this._constants; }
        }

        public List<CellValue> CellValues
        {
            get { return this._cellvals; }
        }

        private void initcells()
        {
            converter.Parse(VisioAutomation.Metadata.Properties.Resources.cells);
            var cells_table = converter.DataSet.Tables[0];
            _cells = new List<Cell>();
            foreach (var item in cells_table.AsEnumerable())
            {
                var c = new Cell();
                _cells.Add(c);
                c.ID = item.Field<string>("ID");
                c.Name = item.Field<string>("Name");
                c.NameFormatString = item.Field<string>("NameFormatString");
                c.Object = item.Field<string>("Object");
                c.NameType = item.Field<string>("NameType");
                c.DataType = item.Field<string>("DataType");
                c.ContentType = item.Field<string>("ContentType");
                c.Unit = item.Field<string>("Unit");
                c.SectionIndex = item.Field<string>("SectionIndex");
                c.RowIndex = item.Field<string>("RowIndex");
                c.MinVersion = item.Field<string>("MinVersion");
                c.MaxVersion = item.Field<string>("MaxVersion");
                c.CellIndex = item.Field<string>("CellIndex");
                c.MSDN = item.Field<string>("MSDN");
            }
        }

        private void initsections()
        {
            converter.Parse(VisioAutomation.Metadata.Properties.Resources.sections);
            var sections_table = converter.DataSet.Tables[0];
            this._sections = new List<Section>();
            foreach (var item in sections_table.AsEnumerable())
            {
                var c = new Section();
                this._sections.Add(c);
                c.ID = item.Field<string>("ID");
                c.DisplayName = item.Field<string>("DisplayName");
                c.Name = item.Field<string>("Name");
            }
        }

        private void initconstants()
        {
            converter.Parse(VisioAutomation.Metadata.Properties.Resources.automationconstants);
            var automationenums_table = converter.DataSet.Tables[0];

            this._constants = new List<AutomationConstant>();
            foreach (var item in automationenums_table.AsEnumerable())
            {
                var c = new AutomationConstant();
                this._constants.Add(c);
                c.ID = item.Field<string>("ID");
                c.Enum = item.Field<string>("EnumName");
                c.Name = item.Field<string>("ValueName");
                c.Value = int.Parse(item.Field<string>("ValueInt"));
            }
        }


        private void initcellvalues()
        {
            converter.Parse(VisioAutomation.Metadata.Properties.Resources.cellvalues);
            var cellvalues_table = converter.DataSet.Tables[0];
            this._cellvals = new List<CellValue>();
            foreach (var item in cellvalues_table.AsEnumerable())
            {
                var c = new CellValue();
                this._cellvals.Add(c);
                c.ID = item.Field<string>("ID");
                c.Enum = item.Field<string>("Enum");
                c.Name = item.Field<string>("Name");

                bool s;
                int v;
                s = int.TryParse(item.Field<string>("Value"), out v);
                if (s)
                {
                    c.Value = v;
                }
                else
                {
                    c.Value = null;
                }

                c.AutomationConstant = item.Field<string>("AutomationConstant");
            }
        }
    }
}