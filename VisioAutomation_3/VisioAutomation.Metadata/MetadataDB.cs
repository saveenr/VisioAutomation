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
        private List<AutomationEnum> _autoenums;
        private Dictionary<string,AutomationEnum> _name_to_autoenums;
        private List<AutomationConstant> _constants;
        private ExcelXmlToDataSetConverter converter;
        private Dictionary<string, AutomationConstant> _name_to_constants;
        
        public MetadataDB()
        {
            this.converter = new ExcelUtil.ExcelXmlToDataSetConverter();
            initconstants();
            this.initautoenums();
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

        public List<AutomationConstant> Constants
        {
            get { return this._constants; }
        }

        public List<AutomationEnum> AutomationEnums
        {
            get { return this._autoenums; }
        }

        public AutomationEnum GetAutomationEnumByName(string name)
        {
            return this._name_to_autoenums[name];
        }

        public AutomationConstant GetAutomationConstantByName(string name)
        {
            return this._name_to_constants[name];
        }

        public List<CellValue> CellValues
        {
            get { return this._cellvals; }
        }

        private void initautoenums()
        {
            this._autoenums = new List<AutomationEnum>();
            this._name_to_autoenums = new Dictionary<string, AutomationEnum>();
            foreach (var constant in this.Constants)
            {
                var enum_name = constant.Enum;
                AutomationEnum a_enum=null;
                this._name_to_autoenums.TryGetValue(constant.Enum, out a_enum);
                if (a_enum ==null)
                {
                    a_enum = new AutomationEnum(constant.Enum);
                    this._name_to_autoenums[a_enum.Name] = a_enum;
                    this._autoenums.Add(a_enum);
                }

                a_enum.Add(constant.Name,constant.Value);
            }
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
                c.NameCode = item.Field<string>("NameCode");
                if (c.NameCode == null || c.NameCode.Length == 0)
                {
                    c.NameCode = c.Name;
                }
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
            this._name_to_constants = new Dictionary<string, AutomationConstant>();
            foreach (var item in automationenums_table.AsEnumerable())
            {
                var c = new AutomationConstant();
                this._constants.Add(c);


                c.ID = item.Field<string>("ID");
                c.Enum = item.Field<string>("EnumName");
                c.Name = item.Field<string>("ValueName");
                c.Value = int.Parse(item.Field<string>("ValueInt"));

                this._name_to_constants[c.Name] = c;
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