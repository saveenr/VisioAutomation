using VA = VisioAutomation;
using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public struct TabStop
    {
        private static readonly VA.ShapeSheet.SRC src_tabstopcount = VA.ShapeSheet.SRCConstants.Tabs_StopCount;
        private static readonly short unitcode_nocast = (short) IVisio.VisUnitCodes.visNoCast;
        private const short tab_section = (short) IVisio.VisSectionIndices.visSectionTab;

        public double Position { get; private set; }
        public TabStopAlignment Alignment { get; private set; }

        public TabStop(double pos, VA.Text.TabStopAlignment align) : this()
        {
            this.Position = pos;
            this.Alignment = align;
        }

        public override string ToString()
        {
            string s = string.Format("(Position={0},Alignment={1})",
                                     this.Position,
                                     this.Alignment);
            return s;
        }


        public static void SetTabStops(IVisio.Shape shape, IList<TabStop> stops)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            if (stops == null)
            {
                throw new ArgumentNullException("stops");
            }

            ClearTabStops(shape);
            if (stops.Count < 1)
            {
                return;
            }

            const short row = 0;
            var invariant_culture = System.Globalization.CultureInfo.InvariantCulture;
            var vis_tab_stop_count = (short) IVisio.VisCellIndices.visTabStopCount;
            var tabstopcountcell = shape.CellsSRC[tab_section, row, vis_tab_stop_count];
            tabstopcountcell.FormulaU = stops.Count.ToString(invariant_culture);

            // set the number of tab stobs allowed for the shape
            var tagtab = GetTabTagForStops(stops.Count);
            shape.RowType[tab_section, (short) IVisio.VisRowIndices.visRowTab] = (short) tagtab;

            // add tab properties for each stop
            var update = new VA.ShapeSheet.Update.SRCUpdate();
            for (int stop_index = 0; stop_index < stops.Count; stop_index++)
            {
                int i = stop_index*3;

                var alignment = ((int) stops[stop_index].Alignment).ToString(invariant_culture);
                var position = ((int) stops[stop_index].Position).ToString(invariant_culture);

                var src_tabpos = new VA.ShapeSheet.SRC(tab_section, row, (short) (i + 1));
                var src_tabalign = new VA.ShapeSheet.SRC(tab_section, row, (short) (i + 2));
                var src_tabother = new VA.ShapeSheet.SRC(tab_section, row, (short) (i + 3));

                update.SetFormula(src_tabpos, position); // tab position
                update.SetFormula(src_tabalign, alignment); // tab alignment
                update.SetFormula(src_tabother, "0"); // tab unknown
            }

            update.Execute(shape);
        }

        public static IVisio.VisRowTags GetTabTagForStops(int stops)
        {
            if (stops < 0)
            {
                throw new ArgumentOutOfRangeException("stops");
            }

            var tagtab = IVisio.VisRowTags.visTagTab0;
            if ((0 <= stops) && (stops <= 2))
            {
                tagtab = IVisio.VisRowTags.visTagTab2;
            }
            else if ((3 <= stops) && (stops <= 10))
            {
                tagtab = IVisio.VisRowTags.visTagTab10;
            }
            else if ((11 <= stops) && (stops <= 60))
            {
                tagtab = IVisio.VisRowTags.visTagTab60;
            }
            else
            {
                throw new ArgumentOutOfRangeException("stops", "unsupported number of tabs");
            }

            return tagtab;
        }

        public static int GetTabStopCount(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            var tcell = shape.GetCell(src_tabstopcount);
            const short rounding = 0;
            return tcell.ResultInt[unitcode_nocast, rounding];
        }

        /// <summary>
        /// Remove all tab stops on the shape
        /// </summary>
        /// <param name="shape"></param>
        public static void ClearTabStops(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            int num_existing_tabstops = GetTabStopCount(shape);

            if (num_existing_tabstops < 1)
            {
                return;
            }

            var cell_tabstopcount = shape.GetCell(src_tabstopcount);
            cell_tabstopcount.FormulaForce = "0";

            const string formula = "0";

            var update = new VA.ShapeSheet.Update.SRCUpdate();
            for (int i = 1; i < num_existing_tabstops*3; i++)
            {
                var src = new VA.ShapeSheet.SRC(tab_section, (short) IVisio.VisRowIndices.visRowTab,
                                                (short) i);
                update.SetFormula(src, formula);
            }

            update.Execute(shape);
        }
    }
}