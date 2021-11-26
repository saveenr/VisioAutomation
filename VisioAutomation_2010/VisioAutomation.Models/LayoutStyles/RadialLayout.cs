﻿using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.LayoutStyles
{
    public class RadialLayoutStyle : LayoutStyleBase
    {
        public RadialLayoutStyle()
        {
            this.ConnectorStyle = ConnectorStyle.RightAngle;
        }

        protected override void _set_page_cells(VisioAutomation.Pages.LayoutCells layout_cells)
        {
            base._set_page_cells(layout_cells);
            layout_cells.PlaceStyle = (int) IVisio.VisCellVals.visPLOPlaceDefault;
        }
    }
}