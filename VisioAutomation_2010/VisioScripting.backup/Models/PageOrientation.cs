﻿using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioScripting.Models
{
    public enum PageOrientation
    {
        SameAsPrinter = IVisio.VisCellVals.visPPOSameAsPrinter,
        Portrait = IVisio.VisCellVals.visPPOPortrait,
        Landscape = IVisio.VisCellVals.visPPOLandscape
    }
}