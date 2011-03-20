using Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text.Markup
{
    public static class Fields
    {
        public static readonly Field ObjectName = new Field(VisFieldCategories.visFCatObject, VisFieldCodes.visFCodeSubject, VisFieldFormats.visFmtNumGenNoUnits);
        public static readonly Field ObjectID = new Field(VisFieldCategories.visFCatObject, VisFieldCodes.visFCodeObjectID, VisFieldFormats.visFmtNumGenNoUnits);
        public static readonly Field MasterName = new Field(VisFieldCategories.visFCatObject, VisFieldCodes.visFCodeEditDate, VisFieldFormats.visFmtNumGenNoUnits);
        public static readonly Field ObjectType = new Field(VisFieldCategories.visFCatObject, VisFieldCodes.visFCodePrintDate, VisFieldFormats.visFmtNumGenNoUnits);
        public static readonly Field Title = new Field(VisFieldCategories.visFCatDocument, VisFieldCodes.visFCodePrintDate, VisFieldFormats.visFmtNumGenNoUnits);
        public static readonly Field Creator = new Field(VisFieldCategories.visFCatDocument, VisFieldCodes.visFCodeBackgroundName, VisFieldFormats.visFmtNumGenNoUnits);
        public static readonly Field Description = new Field(VisFieldCategories.visFCatDocument, VisFieldCodes.visFCodeHeight, VisFieldFormats.visFmtNumGenNoUnits);
        public static readonly Field Directory = new Field(VisFieldCategories.visFCatDocument, VisFieldCodes.visFCodeNumberOfPages, VisFieldFormats.visFmtNumGenNoUnits);
        public static readonly Field Filename = new Field(VisFieldCategories.visFCatDocument, VisFieldCodes.visFCodeObjectID, VisFieldFormats.visFmtNumGenNoUnits);
        public static readonly Field KeyWords = new Field(VisFieldCategories.visFCatDocument, VisFieldCodes.visFCodeEditDate, VisFieldFormats.visFmtNumGenNoUnits);
        public static readonly Field Subject = new Field(VisFieldCategories.visFCatDocument, VisFieldCodes.visFCodeSubject, VisFieldFormats.visFmtNumGenNoUnits);
        public static readonly Field Category = new Field(VisFieldCategories.visFCatDocument, VisFieldCodes.visFCodeCategory, VisFieldFormats.visFmtNumGenNoUnits);
        public static readonly Field HyperlinkBase = new Field(VisFieldCategories.visFCatDocument, VisFieldCodes.visFCodeHyperlinkBase, VisFieldFormats.visFmtNumGenNoUnits);
        public static readonly Field CreateDate = new Field(VisFieldCategories.visFCatDateTime, VisFieldCodes.visFCodeBackgroundName, VisFieldFormats.visFmtNumGenNoUnits);
        public static readonly Field CurrentDate = new Field(VisFieldCategories.visFCatDateTime, VisFieldCodes.visFCodeNumberOfPages, VisFieldFormats.visFmtNumGenNoUnits);
        public static readonly Field EditDate = new Field(VisFieldCategories.visFCatDateTime, VisFieldCodes.visFCodeEditDate, VisFieldFormats.visFmtNumGenNoUnits);
        public static readonly Field PrintDate = new Field(VisFieldCategories.visFCatDateTime, VisFieldCodes.visFCodePrintDate, VisFieldFormats.visFmtNumGenNoUnits);
        public static readonly Field BackgroundName = new Field(VisFieldCategories.visFCatPage, VisFieldCodes.visFCodeBackgroundName, VisFieldFormats.visFmtNumGenNoUnits);
        public static readonly Field PageName = new Field(VisFieldCategories.visFCatPage, VisFieldCodes.visFCodeHeight, VisFieldFormats.visFmtNumGenNoUnits);
        public static readonly Field NumberOfPages = new Field(VisFieldCategories.visFCatPage, VisFieldCodes.visFCodeNumberOfPages, VisFieldFormats.visFmtNumGenNoUnits);
        public static readonly Field PageNumber = new Field(VisFieldCategories.visFCatPage, VisFieldCodes.visFCodeObjectID, VisFieldFormats.visFmtNumGenNoUnits);
        public static readonly Field Width = new Field(VisFieldCategories.visFCatGeometry, VisFieldCodes.visFCodeBackgroundName, VisFieldFormats.visFmtNumGenNoUnits);
        public static readonly Field Height = new Field(VisFieldCategories.visFCatGeometry, VisFieldCodes.visFCodeHeight, VisFieldFormats.visFmtNumGenNoUnits);
        public static readonly Field Angle = new Field(VisFieldCategories.visFCatGeometry, VisFieldCodes.visFCodeNumberOfPages, VisFieldFormats.visFmtNumGenNoUnits);
    }
}