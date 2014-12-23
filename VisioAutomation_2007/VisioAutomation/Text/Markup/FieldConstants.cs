using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text.Markup
{
    public static class FieldConstants
    {
        public static Field ObjectName
        {
            get { return new Field(IVisio.VisFieldCategories.visFCatObject, IVisio.VisFieldCodes.visFCodeSubject, IVisio.VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field ObjectID
        {
            get { return new Field(IVisio.VisFieldCategories.visFCatObject, IVisio.VisFieldCodes.visFCodeObjectID, IVisio.VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field MasterName
        {
            get { return new Field(IVisio.VisFieldCategories.visFCatObject, IVisio.VisFieldCodes.visFCodeEditDate, IVisio.VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field ObjectType
        {
            get { return new Field(IVisio.VisFieldCategories.visFCatObject, IVisio.VisFieldCodes.visFCodePrintDate, IVisio.VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field Title
        {
            get { return new Field(IVisio.VisFieldCategories.visFCatDocument, IVisio.VisFieldCodes.visFCodePrintDate, IVisio.VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field Creator
        {
            get { return new Field(IVisio.VisFieldCategories.visFCatDocument, IVisio.VisFieldCodes.visFCodeBackgroundName, IVisio.VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field Description
        {
            get { return new Field(IVisio.VisFieldCategories.visFCatDocument, IVisio.VisFieldCodes.visFCodeHeight, IVisio.VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field Directory
        {
            get { return new Field(IVisio.VisFieldCategories.visFCatDocument, IVisio.VisFieldCodes.visFCodeNumberOfPages, IVisio.VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field Filename
        {
            get { return new Field(IVisio.VisFieldCategories.visFCatDocument, IVisio.VisFieldCodes.visFCodeObjectID, IVisio.VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field KeyWords
        {
            get { return new Field(IVisio.VisFieldCategories.visFCatDocument, IVisio.VisFieldCodes.visFCodeEditDate, IVisio.VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field Subject
        {
            get { return new Field(IVisio.VisFieldCategories.visFCatDocument, IVisio.VisFieldCodes.visFCodeSubject, IVisio.VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field Category
        {
            get { return new Field(IVisio.VisFieldCategories.visFCatDocument, IVisio.VisFieldCodes.visFCodeCategory, IVisio.VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field HyperlinkBase
        {
            get { return new Field(IVisio.VisFieldCategories.visFCatDocument, IVisio.VisFieldCodes.visFCodeHyperlinkBase, IVisio.VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field CreateDate
        {
            get { return new Field(IVisio.VisFieldCategories.visFCatDateTime, IVisio.VisFieldCodes.visFCodeBackgroundName, IVisio.VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field CurrentDate
        {
            get { return new Field(IVisio.VisFieldCategories.visFCatDateTime, IVisio.VisFieldCodes.visFCodeNumberOfPages, IVisio.VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field EditDate
        {
            get { return new Field(IVisio.VisFieldCategories.visFCatDateTime, IVisio.VisFieldCodes.visFCodeEditDate, IVisio.VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field PrintDate
        {
            get { return new Field(IVisio.VisFieldCategories.visFCatDateTime, IVisio.VisFieldCodes.visFCodePrintDate, IVisio.VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field BackgroundName
        {
            get { return new Field(IVisio.VisFieldCategories.visFCatPage, IVisio.VisFieldCodes.visFCodeBackgroundName, IVisio.VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field PageName
        {
            get { return new Field(IVisio.VisFieldCategories.visFCatPage, IVisio.VisFieldCodes.visFCodeHeight, IVisio.VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field NumberOfPages
        {
            get { return new Field(IVisio.VisFieldCategories.visFCatPage, IVisio.VisFieldCodes.visFCodeNumberOfPages, IVisio.VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field PageNumber
        {
            get { return new Field(IVisio.VisFieldCategories.visFCatPage, IVisio.VisFieldCodes.visFCodeObjectID, IVisio.VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field Width
        {
            get { return new Field(IVisio.VisFieldCategories.visFCatGeometry, IVisio.VisFieldCodes.visFCodeBackgroundName, IVisio.VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field Height
        {
            get { return new Field(IVisio.VisFieldCategories.visFCatGeometry, IVisio.VisFieldCodes.visFCodeHeight, IVisio.VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field Angle
        {
            get { return new Field(IVisio.VisFieldCategories.visFCatGeometry, IVisio.VisFieldCodes.visFCodeNumberOfPages, IVisio.VisFieldFormats.visFmtNumGenNoUnits); }
        }
    }
}