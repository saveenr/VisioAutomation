using Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text.Markup
{
    public static class Fields
    {
        public static Field ObjectName
        {
            get { return new Field(VisFieldCategories.visFCatObject, VisFieldCodes.visFCodeSubject, VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field ObjectID
        {
            get { return new Field(VisFieldCategories.visFCatObject, VisFieldCodes.visFCodeObjectID, VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field MasterName
        {
            get { return new Field(VisFieldCategories.visFCatObject, VisFieldCodes.visFCodeEditDate, VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field ObjectType
        {
            get { return new Field(VisFieldCategories.visFCatObject, VisFieldCodes.visFCodePrintDate, VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field Title
        {
            get { return new Field(VisFieldCategories.visFCatDocument, VisFieldCodes.visFCodePrintDate, VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field Creator
        {
            get { return new Field(VisFieldCategories.visFCatDocument, VisFieldCodes.visFCodeBackgroundName, VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field Description
        {
            get { return new Field(VisFieldCategories.visFCatDocument, VisFieldCodes.visFCodeHeight, VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field Directory
        {
            get { return new Field(VisFieldCategories.visFCatDocument, VisFieldCodes.visFCodeNumberOfPages, VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field Filename
        {
            get { return new Field(VisFieldCategories.visFCatDocument, VisFieldCodes.visFCodeObjectID, VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field KeyWords
        {
            get { return new Field(VisFieldCategories.visFCatDocument, VisFieldCodes.visFCodeEditDate, VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field Subject
        {
            get { return new Field(VisFieldCategories.visFCatDocument, VisFieldCodes.visFCodeSubject, VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field Category
        {
            get { return new Field(VisFieldCategories.visFCatDocument, VisFieldCodes.visFCodeCategory, VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field HyperlinkBase
        {
            get { return new Field(VisFieldCategories.visFCatDocument, VisFieldCodes.visFCodeHyperlinkBase, VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field CreateDate
        {
            get { return new Field(VisFieldCategories.visFCatDateTime, VisFieldCodes.visFCodeBackgroundName, VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field CurrentDate
        {
            get { return new Field(VisFieldCategories.visFCatDateTime, VisFieldCodes.visFCodeNumberOfPages, VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field EditDate
        {
            get { return new Field(VisFieldCategories.visFCatDateTime, VisFieldCodes.visFCodeEditDate, VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field PrintDate
        {
            get { return new Field(VisFieldCategories.visFCatDateTime, VisFieldCodes.visFCodePrintDate, VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field BackgroundName
        {
            get { return new Field(VisFieldCategories.visFCatPage, VisFieldCodes.visFCodeBackgroundName, VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field PageName
        {
            get { return new Field(VisFieldCategories.visFCatPage, VisFieldCodes.visFCodeHeight, VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field NumberOfPages
        {
            get { return new Field(VisFieldCategories.visFCatPage, VisFieldCodes.visFCodeNumberOfPages, VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field PageNumber
        {
            get { return new Field(VisFieldCategories.visFCatPage, VisFieldCodes.visFCodeObjectID, VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field Width
        {
            get { return new Field(VisFieldCategories.visFCatGeometry, VisFieldCodes.visFCodeBackgroundName, VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field Height
        {
            get { return new Field(VisFieldCategories.visFCatGeometry, VisFieldCodes.visFCodeHeight, VisFieldFormats.visFmtNumGenNoUnits); }
        }

        public static Field Angle
        {
            get { return new Field(VisFieldCategories.visFCatGeometry, VisFieldCodes.visFCodeNumberOfPages, VisFieldFormats.visFmtNumGenNoUnits); }
        }
    }
}