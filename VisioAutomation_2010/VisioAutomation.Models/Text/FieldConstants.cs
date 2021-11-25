

namespace VisioAutomation.Models.Text;

public static class FieldConstants
{
    public static Field Angle => new Field(IVisio.VisFieldCategories.visFCatGeometry, IVisio.VisFieldCodes.visFCodeNumberOfPages, IVisio.VisFieldFormats.visFmtNumGenNoUnits);
    public static Field BackgroundName => new Field(IVisio.VisFieldCategories.visFCatPage, IVisio.VisFieldCodes.visFCodeBackgroundName, IVisio.VisFieldFormats.visFmtNumGenNoUnits);
    public static Field Category => new Field(IVisio.VisFieldCategories.visFCatDocument, IVisio.VisFieldCodes.visFCodeCategory, IVisio.VisFieldFormats.visFmtNumGenNoUnits);
    public static Field CreateDate => new Field(IVisio.VisFieldCategories.visFCatDateTime, IVisio.VisFieldCodes.visFCodeBackgroundName, IVisio.VisFieldFormats.visFmtNumGenNoUnits);
    public static Field Creator => new Field(IVisio.VisFieldCategories.visFCatDocument, IVisio.VisFieldCodes.visFCodeBackgroundName, IVisio.VisFieldFormats.visFmtNumGenNoUnits);
    public static Field CurrentDate => new Field(IVisio.VisFieldCategories.visFCatDateTime, IVisio.VisFieldCodes.visFCodeNumberOfPages, IVisio.VisFieldFormats.visFmtNumGenNoUnits);
    public static Field Description => new Field(IVisio.VisFieldCategories.visFCatDocument, IVisio.VisFieldCodes.visFCodeHeight, IVisio.VisFieldFormats.visFmtNumGenNoUnits);
    public static Field Directory => new Field(IVisio.VisFieldCategories.visFCatDocument, IVisio.VisFieldCodes.visFCodeNumberOfPages, IVisio.VisFieldFormats.visFmtNumGenNoUnits);
    public static Field EditDate => new Field(IVisio.VisFieldCategories.visFCatDateTime, IVisio.VisFieldCodes.visFCodeEditDate, IVisio.VisFieldFormats.visFmtNumGenNoUnits);
    public static Field Filename => new Field(IVisio.VisFieldCategories.visFCatDocument, IVisio.VisFieldCodes.visFCodeObjectID, IVisio.VisFieldFormats.visFmtNumGenNoUnits);
    public static Field Height => new Field(IVisio.VisFieldCategories.visFCatGeometry, IVisio.VisFieldCodes.visFCodeHeight, IVisio.VisFieldFormats.visFmtNumGenNoUnits);
    public static Field HyperlinkBase => new Field(IVisio.VisFieldCategories.visFCatDocument, IVisio.VisFieldCodes.visFCodeHyperlinkBase, IVisio.VisFieldFormats.visFmtNumGenNoUnits);
    public static Field KeyWords => new Field(IVisio.VisFieldCategories.visFCatDocument, IVisio.VisFieldCodes.visFCodeEditDate, IVisio.VisFieldFormats.visFmtNumGenNoUnits);
    public static Field MasterName => new Field(IVisio.VisFieldCategories.visFCatObject, IVisio.VisFieldCodes.visFCodeEditDate, IVisio.VisFieldFormats.visFmtNumGenNoUnits);
    public static Field NumberOfPages => new Field(IVisio.VisFieldCategories.visFCatPage, IVisio.VisFieldCodes.visFCodeNumberOfPages, IVisio.VisFieldFormats.visFmtNumGenNoUnits);
    public static Field ObjectID => new Field(IVisio.VisFieldCategories.visFCatObject, IVisio.VisFieldCodes.visFCodeObjectID, IVisio.VisFieldFormats.visFmtNumGenNoUnits);
    public static Field ObjectName => new Field(IVisio.VisFieldCategories.visFCatObject, IVisio.VisFieldCodes.visFCodeSubject, IVisio.VisFieldFormats.visFmtNumGenNoUnits);
    public static Field ObjectType => new Field(IVisio.VisFieldCategories.visFCatObject, IVisio.VisFieldCodes.visFCodePrintDate, IVisio.VisFieldFormats.visFmtNumGenNoUnits);
    public static Field PageName => new Field(IVisio.VisFieldCategories.visFCatPage, IVisio.VisFieldCodes.visFCodeHeight, IVisio.VisFieldFormats.visFmtNumGenNoUnits);
    public static Field PageNumber => new Field(IVisio.VisFieldCategories.visFCatPage, IVisio.VisFieldCodes.visFCodeObjectID, IVisio.VisFieldFormats.visFmtNumGenNoUnits);
    public static Field PrintDate => new Field(IVisio.VisFieldCategories.visFCatDateTime, IVisio.VisFieldCodes.visFCodePrintDate, IVisio.VisFieldFormats.visFmtNumGenNoUnits);
    public static Field Subject => new Field(IVisio.VisFieldCategories.visFCatDocument, IVisio.VisFieldCodes.visFCodeSubject, IVisio.VisFieldFormats.visFmtNumGenNoUnits);
    public static Field Title => new Field(IVisio.VisFieldCategories.visFCatDocument, IVisio.VisFieldCodes.visFCodePrintDate, IVisio.VisFieldFormats.visFmtNumGenNoUnits);
    public static Field Width => new Field(IVisio.VisFieldCategories.visFCatGeometry, IVisio.VisFieldCodes.visFCodeBackgroundName, IVisio.VisFieldFormats.visFmtNumGenNoUnits);
}