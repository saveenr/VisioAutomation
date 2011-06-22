using System;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.ShapeSheet
{
    public static partial class ShapeSheetHelper
    {
        public static string TryGetNameFromSRC(VA.ShapeSheet.SRC src)
        {
            switch (src.Section)
            {
                case ((short)IVisio.VisSectionIndices.visSectionObject):
                    {
                        return TryGetNameFromSRC_Section_Object(src);
                    }
                default:
                    break;
            }
            return null;
        }

        private static string TryGetNameFromSRC_Section_Object(VA.ShapeSheet.SRC src)
        {
            switch (src.Row)
            {
                case ((short)IVisio.VisRowIndices.visRowFill):
                    {
                        switch (src.Cell)
                        {
                            case ((short)IVisio.VisCellIndices.visFillBkgnd): return "FillBkgnd";
                            case ((short)IVisio.VisCellIndices.visFillBkgndTrans): return "BkgndTrans";
                            case ((short)IVisio.VisCellIndices.visFillForegnd): return "FillForegnd";
                            case ((short)IVisio.VisCellIndices.visFillForegndTrans): return "ForegndTrans";
                            case ((short)IVisio.VisCellIndices.visFillPattern): return "FillPattern";
                            case ((short)IVisio.VisCellIndices.visFillShdwBkgnd): return "ShdwBkgnd";
                            case ((short)IVisio.VisCellIndices.visFillShdwBkgndTrans): return "ShdwBkgndTrans";
                            case ((short)IVisio.VisCellIndices.visFillShdwForegnd): return "ShdwForegnd";
                            case ((short)IVisio.VisCellIndices.visFillShdwForegndTrans): return "ShdwForegndTrans";
                            case ((short)IVisio.VisCellIndices.visFillShdwObliqueAngle): return "ShdwObliqueAngle";
                            case ((short)IVisio.VisCellIndices.visFillShdwOffsetX): return "ShdwOffsetX";
                            case ((short)IVisio.VisCellIndices.visFillShdwOffsetY): return "ShdwOffsetY";
                            case ((short)IVisio.VisCellIndices.visFillShdwPattern): return "ShdwPattern";
                            case ((short)IVisio.VisCellIndices.visFillShdwScaleFactor): return "ShdwScaleFactor";
                            case ((short)IVisio.VisCellIndices.visFillShdwType): return "visFillShdwType";
                            default:
                                break;
                        }
                        break;
                    }
                case ((short)IVisio.VisRowIndices.visRowLine):
                    {
                        switch (src.Cell)
                        {
                            case ((short)IVisio.VisCellIndices.visLineBeginArrow): return "BeginArrow";
                            case ((short)IVisio.VisCellIndices.visLineBeginArrowSize): return "BeginArrowSize";
                            case ((short)IVisio.VisCellIndices.visLineColor): return "LineColor";
                            case ((short)IVisio.VisCellIndices.visLineColorTrans): return "LineColorTrans";
                            case ((short)IVisio.VisCellIndices.visLineEndArrow): return "LineEndArrow";
                            case ((short)IVisio.VisCellIndices.visLineEndArrowSize): return "LineEndArrowSize";
                            case ((short)IVisio.VisCellIndices.visLineEndCap): return "LineEndCap";
                            case ((short)IVisio.VisCellIndices.visLinePattern): return "LinePattern";
                            case ((short)IVisio.VisCellIndices.visLineRounding): return "LineRounding";
                            case ((short)IVisio.VisCellIndices.visLineWeight): return "LineWeight";
                            default:
                                break;
                        }
                        break;
                    }
                default:
                    break;
            }
            return null;

        }

        public static SRC? TryGetSRCFromName(string name)
        {
            var dic = NameToSRCDictionary;
            SRC src;
            bool found_value = dic.TryGetValue(name, out src);

            if (found_value)
            {
                return src;
            }

            int dot_pos = name.IndexOf('.');
            if (dot_pos >= 0)
            {
                string left_of_dot = name.Substring(0, dot_pos);

                int left_bracket_pos = name.IndexOf('[');
                if (left_bracket_pos >= 0)
                {
                    string unbracketed_name = name.Substring(0, left_bracket_pos);
                    int len = left_bracket_pos - dot_pos - 1;
                    string between = name.Substring(dot_pos+1, len);

                    int right_bracket_pos = name.IndexOf(']');
                    if (right_bracket_pos > 0)
                    {
                        int between_brackets_len = right_bracket_pos - left_bracket_pos - 1;
                        string between_brackets_str = name.Substring(left_bracket_pos + 1, between_brackets_len);

                        if ((left_of_dot == "Char") || (left_of_dot == "Para"))
                        {
                            var x = TryGetSRCFromName(unbracketed_name);
                            if (x.HasValue)
                            {
                                int bracket_int = int.Parse(between_brackets_str);
                                var y = x.Value.ForRow((short) (bracket_int -1) );
                                return y;

                            }
                        }

                    }




                }

                if (left_of_dot == "Para")
                {
                }
            }

            return null;
        }

        public static SRC GetSRCFromName(string name)
        {
            var src = TryGetSRCFromName(name);
            if (src.HasValue)
            {
                return src.Value;
            }

            string msg = string.Format("Cannot identify indices for cell with name \"{0}\"", name);
            throw new AutomationException(msg);
        }

        private static Dictionary<string, SRC> NameToSRCDictionary
        {
            get
            {
                if (simple_name_to_src_map == null)
                {
                    CreateNameToSRCDictionary();
                }
                return simple_name_to_src_map;
            }
        }

        private static void CreateNameToSRCDictionary()
        {
            simple_name_to_src_map = new Dictionary<string, SRC>(StringComparer.OrdinalIgnoreCase)
                                             {
                                                 {"PinX",SRCConstants.PinX},
                                                 {"PinY",SRCConstants.PinY},
                                                 {"LocPinX",SRCConstants.LocPinX},
                                                 {"LocPinY",SRCConstants.LocPinY},
                                                 {"Width",SRCConstants.Width},
                                                 {"Height",SRCConstants.Height},
                                                 {"Angle",SRCConstants.Angle},
                                                 {"FlipX",SRCConstants.FlipX},
                                                 {"FlipY",SRCConstants.FlipY},
                                                 {"ResizeMode",SRCConstants.ResizeMode},


                                                 {"FillBkgnd",SRCConstants.FillBkgnd},
                                                 {"FillBkgndTrans",SRCConstants.FillBkgndTrans},
                                                 {"FillForegnd",SRCConstants.FillForegnd},
                                                 {"FillForegndTrans",SRCConstants.FillForegndTrans},
                                                 {"FillPattern",SRCConstants.FillPattern},
                                                 {"ShapeShdwObliqueAngle",SRCConstants.ShapeShdwObliqueAngle},
                                                 {"ShapeShdwOffsetX",SRCConstants.ShapeShdwOffsetX},
                                                 {"ShapeShdwOffsetY",SRCConstants.ShapeShdwOffsetY},
                                                 {"ShapeShdwScaleFactor",SRCConstants.ShapeShdwScaleFactor},
                                                 {"ShapeShdwType",SRCConstants.ShapeShdwType},
                                                 {"ShdwBkgnd",SRCConstants.ShdwBkgnd},
                                                 {"ShdwBkgndTrans",SRCConstants.ShdwBkgndTrans},
                                                 {"ShdwForegnd",SRCConstants.ShdwForegnd},
                                                 {"ShdwForegndTrans",SRCConstants.ShdwForegndTrans},
                                                 {"ShdwPattern",SRCConstants.ShdwPattern},
                                                 {"LineCap",SRCConstants.LineCap},
                                                 {"LineColor",SRCConstants.LineColor},
                                                 {"LineColorTrans",SRCConstants.LineColorTrans},
                                                 {"LineWeight",SRCConstants.LineWeight},
                                                 {"LinePattern",SRCConstants.LinePattern},
                                                 {"Rounding",SRCConstants.Rounding},
                                                 {"BeginArrow",SRCConstants.BeginArrow},
                                                 {"BeginArrowSize",SRCConstants.BeginArrowSize},
                                                 {"EndArrow",SRCConstants.EndArrow},
                                                 {"EndArrowSize",SRCConstants.EndArrowSize},
                                                 {"BeginX",SRCConstants.BeginX},
                                                 {"BeginY",SRCConstants.BeginY},
                                                 {"EndX",SRCConstants.EndX},
                                                 {"EndY",SRCConstants.EndY},
                                                 {"Char.Case",SRCConstants.Char_Case},
                                                 {"Char.Color",SRCConstants.Char_Color},
                                                 {"Char.ColorTrans",SRCConstants.Char_ColorTrans},
                                                 {"Char.DblUnderline",SRCConstants.Char_DblUnderline},
                                                 {"Char.DoubleStrikethrough",SRCConstants.Char_DoubleStrikethrough},
                                                 {"Char.Font",SRCConstants.Char_Font},
                                                 {"Char.FontScale",SRCConstants.Char_FontScale},
                                                 {"Char.Letterspace",SRCConstants.Char_Letterspace},
                                                 {"Char.Overline",SRCConstants.Char_Overline},
                                                 {"Char.Size",SRCConstants.Char_Size},
                                                 {"Char.Strikethru",SRCConstants.Char_Strikethru},
                                                 {"Char.Style",SRCConstants.Char_Style},
                                                 {"Char.Pos",SRCConstants.Char_Pos},
                                                 {"Char.RTLText",SRCConstants.RTLText},
                                                 {"Char.UseVertical",SRCConstants.UseVertical},

                                                 //glueinfo                                                                      {"Char.Style",VA.ShapeSheet.CellSRCConstants.Char_Style},
                                                 {"BegTrigger",SRCConstants.BegTrigger},
                                                 {"EndTrigger",SRCConstants.EndTrigger},
                                                 {"GlueType",SRCConstants.GlueType},
                                                 {"WalkPreference",SRCConstants.WalkPreference},

                                                 // group
                                                 {"DisplayMode",SRCConstants.DisplayMode},
                                                 {"DontMoveChildren",SRCConstants.DontMoveChildren},
                                                 {"IsDropTarget",SRCConstants.IsDropTarget},
                                                 {"IsSnapTarget",SRCConstants.IsSnapTarget},
                                                 {"IsTextEditTarget",SRCConstants.IsTextEditTarget},
                                                 {"SelectMode",SRCConstants.SelectMode},

                                                 // misc
                                                 {"Calendar",SRCConstants.Calendar},
                                                 {"Comment",SRCConstants.Comment},
                                                 {"DropOnPageScale",SRCConstants.DropOnPageScale},
                                                 {"DynFeedback",SRCConstants.DynFeedback},
                                                 {"HideText",SRCConstants.HideText},
                                                 {"IsDropSource",SRCConstants.IsDropSource},
                                                 {"LangID",SRCConstants.LangID},
                                                 {"LocalizeMerge",SRCConstants.LocalizeMerge},
                                                 {"NoAlignBox",SRCConstants.NoAlignBox},
                                                 {"NoCtlHandles",SRCConstants.NoCtlHandles},
                                                 {"NoLiveDynamics",SRCConstants.NoLiveDynamics},
                                                 {"NonPrinting",SRCConstants.NonPrinting},
                                                 {"NoObjHandles",SRCConstants.NoObjHandles},
                                                 {"ObjType",SRCConstants.ObjType},
                                                 {"UpdateAlignBox",SRCConstants.UpdateAlignBox},




                                                 {"Para.Bullet",SRCConstants.Para_Bullet},
                                                 {"Para.BulletFont",SRCConstants.Para_BulletFont},
                                                 {"Para.BulletSize",SRCConstants.Para_BulletSize},
                                                 {"Para.BulletStr",SRCConstants.Para_BulletStr},
                                                 {"Para.Flags",SRCConstants.Para_Flags},
                                                 {"Para.HorzAlign",SRCConstants.Para_HorzAlign},
                                                 {"Para.IndFirst",SRCConstants.Para_IndFirst},
                                                 {"Para.IndLeft",SRCConstants.Para_IndLeft},
                                                 {"Para.IndRight",SRCConstants.Para_IndRight},
                                                 {"Para.LocBulletFont",SRCConstants.Para_LocBulletFont},
                                                 {"Para.SpAfter",SRCConstants.Para_SpAfter},
                                                 {"Para.SpBefore",SRCConstants.Para_SpBefore},
                                                 {"Para.SpLine",SRCConstants.Para_SpLine},
                                                 {"Para.TextPosAfterBullet",SRCConstants.Para_TextPosAfterBullet},
                                                                      
                                                 {"LockAspect",SRCConstants.LockAspect},
                                                 {"LockBegin",SRCConstants.LockBegin},
                                                 {"LockCalcWH",SRCConstants.LockCalcWH},
                                                 {"LockCrop",SRCConstants.LockCrop},
                                                 {"LockCustProp",SRCConstants.LockCustProp},
                                                 {"LockDelete",SRCConstants.LockDelete},
                                                 {"LockEnd",SRCConstants.LockEnd},
                                                 {"LockFormat",SRCConstants.LockFormat},
                                                 {"LockFromGroupFormat",SRCConstants.LockFromGroupFormat},
                                                 {"LockGroup",SRCConstants.LockGroup},
                                                 {"LockHeight",SRCConstants.LockHeight},
                                                 {"LockMoveX",SRCConstants.LockMoveX},
                                                 {"LockMoveY",SRCConstants.LockMoveY},
                                                 {"LockRotate",SRCConstants.LockRotate},
                                                 {"LockSelect",SRCConstants.LockSelect},
                                                 {"LockTextEdit",SRCConstants.LockTextEdit},
                                                 {"LockThemeColors",SRCConstants.LockThemeColors},
                                                 {"LockThemeEffects",SRCConstants.LockThemeEffects},
                                                 {"LockVtxEdit",SRCConstants.LockVtxEdit},
                                                 {"LockWidth",SRCConstants.LockWidth},
                                                                      
                                                 {"TxtAngle",SRCConstants.TxtAngle },
                                                 {"TxtHeight",SRCConstants.TxtHeight },
                                                 {"TxtLocPinX",SRCConstants.TxtLocPinX},
                                                 {"TxtLocPinY",SRCConstants.TxtLocPinY},
                                                 {"TxtPinX",SRCConstants.TxtPinX },
                                                 {"TxtPinY",SRCConstants.TxtPinY  },
                                                 {"TxtWidth",SRCConstants.TxtWidth },
                                                                      
                                                 {"BottomMargin",SRCConstants.BottomMargin },
                                                 {"DefaultTabstop",SRCConstants.DefaultTabstop},
                                                 {"LeftMargin",SRCConstants.LeftMargin },
                                                 {"RightMargin",SRCConstants.RightMargin  },
                                                 {"TextBkgnd",SRCConstants.TextBkgnd },
                                                 {"TextBkgndTrans",SRCConstants.TextBkgndTrans},
                                                 {"TextDirection",SRCConstants.TextDirection },
                                                 {"TopMargin",SRCConstants.TopMargin },
                                                 {"VerticalAlign",SRCConstants.VerticalAlign },
                                                                      
                                                 {"ConFixedCode",SRCConstants.ConFixedCode},
                                                 {"ConLineJumpCode",SRCConstants.ConLineJumpCode},
                                                 {"ConLineJumpDirX",SRCConstants.ConLineJumpDirX},
                                                 {"ConLineJumpDirY",SRCConstants.ConLineJumpDirY},
                                                 {"ConLineJumpStyle",SRCConstants.ConLineJumpStyle},
                                                 {"ConLineRouteExt",SRCConstants.ConLineRouteExt},
                                                 {"ShapeFixedCode",SRCConstants.ShapeFixedCode},
                                                 {"ShapePermeablePlace",SRCConstants.ShapePermeablePlace},
                                                 {"ShapePermeableX",SRCConstants.ShapePermeableX},
                                                 {"ShapePermeableY",SRCConstants.ShapePermeableY},
                                                 {"ShapePlaceFlip",SRCConstants.ShapePlaceFlip},
                                                 {"ShapePlaceStyle",SRCConstants.ShapePlaceStyle},
                                                 {"ShapePlowCode",SRCConstants.ShapePlowCode},
                                                 {"ShapeRouteStyle",SRCConstants.ShapeRouteStyle},
                                                 {"ShapeSplit",SRCConstants.ShapeSplit},
                                                 {"ShapeSplittable",SRCConstants.ShapeSplittable},
                                             };
            
        }

        private static Dictionary<string, SRC> simple_name_to_src_map;
    }
}