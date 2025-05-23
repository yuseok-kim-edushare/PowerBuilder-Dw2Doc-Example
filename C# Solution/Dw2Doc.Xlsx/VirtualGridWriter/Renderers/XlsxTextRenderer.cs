﻿using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Models;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Renderers.Attributes;
using Appeon.DotnetDemo.Dw2Doc.Xlsx.Extensions;
using Appeon.DotnetDemo.Dw2Doc.Xlsx.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SixLabors.ImageSharp.PixelFormats;

namespace Appeon.DotnetDemo.Dw2Doc.Xlsx.VirtualGridWriter.Renderers.Xlsx
{
    [RendererFor(typeof(DwTextAttributes), typeof(XlsxTextRenderer))]
    internal class XlsxTextRenderer : AbstractXlsxRenderer
    {
        private const double TextBoxWidthAdjustment = 5;

        private static double ConvertFontSize(DwTextAttributes attributes)
        {
            return attributes.FontSize - 1;
        }

        public override ExportedCellBase? Render(
            ISheet sheet,
            VirtualCell cell,
            DwObjectAttributesBase attribute,
            ICell renderTarget)
        {
            var textAttribute = CheckAttributeType<DwTextAttributes>(attribute);
            
            // Use null-conditional operator and explicitly check for null
            var cellStyle = renderTarget.Row.Sheet.Workbook.CreateCellStyle();
            if (cellStyle is not XSSFCellStyle xssfStyle) 
            {
                throw new InvalidCastException("Could not get XSSFCellStyle");
            }
            
            XSSFCellStyle style = xssfStyle;
            var font = renderTarget.Row.Sheet.Workbook.CreateFont();
            if (font is not XSSFFont xssfFont)
            {
                throw new InvalidCastException("Could not get XSSFFont");
            }
            
            xssfFont.FontHeightInPoints = ConvertFontSize(textAttribute);
            xssfFont.FontName = textAttribute.FontFace;
            xssfFont.SetColor(new XSSFColor(new byte[] {
                textAttribute.FontColor.Value.R,
                textAttribute.FontColor.Value.G,
                textAttribute.FontColor.Value.B}));
            xssfFont.IsBold = textAttribute.FontWeight >= 700;
            xssfFont.IsItalic = textAttribute.Italics;
            xssfFont.IsStrikeout = textAttribute.Strikethrough;
            xssfFont.Underline = textAttribute.Underline ? FontUnderlineType.Single : FontUnderlineType.None;
            style.SetFont(xssfFont);
            style.SetFillForegroundColor(new XSSFColor(new byte[3] {
                textAttribute.BackgroundColor.Value.R,
                textAttribute.BackgroundColor.Value.G,
                textAttribute.BackgroundColor.Value.B,
            }));
            style.FillPattern = textAttribute.BackgroundColor.Value.A == 0 ? FillPattern.NoFill : FillPattern.SolidForeground;
            style.Alignment = textAttribute.Alignment.ToNpoiHorizontalAlignment();
            style.VerticalAlignment = VerticalAlignment.Top;


            // Use null-conditional operator and pattern matching to avoid null warnings
            if (renderTarget is XSSFCell xTarget &&
                textAttribute.FormatString is not null
                && textAttribute.FormatString.ToLower() != "[general]")
            {
                switch (textAttribute.DataType)
                {
                    case Common.Enums.DataType.Money:
                    case Common.Enums.DataType.Number:
                        if (sheet.Workbook?.CreateDataFormat() != null)
                        {
                            style.SetDataFormat(sheet
                                .Workbook
                                .CreateDataFormat()
                                .GetFormat(textAttribute.FormatString ?? string.Empty));
                        }
                        break;
                }
            }

            renderTarget.CellStyle = style;
            if (!string.IsNullOrEmpty(textAttribute.Text) &&
                (textAttribute.DataType is Common.Enums.DataType.Money ||
                textAttribute.DataType is Common.Enums.DataType.Number) &&
                !string.IsNullOrEmpty(textAttribute.RawText))
            {
                renderTarget.SetCellValue(double.Parse(textAttribute.RawText));
            }
            else
                renderTarget.SetCellValue(textAttribute.Text);

            return new ExportedCell(cell, attribute)
            {
                OutputCell = renderTarget,
            };
        }

        public override ExportedCellBase Render(ISheet sheet, FloatingVirtualCell cell, DwObjectAttributesBase attribute, (int x, int y, XSSFDrawing draw) renderTarget)
        {
            var textAttribute = CheckAttributeType<DwTextAttributes>(attribute);

            var anchor = GetAnchor(
                renderTarget.draw,
                cell,
                renderTarget.x,
                renderTarget.y,
                widthAdjustment: TextBoxWidthAdjustment);

            anchor.AnchorType = AnchorType.MoveDontResize;

            var textBox = renderTarget.draw.CreateTextbox(anchor);
            textBox.BottomInset = 0;
            textBox.LeftInset = 0;
            textBox.TopInset = 0;
            textBox.RightInset = 0;
            if (textAttribute.BackgroundColor.Value.A != 0)
            {
                textBox.SetFillColor(textAttribute.BackgroundColor.Value.R,
                textAttribute.BackgroundColor.Value.G,
                textAttribute.BackgroundColor.Value.B);
            }


            var paragraph = textBox.TextParagraphs[0];
            var run = paragraph.AddNewTextRun();
            run.FontColor = new Rgb24(textAttribute.FontColor.Value.R,
                textAttribute.FontColor.Value.G,
                textAttribute.FontColor.Value.B);
            run.Text = textAttribute.Text;
            run.FontSize = ConvertFontSize(textAttribute);
            run.IsItalic = textAttribute.Italics;
            run.IsStrikethrough = textAttribute.Strikethrough;
            run.IsUnderline = textAttribute.Underline;
            run.SetFont(textAttribute.FontFace);
            run.IsBold = textAttribute.FontWeight >= 700;
            paragraph.TextAlign = textAttribute.Alignment.ToNpoiTextAlignment();
            textBox.LineStyle = LineStyle.None;
            //textBox.SetText();

            return new ExportedFloatingCell(cell, attribute)
            {
                OutputShape = textBox,
            };
        }
    }
}
