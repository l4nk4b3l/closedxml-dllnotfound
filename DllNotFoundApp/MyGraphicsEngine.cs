using System;
using System.IO;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using ClosedXML.Graphics;

namespace DllNotFoundApp
{
    public class MyGraphicsEngine : IXLGraphicEngine
    {
        public XLPictureInfo GetPictureInfo(Stream imageStream, XLPictureFormat expectedFormat)
        {
            throw new NotImplementedException();
        }

        public double GetTextHeight(IXLFontBase font, double dpiY)
        {
            throw new NotImplementedException();
        }

        public double GetTextWidth(string text, IXLFontBase font, double dpiX)
        {
            throw new NotImplementedException();
        }

        public double GetMaxDigitWidth(IXLFontBase font, double dpiX)
        {
            return 0;
        }

        public double GetDescent(IXLFontBase font, double dpiY)
        {
            throw new NotImplementedException();
        }

        public GlyphBox GetGlyphBox(ReadOnlySpan<int> graphemeCluster, IXLFontBase font, Dpi dpi)
        {
            throw new NotImplementedException();
        }
    }
}