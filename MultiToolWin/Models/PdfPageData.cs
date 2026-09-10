using System.Collections.Generic;

namespace MultiToolWin.Models
{
    public sealed class PdfPageData
    {
        public string ImagePath { get; set; }
        public int ImageWidth { get; set; }
        public int ImageHeight { get; set; }
        public IList<OcrTextBlock> TextBlocks { get; set; }

        public static PdfPageData FromImage(string imagePath)
        {
            return new PdfPageData { ImagePath = imagePath };
        }

        public static PdfPageData FromOcrResult(OcrResult result)
        {
            return new PdfPageData
            {
                ImagePath = result.ImagePath,
                ImageWidth = result.ImageWidth,
                ImageHeight = result.ImageHeight,
                TextBlocks = result.TextBlocks
            };
        }
    }
}
