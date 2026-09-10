using System.Collections.Generic;

namespace MultiToolWin.Models
{
    public sealed class OcrResult
    {
        public string ImagePath { get; set; }
        public int ImageWidth { get; set; }
        public int ImageHeight { get; set; }
        public List<OcrTextBlock> TextBlocks { get; } = new List<OcrTextBlock>();
    }
}
