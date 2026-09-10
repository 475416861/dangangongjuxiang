namespace MultiToolWin.Models
{
    public sealed class OcrTextBlock
    {
        public string Text { get; set; }
        public float Confidence { get; set; }
        public float X { get; set; }
        public float Y { get; set; }
        public float Width { get; set; }
        public float Height { get; set; }
    }
}
