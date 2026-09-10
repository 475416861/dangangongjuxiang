using System.Collections.Generic;

namespace MultiToolWin.Models
{
    public sealed class PdfScanResult
    {
        public PdfDirectoryLayout Layout { get; set; }
        public List<PdfSourceFolder> Folders { get; } = new List<PdfSourceFolder>();
    }
}
