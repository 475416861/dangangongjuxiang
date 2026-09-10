using System.Collections.Generic;

namespace MultiToolWin.Models
{
    public sealed class PdfSourceFolder
    {
        public string FolderName { get; set; }
        public string FolderPath { get; set; }
        public string RelativeOutputDirectory { get; set; }
        public List<string> ImageFiles { get; set; } = new List<string>();
    }
}
