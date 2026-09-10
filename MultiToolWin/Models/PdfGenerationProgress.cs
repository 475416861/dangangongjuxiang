namespace MultiToolWin.Models
{
    public sealed class PdfGenerationProgress
    {
        public string Message { get; set; }
        public string CurrentFolderName { get; set; }
        public int CurrentImage { get; set; }
        public int TotalImages { get; set; }
        public int CompletedFolders { get; set; }
        public int TotalFolders { get; set; }
        public string LogMessage { get; set; }
    }
}
