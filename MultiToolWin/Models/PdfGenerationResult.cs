namespace MultiToolWin.Models
{
    public sealed class PdfGenerationResult
    {
        public int TotalFolders { get; set; }
        public int GeneratedCount { get; set; }
        public int ConflictSkippedCount { get; set; }
        public int ExistingSkippedCount { get; set; }
        public int FailedCount { get; set; }
        public bool IsCancelled { get; set; }
    }
}
