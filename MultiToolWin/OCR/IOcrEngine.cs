using System.Threading;
using System.Threading.Tasks;
using MultiToolWin.Models;

namespace MultiToolWin.OCR
{
    public interface IOcrEngine
    {
        Task<OcrResult> RecognizeAsync(string imagePath, CancellationToken cancellationToken);
    }
}
