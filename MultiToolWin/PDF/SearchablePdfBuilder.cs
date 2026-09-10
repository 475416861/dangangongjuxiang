using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;
using MultiToolWin.Models;

namespace MultiToolWin.PDF
{
    public sealed class SearchablePdfBuilder
    {
        private static readonly string[] FontCandidates =
        {
            "msyh.ttc,0",
            "simsun.ttc,0",
            "simhei.ttf"
        };

        public Task<WriterSession> CreateSessionAsync(
            string outputPdfPath,
            bool includeTextLayer,
            CancellationToken cancellationToken)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                return new WriterSession(outputPdfPath, includeTextLayer);
            }, cancellationToken);
        }

        public static bool HasUsableOcrFont()
        {
            return ResolveChineseFont() != null;
        }

        public sealed class WriterSession : IDisposable
        {
            private FileStream stream;
            private Document document;
            private PdfWriter writer;
            private readonly BaseFont font;
            private readonly bool includeTextLayer;
            private bool completed;
            private bool disposed;

            internal WriterSession(string outputPdfPath, bool includeTextLayer)
            {
                if (string.IsNullOrWhiteSpace(outputPdfPath))
                    throw new ArgumentException("PDF 输出路径不能为空。", nameof(outputPdfPath));

                this.includeTextLayer = includeTextLayer;
                try
                {
                    stream = new FileStream(
                        outputPdfPath,
                        FileMode.CreateNew,
                        FileAccess.Write,
                        FileShare.None);
                    document = new Document();
                    writer = PdfWriter.GetInstance(document, stream);
                    document.Open();
                    font = includeTextLayer ? CreateChineseFont() : null;
                }
                catch
                {
                    Dispose();
                    throw;
                }
            }

            public Task AddPageAsync(
                PdfPageData page,
                CancellationToken cancellationToken)
            {
                if (page == null) throw new ArgumentNullException(nameof(page));
                return Task.Run(
                    () => AddPage(page, cancellationToken),
                    cancellationToken);
            }

            public Task CompleteAsync(CancellationToken cancellationToken)
            {
                return Task.Run(() =>
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    Complete();
                }, cancellationToken);
            }

            private void AddPage(PdfPageData page, CancellationToken token)
            {
                ThrowIfClosed();
                token.ThrowIfCancellationRequested();

                float width;
                float height;
                int imageWidth;
                int imageHeight;
                using (var source = System.Drawing.Image.FromFile(page.ImagePath))
                {
                    float horizontalResolution = source.HorizontalResolution > 0
                        ? source.HorizontalResolution : 300F;
                    float verticalResolution = source.VerticalResolution > 0
                        ? source.VerticalResolution : 300F;
                    width = source.Width / horizontalResolution * 72F;
                    height = source.Height / verticalResolution * 72F;
                    imageWidth = source.Width;
                    imageHeight = source.Height;
                }

                token.ThrowIfCancellationRequested();
                document.SetPageSize(new Rectangle(width, height));
                document.NewPage();
                PdfContentByte canvas = writer.DirectContent;
                iTextSharp.text.Image pageImage =
                    iTextSharp.text.Image.GetInstance(page.ImagePath);
                pageImage.SetAbsolutePosition(0, 0);
                pageImage.ScaleAbsolute(width, height);
                canvas.AddImage(pageImage);

                token.ThrowIfCancellationRequested();
                if (!includeTextLayer || page.TextBlocks == null ||
                    page.TextBlocks.Count == 0)
                    return;

                int coordinateWidth = page.ImageWidth > 0
                    ? page.ImageWidth : imageWidth;
                int coordinateHeight = page.ImageHeight > 0
                    ? page.ImageHeight : imageHeight;
                float scaleX = width / coordinateWidth;
                float scaleY = height / coordinateHeight;

                canvas.BeginText();
                try
                {
                    canvas.SetTextRenderingMode(PdfContentByte.TEXT_RENDER_MODE_INVISIBLE);
                    foreach (OcrTextBlock block in page.TextBlocks)
                    {
                        token.ThrowIfCancellationRequested();
                        if (string.IsNullOrWhiteSpace(block.Text)) continue;
                        float size = Math.Max(1F, block.Height * scaleY * 0.82F);
                        float textWidth = font.GetWidthPoint(block.Text, size);
                        canvas.SetFontAndSize(font, size);
                        canvas.SetHorizontalScaling(textWidth > 0F
                            ? Math.Max(1F, Math.Min(1000F,
                                block.Width * scaleX / textWidth * 100F))
                            : 100F);
                        canvas.SetTextMatrix(
                            block.X * scaleX,
                            height - (block.Y + block.Height) * scaleY);
                        canvas.ShowText(block.Text);
                    }
                }
                finally
                {
                    canvas.EndText();
                }
            }

            private void Complete()
            {
                ThrowIfClosed();
                try
                {
                    document.Close();
                    completed = true;
                }
                finally
                {
                    DisposeResources();
                }
            }

            private void ThrowIfClosed()
            {
                if (disposed || completed)
                    throw new ObjectDisposedException(nameof(WriterSession));
            }

            public void Dispose()
            {
                if (disposed) return;
                disposed = true;
                DisposeResources();
            }

            private void DisposeResources()
            {
                Document currentDocument = document;
                FileStream currentStream = stream;
                document = null;
                writer = null;
                stream = null;

                if (currentDocument != null)
                {
                    try { currentDocument.Close(); } catch { }
                    try { currentDocument.Dispose(); } catch { }
                }
                if (currentStream != null)
                {
                    try { currentStream.Dispose(); } catch { }
                }
            }
        }

        private static BaseFont CreateChineseFont()
        {
            string path = ResolveChineseFont();
            if (path == null)
                throw new FileNotFoundException(
                    "未找到可用于 OCR 文字层的中文字体（微软雅黑、宋体或黑体）。普通 PDF 仍可正常生成。");
            return BaseFont.CreateFont(path, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
        }

        private static string ResolveChineseFont()
        {
            string fontsDirectory =
                Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
            foreach (string name in FontCandidates)
            {
                string candidate = Path.Combine(fontsDirectory, name);
                if (File.Exists(candidate.Replace(",0", string.Empty)))
                    return candidate;
            }
            return null;
        }
    }
}
