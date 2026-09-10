using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using iTextSharp.text.pdf;
using MultiToolWin.Models;
using MultiToolWin.OCR;
using MultiToolWin.PDF;

namespace MultiToolWin.Services
{
    public sealed class PdfGenerationService
    {
        private readonly SearchablePdfBuilder pdfBuilder = new SearchablePdfBuilder();

        public async Task<PdfGenerationResult> GenerateAsync(
            PdfScanResult scanResult,
            string outputRoot,
            bool includeTextLayer,
            IProgress<PdfGenerationProgress> progress,
            CancellationToken cancellationToken)
        {
            if (scanResult == null) throw new ArgumentNullException(nameof(scanResult));
            if (string.IsNullOrWhiteSpace(outputRoot))
                throw new ArgumentException("PDF 输出目录不能为空。", nameof(outputRoot));

            var result = new PdfGenerationResult
            {
                TotalFolders = scanResult.Folders.Count
            };

            var conflictingFolders = new HashSet<PdfSourceFolder>();
            foreach (IGrouping<string, PdfSourceFolder> group in scanResult.Folders
                .GroupBy(GetOutputFileName, StringComparer.OrdinalIgnoreCase)
                .Where(group => group.Count() > 1))
            {
                List<PdfSourceFolder> folders = group.ToList();
                foreach (PdfSourceFolder folder in folders) conflictingFolders.Add(folder);
                result.ConflictSkippedCount += folders.Count;
                ReportLog(progress, CreateConflictLog(group.Key, folders));
            }

            var existingOutputFolders = new HashSet<PdfSourceFolder>();
            foreach (PdfSourceFolder folder in scanResult.Folders)
            {
                if (conflictingFolders.Contains(folder)) continue;
                string outputFileName = GetOutputFileName(folder);
                if (!File.Exists(Path.Combine(outputRoot, outputFileName))) continue;

                existingOutputFolders.Add(folder);
                result.ExistingSkippedCount++;
                ReportLog(progress,
                    "目标PDF已存在，已跳过：" + Environment.NewLine + outputFileName);
            }

            int foldersToGenerate = scanResult.Folders.Count -
                conflictingFolders.Count - existingOutputFolders.Count;
            if (foldersToGenerate == 0)
            {
                ReportAllSkipped(progress, scanResult.Folders.Count);
                return result;
            }

            if (includeTextLayer)
            {
                if (!Environment.Is64BitOperatingSystem)
                    throw new PlatformNotSupportedException(
                        "OCR双层 PDF 需要64位 Windows。普通 PDF 仍可使用。");

                OcrResourceValidationResult validation = OcrResourceValidator.Validate();
                if (!validation.IsValid)
                    throw new InvalidOperationException(OcrResourceValidator.UserMessage);

                if (!SearchablePdfBuilder.HasUsableOcrFont())
                    throw new InvalidOperationException(
                        "未找到可用于 OCR 文字层的中文字体（微软雅黑、宋体或黑体）。普通 PDF 仍可正常生成。");
            }

            PaddleOcrEngine ocrEngine = includeTextLayer ? new PaddleOcrEngine() : null;
            try
            {
                int completedFolders = 0;
                foreach (PdfSourceFolder folder in scanResult.Folders)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    if (conflictingFolders.Contains(folder) ||
                        existingOutputFolders.Contains(folder))
                    {
                        completedFolders++;
                        ReportSkipped(progress, folder, completedFolders,
                            scanResult.Folders.Count);
                        continue;
                    }

                    string outputFileName = GetOutputFileName(folder);
                    string outputPath = Path.Combine(outputRoot, outputFileName);
                    string temporaryPath = null;
                    string currentImagePath = null;
                    SearchablePdfBuilder.WriterSession session = null;
                    string failureStage = "PdfWrite";
                    string operationPath = null;
                    bool validationPassed = false;
                    bool preserveValidatedTemp = false;

                    try
                    {
                        Directory.CreateDirectory(outputRoot);
                        temporaryPath = CreateTemporaryPdfPath(outputRoot);
                        operationPath = temporaryPath;
                        ReportLog(progress,
                            "[PDF] 开始生成" + Environment.NewLine +
                            "源目录：" + GetSourceDisplayPath(folder) + Environment.NewLine +
                            "临时PDF：" + temporaryPath + Environment.NewLine +
                            "正式PDF：" + outputPath);
                        session = await pdfBuilder.CreateSessionAsync(
                                temporaryPath,
                                includeTextLayer,
                                cancellationToken)
                            .ConfigureAwait(false);

                        for (int imageIndex = 0;
                            imageIndex < folder.ImageFiles.Count;
                            imageIndex++)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            currentImagePath = folder.ImageFiles[imageIndex];
                            failureStage = includeTextLayer
                                ? "SourceImageRead"
                                : "PdfWrite";
                            operationPath = currentImagePath;
                            Report(progress, folder, completedFolders,
                                scanResult.Folders.Count,
                                imageIndex + 1,
                                folder.ImageFiles.Count,
                                includeTextLayer ? "正在识别" : "正在写入");

                            PdfPageData page;
                            if (includeTextLayer)
                            {
                                OcrResult ocrResult = await ocrEngine
                                    .RecognizeAsync(currentImagePath, cancellationToken)
                                    .ConfigureAwait(false);
                                page = PdfPageData.FromOcrResult(ocrResult);
                            }
                            else
                            {
                                page = PdfPageData.FromImage(currentImagePath);
                            }

                            failureStage = "PdfWrite";
                            operationPath = temporaryPath;
                            await session.AddPageAsync(page, cancellationToken)
                                .ConfigureAwait(false);
                            page = null;
                        }

                        ReportLog(progress, "[PDF] 页面写入完成：" +
                            folder.ImageFiles.Count + "/" + folder.ImageFiles.Count);
                        cancellationToken.ThrowIfCancellationRequested();
                        failureStage = "PdfClose";
                        operationPath = temporaryPath;
                        ReportLog(progress, "[PDF] Writer关闭开始");
                        await session.CompleteAsync(cancellationToken)
                            .ConfigureAwait(false);
                        session.Dispose();
                        session = null;
                        ReportLog(progress, "[PDF] Writer关闭完成");

                        bool temporaryExists = File.Exists(temporaryPath);
                        long temporaryLength = temporaryExists
                            ? new FileInfo(temporaryPath).Length
                            : 0L;
                        ReportLog(progress,
                            "[PDF] Temp存在：" + temporaryExists + Environment.NewLine +
                            "[PDF] Temp大小：" + temporaryLength + " bytes");

                        cancellationToken.ThrowIfCancellationRequested();
                        failureStage = "Validate";
                        operationPath = temporaryPath;
                        ReportLog(progress, "[PDF] 开始完整性验证");
                        int actualPageCount = ValidateTemporaryPdf(
                            temporaryPath,
                            folder.ImageFiles.Count,
                            GetSourceDisplayPath(folder));
                        validationPassed = true;
                        ReportLog(progress,
                            "[PDF] 验证页数：" + actualPageCount + "/" +
                            folder.ImageFiles.Count + Environment.NewLine +
                            "[PDF] 完整性验证通过" + Environment.NewLine +
                            "[PDF] Validation Reader已释放");

                        cancellationToken.ThrowIfCancellationRequested();
                        failureStage = "Commit";
                        operationPath = outputPath;
                        bool outputExists = File.Exists(outputPath);
                        ReportLog(progress, "[PDF] 正式目标存在：" + outputExists);
                        if (outputExists)
                        {
                            result.ExistingSkippedCount++;
                            ReportLog(progress,
                                "目标PDF已存在，已跳过：" +
                                Environment.NewLine + outputFileName);
                            completedFolders++;
                            ReportSkipped(progress, folder, completedFolders,
                                scanResult.Folders.Count);
                            continue;
                        }

                        ReportLog(progress, "[PDF] Commit开始");
                        CommitTemporaryPdf(
                            temporaryPath,
                            outputPath,
                            progress,
                            cancellationToken);
                        temporaryPath = null;
                        ReportLog(progress,
                            "[PDF] 正式文件存在：" + File.Exists(outputPath) +
                            Environment.NewLine + "[PDF] Temp存在：False");
                        completedFolders++;
                        result.GeneratedCount++;
                        Report(progress, folder, completedFolders,
                            scanResult.Folders.Count,
                            folder.ImageFiles.Count,
                            folder.ImageFiles.Count,
                            "已生成");
                    }
                    catch (OperationCanceledException ex)
                    {
                        failureStage = "Cancelled";
                        operationPath = temporaryPath;
                        preserveValidatedTemp = validationPassed &&
                            !string.IsNullOrWhiteSpace(temporaryPath) &&
                            File.Exists(temporaryPath);
                        ReportFailure(progress, folder, currentImagePath,
                            failureStage, operationPath, temporaryPath,
                            outputPath, ex);
                        throw;
                    }
                    catch (Exception ex)
                    {
                        preserveValidatedTemp = validationPassed &&
                            string.Equals(failureStage, "Commit",
                                StringComparison.Ordinal) &&
                            !string.IsNullOrWhiteSpace(temporaryPath) &&
                            File.Exists(temporaryPath);
                        completedFolders++;
                        result.FailedCount++;
                        ReportFailure(progress, folder, currentImagePath,
                            failureStage, operationPath, temporaryPath,
                            outputPath, ex);
                        ReportFailed(progress, folder, completedFolders,
                            scanResult.Folders.Count);

                        if (includeTextLayer && ocrEngine != null &&
                            !ocrEngine.IsHealthy)
                        {
                            ocrEngine.Dispose();
                            ocrEngine = new PaddleOcrEngine();
                            ReportLog(progress,
                                "OCR引擎已重新创建，将继续处理下一件。");
                        }
                    }
                    finally
                    {
                        if (session != null)
                        {
                            try
                            {
                                session.Dispose();
                            }
                            catch (Exception ex)
                            {
                                ReportCleanupFailure(progress, folder,
                                    temporaryPath, outputPath,
                                    "释放PDF Writer/Stream失败", ex);
                            }
                            session = null;
                        }

                        if (preserveValidatedTemp)
                        {
                            ReportPreservedTemporaryPdf(progress, folder,
                                temporaryPath, outputPath, failureStage);
                        }
                        else
                        {
                            DeleteTemporaryPdf(temporaryPath, outputPath,
                                folder, progress);
                        }
                    }
                }

                return result;
            }
            catch (OperationCanceledException)
            {
                result.IsCancelled = true;
                return result;
            }
            finally
            {
                if (ocrEngine != null) ocrEngine.Dispose();
            }
        }

        private static string CreateTemporaryPdfPath(string outputRoot)
        {
            return Path.Combine(
                outputRoot,
                ".pdfgen-" + Guid.NewGuid().ToString("N") + ".tmp.pdf");
        }

        private static int ValidateTemporaryPdf(
            string temporaryPath,
            int expectedPageCount,
            string sourceDisplayPath)
        {
            if (string.IsNullOrWhiteSpace(temporaryPath) ||
                !File.Exists(temporaryPath))
            {
                throw new InvalidDataException(
                    "PDF完整性验证失败：" + Environment.NewLine +
                    sourceDisplayPath + Environment.NewLine +
                    "生成的临时PDF不存在。");
            }

            if (new FileInfo(temporaryPath).Length <= 0)
            {
                throw new InvalidDataException(
                    "PDF完整性验证失败：" + Environment.NewLine +
                    sourceDisplayPath + Environment.NewLine +
                    "生成的临时PDF为空文件。");
            }

            int actualPageCount;
            try
            {
                using (var stream = new FileStream(
                    temporaryPath,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.Read))
                using (var reader = new PdfReader(stream))
                {
                    actualPageCount = reader.NumberOfPages;
                }
            }
            catch (Exception ex)
            {
                throw new InvalidDataException(
                    "PDF完整性验证失败：" + Environment.NewLine +
                    sourceDisplayPath + Environment.NewLine +
                    "无法重新打开生成的临时PDF。",
                    ex);
            }

            if (actualPageCount != expectedPageCount)
            {
                throw new InvalidDataException(
                    "PDF完整性验证失败：" + Environment.NewLine +
                    sourceDisplayPath + Environment.NewLine +
                    "预期页数：" + expectedPageCount + Environment.NewLine +
                    "实际页数：" + actualPageCount);
            }

            return actualPageCount;
        }

        private static void DeleteTemporaryPdf(
            string temporaryPath,
            string outputPath,
            PdfSourceFolder folder,
            IProgress<PdfGenerationProgress> progress)
        {
            if (string.IsNullOrWhiteSpace(temporaryPath) ||
                !File.Exists(temporaryPath))
                return;

            try
            {
                File.Delete(temporaryPath);
            }
            catch (Exception ex)
            {
                ReportCleanupFailure(progress, folder, temporaryPath,
                    outputPath, "临时PDF删除失败", ex);
            }
        }

        private static void CommitTemporaryPdf(
            string temporaryPath,
            string outputPath,
            IProgress<PdfGenerationProgress> progress,
            CancellationToken cancellationToken)
        {
            int[] retryDelays = { 200, 500, 1000 };
            int retry = 0;
            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                try
                {
                    ReportLog(progress, "[PDF] File.Move开始" +
                        (retry == 0 ? string.Empty :
                            "（重试 " + retry + "/" + retryDelays.Length + "）"));
                    File.Move(temporaryPath, outputPath);
                    ReportLog(progress, "[PDF] File.Move成功");
                    return;
                }
                catch (IOException ex)
                {
                    if (File.Exists(outputPath))
                    {
                        throw new IOException(
                            "Commit失败：正式目标文件已存在，不会覆盖：" +
                            outputPath,
                            ex);
                    }

                    if (!IsSharingViolation(ex) || retry >= retryDelays.Length)
                        throw;

                    int delay = retryDelays[retry];
                    retry++;
                    ReportLog(progress,
                        "[PDF] Commit重试 " + retry + "/" +
                        retryDelays.Length + "，等待 " + delay + "ms" +
                        Environment.NewLine +
                        "临时PDF：" + temporaryPath + Environment.NewLine +
                        "正式PDF：" + outputPath + Environment.NewLine +
                        "异常：" + ex.Message);
                    if (cancellationToken.WaitHandle.WaitOne(delay))
                        cancellationToken.ThrowIfCancellationRequested();
                }
            }
        }

        private static bool IsSharingViolation(IOException exception)
        {
            int errorCode = exception.HResult & 0xFFFF;
            return errorCode == 32 || errorCode == 33;
        }

        private static void ReportFailure(
            IProgress<PdfGenerationProgress> progress,
            PdfSourceFolder folder,
            string imagePath,
            string failureStage,
            string operationPath,
            string temporaryPath,
            string outputPath,
            Exception exception)
        {
            var message = new StringBuilder();
            if (exception is OcrTimeoutException)
                message.AppendLine("OCR超时，当前件已跳过：" +
                    GetSourceDisplayPath(folder));
            else
                message.AppendLine("生成失败，当前件已跳过：" +
                    GetSourceDisplayPath(folder));

            message.AppendLine("失败阶段：" + failureStage);
            message.AppendLine("源目录：" + GetSourceDisplayPath(folder));
            if (!string.IsNullOrWhiteSpace(imagePath))
                message.AppendLine("当前源图片：" + imagePath);
            if (!string.IsNullOrWhiteSpace(operationPath))
                message.AppendLine("实际操作文件：" + operationPath);
            if (!string.IsNullOrWhiteSpace(temporaryPath))
                message.AppendLine("临时路径：" + temporaryPath);
            message.AppendLine("正式路径：" + outputPath);
            message.AppendLine("异常类型：" + exception.GetType().FullName);
            message.AppendLine("HResult：0x" +
                exception.HResult.ToString("X8"));
            message.Append("异常详情：" + exception);
            ReportLog(progress, message.ToString());
        }

        private static void ReportCleanupFailure(
            IProgress<PdfGenerationProgress> progress,
            PdfSourceFolder folder,
            string temporaryPath,
            string outputPath,
            string action,
            Exception exception)
        {
            var message = new StringBuilder();
            message.AppendLine("Cleanup失败：" + action);
            message.AppendLine("失败阶段：Cleanup");
            message.AppendLine("源目录：" + GetSourceDisplayPath(folder));
            message.AppendLine("实际操作文件：" + temporaryPath);
            message.AppendLine("临时路径：" + temporaryPath);
            message.AppendLine("正式路径：" + outputPath);
            message.AppendLine("异常类型：" + exception.GetType().FullName);
            message.AppendLine("HResult：0x" +
                exception.HResult.ToString("X8"));
            message.Append("异常详情：" + exception);
            ReportLog(progress, message.ToString());
        }

        private static void ReportPreservedTemporaryPdf(
            IProgress<PdfGenerationProgress> progress,
            PdfSourceFolder folder,
            string temporaryPath,
            string outputPath,
            string failureStage)
        {
            string reason = string.Equals(failureStage, "Cancelled",
                StringComparison.Ordinal)
                ? "用户在完整性验证通过后取消，完整临时PDF已保留。"
                : "PDF已完整生成并验证通过，但正式提交失败，完整临时PDF已保留。";
            ReportLog(progress,
                reason + Environment.NewLine +
                "失败阶段：" + failureStage + Environment.NewLine +
                "源目录：" + GetSourceDisplayPath(folder) + Environment.NewLine +
                "临时PDF：" + temporaryPath + Environment.NewLine +
                "应生成正式PDF：" + outputPath + Environment.NewLine +
                "该临时PDF可用于人工恢复。");
        }

        private static string GetOutputFileName(PdfSourceFolder folder)
        {
            return folder.FolderName + ".pdf";
        }

        private static string CreateConflictLog(
            string outputFileName,
            IList<PdfSourceFolder> folders)
        {
            var message = new StringBuilder();
            message.AppendLine("发现同名件，以下文件夹将生成相同PDF：");
            foreach (PdfSourceFolder folder in folders)
                message.AppendLine(GetSourceDisplayPath(folder));
            message.AppendLine();
            message.AppendLine("目标PDF：" + outputFileName);
            message.Append("已跳过以上" + folders.Count + "件。");
            return message.ToString();
        }

        private static string GetSourceDisplayPath(PdfSourceFolder folder)
        {
            return string.IsNullOrWhiteSpace(folder.RelativeOutputDirectory)
                ? folder.FolderName
                : Path.Combine(folder.RelativeOutputDirectory, folder.FolderName);
        }

        private static void ReportLog(
            IProgress<PdfGenerationProgress> progress,
            string message)
        {
            if (progress == null) return;
            progress.Report(new PdfGenerationProgress { LogMessage = message });
        }

        private static void ReportSkipped(
            IProgress<PdfGenerationProgress> progress,
            PdfSourceFolder folder,
            int completedFolders,
            int totalFolders)
        {
            if (progress == null) return;
            progress.Report(new PdfGenerationProgress
            {
                Message = "已跳过",
                CurrentFolderName = GetSourceDisplayPath(folder),
                CompletedFolders = completedFolders,
                TotalFolders = totalFolders
            });
        }

        private static void ReportFailed(
            IProgress<PdfGenerationProgress> progress,
            PdfSourceFolder folder,
            int completedFolders,
            int totalFolders)
        {
            if (progress == null) return;
            progress.Report(new PdfGenerationProgress
            {
                Message = "失败",
                CurrentFolderName = GetSourceDisplayPath(folder),
                CompletedFolders = completedFolders,
                TotalFolders = totalFolders
            });
        }

        private static void ReportAllSkipped(
            IProgress<PdfGenerationProgress> progress,
            int totalFolders)
        {
            if (progress == null) return;
            progress.Report(new PdfGenerationProgress
            {
                Message = "已全部跳过",
                CurrentFolderName = "-",
                CompletedFolders = totalFolders,
                TotalFolders = totalFolders
            });
        }

        private static void Report(
            IProgress<PdfGenerationProgress> progress,
            PdfSourceFolder folder,
            int completedFolders,
            int totalFolders,
            int currentImage,
            int totalImages,
            string stage)
        {
            if (progress == null) return;
            progress.Report(new PdfGenerationProgress
            {
                Message = stage,
                CurrentFolderName = GetSourceDisplayPath(folder),
                CurrentImage = currentImage,
                TotalImages = totalImages,
                CompletedFolders = completedFolders,
                TotalFolders = totalFolders
            });
        }
    }
}
