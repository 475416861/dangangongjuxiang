using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;
using System.Threading;
using System.Threading.Tasks;
using MultiToolWin.Models;
using PaddleOCRJson;

namespace MultiToolWin.OCR
{
    public sealed class OcrTimeoutException : TimeoutException
    {
        public OcrTimeoutException(string imagePath, TimeSpan timeout)
            : base(string.Format(
                "图片 OCR 超时（{0}秒）：{1}",
                (int)timeout.TotalSeconds,
                Path.GetFileName(imagePath)))
        {
            ImagePath = imagePath;
        }

        public string ImagePath { get; }
    }

    public sealed class PaddleOcrEngine : IOcrEngine, IDisposable
    {
        private static readonly TimeSpan DefaultPageTimeout = TimeSpan.FromSeconds(120);
        private static readonly TimeSpan StartupTimeout = TimeSpan.FromSeconds(30);
        private const int MaxDiagnosticLength = 2048;
        private const int MaxProtocolSnippetLength = 512;

        private readonly object stateLock = new object();
        private readonly SemaphoreSlim requestGate = new SemaphoreSlim(1, 1);
        private readonly AutoResetEvent startupEvent = new AutoResetEvent(false);
        private readonly AutoResetEvent responseEvent = new AutoResetEvent(false);
        private readonly AutoResetEvent processExitedEvent = new AutoResetEvent(false);
        private readonly StringBuilder startupOutput = new StringBuilder();
        private readonly StringBuilder diagnosticOutput = new StringBuilder();
        private readonly TimeSpan pageTimeout;

        private Process process;
        private string response;
        private bool waitingForResponse;
        private bool started;
        private bool unhealthy;
        private bool disposed;

        public PaddleOcrEngine()
            : this(DefaultPageTimeout)
        {
        }

        internal PaddleOcrEngine(TimeSpan pageTimeout)
        {
            if (pageTimeout <= TimeSpan.Zero)
                throw new ArgumentOutOfRangeException(nameof(pageTimeout));
            this.pageTimeout = pageTimeout;
        }

        public bool IsHealthy
        {
            get
            {
                lock (stateLock) return !disposed && !unhealthy;
            }
        }

        public Task<OcrResult> RecognizeAsync(string imagePath, CancellationToken cancellationToken)
        {
            return Task.Run(() => Recognize(imagePath, cancellationToken));
        }

        private OcrResult Recognize(string imagePath, CancellationToken cancellationToken)
        {
            bool gateEntered = false;
            try
            {
                requestGate.Wait(cancellationToken);
                gateEntered = true;
                cancellationToken.ThrowIfCancellationRequested();
                EnsureStarted(cancellationToken);

                byte[] imageBytes = File.ReadAllBytes(imagePath);
                string command = "{\"image_base64\":\"" +
                    Convert.ToBase64String(imageBytes) + "\"}";
                imageBytes = null;

                Process currentProcess;
                lock (stateLock)
                {
                    ThrowIfUnavailable();
                    currentProcess = process;
                    response = null;
                    waitingForResponse = true;
                    responseEvent.Reset();
                    processExitedEvent.Reset();
                }

                try
                {
                    currentProcess.StandardInput.WriteLine(command);
                    currentProcess.StandardInput.Flush();
                }
                catch
                {
                    MarkUnhealthy();
                    throw;
                }
                finally
                {
                    command = null;
                }

                int signal = WaitHandle.WaitAny(
                    new WaitHandle[]
                    {
                        responseEvent,
                        cancellationToken.WaitHandle,
                        processExitedEvent
                    },
                    pageTimeout);

                lock (stateLock) waitingForResponse = false;

                if (signal == 1)
                {
                    AbortProcess();
                    throw new OperationCanceledException(cancellationToken);
                }
                if (signal == 2)
                {
                    string diagnostics = GetDiagnosticSuffix();
                    AbortProcess();
                    throw new IOException(
                        "PaddleOCR-json 进程意外退出。" + diagnostics);
                }
                if (signal == WaitHandle.WaitTimeout)
                {
                    AbortProcess();
                    throw new OcrTimeoutException(imagePath, pageTimeout);
                }

                cancellationToken.ThrowIfCancellationRequested();
                string json;
                lock (stateLock) json = response;
                return Parse(imagePath, json);
            }
            finally
            {
                lock (stateLock) waitingForResponse = false;
                if (gateEntered) requestGate.Release();
            }
        }

        private void EnsureStarted(CancellationToken cancellationToken)
        {
            lock (stateLock)
            {
                ThrowIfUnavailable();
                if (started) return;
                StartProcess();
            }

            int signal = WaitHandle.WaitAny(
                new WaitHandle[]
                {
                    startupEvent,
                    cancellationToken.WaitHandle,
                    processExitedEvent
                },
                StartupTimeout);

            if (signal == 0)
            {
                lock (stateLock)
                {
                    if (started) return;
                }
            }

            if (signal == 1)
            {
                AbortProcess();
                throw new OperationCanceledException(cancellationToken);
            }

            string details = GetStartupDetails();
            AbortProcess();
            if (signal == 2)
                throw new IOException("PaddleOCR-json 启动期间意外退出。" +
                    (details.Length == 0 ? string.Empty : Environment.NewLine + details));
            throw new TimeoutException("PaddleOCR-json 启动超时。" +
                (details.Length == 0 ? string.Empty : Environment.NewLine + details));
        }

        private void StartProcess()
        {
            string executablePath = Path.Combine(
                OcrResourceValidator.GetRuntimeDirectory(),
                "PaddleOCR-json.exe");

            var args = OcrEngineStartupArgs.WithPipeMode(executablePath)
                .WithConfigPath("models/config_chinese.txt")
                .WithCpuThreads(Math.Max(1, Environment.ProcessorCount - 1))
                .WithEnableMkldnn(true)
                .WithLimitSideLen(1920)
                .WithUseAngleCls(true)
                .WithEnsureAscii(true);

            startupEvent.Reset();
            responseEvent.Reset();
            processExitedEvent.Reset();
            startupOutput.Clear();
            diagnosticOutput.Clear();
            started = false;

            var nextProcess = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = executablePath,
                    WorkingDirectory = Path.GetDirectoryName(executablePath),
                    Arguments = args.ToString(),
                    CreateNoWindow = true,
                    RedirectStandardInput = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false
                },
                EnableRaisingEvents = true
            };
            nextProcess.OutputDataReceived += OnEngineOutputDataReceived;
            nextProcess.ErrorDataReceived += OnEngineErrorDataReceived;
            nextProcess.Exited += OnEngineExited;
            process = nextProcess;

            try
            {
                if (!nextProcess.Start())
                    throw new InvalidOperationException("无法启动 PaddleOCR-json 进程。");
                nextProcess.StandardInput.AutoFlush = true;
                nextProcess.BeginOutputReadLine();
                nextProcess.BeginErrorReadLine();
            }
            catch
            {
                unhealthy = true;
                process = null;
                try { nextProcess.Dispose(); } catch { }
                throw;
            }
        }

        private void OnEngineOutputDataReceived(
            object sender,
            DataReceivedEventArgs args)
        {
            if (string.IsNullOrWhiteSpace(args.Data)) return;

            lock (stateLock)
            {
                if (disposed || unhealthy || !ReferenceEquals(process, sender)) return;
                if (!started)
                {
                    startupOutput.AppendLine(args.Data);
                    if (string.Equals(args.Data, "OCR init completed.", StringComparison.Ordinal))
                    {
                        started = true;
                        startupEvent.Set();
                    }
                    return;
                }

                if (!waitingForResponse) return;
                response = args.Data;
                responseEvent.Set();
            }
        }

        private void OnEngineErrorDataReceived(
            object sender,
            DataReceivedEventArgs args)
        {
            if (string.IsNullOrWhiteSpace(args.Data)) return;

            string diagnostic = Truncate(args.Data, MaxDiagnosticLength);
            lock (stateLock)
            {
                if (disposed || !ReferenceEquals(process, sender)) return;
                AppendBoundedDiagnostic(diagnosticOutput, diagnostic);
            }

            Trace.TraceWarning("PaddleOCR-json stderr: {0}", diagnostic);
        }

        private void OnEngineExited(object sender, EventArgs args)
        {
            lock (stateLock)
            {
                if (!ReferenceEquals(process, sender)) return;
                unhealthy = true;
                processExitedEvent.Set();
            }
        }

        private void ThrowIfUnavailable()
        {
            if (disposed) throw new ObjectDisposedException(nameof(PaddleOcrEngine));
            if (unhealthy) throw new InvalidOperationException("OCR 引擎已失效，需要重新创建。");
        }

        private void MarkUnhealthy()
        {
            lock (stateLock) unhealthy = true;
        }

        private void AbortProcess()
        {
            StopProcess(false);
        }

        private void StopProcess(bool graceful)
        {
            Process currentProcess;
            lock (stateLock)
            {
                unhealthy = true;
                started = false;
                waitingForResponse = false;
                currentProcess = process;
                process = null;
                startupEvent.Set();
                responseEvent.Set();
                processExitedEvent.Set();
            }

            if (currentProcess == null) return;
            try
            {
                if (!currentProcess.HasExited)
                {
                    if (graceful && TryRequestGracefulExit(currentProcess))
                    {
                        if (!WaitForExit(currentProcess, 3000, "正常退出"))
                            ForceTerminate(currentProcess);
                    }
                    else
                    {
                        ForceTerminate(currentProcess);
                    }
                }
            }
            catch (Exception ex)
            {
                Trace.TraceWarning(
                    "清理 PaddleOCR-json 进程时发生异常：{0}", ex.Message);
            }
            finally
            {
                try { currentProcess.CancelOutputRead(); } catch { }
                try { currentProcess.CancelErrorRead(); } catch { }
                try { currentProcess.Dispose(); } catch { }
            }
        }

        private static bool TryRequestGracefulExit(Process currentProcess)
        {
            try
            {
                currentProcess.StandardInput.WriteLine("exit");
                currentProcess.StandardInput.Flush();
                return true;
            }
            catch (Exception ex)
            {
                Trace.TraceWarning(
                    "向 PaddleOCR-json 发送 exit 失败，将强制终止：{0}",
                    ex.Message);
                return false;
            }
        }

        private static void ForceTerminate(Process currentProcess)
        {
            try
            {
                if (!currentProcess.HasExited) currentProcess.Kill();
            }
            catch (Exception ex)
            {
                Trace.TraceWarning(
                    "强制终止 PaddleOCR-json 进程失败：{0}", ex.Message);
                return;
            }

            WaitForExit(currentProcess, 3000, "强制终止");
        }

        private static bool WaitForExit(
            Process currentProcess,
            int milliseconds,
            string operation)
        {
            try
            {
                bool exited = currentProcess.WaitForExit(milliseconds);
                if (!exited)
                {
                    Trace.TraceWarning(
                        "OCR进程在{0}后{1}秒内仍未退出。",
                        operation,
                        milliseconds / 1000);
                }
                return exited;
            }
            catch (Exception ex)
            {
                Trace.TraceWarning(
                    "等待 PaddleOCR-json 进程{0}时发生异常：{1}",
                    operation,
                    ex.Message);
                return false;
            }
        }

        private static OcrResult Parse(string imagePath, string json)
        {
            if (string.IsNullOrWhiteSpace(json))
                throw CreateProtocolException("stdout 未返回JSON。", json, null);

            var serializer = new JavaScriptSerializer
            {
                MaxJsonLength = int.MaxValue
            };
            PaddleOcrResponse responseModel;
            try
            {
                responseModel = serializer.Deserialize<PaddleOcrResponse>(json);
            }
            catch (Exception ex)
            {
                throw CreateProtocolException("JSON无法解析。", json, ex);
            }

            if (responseModel == null)
                throw CreateProtocolException("JSON响应为空。", json, null);

            var result = CreateResult(imagePath);
            if (responseModel.code == 101) return result;
            if (responseModel.code != 100)
            {
                throw new InvalidOperationException(string.Format(
                    "PaddleOCR-json 识别失败（code={0}）：{1}",
                    responseModel.code,
                    SummarizeValue(serializer, responseModel.data)));
            }

            PaddleOcrData[] rows;
            try
            {
                rows = responseModel.data == null
                    ? new PaddleOcrData[0]
                    : serializer.ConvertToType<PaddleOcrData[]>(
                        responseModel.data);
            }
            catch (Exception ex)
            {
                throw CreateProtocolException(
                    "code=100 的 data 结构无效。", json, ex);
            }

            if (rows == null) rows = new PaddleOcrData[0];
            foreach (PaddleOcrData row in rows)
            {
                if (row == null || row.box == null) continue;

                double minX = double.MaxValue;
                double minY = double.MaxValue;
                double maxX = double.MinValue;
                double maxY = double.MinValue;
                foreach (double[] point in row.box)
                {
                    if (point == null || point.Length < 2) continue;
                    minX = Math.Min(minX, point[0]);
                    minY = Math.Min(minY, point[1]);
                    maxX = Math.Max(maxX, point[0]);
                    maxY = Math.Max(maxY, point[1]);
                }
                if (maxX <= minX || maxY <= minY) continue;

                result.TextBlocks.Add(new OcrTextBlock
                {
                    Text = row.text,
                    Confidence = (float)row.score,
                    X = (float)minX,
                    Y = (float)minY,
                    Width = (float)(maxX - minX),
                    Height = (float)(maxY - minY)
                });
            }

            return result;
        }

        private static OcrResult CreateResult(string imagePath)
        {
            var result = new OcrResult { ImagePath = imagePath };
            using (var image = Image.FromFile(imagePath))
            {
                result.ImageWidth = image.Width;
                result.ImageHeight = image.Height;
            }
            return result;
        }

        private static InvalidOperationException CreateProtocolException(
            string reason,
            string json,
            Exception innerException)
        {
            string message = "PaddleOCR-json 协议错误：" + reason;
            string snippet = Truncate(json, MaxProtocolSnippetLength);
            if (!string.IsNullOrWhiteSpace(snippet))
                message += " 响应片段：" + snippet;
            return new InvalidOperationException(message, innerException);
        }

        private static string SummarizeValue(
            JavaScriptSerializer serializer,
            object value)
        {
            if (value == null) return "无返回数据";
            try
            {
                return Truncate(
                    serializer.Serialize(value), MaxProtocolSnippetLength);
            }
            catch
            {
                return Truncate(value.ToString(), MaxProtocolSnippetLength);
            }
        }

        private string GetStartupDetails()
        {
            lock (stateLock)
            {
                string standardOutput = startupOutput.ToString().Trim();
                string standardError = diagnosticOutput.ToString().Trim();
                if (standardOutput.Length == 0) return standardError;
                if (standardError.Length == 0) return standardOutput;
                return standardOutput + Environment.NewLine + standardError;
            }
        }

        private string GetDiagnosticSuffix()
        {
            lock (stateLock)
            {
                string details = diagnosticOutput.ToString().Trim();
                return details.Length == 0
                    ? string.Empty
                    : Environment.NewLine + "OCR诊断：" + details;
            }
        }

        private static void AppendBoundedDiagnostic(
            StringBuilder builder,
            string value)
        {
            if (builder.Length > 0) builder.AppendLine();
            builder.Append(value);
            if (builder.Length > MaxDiagnosticLength)
                builder.Remove(0, builder.Length - MaxDiagnosticLength);
        }

        private static string Truncate(string value, int maximumLength)
        {
            if (string.IsNullOrEmpty(value) || value.Length <= maximumLength)
                return value;
            return value.Substring(0, maximumLength) + "…";
        }

        public sealed class PaddleOcrResponse
        {
            public int code { get; set; }
            public object data { get; set; }
        }

        public sealed class PaddleOcrData
        {
            public double[][] box { get; set; }
            public double score { get; set; }
            public string text { get; set; }
        }

        public void Dispose()
        {
            bool graceful;
            lock (stateLock)
            {
                if (disposed) return;
                graceful = process != null && started && !unhealthy;
                disposed = true;
            }

            StopProcess(graceful);
            requestGate.Dispose();
            startupEvent.Dispose();
            responseEvent.Dispose();
            processExitedEvent.Dispose();
        }
    }
}
