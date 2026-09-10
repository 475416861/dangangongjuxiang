using System;
using System.Collections.Generic;
using System.IO;

namespace MultiToolWin.OCR
{
    public sealed class OcrResourceValidationResult
    {
        public string RuntimeDirectory { get; set; }
        public List<string> MissingFiles { get; } = new List<string>();
        public bool IsValid => MissingFiles.Count == 0;
    }

    public static class OcrResourceValidator
    {
        public const string RuntimeFolderName = "PaddleOCR-json";

        public static readonly string UserMessage =
            "未检测到完整的 OCR 组件或模型文件。" + Environment.NewLine +
            "普通 PDF 仍可正常生成。" + Environment.NewLine +
            "如需生成可搜索双层 PDF，请将 OCR 资源文件夹与程序放在一起。";

        private static readonly string[] RequiredFiles =
        {
            "PaddleOCR-json.exe",
            "api-ms-win-core-libraryloader-l1-2-0.dll",
            "api-ms-win-core-processtopology-obsolete-l1-1-0.dll",
            "api-ms-win-eventing-provider-l1-1-0.dll",
            "concrt140.dll",
            "libiomp5md.dll",
            "mkldnn.dll",
            "mklml.dll",
            "msvcp140.dll",
            "onnxruntime.dll",
            "opencv_world4100.dll",
            "paddle_inference.dll",
            "paddle2onnx.dll",
            "vcomp140.dll",
            "vcruntime140.dll",
            "vcruntime140_1.dll",
            "models\\config_chinese.txt",
            "models\\dict_chinese.txt",
            "models\\ch_PP-OCRv3_det_infer\\inference.pdmodel",
            "models\\ch_PP-OCRv3_det_infer\\inference.pdiparams",
            "models\\ch_ppocr_mobile_v2.0_cls_infer\\inference.pdmodel",
            "models\\ch_ppocr_mobile_v2.0_cls_infer\\inference.pdiparams",
            "models\\ch_PP-OCRv3_rec_infer\\inference.pdmodel",
            "models\\ch_PP-OCRv3_rec_infer\\inference.pdiparams"
        };

        public static string GetRuntimeDirectory()
        {
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, RuntimeFolderName);
        }

        public static OcrResourceValidationResult Validate()
        {
            string runtimeDirectory = GetRuntimeDirectory();
            var result = new OcrResourceValidationResult { RuntimeDirectory = runtimeDirectory };

            foreach (string relativePath in RequiredFiles)
            {
                if (!File.Exists(Path.Combine(runtimeDirectory, relativePath)))
                    result.MissingFiles.Add(relativePath);
            }

            return result;
        }
    }
}
