using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MultiToolWin.Models;

namespace MultiToolWin.Services
{
    public sealed class PdfFolderScanner
    {
        private static readonly HashSet<string> ImageExtensions =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                ".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".gif"
            };

        private readonly NaturalStringComparer comparer = new NaturalStringComparer();

        public PdfScanResult Scan(string rootPath)
        {
            if (string.IsNullOrWhiteSpace(rootPath))
                throw new ArgumentException("图片根目录不能为空。", nameof(rootPath));
            if (!Directory.Exists(rootPath))
                throw new DirectoryNotFoundException("图片根目录不存在：" + rootPath);

            var result = new PdfScanResult();
            bool foundDirectSource = false;
            bool foundSecondLevelSource = false;

            List<string> rootImages = GetSortedImages(rootPath);
            if (rootImages.Count > 0)
            {
                foundDirectSource = true;
                result.Folders.Add(CreateSourceFolder(
                    rootPath,
                    string.Empty,
                    rootImages));
            }

            foreach (string firstLevelPath in GetSortedDirectories(rootPath))
            {
                List<string> directImages = GetSortedImages(firstLevelPath);
                if (directImages.Count > 0)
                {
                    foundDirectSource = true;
                    result.Folders.Add(CreateSourceFolder(
                        firstLevelPath,
                        string.Empty,
                        directImages));
                }

                string volumeName = Path.GetFileName(firstLevelPath);
                foreach (string secondLevelPath in
                    GetSortedDirectories(firstLevelPath))
                {
                    List<string> secondLevelImages =
                        GetSortedImages(secondLevelPath);
                    if (secondLevelImages.Count == 0) continue;

                    foundSecondLevelSource = true;
                    result.Folders.Add(CreateSourceFolder(
                        secondLevelPath,
                        volumeName,
                        secondLevelImages));
                }
            }

            result.Layout = foundSecondLevelSource
                ? PdfDirectoryLayout.VolumesAndCases
                : foundDirectSource
                    ? PdfDirectoryLayout.Cases
                    : PdfDirectoryLayout.None;
            return result;
        }

        private List<string> GetSortedDirectories(string path)
        {
            return Directory.GetDirectories(path)
                .OrderBy(Path.GetFileName, comparer)
                .ToList();
        }

        private List<string> GetSortedImages(string path)
        {
            return Directory.GetFiles(path)
                .Where(file => ImageExtensions.Contains(Path.GetExtension(file)))
                .OrderBy(Path.GetFileName, comparer)
                .ToList();
        }

        private static PdfSourceFolder CreateSourceFolder(
            string folderPath,
            string relativeOutputDirectory,
            List<string> imageFiles)
        {
            return new PdfSourceFolder
            {
                FolderName = Path.GetFileName(folderPath),
                FolderPath = folderPath,
                RelativeOutputDirectory = relativeOutputDirectory,
                ImageFiles = imageFiles
            };
        }

        private sealed class NaturalStringComparer : IComparer<string>
        {
            public int Compare(string x, string y)
            {
                if (ReferenceEquals(x, y)) return 0;
                if (x == null) return -1;
                if (y == null) return 1;

                int xIndex = 0;
                int yIndex = 0;
                while (xIndex < x.Length && yIndex < y.Length)
                {
                    bool xDigit = char.IsDigit(x[xIndex]);
                    bool yDigit = char.IsDigit(y[yIndex]);
                    if (xDigit && yDigit)
                    {
                        int numericResult = CompareNumberParts(x, ref xIndex, y, ref yIndex);
                        if (numericResult != 0) return numericResult;
                        continue;
                    }

                    int characterResult = char.ToUpperInvariant(x[xIndex])
                        .CompareTo(char.ToUpperInvariant(y[yIndex]));
                    if (characterResult != 0) return characterResult;
                    xIndex++;
                    yIndex++;
                }

                return x.Length.CompareTo(y.Length);
            }

            private static int CompareNumberParts(
                string x,
                ref int xIndex,
                string y,
                ref int yIndex)
            {
                int xStart = xIndex;
                int yStart = yIndex;
                while (xIndex < x.Length && char.IsDigit(x[xIndex])) xIndex++;
                while (yIndex < y.Length && char.IsDigit(y[yIndex])) yIndex++;

                string xNumber = TrimLeadingZeroes(x.Substring(xStart, xIndex - xStart));
                string yNumber = TrimLeadingZeroes(y.Substring(yStart, yIndex - yStart));
                int lengthResult = xNumber.Length.CompareTo(yNumber.Length);
                if (lengthResult != 0) return lengthResult;

                int valueResult = string.CompareOrdinal(xNumber, yNumber);
                if (valueResult != 0) return valueResult;
                return (xIndex - xStart).CompareTo(yIndex - yStart);
            }

            private static string TrimLeadingZeroes(string value)
            {
                string trimmed = value.TrimStart('0');
                return trimmed.Length == 0 ? "0" : trimmed;
            }
        }
    }
}
