using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace moshushou
{
    public enum SendAttemptOutcome
    {
        ConfirmedFailure = 0,
        Success = 1,
        Ambiguous = 2
    }

    public static class SendAttemptStatuses
    {
        public const string Prepared = "Prepared";
        public const string Triggered = "Triggered";
        public const string Delivered = "Delivered";
        public const string ConfirmedFailure = "ConfirmedFailure";
        public const string Ambiguous = "Ambiguous";
    }

    public sealed record SendAttemptContext(
        Guid RunId,
        Guid AttemptId,
        string StoreName,
        string ExpectedGroupName,
        bool IsWework,
        int SelectionVersion,
        int SegmentNumber,
        string PayloadHash)
    {
        public static SendAttemptContext Create(
            Guid runId,
            string storeName,
            string expectedGroupName,
            bool isWework,
            int selectionVersion,
            int segmentNumber = 0)
        {
            return new SendAttemptContext(
                runId == Guid.Empty ? Guid.NewGuid() : runId,
                Guid.NewGuid(),
                storeName?.Trim() ?? string.Empty,
                expectedGroupName?.Trim() ?? string.Empty,
                isWework,
                selectionVersion,
                Math.Max(0, segmentNumber),
                string.Empty);
        }

        public SendAttemptContext ForPayload(string content, bool isFile, int? segmentNumber = null)
        {
            return this with
            {
                AttemptId = Guid.NewGuid(),
                SegmentNumber = Math.Max(0, segmentNumber ?? SegmentNumber),
                PayloadHash = SendReliabilityPolicy.ComputePayloadHash(content, isFile)
            };
        }
    }

    public static class SendReliabilityPolicy
    {
        private const string UnshippedWarning = "\u672A\u53D1\u8D27\u9884\u8B66";
        private const string AssessmentPenalty = "\u8003\u6838\u5904\u7F5A";

        private static readonly Regex GroupNoiseRegex = new(
            @"\((?:\d+|\u5916\u90E8)\)|\uFF08(?:\d+|\u5916\u90E8)\uFF09|\s+",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static readonly Regex ContentNoiseRegex = new(
            @"[\s\p{P}\p{S}]+",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        public static bool IsStrictIdentityMatch(string? expected, string? actual)
        {
            string expectedValue = NormalizeIdentity(expected);
            string actualValue = NormalizeIdentity(actual);
            if (expectedValue.Length == 0 || actualValue.Length == 0)
            {
                return false;
            }

            if (string.Equals(expectedValue, actualValue, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            // Identity checks fail closed. Known UI suffixes are removed during
            // normalization; arbitrary containment or edit-distance matching is unsafe.
            return false;
        }

        public static bool IsStrictContentMatch(string? expectedKeyword, string? actual)
        {
            string expectedValue = NormalizeContent(expectedKeyword);
            string actualValue = NormalizeContent(actual);
            if (expectedValue.Length == 0 || actualValue.Length == 0)
            {
                return false;
            }

            if (actualValue.Contains(expectedValue, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return false;
        }

        public static bool IsInputBoxClearSignal(string? expectedKeyword, string? inputBoxText)
        {
            if (string.IsNullOrWhiteSpace(inputBoxText) ||
                string.Equals(inputBoxText.Trim(), "\u672A\u8BC6\u522B\u5230\u6587\u5B57", StringComparison.Ordinal))
            {
                return true;
            }

            if (IsStrictContentMatch(expectedKeyword, inputBoxText))
            {
                return false;
            }

            // OCR of an empty WeChat/WeCom editor commonly returns only the
            // send-button caption. Treat only these exact UI-only forms as clear.
            string normalized = NormalizeContent(inputBoxText);
            return normalized is "\u53D1\u9001" or "\u53D1\u9001s" or "send" or "sends";
        }

        public static string BuildVerificationKeyword(string content, bool isFile)
        {
            if (isFile)
            {
                return Path.GetFileName(content ?? string.Empty);
            }

            string[] lines = (content ?? string.Empty)
                .Replace("\r\n", "\n", StringComparison.Ordinal)
                .Replace('\r', '\n')
                .Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

            string candidate = lines.LastOrDefault() ?? string.Empty;
            if (candidate.Contains(UnshippedWarning, StringComparison.Ordinal))
            {
                return UnshippedWarning;
            }

            if (candidate.Contains(AssessmentPenalty, StringComparison.Ordinal))
            {
                return AssessmentPenalty;
            }

            string normalized = NormalizeContent(candidate);
            if (normalized.Length <= 12)
            {
                return normalized;
            }

            // The input box commonly scrolls to its end, so use a trailing fragment.
            return normalized[^10..];
        }

        public static string ComputePayloadHash(string content, bool isFile)
        {
            string normalized = isFile
                ? NormalizeFilePayload(content)
                : (content ?? string.Empty)
                    .Replace("\r\n", "\n", StringComparison.Ordinal)
                    .Replace('\r', '\n');
            byte[] bytes = SHA256.HashData(Encoding.UTF8.GetBytes($"{(isFile ? "F" : "T")}|{normalized}"));
            return Convert.ToHexString(bytes);
        }

        private static string NormalizeIdentity(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            string normalized = value
                .Replace("\uFF08\u5916\u90E8\uFF09", string.Empty, StringComparison.Ordinal)
                .Replace("(\u5916\u90E8)", string.Empty, StringComparison.Ordinal);
            return GroupNoiseRegex.Replace(normalized, string.Empty).ToLowerInvariant();
        }

        private static string NormalizeContent(string? value)
        {
            return string.IsNullOrWhiteSpace(value)
                ? string.Empty
                : ContentNoiseRegex.Replace(value, string.Empty).ToLowerInvariant();
        }

        private static string NormalizeFilePayload(string? path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return string.Empty;
            }

            try
            {
                var info = new FileInfo(path);
                return $"{info.FullName}|{(info.Exists ? info.Length : -1)}|{(info.Exists ? info.LastWriteTimeUtc.Ticks : 0)}";
            }
            catch
            {
                return path.Trim();
            }
        }

    }
}
