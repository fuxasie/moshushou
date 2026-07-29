using moshushou;
using moshushou.Ocr;
using moshushou.Yolo;
using System.Drawing;
using System.Drawing.Imaging;

static void Assert(bool condition, string message)
{
    if (!condition)
    {
        throw new InvalidOperationException(message);
    }
}

const string expectedGroup = "\u661F\u6D77\u65D7\u8230\u5E97";
const string otherGroup = "\u661F\u6CB3\u65D7\u8230\u5E97";
const string shortFragment = "\u65D7\u8230\u5E97";
const string highCoverageFragment = "\u661F\u6D77\u65D7\u8230";

Assert(
    SendReliabilityPolicy.IsStrictIdentityMatch(expectedGroup, expectedGroup),
    "An exact group title must match.");
Assert(
    SendReliabilityPolicy.IsStrictIdentityMatch(
        expectedGroup,
        expectedGroup + "\uFF08123\uFF09\uFF08\u5916\u90E8\uFF09"),
    "Member counts and external markers must be ignored.");
Assert(
    !SendReliabilityPolicy.IsStrictIdentityMatch(expectedGroup, shortFragment),
    "A short reverse-contained title must not match.");
Assert(
    !SendReliabilityPolicy.IsStrictIdentityMatch(expectedGroup, highCoverageFragment),
    "A truncated OCR title must fail closed.");
Assert(
    !SendReliabilityPolicy.IsStrictIdentityMatch(expectedGroup, otherGroup),
    "A different group title must not match.");
Assert(
    !SendReliabilityPolicy.IsStrictIdentityMatch(expectedGroup, expectedGroup + "2"),
    "A similarly named numbered group must not match.");
Assert(
    !SendReliabilityPolicy.IsStrictIdentityMatch(
        "\u661F\u6D77-\u65D7\u8230\u5E97",
        "\u661F\u6D77\u65D7\u8230\u5E97"),
    "Meaningful punctuation in a group name must not be discarded.");
Assert(
    !SendReliabilityPolicy.IsStrictIdentityMatch(
        "\u661F\u6D77\u65D7\u8230\u5E97\u5BA2\u670D\u7FA41",
        "\u661F\u6D77\u65D7\u8230\u5E97\u5BA2\u670D\u7FA42"),
    "Near-identical group titles must not be accepted through edit-distance matching.");

const string warning = "\u672A\u53D1\u8D27\u9884\u8B66";
Assert(
    SendReliabilityPolicy.IsStrictContentMatch(warning, "\u4ECA\u65E5" + warning + "\uFF1A3\u5355"),
    "The verification keyword must be found in the message region.");
Assert(
    !SendReliabilityPolicy.IsStrictContentMatch(warning, "\u53D1\u8D27\u5B8C\u6210"),
    "Unrelated message text must not pass verification.");
Assert(
    !SendReliabilityPolicy.IsStrictContentMatch("SF1234567890", "SF1234567891"),
    "Near-identical payloads must not be accepted through edit-distance matching.");
Assert(
    SendReliabilityPolicy.IsInputBoxClearSignal(warning, "\u53D1\u9001(S)"),
    "A send-button-only OCR result must count as a cleared editor.");
Assert(
    SendReliabilityPolicy.IsInputBoxClearSignal(warning, "\u672A\u8BC6\u522B\u5230\u6587\u5B57"),
    "An empty OCR result must count as a cleared editor.");
Assert(
    !SendReliabilityPolicy.IsInputBoxClearSignal(warning, warning),
    "An editor that still contains the verification keyword must not count as cleared.");
Assert(
    !SendReliabilityPolicy.IsInputBoxClearSignal(warning, "\u5F85\u5904\u7406\u5185\u5BB9"),
    "Arbitrary unrelated editor text must not count as a definite clear signal.");
Assert(
    SendReliabilityPolicy.BuildVerificationKeyword(
        "\u8BA2\u5355\u660E\u7EC6\r\n\u8BF7\u53CA\u65F6\u5904\u7406" + warning,
        false) == warning,
    "Known warning text must use a stable verification keyword.");
Assert(
    SendReliabilityPolicy.BuildVerificationKeyword(@"C:\temp\orders.xlsx", true) == "orders.xlsx",
    "File verification must use the file name.");

string crlfHash = SendReliabilityPolicy.ComputePayloadHash("line1\r\nline2", false);
string lfHash = SendReliabilityPolicy.ComputePayloadHash("line1\nline2", false);
string changedHash = SendReliabilityPolicy.ComputePayloadHash("line1\nchanged", false);
Assert(crlfHash == lfHash, "Line ending differences must not change a text payload hash.");
Assert(crlfHash != changedHash, "Different text payloads must have different hashes.");

Guid runId = Guid.NewGuid();
SendAttemptContext baseContext = SendAttemptContext.Create(
    runId,
    "\u5E97\u94FAA",
    expectedGroup,
    false,
    7);
SendAttemptContext payloadContext = baseContext.ForPayload("payload", false, 2);
Assert(payloadContext.RunId == runId, "A payload attempt must retain its run ID.");
Assert(payloadContext.AttemptId != baseContext.AttemptId, "A payload attempt must get a unique attempt ID.");
Assert(payloadContext.SegmentNumber == 2, "A payload attempt must retain its segment number.");
Assert(payloadContext.PayloadHash.Length == 64, "A payload attempt must have a SHA-256 hash.");

var defaultConfig = new SearchConfig();
Assert(defaultConfig.EnablePpOcrV6, "PP-OCRv6 must be enabled by default.");
Assert(defaultConfig.EnableLegacyOcrFallback, "Legacy OCR fallback must be enabled by default.");

using (var ocrEngine = new PpOcrV6Engine())
{
    Assert(ocrEngine.ModelFilesAvailable, "Bundled PP-OCRv6 model files must be available.");
    await ocrEngine.WarmUpAsync();

    using var testImage = new Bitmap(420, 110, PixelFormat.Format24bppRgb);
    using (Graphics graphics = Graphics.FromImage(testImage))
    using (var font = new Font("Arial", 48, FontStyle.Bold, GraphicsUnit.Pixel))
    {
        graphics.Clear(Color.White);
        graphics.DrawString("ABC123", font, Brushes.Black, new PointF(14, 22));
    }

    string recognized = await ocrEngine.RecognizeAsync(testImage);
    Assert(
        recognized.Contains("ABC123", StringComparison.OrdinalIgnoreCase),
        $"PP-OCRv6 must recognize the generated smoke-test text. Actual: '{recognized}'");

    using var chineseTestImage = new Bitmap(520, 120, PixelFormat.Format24bppRgb);
    using (Graphics graphics = Graphics.FromImage(chineseTestImage))
    using (var font = new Font("Microsoft YaHei", 46, FontStyle.Bold, GraphicsUnit.Pixel))
    {
        graphics.Clear(Color.White);
        graphics.DrawString(warning, font, Brushes.Black, new PointF(14, 24));
    }

    string recognizedChinese = await ocrEngine.RecognizeAsync(chineseTestImage);
    Assert(
        recognizedChinese.Contains(warning, StringComparison.Ordinal),
        $"PP-OCRv6 must recognize the Chinese verification keyword. Actual: '{recognizedChinese}'");
}

using (var yoloDetector = new YoloWindowDetector())
using (var blankLayoutImage = new Bitmap(640, 640, PixelFormat.Format24bppRgb))
{
    using Graphics graphics = Graphics.FromImage(blankLayoutImage);
    graphics.Clear(Color.White);
    _ = yoloDetector.Detect(blankLayoutImage);
}

Console.WriteLine("Reliability smoke tests passed.");
