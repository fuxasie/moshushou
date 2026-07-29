using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Text;
using moshushou.Ocr;
using WeChatOcr;

const int measuredIterations = 5;

Console.OutputEncoding = Encoding.UTF8;
Console.WriteLine("OCR engine benchmark");
Console.WriteLine($"Iterations per engine/sample: {measuredIterations}");
Console.WriteLine($"CPU logical cores: {Environment.ProcessorCount}");
Console.WriteLine();

using Bitmap singleLine = CreateSingleLineImage();
using Bitmap multiLine = CreateMultiLineImage();
using var ppOcr = new PpOcrV6Engine();

Console.WriteLine("Cold start (single-line image)");
Measurement ppCold = await MeasureAsync(
    "PP-OCRv6",
    () => ppOcr.RecognizeAsync(singleLine));
Measurement legacyCold = await MeasureAsync(
    "WeChat OCR",
    () => RunWeChatOcrAsync(singleLine));
PrintCold(ppCold);
PrintCold(legacyCold);
Console.WriteLine();

// Run an untimed extra call before the steady-state loop so both native
// runtimes, JIT paths, and model sessions have completed initialization.
_ = await ppOcr.RecognizeAsync(singleLine);
_ = await RunWeChatOcrAsync(singleLine);

await BenchmarkSampleAsync("单行中文（720×140）", singleLine, ppOcr);
await BenchmarkSampleAsync("多行消息区（1280×720）", multiLine, ppOcr);

async Task BenchmarkSampleAsync(string sampleName, Bitmap image, PpOcrV6Engine engine)
{
    var ppMeasurements = new List<Measurement>();
    var legacyMeasurements = new List<Measurement>();

    for (int iteration = 0; iteration < measuredIterations; iteration++)
    {
        // Alternate order to reduce systematic bias from CPU temperature or
        // background load.
        if ((iteration & 1) == 0)
        {
            ppMeasurements.Add(await MeasureAsync("PP-OCRv6", () => engine.RecognizeAsync(image)));
            legacyMeasurements.Add(await MeasureAsync("WeChat OCR", () => RunWeChatOcrAsync(image)));
        }
        else
        {
            legacyMeasurements.Add(await MeasureAsync("WeChat OCR", () => RunWeChatOcrAsync(image)));
            ppMeasurements.Add(await MeasureAsync("PP-OCRv6", () => engine.RecognizeAsync(image)));
        }
    }

    Statistics ppStats = Statistics.From(ppMeasurements);
    Statistics legacyStats = Statistics.From(legacyMeasurements);
    double medianRatio = ppStats.MedianMs / Math.Max(0.001, legacyStats.MedianMs);

    Console.WriteLine(sampleName);
    PrintStatistics("PP-OCRv6", ppStats);
    PrintStatistics("WeChat OCR", legacyStats);
    Console.WriteLine(
        medianRatio >= 1
            ? $"  中位数对比：PP-OCRv6 慢 {medianRatio:F2}×"
            : $"  中位数对比：PP-OCRv6 快 {(1 / medianRatio):F2}×");
    Console.WriteLine($"  PP-OCRv6结果：{Preview(ppMeasurements[^1].Text)}");
    Console.WriteLine($"  微信OCR结果：{Preview(legacyMeasurements[^1].Text)}");
    Console.WriteLine();
}

static async Task<Measurement> MeasureAsync(string engine, Func<Task<string>> operation)
{
    var stopwatch = Stopwatch.StartNew();
    string text = await operation();
    stopwatch.Stop();
    return new Measurement(engine, stopwatch.Elapsed.TotalMilliseconds, text);
}

static async Task<string> RunWeChatOcrAsync(Bitmap bitmap)
{
    byte[] bytes;
    using (var stream = new MemoryStream())
    {
        bitmap.Save(stream, ImageFormat.Png);
        bytes = stream.ToArray();
    }

    var completion = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
    var ocr = new ImageOcr();
    ocr.Run(bytes, (path, result) =>
    {
        try
        {
            if (result?.OcrResult?.SingleResult == null)
            {
                completion.TrySetResult(string.Empty);
                return;
            }

            var text = new StringBuilder();
            foreach (var item in result.OcrResult.SingleResult)
            {
                if (item != null && !string.IsNullOrEmpty(item.SingleStrUtf8))
                {
                    text.Append(item.SingleStrUtf8);
                }
            }
            completion.TrySetResult(text.ToString().Trim());
        }
        catch (Exception ex)
        {
            completion.TrySetException(ex);
        }
        finally
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
            }
        }
    });

    Task completed = await Task.WhenAny(completion.Task, Task.Delay(TimeSpan.FromSeconds(20)));
    if (completed != completion.Task)
    {
        return "[TIMEOUT]";
    }
    return await completion.Task;
}

static Bitmap CreateSingleLineImage()
{
    var image = new Bitmap(720, 140, PixelFormat.Format24bppRgb);
    using Graphics graphics = Graphics.FromImage(image);
    using var font = new Font("Microsoft YaHei", 44, FontStyle.Bold, GraphicsUnit.Pixel);
    graphics.Clear(Color.White);
    graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
    graphics.DrawString("未发货预警  订单ABC123", font, Brushes.Black, new PointF(18, 35));
    return image;
}

static Bitmap CreateMultiLineImage()
{
    var image = new Bitmap(1280, 720, PixelFormat.Format24bppRgb);
    using Graphics graphics = Graphics.FromImage(image);
    using var titleFont = new Font("Microsoft YaHei", 38, FontStyle.Bold, GraphicsUnit.Pixel);
    using var bodyFont = new Font("Microsoft YaHei", 31, FontStyle.Regular, GraphicsUnit.Pixel);
    graphics.Clear(Color.White);
    graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
    graphics.DrawString("未发货预警", titleFont, Brushes.Black, new PointF(36, 28));

    string[] lines =
    {
        "订单号：ABC1234567890，请及时处理。",
        "商家：星海旗舰店，当前状态：等待发货。",
        "超时未交件将按平台规则考核处罚。",
        "已经售后的订单请及时发起拦截。",
        "核实虚假拦截将正常考核处罚。",
        "处理完成后请在群内回复，谢谢。"
    };

    float y = 105;
    foreach (string line in lines)
    {
        graphics.DrawString(line, bodyFont, Brushes.Black, new PointF(36, y));
        y += 82;
    }
    return image;
}

static void PrintCold(Measurement measurement)
{
    Console.WriteLine(
        $"  {measurement.Engine,-12} {measurement.ElapsedMs,9:F1} ms  结果={Preview(measurement.Text)}");
}

static void PrintStatistics(string engine, Statistics stats)
{
    Console.WriteLine(
        $"  {engine,-12} 平均={stats.AverageMs,8:F1} ms  " +
        $"中位={stats.MedianMs,8:F1} ms  最小={stats.MinMs,8:F1} ms  最大={stats.MaxMs,8:F1} ms");
}

static string Preview(string text)
{
    if (string.IsNullOrWhiteSpace(text))
    {
        return "<empty>";
    }

    string oneLine = text.Replace("\r", string.Empty).Replace("\n", " ");
    return oneLine.Length <= 72 ? oneLine : oneLine[..72] + "…";
}

readonly record struct Measurement(string Engine, double ElapsedMs, string Text);

readonly record struct Statistics(double AverageMs, double MedianMs, double MinMs, double MaxMs)
{
    public static Statistics From(IReadOnlyList<Measurement> measurements)
    {
        double[] values = measurements.Select(item => item.ElapsedMs).OrderBy(value => value).ToArray();
        double median = values.Length % 2 == 1
            ? values[values.Length / 2]
            : (values[values.Length / 2 - 1] + values[values.Length / 2]) / 2;
        return new Statistics(values.Average(), median, values[0], values[^1]);
    }
}
