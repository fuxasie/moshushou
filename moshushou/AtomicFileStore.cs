using System;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace moshushou
{
    internal static class AtomicFileStore
    {
        public static void WriteAllText(string path, string content, Encoding encoding)
        {
            string? directory = Path.GetDirectoryName(path);
            if (string.IsNullOrWhiteSpace(directory))
            {
                throw new InvalidOperationException($"无法确定状态文件目录: {path}");
            }

            Directory.CreateDirectory(directory);
            string tempPath = Path.Combine(directory, $".{Path.GetFileName(path)}.{Guid.NewGuid():N}.tmp");

            try
            {
                using (var stream = new FileStream(
                    tempPath,
                    FileMode.CreateNew,
                    FileAccess.Write,
                    FileShare.None,
                    4096,
                    FileOptions.WriteThrough))
                using (var writer = new StreamWriter(stream, encoding))
                {
                    writer.Write(content);
                    writer.Flush();
                    stream.Flush(flushToDisk: true);
                }

                if (File.Exists(path))
                {
                    File.Replace(tempPath, path, null, ignoreMetadataErrors: true);
                }
                else
                {
                    File.Move(tempPath, path);
                }
            }
            catch
            {
                TryDelete(tempPath);
                throw;
            }
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[AtomicFileStore] 清理临时文件失败: {ex.Message}");
            }
        }
    }
}
