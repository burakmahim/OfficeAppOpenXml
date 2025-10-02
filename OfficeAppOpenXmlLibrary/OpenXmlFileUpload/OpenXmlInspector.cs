using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace OfficeAppOpenXmlLibrary.OpenXmlFileUpload
{
    public static class OpenXmlInspector
    {
        public static IDictionary<string, string> ExtractXmlParts(
            Stream officeFileStream, bool includeRels = true, int maxPartSizeBytes = 2_000_000)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (officeFileStream == null || !officeFileStream.CanRead) return result;

            using var archive = new ZipArchive(officeFileStream, ZipArchiveMode.Read, leaveOpen: true);
            foreach (var entry in archive.Entries)
            {
                var name = entry.FullName.Replace('\\', '/');
                var isXml = name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase);
                var isRels = name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase);
                if (!isXml && !(includeRels && isRels)) continue;
                if (entry.Length > maxPartSizeBytes) continue;

                using var s = entry.Open();
                using var ms = new MemoryStream();
                s.CopyTo(ms);
                var bytes = ms.ToArray();
                var enc = DetectEncoding(bytes) ?? new UTF8Encoding(false, true);
                result[name] = enc.GetString(bytes);
            }
            return result;
        }

        public static byte[] ToZip(IDictionary<string, string> parts)
        {
            using var ms = new MemoryStream();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                foreach (var kv in parts)
                {
                    var e = zip.CreateEntry(kv.Key, CompressionLevel.NoCompression);
                    using var es = e.Open();
                    var data = Encoding.UTF8.GetBytes(kv.Value ?? string.Empty);
                    es.Write(data, 0, data.Length);
                }
            }
            return ms.ToArray();
        }

        private static Encoding? DetectEncoding(byte[] b)
        {
            if (b.Length >= 3 && b[0] == 0xEF && b[1] == 0xBB && b[2] == 0xBF) return new UTF8Encoding(true);
            if (b.Length >= 2 && b[0] == 0xFF && b[1] == 0xFE) return Encoding.Unicode;
            if (b.Length >= 2 && b[0] == 0xFE && b[1] == 0xFF) return Encoding.BigEndianUnicode;
            if (b.Length >= 4 && b[0] == 0xFF && b[1] == 0xFE && b[2] == 0x00 && b[3] == 0x00) return Encoding.UTF32;
            if (b.Length >= 4 && b[0] == 0x00 && b[1] == 0x00 && b[2] == 0xFE && b[3] == 0xFF) return new UTF32Encoding(true, true);
            return null;
        }
    }
}
