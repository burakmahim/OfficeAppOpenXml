using Microsoft.Data.SqlClient;
using OfficeAppOpenXmlLibrary.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using UglyToad.PdfPig.Content;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using Syncfusion.Pdf;

namespace OfficeAppOpenXmlLibrary.Services
{
    public class FileService
    {
        private readonly string _cs;
        public FileService(string connectionString) => _cs = connectionString;

        // --------------------------------------------------------------------
        // LIST
        // --------------------------------------------------------------------
        public IEnumerable<FileRow> List(int top = 100)
        {
            List<FileRow> list = new List<FileRow>();

            using SqlConnection con = new SqlConnection(_cs);
            using SqlCommand cmd = new SqlCommand(@"
                SELECT TOP (@top)
                    Id, FileName, FileType, MIMEType, Created, Updated
                FROM dbo.Files
                ORDER BY Id DESC;", con);

            cmd.Parameters.AddWithValue("@top", top);

            con.Open();
            using SqlDataReader r = cmd.ExecuteReader();

            while (r.Read())
            {
                list.Add(new FileRow
                {
                    Id = r.GetInt32(0),
                    FileName = r.GetString(1),
                    FileType = r.GetString(2),
                    MIMEType = r.GetString(3),
                    Created = r.GetDateTime(4),
                    Updated = r.IsDBNull(5) ? (DateTime?)null : r.GetDateTime(5)
                });
            }

            return list;
        }

        // --------------------------------------------------------------------
        // UPLOAD
        // --------------------------------------------------------------------
        public void Upload(string fileName, string? mime, byte[] bytes)
        {
            string? ext = Path.GetExtension(fileName)?.ToLowerInvariant();

            HashSet<string> allowed = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                ".docx", ".xlsx", ".xls", ".pptx", ".pdf", ".txt"
            };

            if (ext is null || !allowed.Contains(ext))
                throw new InvalidOperationException("Yalnızca DOCX/XLSX/XLS/PPTX/PDF/TXT kabul edilir.");

            // Extract text
            string? text = null;

            if (ext == ".docx")
                text = ExtractDocxText(bytes);

            else if (ext == ".xlsx" )
                text = ExtractXlsxText(bytes);

            else if (ext == ".xls")
                text = ExtractXlsTextWithSyncfusion(bytes);    // ✅ xls için Syncfusion

            else if (ext == ".pdf")
                text = ExtractPdfText(bytes);

            else if (ext == ".txt")
                text = Encoding.UTF8.GetString(bytes);

            using SqlConnection conn = new SqlConnection(_cs);
            conn.Open();

            using SqlCommand cmd = new SqlCommand(@"
                INSERT INTO dbo.Files
                (FileName, FileType, MIMEType, FileContent, ContentText, Created)
                VALUES (@n, @ft, @mt, @fc, @ct, SYSUTCDATETIME());", conn);

            cmd.Parameters.AddWithValue("@n", fileName);
            cmd.Parameters.AddWithValue("@ft", ext);
            cmd.Parameters.AddWithValue("@mt", (object?)mime ?? "");
            cmd.Parameters.Add("@fc", SqlDbType.VarBinary, -1).Value = bytes;
            cmd.Parameters.AddWithValue("@ct", (object?)text ?? DBNull.Value);

            cmd.ExecuteNonQuery();
        }

        // --------------------------------------------------------------------
        // DOWNLOAD
        // --------------------------------------------------------------------  
        public (string FileName, string Mime, byte[] Bytes)? GetFile(int id)
        {
            using SqlConnection con = new SqlConnection(_cs);
            using SqlCommand cmd = new SqlCommand(@"
                SELECT FileName, MIMEType, FileContent
                FROM dbo.Files
                WHERE Id = @Id;", con);

            cmd.Parameters.AddWithValue("@Id", id);

            con.Open();
            using SqlDataReader r = cmd.ExecuteReader();

            if (!r.Read())
                return null;

            return (
                r.GetString(0),
                r.GetString(1),
                (byte[])r["FileContent"]
            );
        }

        // --------------------------------------------------------------------
        // SEARCH
        // --------------------------------------------------------------------
        public IEnumerable<SearchResult> Search(
            string query,
            int top = 50,
            bool freeText = false,
            Guid? modelsId = null,
            string? ext = null,
            DateTime? from = null,
            DateTime? to = null)
        {
            string sqlFree = @"
                SELECT TOP (@top)
                    f.Id, f.FileName, f.FileType, f.Created, t.[RANK]
                FROM FREETEXTTABLE(dbo.Files,
                    (ContentText, FileContent, FileName), @q, LANGUAGE 1055) t
                JOIN dbo.Files f ON f.Id = t.[KEY]
                WHERE 1=1
                  AND (@sid IS NULL OR f.ModelsId = @sid)
                  AND (@ext IS NULL OR f.FileType = @ext)
                  AND (@from IS NULL OR f.Created >= @from)
                  AND (@to   IS NULL OR f.Created <  @to)
                ORDER BY t.[RANK] DESC, f.Created DESC;";

            string sqlContains = @"
                SELECT TOP (@top)
                    f.Id, f.FileName, f.FileType, f.Created, t.[RANK]
                FROM CONTAINSTABLE(dbo.Files,
                    (ContentText, FileContent, FileName), @q, LANGUAGE 1055) t
                JOIN dbo.Files f ON f.Id = t.[KEY]
                WHERE 1=1
                  AND (@sid IS NULL OR f.ModelsId = @sid)
                  AND (@ext IS NULL OR f.FileType = @ext)
                  AND (@from IS NULL OR f.Created >= @from)
                  AND (@to   IS NULL OR f.Created <  @to)
                ORDER BY t.[RANK] DESC, f.Created DESC;";

            using SqlConnection con = new SqlConnection(_cs);
            con.Open();

            using SqlCommand cmd = new SqlCommand(freeText ? sqlFree : sqlContains, con);

            cmd.Parameters.AddWithValue("@q", query);
            cmd.Parameters.AddWithValue("@top", top);
            cmd.Parameters.AddWithValue("@sid", (object?)modelsId ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ext", (object?)ext ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@from", (object?)from ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@to", (object?)to ?? DBNull.Value);

            using SqlDataReader rd = cmd.ExecuteReader();

            while (rd.Read())
            {
                yield return new SearchResult
                {
                    Id = rd.GetInt32(0),
                    FileName = rd.GetString(1),
                    FileType = rd.GetString(2),
                    Created = rd.GetDateTime(3),
                    Rank = rd.GetInt32(4)
                };
            }
        }

        // --------------------------------------------------------------------
        // MARK SET BY FILTER
        // --------------------------------------------------------------------
        public Guid MarkSetByFilter(string? ext = null, int lastDays = 30, int take = 1000)
        {
            Guid setId = Guid.NewGuid();

            using SqlConnection con = new SqlConnection(_cs);
            con.Open();

            using SqlCommand cmd = new SqlCommand(@"
                WITH Pick AS (
                    SELECT TOP (@take) Id
                    FROM dbo.Files
                    WHERE (@ext IS NULL OR FileType = @ext)
                      AND Created >= DATEADD(day, -@days, SYSUTCDATETIME())
                    ORDER BY Created DESC
                )
                UPDATE f SET ModelsId = @sid
                FROM dbo.Files f
                JOIN Pick p ON p.Id = f.Id;", con);

            cmd.Parameters.AddWithValue("@sid", setId);
            cmd.Parameters.AddWithValue("@ext", (object?)ext ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@days", lastDays);
            cmd.Parameters.AddWithValue("@take", take);

            cmd.ExecuteNonQuery();
            return setId;
        }

        // --------------------------------------------------------------------
        // MARK SET BY IDS
        // --------------------------------------------------------------------
        public Guid MarkSetByIds(IEnumerable<int> ids)
        {
            Guid setId = Guid.NewGuid();

            using SqlConnection con = new SqlConnection(_cs);
            con.Open();

            using SqlCommand c = new SqlCommand("CREATE TABLE #Ids(Id int PRIMARY KEY);", con);
            c.ExecuteNonQuery();

            using SqlBulkCopy bcp = new SqlBulkCopy(con)
            {
                DestinationTableName = "#Ids"
            };
            using DataTable dt = new DataTable();
            dt.Columns.Add("Id", typeof(int));

            foreach (int id in ids)
                dt.Rows.Add(id);

            bcp.WriteToServer(dt);

            using SqlCommand u = new SqlCommand(@"
                UPDATE f SET ModelsId = @sid
                FROM dbo.Files f
                JOIN #Ids i ON i.Id = f.Id;", con);

            u.Parameters.AddWithValue("@sid", setId);
            u.ExecuteNonQuery();

            return setId;
        }

        // --------------------------------------------------------------------
        // CLEAR SET
        // --------------------------------------------------------------------
        public int ClearSet(Guid modelsId)
        {
            using SqlConnection con = new SqlConnection(_cs);
            using SqlCommand cmd = new SqlCommand(
                "UPDATE dbo.Files SET ModelsId = NULL WHERE ModelsId = @sid;", con);

            cmd.Parameters.AddWithValue("@sid", modelsId);

            con.Open();
            return cmd.ExecuteNonQuery();
        }

        // --------------------------------------------------------------------
        // BACKFILL CONTENTTEXT
        // --------------------------------------------------------------------
        public int BackfillContentText()
        {
            List<(int Id, byte[] Bytes, string Ext)> items = new List<(int Id, byte[] Bytes, string Ext)>();

            using SqlConnection con = new SqlConnection(_cs);
            using SqlCommand cmd = new SqlCommand(@"
                SELECT Id, FileContent, FileType
                FROM dbo.Files
                WHERE ContentText IS NULL;", con);

            con.Open();
            using SqlDataReader r = cmd.ExecuteReader();

            while (r.Read())
                items.Add((r.GetInt32(0), (byte[])r["FileContent"], r.GetString(2)));

            int updated = 0;

            using SqlConnection con2 = new SqlConnection(_cs);
            con2.Open();

            foreach ((int Id, byte[] Bytes, string Ext) it in items)
            {
                string? text = null;

                if (it.Ext.Equals(".docx", StringComparison.OrdinalIgnoreCase))
                    text = ExtractDocxText(it.Bytes);

                else if (it.Ext.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                    text = ExtractXlsxText(it.Bytes);

                else if (it.Ext.Equals(".txt", StringComparison.OrdinalIgnoreCase))
                    text = ExtractTxtText(it.Bytes);

                using SqlCommand up = new SqlCommand(
                    "UPDATE dbo.Files SET ContentText=@t WHERE Id=@id;", con2);

                up.Parameters.AddWithValue("@t", (object?)text ?? DBNull.Value);
                up.Parameters.AddWithValue("@id", it.Id);

                updated += up.ExecuteNonQuery();
            }

            return updated;
        }

        // --------------------------------------------------------------------
        // DOCX EXTRACT
        // --------------------------------------------------------------------
        private static string? ExtractDocxText(byte[] bytes)
        {
            using MemoryStream ms = new MemoryStream(bytes);
            using ZipArchive zip = new ZipArchive(ms, ZipArchiveMode.Read);

            ZipArchiveEntry entry = zip.GetEntry("word/document.xml");
            if (entry == null) return null;

            using StreamReader sr = new StreamReader(entry.Open(), Encoding.UTF8);

            string xml = sr.ReadToEnd();

            // ✔ Sadece Office XML TAG'lerini sil (kullanıcının yazdığı <burak> kalır)
            string text = Regex.Replace(xml, "</?w:[^>]+>", " ");

            text = System.Net.WebUtility.HtmlDecode(text);
            return Regex.Replace(text, @"\s+", " ").Trim();
        }

        // --------------------------------------------------------------------
        // XLSX EXTRACT
        // --------------------------------------------------------------------
        public static string ExtractXlsxText(byte[] bytes)
        {
            using MemoryStream ms = new MemoryStream(bytes);
            using ZipArchive zip = new ZipArchive(ms, ZipArchiveMode.Read);

            StringBuilder result = new StringBuilder();

            // SharedStrings
            List<string> sharedStrings = new List<string>();
            ZipArchiveEntry sst = zip.GetEntry("xl/sharedStrings.xml");

            if (sst != null)
            {
                using StreamReader reader = new StreamReader(sst.Open(), Encoding.UTF8);
                string xml = reader.ReadToEnd();

                foreach (Match m in Regex.Matches(xml, "<t[^>]*>(.*?)</t>"))
                    sharedStrings.Add(m.Groups[1].Value);
            }

            // SheetX.xml dosyalarını tara
            foreach (ZipArchiveEntry? entry in zip.Entries)
            {
                if (!entry.FullName.StartsWith("xl/worksheets/sheet"))
                    continue;

                using StreamReader sr = new StreamReader(entry.Open(), Encoding.UTF8);
                string sheetXml = sr.ReadToEnd();

                // Inline string
                foreach (Match m in Regex.Matches(sheetXml, "<t[^>]*>(.*?)</t>"))
                {
                    string t = System.Net.WebUtility.HtmlDecode(m.Groups[1].Value);
                    result.AppendLine(t);
                }

                // Shared string reference
                foreach (Match m in Regex.Matches(sheetXml, "<c[^>]*t=\"s\"[^>]*>\\s*<v>(\\d+)</v>"))
                {
                    if (int.TryParse(m.Groups[1].Value, out int idx))
                    {
                        if (idx >= 0 && idx < sharedStrings.Count)
                            result.AppendLine(sharedStrings[idx]);
                    }
                }
            }

            return result.ToString();
        }

        // --------------------------------------------------------------------
        // TXT EXTRACT
        // --------------------------------------------------------------------
        public static string ExtractTxtText(byte[] bytes)
        {
            return Encoding.UTF8.GetString(bytes);
        }


        private static byte[] ConvertDocxToPdf(byte[] docxBytes)
        {
            using MemoryStream inMs = new MemoryStream(docxBytes);
            using MemoryStream outMs = new MemoryStream();

            // Syncfusion.DocIO
            using WordDocument wordDoc = new WordDocument(inMs, Syncfusion.DocIO.FormatType.Docx);

            // Syncfusion.DocIORenderer
            using DocIORenderer renderer = new DocIORenderer();
            using PdfDocument pdfDoc = renderer.ConvertToPDF(wordDoc);

            pdfDoc.Save(outMs);
            return outMs.ToArray();
        }

        private static byte[] ConvertXlsxToPdf(byte[] xlsxBytes)
        {
            using MemoryStream inMs = new MemoryStream(xlsxBytes);
            inMs.Position = 0;

            using ExcelEngine engine = new ExcelEngine();
            IApplication app = engine.Excel;
            app.DefaultVersion = ExcelVersion.Excel2016;

            IWorkbook wb = null;
            try
            {
                wb = app.Workbooks.Open(inMs);

                XlsIORenderer renderer = new XlsIORenderer(); // ❌ using YOK
                PdfDocument pdfDoc = renderer.ConvertToPDF(wb);

                using MemoryStream outMs = new MemoryStream();
                pdfDoc.Save(outMs);
                pdfDoc.Close(true);

                return outMs.ToArray();
            }
            finally
            {
                if (wb != null)
                    wb.Close();
            }
        }


        private static string ExtractXlsTextWithSyncfusion(byte[] xlsBytes)
        {
            using MemoryStream inMs = new MemoryStream(xlsBytes);
            inMs.Position = 0;

            using ExcelEngine engine = new ExcelEngine();
            IApplication app = engine.Excel;

            IWorkbook wb = null;
            try
            {
                wb = app.Workbooks.Open(inMs); // .xls / .xlsx

                StringBuilder sb = NewMethod();

                foreach (IWorksheet ws in wb.Worksheets)
                {
                    IRange used = ws.UsedRange;
                    if (used == null) continue;

                    for (int r = 1; r <= used.LastRow; r++)
                    {
                        for (int c = 1; c <= used.LastColumn; c++)
                        {
                            string? v = ws[r, c]?.DisplayText;
                            if (!string.IsNullOrWhiteSpace(v))
                                sb.AppendLine(v);
                        }
                    }
                }

                return sb.ToString();
            }
            finally
            {
                // wb null olabilir, o yüzden kontrol
                if (wb != null)
                    wb.Close();
            }
        }

        private static StringBuilder NewMethod()
        {
            return new StringBuilder();
        }

        private static string ExtractPdfText(byte[] pdfBytes)
        {
            try
            {
                using MemoryStream ms = new MemoryStream(pdfBytes);

                // 👇 PdfPig PdfDocument – tam isim
                using UglyToad.PdfPig.PdfDocument doc = UglyToad.PdfPig.PdfDocument.Open(ms);

                StringBuilder sb = new StringBuilder();

                foreach (Page page in doc.GetPages())
                {
                    string text = page.Text;
                    if (!string.IsNullOrWhiteSpace(text))
                        sb.AppendLine(text);
                }

                return sb.ToString();
            }
            catch
            {
                return "";
            }
        }


        private static byte[] ConvertPptxToPdf(byte[] pptxBytes)
        {
            using MemoryStream inMs = new MemoryStream(pptxBytes);
            using MemoryStream outMs = new MemoryStream();

            // Syncfusion.Presentation
            using IPresentation presentation = Presentation.Open(inMs);

            // Syncfusion.PresentationRenderer
            PresentationToPdfConverterSettings settings = new PresentationToPdfConverterSettings();
            using PdfDocument pdfDoc = PresentationToPdfConverter.Convert(presentation, settings);

            pdfDoc.Save(outMs);
            return outMs.ToArray();
        }


        public (string FileName, string Mime, byte[] Bytes)? GetFileAsPdf(int id)
        {
            (string FileName, string Mime, byte[] Bytes)? file = GetFile(id);
            if (file == null)
                return null;

            (string fileName, string mime, byte[] bytes) = file.Value;
            string? ext = Path.GetExtension(fileName)?.ToLowerInvariant();

            if (ext == ".pdf")
                return (fileName, "application/pdf", bytes);

            if (ext == ".docx")
                return (
                    Path.ChangeExtension(fileName, ".pdf"),
                    "application/pdf",
                    ConvertDocxToPdf(bytes)
                );

            if (ext == ".xls")
                return (
                    Path.ChangeExtension(fileName, ".pdf"),
                    "application/pdf",
                    ConvertXlsxToPdf(bytes) 
                );


            if (ext == ".xlsx")
                return (
                    Path.ChangeExtension(fileName, ".pdf"),
                    "application/pdf",
                    ConvertXlsxToPdf(bytes)
                );

            if (ext == ".pptx")
                return (
                    Path.ChangeExtension(fileName, ".pdf"),
                    "application/pdf",
                    ConvertPptxToPdf(bytes)
                );

            throw new NotSupportedException("Bu dosya türü PDF'e çevrilemiyor.");
        }

    }
}
