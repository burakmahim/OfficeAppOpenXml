using DocumentFormat.OpenXml.Office.CustomXsn;
using Microsoft.Data.SqlClient;
using OfficeAppOpenXmlLibrary.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeAppOpenXmlLibrary.Services
{
    public class FileService
    {
        private readonly string _cs;
        public FileService(string connectionString) => _cs = connectionString;

        // --------------------------------------------------------------------
        // LISTE
        // --------------------------------------------------------------------
        public IEnumerable<FileRow> List(int top = 100)
        {
            List<FileRow> list = new List<FileRow>();

            using (SqlConnection con = new SqlConnection(_cs))
            using (SqlCommand cmd = new SqlCommand(@"
                SELECT TOP (@top) Id, FileName, FileType, MIMEType, Created, Updated
                FROM dbo.Files
                ORDER BY Id DESC;", con))
            {
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
            }
            return list;
        }

        // --------------------------------------------------------------------
        // YÜKLEME 
        // --------------------------------------------------------------------
        public void Upload(string fileName, string? mime, byte[] bytes)
        {
            string? ext = Path.GetExtension(fileName)?.ToLowerInvariant();  // ".docx" vb.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      
            HashSet<string> allowed = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { ".docx", ".xlsx", ".pptx", ".pdf" };
            if (ext is null || !allowed.Contains(ext))
                throw new InvalidOperationException("Yalnızca DOCX/XLSX/PPTX/PDF kabul edilir.");

            using SqlConnection conn = new SqlConnection(_cs);
            conn.Open();

            using SqlCommand cmd = new SqlCommand(@"
            INSERT INTO dbo.Files (FileName, FileType, MIMEType, FileContent, Created)
            VALUES (@n, @ft, @mt, @fc, SYSUTCDATETIME());", conn);

            cmd.Parameters.AddWithValue("@n", fileName);
            cmd.Parameters.AddWithValue("@ft", ext);                   // TYPE COLUMN için önemli
            cmd.Parameters.AddWithValue("@mt", (object?)mime ?? "");
            SqlParameter p = cmd.Parameters.Add("@fc", SqlDbType.VarBinary, -1);
            p.Value = bytes;

            cmd.ExecuteNonQuery();
        }

        // -------------------------------------------------------------------- 
        // INDIRME 
        // -------------------------------------------------------------------- 
        public (string FileName, string Mime, byte[] Bytes)? GetFile(int id)
        {
            using SqlConnection con = new SqlConnection(_cs);
            using SqlCommand cmd = new SqlCommand(@"
            SELECT FileName, MIMEType, FileContent
            FROM dbo.Files WHERE Id = @Id;", con);

            cmd.Parameters.AddWithValue("@Id", id);
            con.Open();

            using SqlDataReader r = cmd.ExecuteReader();
            if (!r.Read()) return null;

            return (r.GetString(0), r.GetString(1), (byte[])r["FileContent"]);
        }
        
        // --------------------------------------------------------------------
        // FULL-TEXT ARAMA (MODELSID / UZANTI / TARIH FILTRELI)
        // --------------------------------------------------------------------
        public IEnumerable<SearchResult> Search(
            string query,
            int top = 50,
            bool freeText = false,
            Guid? modelsId = null,     // set filtresi
            string? ext = null,        // ".docx" / ".xlsx" / ".pptx" / ".pdf"
            DateTime? @from = null,
            DateTime? to = null)
        {
            string sqlFree = @"
            SELECT TOP (@top) f.Id, f.FileName, f.FileType, f.Created, t.[RANK]
            FROM FREETEXTTABLE(dbo.Files, (ContentText, FileContent, FileName), @q, LANGUAGE 1055) t
            JOIN dbo.Files f ON f.Id = t.[KEY]
            WHERE 1=1
              AND (@sid  IS NULL OR f.ModelsId = @sid)
              AND (@ext  IS NULL OR f.FileType = @ext)
              AND (@from IS NULL OR f.Created >= @from)
              AND (@to   IS NULL OR f.Created <  @to)
            ORDER BY t.[RANK] DESC, f.Created DESC;";

                        string sqlContains = @"
            SELECT TOP (@top) f.Id, f.FileName, f.FileType, f.Created, t.[RANK]
            FROM CONTAINSTABLE(dbo.Files, (ContentText, FileContent, FileName), @q, LANGUAGE 1055) t
            JOIN dbo.Files f ON f.Id = t.[KEY]
            WHERE 1=1
              AND (@sid  IS NULL OR f.ModelsId = @sid)
              AND (@ext  IS NULL OR f.FileType = @ext)
              AND (@from IS NULL OR f.Created >= @from)
              AND (@to   IS NULL OR f.Created <  @to)
            ORDER BY t.[RANK] DESC, f.Created DESC;";

            using SqlConnection conn = new SqlConnection(_cs);
            conn.Open();

            using SqlCommand cmd = new SqlCommand(freeText ? sqlFree : sqlContains, conn);
            cmd.Parameters.AddWithValue("@q", query);
            cmd.Parameters.AddWithValue("@top", top);
            cmd.Parameters.AddWithValue("@sid", (object?)modelsId ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ext", (object?)ext ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@from", (object?)@from ?? DBNull.Value);
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
        // MODELSID: SORGULA VE SET OLUSTUR (ör: son X gün + uzantı) 
        // --------------------------------------------------------------------
        public Guid MarkSetByFilter(string? ext = null, int lastDays = 30, int take = 1000)
        {
            Guid setId = Guid.NewGuid();

            using SqlConnection con = new SqlConnection(_cs);
            con.Open();

            string sql = @"
            WITH Pick AS (
              SELECT TOP (@take) Id
              FROM dbo.Files
              WHERE (@ext IS NULL OR FileType = @ext)
                AND Created >= DATEADD(day, -@days, SYSUTCDATETIME())
              ORDER BY Created DESC
            )
            UPDATE f SET ModelsId = @sid
            FROM dbo.Files f
            JOIN Pick p ON p.Id = f.Id;";

            using SqlCommand cmd = new SqlCommand(sql, con);
            cmd.Parameters.AddWithValue("@sid", setId);
            cmd.Parameters.AddWithValue("@ext", (object?)ext ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@days", lastDays);
            cmd.Parameters.AddWithValue("@take", take);
            cmd.ExecuteNonQuery();

            return setId;
        }

        // --------------------------------------------------------------------
        // MODELSID: ID LISTESIYLE SET OLUSTUR
        // --------------------------------------------------------------------
        public Guid MarkSetByIds(IEnumerable<int> ids)
        {
            Guid setId = Guid.NewGuid();

            using SqlConnection con = new SqlConnection(_cs);
            con.Open();

            // Geçici tablo
            using (SqlCommand c = new SqlCommand("CREATE TABLE #Ids(Id int PRIMARY KEY);", con))
                c.ExecuteNonQuery();

            // Bulk copy ile hızlı yükleme
            using (SqlBulkCopy bcp = new SqlBulkCopy(con))
            {
                bcp.DestinationTableName = "#Ids";
                using DataTable dt = new DataTable();
                dt.Columns.Add("Id", typeof(int));
                foreach (int id in ids) dt.Rows.Add(id);
                bcp.WriteToServer(dt);
            }

            using (SqlCommand u = new SqlCommand(@"
                UPDATE f SET ModelsId = @sid
                FROM dbo.Files f
                JOIN #Ids i ON i.Id = f.Id;", con))
            {
                u.Parameters.AddWithValue("@sid", setId);
                u.ExecuteNonQuery();
            }

            return setId;
        }

        // --------------------------------------------------------------------
        // MODELSID: TEMIZLE
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
        // GERİYE DÖNÜK ContentText DOLDURMA (opsiyonel)
        // --------------------------------------------------------------------
        public int BackfillContentText()
        {
            List<(int Id, byte[] Bytes, string Ext)> items = new List<(int Id, byte[] Bytes, string Ext)>();

            using (SqlConnection con = new SqlConnection(_cs))
            using (SqlCommand cmd = new SqlCommand(@"
            SELECT Id, FileContent, FileType
            FROM dbo.Files
            WHERE ContentText IS NULL;", con))
            {
                con.Open();
                using SqlDataReader r = cmd.ExecuteReader();
                while (r.Read())
                    items.Add((r.GetInt32(0), (byte[])r["FileContent"], r.GetString(2)));
            }

            int updated = 0;

            using (SqlConnection con = new SqlConnection(_cs))
            {
                con.Open();
                foreach ((int Id, byte[] Bytes, string Ext) it in items)
                {
                    string? text = null;
                    if (string.Equals(it.Ext, ".docx", StringComparison.OrdinalIgnoreCase))
                    {
                        try { text = ExtractDocxText(it.Bytes); } catch { /* yut */ }
                    }

                    using SqlCommand up = new SqlCommand(
                        "UPDATE dbo.Files SET ContentText=@t WHERE Id=@id;", con);
                    up.Parameters.AddWithValue("@t", (object?)text ?? DBNull.Value);
                    up.Parameters.AddWithValue("@id", it.Id);
                    updated += up.ExecuteNonQuery();
                }
            }

            return updated;
        }

        // Basit DOCX metin çıkarma (zip + xml)
        private static string? ExtractDocxText(byte[] bytes)
        {
            using MemoryStream ms = new MemoryStream(bytes);
            using ZipArchive zip = new ZipArchive(ms, ZipArchiveMode.Read, leaveOpen: false);

            ZipArchiveEntry? entry = zip.GetEntry("word/document.xml");
            if (entry == null) return null;

            using StreamReader sr = new StreamReader(entry.Open(), Encoding.UTF8, true);
            string xml = sr.ReadToEnd();

            string text = Regex.Replace(xml, "<[^>]+>", " ");
            text = System.Net.WebUtility.HtmlDecode(text);
            return Regex.Replace(text, @"\s+", " ").Trim();
        }
    }
}
