using System;
using System.Collections.Generic;
using System.Data.OleDb;

namespace OfficeAppOpenXmlLibrary
{
    public class SearchResult
    {
        public string FileName      { get; set; }
        public string FilePath      { get; set; }
        public string FileExtension { get; set; }
        public string ItemType      { get; set; }
        public string AutoSummary   { get; set; }
    }

    public class WindowsSearchService   
    {
        public  List<SearchResult>  Search     (string query)
        {
            List<SearchResult> results = new List<SearchResult>();

            if (string.IsNullOrWhiteSpace(query))
                return results;

            string connectionString = "Provider=Search.CollatorDSO;Extended Properties='Application=Windows';";

            string sql = @"SELECT System.ItemName,
                          System.ItemPathDisplay,
                          System.ItemType,
                          System.Search.AutoSummary
                          FROM SystemIndex
                          WHERE CONTAINS(*, '""" + query.Replace("'", "''").Replace("\"", "\"\"") + @"""')
                          AND (
                              System.FileExtension = '.doc'  OR
                              System.FileExtension = '.docx' OR
                              System.FileExtension = '.docm' OR
                              System.FileExtension = '.xls'  OR
                              System.FileExtension = '.xlsx' OR
                              System.FileExtension = '.xlsm' OR
                              System.FileExtension = '.xlsb' OR
                              System.FileExtension = '.ppt'  OR
                              System.FileExtension = '.pptx' OR
                              System.FileExtension = '.pptm' OR
                              System.FileExtension = '.pdf'  OR
                              System.FileExtension = '.txt'  OR
                              System.FileExtension = '.log'  OR
                              System.FileExtension = '.csv'  OR
                              System.FileExtension = '.rtf'  OR
                              System.FileExtension = '.odt'  OR
                              System.FileExtension = '.ods'  OR
                              System.FileExtension = '.odp'  OR
                              System.FileExtension = '.md'   OR
                              System.FileExtension = '.xml'  OR
                              System.FileExtension = '.json'
                          )";

            try
            {
                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                    {
                        using (OleDbDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string fileName = reader["System.ItemName"]?.ToString() ?? "";

                                SearchResult result = new SearchResult
                                {
                                    FileName        = fileName,
                                    FilePath        = reader["System.ItemPathDisplay"]?.ToString() ?? "",
                                    FileExtension   = System.IO.Path.GetExtension(fileName),
                                    ItemType        = reader["System.ItemType"]?.ToString() ?? "",
                                    AutoSummary     = reader["System.Search.AutoSummary"]?.ToString() ?? ""
                                };

                                results.Add(result);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Windows Search hatası: " + ex.Message, ex);
            }

            return results;
        }
    }
}
