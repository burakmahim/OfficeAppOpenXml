using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using Lucene.Net.Analysis.Standard;
using Lucene.Net.Documents;
using Lucene.Net.Index;
using Lucene.Net.QueryParsers.Classic;
using Lucene.Net.Search;
using Lucene.Net.Store;
using Lucene.Net.Util;

namespace OfficeAppOpenXmlLibrary
{
    public class LuceneSearchService
    {
        private readonly string        _indexPath;
        private readonly LuceneVersion _version    = LuceneVersion.LUCENE_48;

        public LuceneSearchService          (string indexPath)                   
        {
            _indexPath = indexPath;
        }
        public  void   CreateIndex          (string folderPath)                  
        {
            FSDirectory       dir      = FSDirectory.Open(_indexPath);
            StandardAnalyzer  analyzer = new StandardAnalyzer(_version);
            IndexWriterConfig config   = new IndexWriterConfig(_version, analyzer);
            config.OpenMode = OpenMode.CREATE;

            using (IndexWriter writer  = new IndexWriter(dir, config))
            {
                string[] files = System.IO.Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories);

                foreach (string file in files)
                {
                    string ext = Path.GetExtension(file).ToLower();

                    try
                    {
                        if      (ext == ".xlsx")
                        {
                            IndexExcelFile(writer, file);
                        }
                        else if (ext == ".docx")
                        {
                            IndexWordFile(writer, file);
                        }
                        else if (ext == ".pptx")
                        {
                            IndexPowerPointFile(writer, file);
                        }
                    }
                    catch { }
                }

                writer.Commit();
            }
        }
        private void   IndexExcelFile       (IndexWriter writer, string filePath)
        {
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filePath, false))
            {
                DocumentFormat.OpenXml.Packaging.WorkbookPart workbookPart = doc.WorkbookPart;

                foreach (DocumentFormat.OpenXml.Spreadsheet.Sheet sheet in workbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>())
                {
                    DocumentFormat.OpenXml.Packaging.WorksheetPart worksheetPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)workbookPart.GetPartById(sheet.Id);
                    DocumentFormat.OpenXml.Spreadsheet.SheetData   sheetData     = worksheetPart.Worksheet.Elements<DocumentFormat.OpenXml.Spreadsheet.SheetData>().First();

                    foreach (DocumentFormat.OpenXml.Spreadsheet.Row row in sheetData.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>())
                    {
                        foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>())
                        {
                            string cellValue = GetCellValue(cell, workbookPart);
                            if (!string.IsNullOrWhiteSpace(cellValue))
                            {
                                Document luceneDoc = new Document();
                                luceneDoc.Add(new StringField("path"    , filePath, Field.Store.YES));
                                luceneDoc.Add(new StringField("filename", Path.GetFileName(filePath), Field.Store.YES));
                                luceneDoc.Add(new StringField("filetype", "Excel", Field.Store.YES));
                                luceneDoc.Add(new StringField("location", "Sheet: " + sheet.Name + ", Cell: " + cell.CellReference, Field.Store.YES));
                                luceneDoc.Add(new TextField  ("content" , cellValue, Field.Store.YES));

                                writer.AddDocument(luceneDoc);
                            }
                        }
                    }
                }
            }
        }
        private void   IndexWordFile        (IndexWriter writer, string filePath)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
            {
                int paragraphIndex = 0;
                foreach (DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph in doc.MainDocumentPart.Document.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
                {
                    paragraphIndex++;
                    string text = paragraph.InnerText;
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        Document luceneDoc = new Document();
                        luceneDoc.Add(new StringField("path"    , filePath, Field.Store.YES));
                        luceneDoc.Add(new StringField("filename", Path.GetFileName(filePath), Field.Store.YES));
                        luceneDoc.Add(new StringField("filetype", "Word", Field.Store.YES));
                        luceneDoc.Add(new StringField("location", "Paragraph " + paragraphIndex, Field.Store.YES));
                        luceneDoc.Add(new TextField  ("content" , text, Field.Store.YES));

                        writer.AddDocument(luceneDoc);
                    }
                }
            }
        }
        private void   IndexPowerPointFile  (IndexWriter writer, string filePath)
        {
            using (PresentationDocument ppt = PresentationDocument.Open(filePath, false))
            {
                int slideIndex = 0;
                foreach (SlidePart slidePart in ppt.PresentationPart.SlideParts)
                {
                    slideIndex++;
                    IEnumerable<DocumentFormat.OpenXml.Drawing.Text> texts = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>();

                    foreach (DocumentFormat.OpenXml.Drawing.Text t in texts)
                    {
                        if (!string.IsNullOrWhiteSpace(t.Text))
                        {
                            Document luceneDoc = new Document();
                            luceneDoc.Add(new StringField("path"    , filePath, Field.Store.YES));
                            luceneDoc.Add(new StringField("filename", Path.GetFileName(filePath), Field.Store.YES));
                            luceneDoc.Add(new StringField("filetype", "PowerPoint", Field.Store.YES));
                            luceneDoc.Add(new StringField("location", "Slide " + slideIndex, Field.Store.YES));
                            luceneDoc.Add(new TextField  ("content" , t.Text, Field.Store.YES));

                            writer.AddDocument(luceneDoc);
                        }
                    }
                }
            }
        }
		private string GetCellValue(DocumentFormat.OpenXml.Spreadsheet.Cell cell, DocumentFormat.OpenXml.Packaging.WorkbookPart workbookPart)
		{
			if (cell == null || cell.CellValue == null)
				return string.Empty;

			string value = cell.CellValue.InnerText;

			if (cell.DataType != null && cell.DataType.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString)
			{
				DocumentFormat.OpenXml.Spreadsheet.SharedStringTable stringTable = workbookPart.SharedStringTablePart.SharedStringTable;
				if (stringTable != null)
				{
					value = stringTable.ElementAt(int.Parse(value)).InnerText;
				}
			}

			return value;
		}

		public List<SearchResult> Search(string queryText, int maxResults = 1000)
        {
            List<SearchResult> results  = new List<SearchResult>();
            FSDirectory        dir      = FSDirectory.Open(_indexPath);
            StandardAnalyzer   analyzer = new StandardAnalyzer(_version);

            using (DirectoryReader reader = DirectoryReader.Open(dir))
            {
                IndexSearcher searcher = new IndexSearcher(reader);

                string wildcardQuery = "*" + queryText.Trim().ToLower() + "*";
                WildcardQuery query  = new WildcardQuery(new Term("content", wildcardQuery));

                ScoreDoc[] hits = searcher.Search(query, maxResults).ScoreDocs;

                foreach (ScoreDoc hit in hits)
                {
                    Document doc = searcher.Doc(hit.Doc);

                    SearchResult result = new SearchResult();

                    result.FileName = doc.Get("filename");
                    result.FilePath = doc.Get("path"    );
                    result.FileType = doc.Get("filetype");
                    result.Location = doc.Get("location");
                    result.Content  = doc.Get("content" );

                    results.Add(result);
                }
            }

            return results;
        }
    }
    public class SearchResult       
    {
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public string FileType { get; set; }
        public string Location { get; set; }
        public string Content { get; set; }
    }
}
