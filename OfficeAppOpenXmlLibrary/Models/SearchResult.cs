using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeAppOpenXmlLibrary.Models
{


    public class SearchResult
    {
        public int Id { get; set; }
        public string FileName { get; set; }
        public string FileType { get; set; }
        public DateTime Created { get; set; }
        public int Rank { get; set; }   // FREETEXTTABLE/CONTAINSTABLE RANK
    }

}
