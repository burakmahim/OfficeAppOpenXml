using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeAppOpenXmlLibrary.Models
{

    public class FileRow
    {
        public int Id { get; set; }
        public string FileName { get; set; }
        public string FileType { get; set; }
        public string MIMEType { get; set; }
        public DateTime Created { get; set; }
        public DateTime? Updated { get; set; }
    }

}
