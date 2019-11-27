using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CommissionsCreate
{
   public class Attachment
    {
        public string Description { get; set; }

        public string FileName { get; set; }

        public bool HasManiaFlag { get; set; }

        public bool HasNewBusinessFlag { get; set; }

        public bool HasProductsFlag { get; set; }

        public string FileNameExtension { get; set; }

        public string FileNamePrefix { get; set; }

        public bool PlaybookFlag { get; set; }

        public string Salesperson { get; set; }

        public Int64 SalespersonGroupId { get; set; }
    }
}
