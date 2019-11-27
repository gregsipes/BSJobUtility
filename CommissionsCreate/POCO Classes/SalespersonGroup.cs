using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CommissionsCreate
{
    public class SalespersonGroup
    {
        public int SalespersonGroupsId { get; set; }

        public string WorksheetName { get; set; }

        public string SalespersonName { get; set; }

        public int EmployeeNumber { get; set; }

        public int TerritoriesId { get; set; }

        public string Territory { get; set; }

        public string Manager { get; set; }

        public string BARCForExcelStoredProcedure { get; set; }

        public int SalespersonCount { get; set; }
    }
}
