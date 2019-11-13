using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PBSMacrosLoad
{
    class CommissionRecord
    {

        public DateTime EndDate { get; set; }

        public DateTime MonthStartDate { get; set; }

        public DateTime PriorEndDate { get; set; }

        public DateTime PriorYearStartDate { get; set; }

        public DateTime YearStartDate { get; set; }

        public Int32 Month { get; set; }

        public Int32 Year { get; set; }

        public String GainsLossesTopCount { get; set; }

        public Int32 SpreadsheetStyle { get; set; }

        public Int64 StructuresId { get; set; }

        public string PerformanceForBARCInsertStoredProcedure { get; set; }

        public string PlaybookForBARCInsertStoredProcedure { get; set; }

        public string PlaybookForBARCUpdateStoredProcedure { get; set; }
    }
}
