using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CommissionsCreate
{
   public class DataMiningProductForExcel
    {
        public string Salesperson { get; set; }

        public string GroupDescription { get; set; }

        public string EDNNumber { get; set; }

        public string Description { get; set; }

        public DateTime TranDate { get; set; }

        public Decimal AmountPreTax { get; set; }

        public Int32 HistoryCoreAccount { get; set; }

        public string ClientName { get; set; }

        public Int32 HistoryCoreTicket { get; set; }
    }
}
