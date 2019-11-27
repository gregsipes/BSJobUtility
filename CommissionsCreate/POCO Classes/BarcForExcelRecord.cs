using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CommissionsCreate
{
    public class BarcForExcelRecord
    {
        public string MarkerFlagName { get; set; }

        public string Salesperson { get; set; }

        public string GroupDescription { get; set; }

        public string PrintDivisionDescription { get; set; }

        public decimal RevenueWithoutTaxes { get; set; }

        public DateTime TranDate { get; set; }

        public string Pub { get; set; }

        public string TranCode { get; set; }

        public string TranType { get; set; }

        public int Account { get; set; }

        public string ClientName { get; set; }

        public int Ticket { get; set; }

        public string SelectSource { get; set; }
    }
}
