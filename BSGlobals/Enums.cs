using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BSGlobals
{
   public class Enums
    {
        public enum DatabaseConnectionStringNames
        {
            EventLogs,
            Parking,
            SBSReports,
            PBS2Macro,
            Commissions,
            BuffNewsForBW,
            Brainworks,
            CommissionsRelated,
            BARC,
            Wrappers,
            Manifests,
            PBSInvoiceExportLoad,
            QualificationReportLoad,
            OfficePay,
            AutoRenew,
            PressRoom,
            PressRoomFree,
            PBSInvoiceTotals,
            PBSInvoices,
            DMMail,
            DMMailFree
        }

        public enum JobLogMessageType
        {
            INFO = 1,
            WARNING = 2,
            ERROR = 3
        }

    }
}
