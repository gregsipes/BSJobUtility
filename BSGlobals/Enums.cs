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
            BSJobUtility,
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
            ManifestsFree,
            PBSInvoiceExportLoad,
            QualificationReportLoad,
            OfficePay,
            AutoRenew,
            PressRoom,
            PressRoomFree,
            PBSInvoiceTotals,
            PBSInvoices,
            DMMail,
            PayByScan,
            PrepackInsertLoad,
            CircDumpWorkLoad,
            CircDumpWorkPopulate,
            CircDumpPost,
            PBSDumpAWorkLoad,
            PBSDumpAWorkPopulate,
            PBSDumpPost,
            PBSDumpBWork,
            PBSDumpCWork,
            Purchasing,
            SuppliesWorkLoad,
            PBSDump,
            BNTransactions,
            TradeWorkLoad,
            SubBalanceLoad,
            Feeds
        }

        public enum JobLogMessageType
        {
            STARTSTOP = 0,
            INFO = 1,
            WARNING = 2,
            ERROR = 3
        }

    }
}
