using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BSJobBase
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
        PBSInvoices
    }

    public enum JobLogMessageType
    {
        INFO = 1,
        WARNING = 2,
        ERROR = 3
    }

    #region Excel

    public enum ExcelHorizontalAlignment
    {
        Center = 3,
        Left = -4131,
        Right = -4152,

    }

    public enum ExcelUnderLines
    {
        SingleUnderline = 2
    }

    public enum ExcelColor
    {
        NoFill = 0,
        Black = 1,
        LightOrange = 2,
        LightGray5 = 3,
        LightGray15 = 4,
        LightGray25 = 5,
        LightGray35 = 6,
        White = 7
    }

    #endregion
}
