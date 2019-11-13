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
        BARC
    }

    public enum JobLogMessageType
    {
        INFO = 1,
        WARNING = 2,
        ERROR = 3
    }
}
