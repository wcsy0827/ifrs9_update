using System.Collections.Generic;
using System.Data;
using Transfer.Utility;

namespace Transfer.Report.Interface
{
    public interface IReportData
    {
        DataSet GetData(List<reportParm> parms);
    }
}