using System.Data;
using Transfer.Utility;
using System.Data.SqlClient;
using System.Collections.Generic;

namespace Transfer.Report.Data
{
    public class SignOff : ReportData
    {
        public override DataSet GetData(List<reportParm> parms)
        {
            DataSet resultsTable = new DataSet();

            string sql = string.Empty;
            sql = "SELECT ''";

            using (SqlConnection conn = new SqlConnection(defaultConnection))
            {
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    Extension.NlogSet(cmd.CommandText);
                    conn.Open();
                    SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                    adapter.Fill(resultsTable);
                }
            }
            return resultsTable;
        }
    }
}